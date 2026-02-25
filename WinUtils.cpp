// WinUtils.cpp : helper functions for Windows programming
//
#include "framework.h"
#include "WinUtils.h"

#include <objbase.h>             // For COM (CoCreateInstance, etc.)
#include <uiautomationclient.h>  // For IUIAutomation / TextPattern2

#include <oleacc.h>  // For IAccessible (MSAA fallback)
#pragma comment(lib, "Oleacc.lib")

#include <dwmapi.h>  // For EnableRoundedCorners()
#pragma comment(lib, "dwmapi.lib")

static IUIAutomation* g_pUIAutomation = nullptr;

// ---------------------------------------------------------------------------
// Per-app caret detection config lookup
// ---------------------------------------------------------------------------
int GetCaretMethodsForWindow(HWND hwnd)
{
#ifdef USE_PER_APP_CARET_CONFIG
    static HWND  cachedHwnd = nullptr;
    static int   cachedMethods = CARET_METHOD_ALL;

    if (hwnd == cachedHwnd)
        return cachedMethods;

    cachedHwnd = hwnd;
    cachedMethods = CARET_METHOD_ALL; // default for unknown apps

    // Get the exe name of the process that owns this window
    DWORD pid = 0;
    GetWindowThreadProcessId(hwnd, &pid);
    if (!pid)
        return cachedMethods;

    HANDLE hProc = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, pid);
    if (!hProc)
        return cachedMethods;

    char fullPath[MAX_PATH] = {};
    DWORD pathLen = MAX_PATH;
    QueryFullProcessImageNameA(hProc, 0, fullPath, &pathLen);
    CloseHandle(hProc);

    // Extract filename after last backslash, lowercase it
    char exeName[MAX_PATH] = {};
    const char* lastSlash = strrchr(fullPath, '\\');
    const char* name = lastSlash ? lastSlash + 1 : fullPath;
    for (int i = 0; name[i] && i < MAX_PATH - 1; i++)
        exeName[i] = (char)tolower((unsigned char)name[i]);

    strncpy_s(g_szAppExeName, exeName, _TRUNCATE);

    // Match against config table
    #define APP_ENTRY(exe, caretM, containerM) if (strcmp(exeName, exe) == 0) { cachedMethods = caretM; return cachedMethods; }
    APP_CONFIG_TABLE
    #undef APP_ENTRY

    OutputDebugFormatA("App '%s': no config, using all methods\n", exeName);
    return cachedMethods;
#else
    return CARET_METHOD_ALL;
#endif
}

void OutputDebugFormatA(const char* format, ...)
{
    if (GetKeyState(VK_CAPITAL) & 0x0001)  // CAPS LOCK on = suppress debug output
        return;

    char buffer[256];
    va_list args;
    va_start(args, format);
    vsnprintf(buffer, sizeof(buffer), format, args);
    va_end(args);
    OutputDebugStringA(buffer);
}

void DebugTraceA(const char* format, ...)
{
    // Always prints (ignores CAPS LOCK). Prefixes with seconds since app start.
    DWORD elapsed = GetTickCount() - g_dwAppStartTime;
    char prefix[32];
    snprintf(prefix, sizeof(prefix), "[%7.3f] ", elapsed / 1000.0);

    char buffer[256];
    va_list args;
    va_start(args, format);
    vsnprintf(buffer, sizeof(buffer), format, args);
    va_end(args);

    char full[300];
    snprintf(full, sizeof(full), "%s%s", prefix, buffer);
    OutputDebugStringA(full);
}

float GetDpiScaleFactor(HWND hWnd)
{
    int dpi = GetDpiForWindow(hWnd);  // Windows 8.1+ (Use 96 DPI as base)
    float fDpiScaleFactor = dpi / 96.0f;
    return fDpiScaleFactor;
}

POINT GetCaretPosition(HWND* pCaretWnd)
{
    POINT retPoint  = {};

    GUITHREADINFO gti;
    gti.cbSize = sizeof(GUITHREADINFO);

    // Get caret info from the active thread
    if (GetGUIThreadInfo(0, &gti))
    if (gti.hwndCaret)
    {
        POINT caretPos_LT = { 0, 0 };
        POINT caretPos_RB = { 0, 0 };

        // Get caret position in local (client) coordinates
        caretPos_LT.x = gti.rcCaret.left;
        caretPos_LT.y = gti.rcCaret.top;
        caretPos_RB.x = gti.rcCaret.right;
        caretPos_RB.y = gti.rcCaret.bottom;

        // Convert client coordinates to screen coordinates
        ClientToScreen(gti.hwndCaret, &caretPos_LT);
        ClientToScreen(gti.hwndCaret, &caretPos_RB);

        retPoint.x = (caretPos_LT.x + caretPos_RB.x + 1) / 2;
        retPoint.y = (caretPos_LT.y + caretPos_RB.y + 1) / 2;

        // Return the caret-owning HWND to the caller
        if (pCaretWnd)
            *pCaretWnd = gti.hwndCaret;

        if( retPoint.y > 2160*3 )
        { int i = 5; }

        if( retPoint.y < 0 )
        { int i = 5; }
    }

    return retPoint;
}

void InitUIAutomation()
{
    CoCreateInstance(__uuidof(CUIAutomation), nullptr, CLSCTX_INPROC_SERVER,
                     __uuidof(IUIAutomation), (void**)&g_pUIAutomation);
}

void CleanupUIAutomation()
{
    if (g_pUIAutomation)
    {
        g_pUIAutomation->Release();
        g_pUIAutomation = nullptr;
    }
}

// Helper: extract screen position from a UIA text range's bounding rectangles
static bool GetPositionFromTextRange(IUIAutomationTextRange* pRange, POINT& result)
{
    SAFEARRAY* pRects = nullptr;
    if (FAILED(pRange->GetBoundingRectangles(&pRects)) || !pRects)
        return false;

    bool found = false;
    LONG lBound, uBound;
    SafeArrayGetLBound(pRects, 1, &lBound);
    SafeArrayGetUBound(pRects, 1, &uBound);
    LONG count = uBound - lBound + 1;

    bool useRightEdge = false;

    OutputDebugFormatA("UIA: GetPositionFromTextRange initial count=%d\n", count);

    if (count < 4)
    {
        // Collapsed (zero-length) range — try expanding to get bounding rects
        SafeArrayDestroy(pRects);
        pRects = nullptr;

        // Strategy 1: character-level expansion (works for most UIA providers)
        struct { TextPatternRangeEndpoint ep; TextUnit unit; int dir; const char* desc; }
        charAttempts[] = {
            { TextPatternRangeEndpoint_End,   TextUnit_Character,  1, "fwd char"  },
            { TextPatternRangeEndpoint_Start, TextUnit_Character, -1, "bwd char"  },
        };

        for (int i = 0; i < _countof(charAttempts) && count < 4; i++)
        {
            IUIAutomationTextRange* pExpanded = nullptr;
            if (!SUCCEEDED(pRange->Clone(&pExpanded)) || !pExpanded)
                continue;

            int moved = 0;
            HRESULT hrMove = pExpanded->MoveEndpointByUnit(charAttempts[i].ep,
                                          charAttempts[i].unit, charAttempts[i].dir, &moved);
            OutputDebugFormatA("UIA: %s: moved=%d hr=0x%08X\n", charAttempts[i].desc, moved, hrMove);

            if ((charAttempts[i].dir > 0 && moved > 0) || (charAttempts[i].dir < 0 && moved < 0))
            {
                if (pRects) { SafeArrayDestroy(pRects); pRects = nullptr; }
                if (SUCCEEDED(pExpanded->GetBoundingRectangles(&pRects)) && pRects)
                {
                    SafeArrayGetLBound(pRects, 1, &lBound);
                    SafeArrayGetUBound(pRects, 1, &uBound);
                    count = uBound - lBound + 1;
                    OutputDebugFormatA("UIA: %s: rect count=%d\n", charAttempts[i].desc, count);
                    if (count >= 4 && charAttempts[i].dir < 0)
                        useRightEdge = true;
                }
            }
            pExpanded->Release();
        }

        // Strategy 2: line-level expansion (for providers like VSCode .cpp
        // where character-level rects don't work but line-level do)
        if (count < 4)
        {
            struct { TextPatternRangeEndpoint ep; TextUnit unit; int dir; const char* desc; }
            lineAttempts[] = {
                { TextPatternRangeEndpoint_End,   TextUnit_Line,  1, "fwd line"  },
                { TextPatternRangeEndpoint_Start, TextUnit_Line, -1, "bwd line"  },
            };

            for (int i = 0; i < _countof(lineAttempts) && count < 4; i++)
            {
                IUIAutomationTextRange* pExpanded = nullptr;
                if (!SUCCEEDED(pRange->Clone(&pExpanded)) || !pExpanded)
                    continue;

                int moved = 0;
                pExpanded->MoveEndpointByUnit(lineAttempts[i].ep,
                                              lineAttempts[i].unit, lineAttempts[i].dir, &moved);
                OutputDebugFormatA("UIA: %s: moved=%d\n", lineAttempts[i].desc, moved);

                if ((lineAttempts[i].dir > 0 && moved > 0) || (lineAttempts[i].dir < 0 && moved < 0))
                {
                    if (pRects) { SafeArrayDestroy(pRects); pRects = nullptr; }
                    if (SUCCEEDED(pExpanded->GetBoundingRectangles(&pRects)) && pRects)
                    {
                        SafeArrayGetLBound(pRects, 1, &lBound);
                        SafeArrayGetUBound(pRects, 1, &uBound);
                        count = uBound - lBound + 1;
                        OutputDebugFormatA("UIA: %s: rect count=%d\n", lineAttempts[i].desc, count);
                        if (count >= 4 && lineAttempts[i].dir < 0)
                            useRightEdge = true;
                    }
                }
                pExpanded->Release();
            }
        }
    }

    if (count >= 4)
    {
        double* pData = nullptr;
        if (SUCCEEDED(SafeArrayAccessData(pRects, (void**)&pData)))
        {
            // Use last rect when right-edge (line-to-caret range may have multiple rects)
            int rectIdx = useRightEdge ? (count / 4 - 1) * 4 : 0;
            double x = pData[rectIdx], y = pData[rectIdx+1], w = pData[rectIdx+2], h = pData[rectIdx+3];

            if (useRightEdge)
                result.x = (LONG)(x + w);   // Right edge = caret X position
            else
                result.x = (LONG)(x);       // Left edge of next char

            result.y = (LONG)(y + h / 2);
            OutputDebugFormatA("UIA: caret at %d, %d (rect[%d]: %.0f,%.0f,%.0f,%.0f)%s\n",
                               result.x, result.y, rectIdx/4, x, y, w, h,
                               useRightEdge ? " [right edge]" : "");
            SafeArrayUnaccessData(pRects);
            found = true;
        }
    }
    if (!found)
        OutputDebugFormatA("UIA: GetPositionFromTextRange FAILED (final count=%d)\n", count);
    if (pRects) SafeArrayDestroy(pRects);
    return found;
}

// Try to get caret position from a UIA element — tries TextPattern2 then TextPattern
static bool TryGetCaretFromElement(IUIAutomationElement* pElement, POINT& result)
{
    // Try TextPattern2::GetCaretRange first (best method)
    IUIAutomationTextPattern2* pTextPattern2 = nullptr;
    HRESULT hr = pElement->GetCurrentPatternAs(UIA_TextPattern2Id,
                     __uuidof(IUIAutomationTextPattern2), (void**)&pTextPattern2);
    if (SUCCEEDED(hr) && pTextPattern2)
    {
        BOOL isActive = FALSE;
        IUIAutomationTextRange* pCaretRange = nullptr;
        if (SUCCEEDED(pTextPattern2->GetCaretRange(&isActive, &pCaretRange)) && pCaretRange)
        {
            bool found = GetPositionFromTextRange(pCaretRange, result);
            pCaretRange->Release();
            pTextPattern2->Release();
            if (found) return true;
        }
        else
            pTextPattern2->Release();
    }

    // Try TextPattern::GetSelection (v1 fallback — caret is a collapsed selection)
    IUIAutomationTextPattern* pTextPattern = nullptr;
    hr = pElement->GetCurrentPatternAs(UIA_TextPatternId,
             __uuidof(IUIAutomationTextPattern), (void**)&pTextPattern);
    if (SUCCEEDED(hr) && pTextPattern)
    {
        bool found = false;
        IUIAutomationTextRangeArray* pSelections = nullptr;
        hr = pTextPattern->GetSelection(&pSelections);
        if (SUCCEEDED(hr) && pSelections)
        {
            int length = 0;
            pSelections->get_Length(&length);
            OutputDebugFormatA("UIA: TextPattern v1 GetSelection returned %d ranges\n", length);
            if (length > 0)
            {
                IUIAutomationTextRange* pRange = nullptr;
                if (SUCCEEDED(pSelections->GetElement(0, &pRange)) && pRange)
                {
                    found = GetPositionFromTextRange(pRange, result);
                    pRange->Release();
                }
            }
            pSelections->Release();
        }
        else
        {
            OutputDebugFormatA("UIA: TextPattern v1 GetSelection failed hr=0x%08X\n", hr);
        }

        pTextPattern->Release();
        if (found)
        {
            OutputDebugFormatA("UIA: (via TextPattern v1 GetSelection)\n");
            return true;
        }
    }

    return false;
}

POINT GetCaretPositionFromUIA()
{
    POINT result = {};

    if (!g_pUIAutomation)
        return result;

    // Throttle: UIA calls are expensive (especially FindFirst on large trees)
    // COM/UIA creates worker threads on each call — must limit aggressively
    static DWORD lastCallTime = 0;
    static POINT cachedResult = {};
    DWORD now = GetTickCount();
    if (now - lastCallTime < 100)  // Max ~10 calls/sec
        return cachedResult;
    lastCallTime = now;

    // 1. Try the focused element directly
    IUIAutomationElement* pFocused = nullptr;
    if (SUCCEEDED(g_pUIAutomation->GetFocusedElement(&pFocused)) && pFocused)
    {
        // Log what element we got
        BSTR name = nullptr;
        CONTROLTYPEID ctrlType = 0;
        pFocused->get_CurrentName(&name);
        pFocused->get_CurrentControlType(&ctrlType);
        OutputDebugFormatA("UIA: focused element: type=%d name='%S'\n", ctrlType, name ? name : L"(null)");
        if (name) SysFreeString(name);

        if (TryGetCaretFromElement(pFocused, result))
        {
            pFocused->Release();
            cachedResult = result;
            return result;
        }
        pFocused->Release();
    }

    // 2. Try from the foreground window's root element — search entire window tree
    if (!g_hForegroundWindow)
        return result;

    IUIAutomationElement* pWindow = nullptr;
    if (FAILED(g_pUIAutomation->ElementFromHandle(g_hForegroundWindow, &pWindow)) || !pWindow)
        return result;

    VARIANT varTrue;
    varTrue.vt = VT_BOOL;
    varTrue.boolVal = VARIANT_TRUE;

    // Search for TextPattern2 first, then TextPattern (v1)
    PROPERTYID patternProps[] = { UIA_IsTextPattern2AvailablePropertyId, UIA_IsTextPatternAvailablePropertyId };
    const char* patternNames[] = { "TextPattern2", "TextPattern" };

    for (int i = 0; i < 2 && result.y == 0; i++)
    {
        IUIAutomationCondition* pCondition = nullptr;
        HRESULT hr = g_pUIAutomation->CreatePropertyCondition(patternProps[i], varTrue, &pCondition);
        if (SUCCEEDED(hr) && pCondition)
        {
            IUIAutomationElement* pChild = nullptr;
            hr = pWindow->FindFirst(TreeScope_Descendants, pCondition, &pChild);
            if (SUCCEEDED(hr) && pChild)
            {
                BSTR childName = nullptr;
                pChild->get_CurrentName(&childName);
                OutputDebugFormatA("UIA: found %s element: '%S'\n", patternNames[i], childName ? childName : L"(null)");
                if (childName) SysFreeString(childName);
                TryGetCaretFromElement(pChild, result);
                pChild->Release();
            }
            else
            {
                OutputDebugFormatA("UIA: no %s element in window tree\n", patternNames[i]);
            }
            pCondition->Release();
        }
    }

    pWindow->Release();
    cachedResult = result;
    return result;
}

POINT GetCaretPositionFromAccessibility(HWND* pCaretWnd)
{
    long x=0,y=0,w=0,h=0;

    IAccessible* pAcc = nullptr;
    if (AccessibleObjectFromWindow(g_hForegroundWindow, OBJID_CARET, IID_IAccessible, (void**)&pAcc) == S_OK)
    {
        if (pAcc)
        {
            pAcc->accLocation(&x, &y, &w, &h, {CHILDID_SELF} );

            // Ask IAccessible which HWND owns this caret
            if (pCaretWnd)
            {
                HWND hwnd = nullptr;
                WindowFromAccessibleObject(pAcc, &hwnd);
                *pCaretWnd = hwnd;
            }

            pAcc->Release();
        }
    }

    return {x + w/2, y + h/2};
}

int GetTaskbarHeight()
{
    // Get the work area dimensions
    RECT workArea = {0};
    SystemParametersInfo(SPI_GETWORKAREA, 0, &workArea, 0);
    int workAreaHeight = workArea.bottom - workArea.top;

    // Calculate the taskbar height
    int taskbarHeight = g_iScreenHeight - workAreaHeight;
    return taskbarHeight;
}

// ---------------------------------------------------------------------------
// Per-app container detection config lookup (mirrors GetCaretMethodsForWindow)
// ---------------------------------------------------------------------------
int GetContainerMethodsForWindow(HWND hwnd)
{
#ifdef USE_PER_APP_CARET_CONFIG
    static HWND  cachedHwnd = nullptr;
    static int   cachedMethods = CONTAINER_METHOD_ALL;

    if (hwnd == cachedHwnd)
        return cachedMethods;

    cachedHwnd = hwnd;
    cachedMethods = CONTAINER_METHOD_ALL;

    DWORD pid = 0;
    GetWindowThreadProcessId(hwnd, &pid);
    if (!pid) return cachedMethods;

    HANDLE hProc = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, pid);
    if (!hProc) return cachedMethods;

    char fullPath[MAX_PATH] = {};
    DWORD pathLen = MAX_PATH;
    QueryFullProcessImageNameA(hProc, 0, fullPath, &pathLen);
    CloseHandle(hProc);

    char exeName[MAX_PATH] = {};
    const char* lastSlash = strrchr(fullPath, '\\');
    const char* name = lastSlash ? lastSlash + 1 : fullPath;
    for (int i = 0; name[i] && i < MAX_PATH - 1; i++)
        exeName[i] = (char)tolower((unsigned char)name[i]);

    #define APP_ENTRY(exe, caretM, containerM) if (strcmp(exeName, exe) == 0) { cachedMethods = containerM; return cachedMethods; }
    APP_CONFIG_TABLE
    #undef APP_ENTRY

    return cachedMethods;
#else
    return CONTAINER_METHOD_ALL;
#endif
}

// ---------------------------------------------------------------------------
// GetContainerRectFromUIA — use UIA ElementFromPoint to get the bounding
// rectangle of the UI element at the given screen coordinates.
// Returns a non-empty RECT on success, or {0,0,0,0} on failure.
// ---------------------------------------------------------------------------
RECT GetContainerRectFromUIA(int x, int y)
{
    RECT result = {};
    if (!g_pUIAutomation) return result;

    POINT pt = { x, y };
    IUIAutomationElement* pElement = nullptr;
    HRESULT hr = g_pUIAutomation->ElementFromPoint(pt, &pElement);
    if (FAILED(hr) || !pElement) return result;

    RECT rc;
    hr = pElement->get_CurrentBoundingRectangle(&rc);
    if (SUCCEEDED(hr) && (rc.right - rc.left) > 50 && (rc.bottom - rc.top) > 10)
        result = rc;

    pElement->Release();
    return result;
}

// ---------------------------------------------------------------------------
// FindSmallestChildContainingXY — find the smallest child window (by area)
// that contains the given (x,y) point. Skips children as large as the parent.
// ---------------------------------------------------------------------------
struct FindChildByXY_Ctx
{
    int      x;
    int      y;
    int      parentArea;
    HWND     best;
    int      bestArea;
    int      totalChildren;
};

static BOOL CALLBACK EnumChildProc_FindByXY(HWND hwnd, LPARAM lParam)
{
    auto* ctx = reinterpret_cast<FindChildByXY_Ctx*>(lParam);
    ctx->totalChildren++;

    RECT rc;
    GetWindowRect(hwnd, &rc);
    int w = rc.right - rc.left;
    int h = rc.bottom - rc.top;
    int area = w * h;

    if (w < 50 || h < 10)       return TRUE;  // skip tiny windows
    if (area >= ctx->parentArea) return TRUE;  // skip if as large as parent

    if (ctx->x >= rc.left && ctx->x <= rc.right &&
        ctx->y >= rc.top  && ctx->y <= rc.bottom)
    {
        if (!ctx->best || area < ctx->bestArea) // smallest wins
        {
            ctx->best = hwnd;
            ctx->bestArea = area;
        }
    }
    return TRUE;
}

HWND FindSmallestChildContainingXY(HWND hwndParent, int x, int y, int* pChildCount)
{
    RECT rcParent;
    GetWindowRect(hwndParent, &rcParent);
    int parentArea = (rcParent.right - rcParent.left) * (rcParent.bottom - rcParent.top);

    FindChildByXY_Ctx ctx = { x, y, parentArea, nullptr, 0, 0 };
    EnumChildWindows(hwndParent, EnumChildProc_FindByXY, (LPARAM)&ctx);
    if (pChildCount) *pChildCount = ctx.totalChildren;
    return ctx.best;
}

#ifdef RENDER_CURSOR
void RenderScaledCursor(HDC hTargetDC, HWND hWnd, int width, int height)
{
    // Get cursor info
    CURSORINFO ci;
    ci.cbSize = sizeof(CURSORINFO);
    if (GetCursorInfo(&ci) && ci.flags == CURSOR_SHOWING)
//    if (ci.hCursor == LoadCursor(NULL, IDC_IBEAM)) // or - "if not arrow, not size, etc...."
    {
        // Get the cursor bitmap
        ICONINFO iconInfo;
        GetIconInfo(ci.hCursor, &iconInfo);

        // Get original cursor size
        BITMAP bmp;
        GetObject(iconInfo.hbmMask, sizeof(BITMAP), &bmp);
        int cursorWidth = bmp.bmWidth;
        int cursorHeight = bmp.bmHeight;

        // Scale the cursor size
        int scaledWidth = (int)(cursorWidth * g_fScaleFactor);
        int scaledHeight = (int)(cursorHeight * g_fScaleFactor);

        // Adjust cursor position based on the hotspot
        int hotspotX = (int)(iconInfo.xHotspot * g_fScaleFactor);
        int hotspotY = (int)(iconInfo.yHotspot * g_fScaleFactor);

        // Convert cursor screen position to local coordinates, accounting for the hotspot
        POINT cursorPos = ci.ptScreenPos;
        int drawX = cursorPos.x - hotspotX;
        int drawY = cursorPos.y - hotspotY;

        // Draw the scaled cursor
        DrawIconEx(hTargetDC, drawX, drawY, ci.hCursor,
                   scaledWidth, scaledHeight, 0, NULL, DI_NORMAL);

        // Clean up
        DeleteObject(iconInfo.hbmMask);
        DeleteObject(iconInfo.hbmColor);
    }
}
#endif // RENDER_CURSOR

#ifdef USE_FONT_HEIGHT
int GetFontHeight(HWND hWnd)
{
    HDC hdc = GetDC(hWnd);

    // Get the font used by the window
    HFONT hFont = (HFONT)SendMessage(hWnd, WM_GETFONT, 0, 0);
    if (hFont) {
        SelectObject(hdc, hFont);
    }

    // Retrieve font metrics
    TEXTMETRIC tm;
    GetTextMetrics(hdc, &tm);

    ReleaseDC(hWnd, hdc);
    return (int)(tm.tmHeight * g_fDpiScaleFactor + 0.5f);
}
#endif

void EnableRoundedCorners(HWND hwnd)
{
    // Define the attribute and the preference for rounded corners
    const DWORD DWMWA_WINDOW_CORNER_PREFERENCE = 33;
    enum DWM_WINDOW_CORNER_PREFERENCE
    {
        DWMWCP_DEFAULT = 0,
        DWMWCP_DONOTROUND = 1,
        DWMWCP_ROUND = 2,
        DWMWCP_ROUNDSMALL = 3
    };

    DWM_WINDOW_CORNER_PREFERENCE preference = DWMWCP_ROUND;

    // Apply the rounded corners preference to the window
    HRESULT hr = DwmSetWindowAttribute(
        hwnd,
        DWMWA_WINDOW_CORNER_PREFERENCE,
        &preference,
        sizeof(preference)
    );

    if (FAILED(hr))
    {
        // Handle the error if needed
    }
}