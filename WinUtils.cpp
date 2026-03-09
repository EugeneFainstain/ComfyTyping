// WinUtils.cpp : helper functions for Windows programming
//
#include "framework.h"
#include "WinUtils.h"
#include "IniReader.h"

#include <objbase.h>             // For COM (CoCreateInstance, etc.)
#include <uiautomationclient.h>  // For IUIAutomation / TextPattern2

#include <oleacc.h>  // For IAccessible (MSAA fallback)
#pragma comment(lib, "Oleacc.lib")

#include <dwmapi.h>  // For EnableRoundedCorners()
#pragma comment(lib, "dwmapi.lib")

static IUIAutomation* g_pUIAutomation = nullptr;

// ---------------------------------------------------------------------------
// Per-app detection configuration
// ---------------------------------------------------------------------------
#ifdef USE_PER_APP_CARET_CONFIG
struct AppConfig
{
    const char* exeName;   // lowercase exe name
    int         methods;   // CARET_* | CONTAINER_* flags
};

// Per-app detection methods. Exe names are lowercase, matched against the
// foreground window's process.
//
// Caret flags (tried in order: GuiThreadInfo -> IAccessible -> UIA):
//   CARET_GTHI (0x001), CARET_IACC (0x002),
//   CARET_UIA (0x004), CARET_ALL (0x007)
//
// Container flags (all tried, narrowest width wins):
//   CONTAINER_HOOK (0x010), CONTAINER_ENUM (0x020),
//   CONTAINER_UIA (0x040), CONTAINER_ALL (0x070)
//
// Unknown apps get METHOD_ALL (0x077).
//
static const AppConfig g_appConfigTable[] = {
    { "dummy_app.exe", METHOD_ALL },

    { "devenv.exe"   , METHOD_ALL | APP_ID_DEVENV  },
    { "code.exe"     , METHOD_ALL | APP_ID_CODE    },
    { "msedge.exe"   , METHOD_ALL | APP_ID_BROWSER },
    { "chrome.exe"   , METHOD_ALL | APP_ID_BROWSER },
    { "firefox.exe"  , METHOD_ALL | APP_ID_BROWSER },

//         { "devenv.exe",          CARET_GTHI | CARET_IACC |                              CONTAINER_ENUM | CONTAINER_UIA },
//         { "code.exe", 0xFF },
//     //    { "code.exe", 0xFF & (~CONTAINER_HOOK) & (~CONTAINER_ENUM) },
//     //    { "code.exe", 0x0F | CONTAINER_UIA },
//         { "cmd.exe",                                       CARET_UIA |                                   CONTAINER_UIA },
//         { "notepad++.exe",       CARET_GTHI | CARET_IACC             | CONTAINER_HOOK | CONTAINER_ENUM                 },
//         { "windowsterminal.exe",                           CARET_UIA |                                   CONTAINER_UIA },
};

// Helper: extract lowercase exe name from a window's owning process.
// Returns true on success, false if pid/process couldn't be queried.
static bool GetExeNameForWindow(HWND hwnd, char* outExeName, int outSize)
{
    DWORD pid = 0;
    GetWindowThreadProcessId(hwnd, &pid);
    if (!pid) return false;

    HANDLE hProc = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, pid);
    if (!hProc) return false;

    char fullPath[MAX_PATH] = {};
    DWORD pathLen = MAX_PATH;
    QueryFullProcessImageNameA(hProc, 0, fullPath, &pathLen);
    CloseHandle(hProc);

    const char* lastSlash = strrchr(fullPath, '\\');
    const char* name = lastSlash ? lastSlash + 1 : fullPath;
    for (int i = 0; name[i] && i < outSize - 1; i++)
        outExeName[i] = (char)tolower((unsigned char)name[i]);

    return true;
}

// Lookup: returns the combined methods mask for a window's owning process.
// Caches by HWND; also sets g_szAppExeName as a side effect.
static int GetMethodsForWindow(HWND hwnd)
{
    static HWND  cachedHwnd = nullptr;
    static int   cachedMethods = METHOD_ALL;

    if (hwnd == cachedHwnd)
        return cachedMethods;

    cachedHwnd = hwnd;
    cachedMethods = METHOD_ALL; // default for unknown apps

    char exeName[MAX_PATH] = {};
    if (!GetExeNameForWindow(hwnd, exeName, MAX_PATH))
        return cachedMethods;

    strncpy_s(g_szAppExeName, exeName, _TRUNCATE);

    // Strip ".exe" suffix for INI lookups (INI keys are base names only)
    char baseName[MAX_PATH] = {};
    strncpy_s(baseName, exeName, _TRUNCATE);
    char* dot = strrchr(baseName, '.');
    if (dot) *dot = '\0';

    // INI access control: blacklist / whitelist / only-whitelisted mode
    if (IsAppBlacklisted(baseName))
    {
        cachedMethods = 0;
        OutputDebugFormatA("App '%s': blacklisted, skipping\n", exeName);
        return cachedMethods;
    }

    if (IsOnlyWhitelistedMode() && !IsAppWhitelisted(baseName))
    {
        cachedMethods = 0;
        OutputDebugFormatA("App '%s': not whitelisted (ONLY_WHITELISTED=1), skipping\n", exeName);
        return cachedMethods;
    }

    // Per-app method table lookup (for fine-tuning detection flags)
    for (int i = 0; i < _countof(g_appConfigTable); i++)
    {
        if (strcmp(exeName, g_appConfigTable[i].exeName) == 0)
        {
            cachedMethods = g_appConfigTable[i].methods;
            return cachedMethods;
        }
    }

    OutputDebugFormatA("App '%s': no config, using all methods\n", exeName);
    return cachedMethods;
}
#endif

// ---------------------------------------------------------------------------
// Public accessors - extract caret or container bits from the combined mask
// ---------------------------------------------------------------------------
int GetCaretMethodsForWindow(HWND hwnd)
{
#ifdef USE_PER_APP_CARET_CONFIG
    return GetMethodsForWindow(hwnd) & CARET_ALL;
#else
    return CARET_ALL;
#endif
}

int GetAppId()
{
    return GetMethodsForWindow(g_hForegroundWindow) & APP_ID_MASK;
}

bool ShouldReuseCaretOnTypingLoss(HWND hwnd)
{
    return (GetMethodsForWindow(hwnd) & REUSE_CARET_ON_TYPING_LOSS) != 0;
}

bool ShouldSkipContainerUpdateOnTyping(HWND hwnd)
{
    return (GetMethodsForWindow(hwnd) & SKIP_CONTAINER_UPDATE_ON_TYPING) != 0;
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
        // Collapsed (zero-length) range - try expanding to get bounding rects
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

#ifdef UIA_LINE_LEVEL_EXPANSION
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
#endif // UIA_LINE_LEVEL_EXPANSION
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

// Try to get caret position from a UIA element - tries TextPattern2 then TextPattern
// Sets *pHadTextPattern = true if the element exposed a TextPattern (even if no rects)
static bool TryGetCaretFromElement(IUIAutomationElement* pElement, POINT& result, bool* pHadTextPattern = nullptr)
{
    if (pHadTextPattern) *pHadTextPattern = false;
    // Try TextPattern2::GetCaretRange first (best method)
    IUIAutomationTextPattern2* pTextPattern2 = nullptr;
    HRESULT hr = pElement->GetCurrentPatternAs(UIA_TextPattern2Id,
                     __uuidof(IUIAutomationTextPattern2), (void**)&pTextPattern2);
    if (SUCCEEDED(hr) && pTextPattern2)
    {
        if (pHadTextPattern) *pHadTextPattern = true;
        BOOL isActive = FALSE;
        IUIAutomationTextRange* pCaretRange = nullptr;
        hr = pTextPattern2->GetCaretRange(&isActive, &pCaretRange);
        OutputDebugFormatA("UIA: TextPattern2 GetCaretRange hr=0x%08X isActive=%d pRange=%p\n",
                           hr, (int)isActive, pCaretRange);
        if (SUCCEEDED(hr) && pCaretRange)
        {
            bool found = GetPositionFromTextRange(pCaretRange, result);
            pCaretRange->Release();
            pTextPattern2->Release();
            if (found)
            {
                OutputDebugFormatA("UIA: (via TextPattern2 GetCaretRange)\n");
                return true;
            }
            OutputDebugFormatA("UIA: TextPattern2 GetCaretRange returned range but GetPositionFromTextRange failed\n");
        }
        else
            pTextPattern2->Release();
    }
    else
    {
        OutputDebugFormatA("UIA: TextPattern2 not available (hr=0x%08X)\n", hr);
    }

    // EXPERIMENT: prefer v2, fall back to v1 if v2 didn't find anything
    bool v2found = (result.y != 0);

    // Try TextPattern::GetSelection (v1 fallback)
    IUIAutomationTextPattern* pTextPattern = nullptr;
    hr = pElement->GetCurrentPatternAs(UIA_TextPatternId,
             __uuidof(IUIAutomationTextPattern), (void**)&pTextPattern);
    if (SUCCEEDED(hr) && pTextPattern)
    {
        if (pHadTextPattern) *pHadTextPattern = true;
        POINT v1result = {};
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
                    GetPositionFromTextRange(pRange, v1result);
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
        if (v2found)
        {
            if (v1result.y != 0)
                OutputDebugFormatA("UIA: (v1 also found (%d,%d) - not used, v2 wins)\n", v1result.x, v1result.y);
        }
        else if (v1result.y != 0)
        {
            result = v1result;
            OutputDebugFormatA("UIA: (via TextPattern v1 GetSelection - v2 fallback)\n");
            return true;
        }
    }

    return v2found;
}

// Send Shift+Alt+F1 keystroke to the foreground window
static void SendShiftAltF1()
{
    INPUT inputs[6] = {};
    inputs[0].type = INPUT_KEYBOARD;  inputs[0].ki.wVk = VK_SHIFT;
    inputs[1].type = INPUT_KEYBOARD;  inputs[1].ki.wVk = VK_MENU;
    inputs[2].type = INPUT_KEYBOARD;  inputs[2].ki.wVk = VK_F1;
    inputs[3].type = INPUT_KEYBOARD;  inputs[3].ki.wVk = VK_F1;    inputs[3].ki.dwFlags = KEYEVENTF_KEYUP;
    inputs[4].type = INPUT_KEYBOARD;  inputs[4].ki.wVk = VK_MENU;  inputs[4].ki.dwFlags = KEYEVENTF_KEYUP;
    inputs[5].type = INPUT_KEYBOARD;  inputs[5].ki.wVk = VK_SHIFT; inputs[5].ki.dwFlags = KEYEVENTF_KEYUP;
    SendInput(6, inputs, sizeof(INPUT));
}

POINT GetCaretPositionFromUIA()
{
    POINT result = {};
    static POINT s_lastResult = {};

    if (!g_pUIAutomation)
        return result;

    // --- Auto-toggle VSCode screen reader mode (one-time per process) ---
    // VSCode doesn't expose text geometry without screen reader mode.
    // Toggling it ON then OFF primes the accessibility layer permanently.
    static DWORD s_toggledPids[16] = {};
    static int   s_toggledPidCount = 0;
    static DWORD s_pendingTogglePid = 0;
    static int   s_toggleState = 0;  // 0=idle, 1=first keystroke sent

    // After toggle, wait for real user input before running UIA detection
    if (g_bWaitForInputAfterToggle)
    {
        OutputDebugFormatA("UIA: waiting for user input after toggle\n");
        return result;
    }

    if (s_toggleState == 1)
    {
        // Second step: send Shift+Alt+F1 again to toggle screen reader OFF
        SendShiftAltF1();
        OutputDebugFormatA("UIA: sending Shift+Alt+F1 (2/2, pid=%d) - accessibility primed, waiting for input...\n", s_pendingTogglePid);
        if (s_toggledPidCount < _countof(s_toggledPids))
            s_toggledPids[s_toggledPidCount++] = s_pendingTogglePid;
        s_pendingTogglePid = 0;
        s_toggleState = 0;
        g_bWaitForInputAfterToggle = true;
        return result;
    }

    // Throttle: UIA calls are expensive (especially FindFirst on large trees)
    // COM/UIA creates worker threads on each call - must limit aggressively
    static DWORD lastCallTime = 0;
    static POINT cachedResult = {};
    if (!g_bAllowOptimizations)
    {
        cachedResult = {};
        lastCallTime = 0;
    }
    DWORD now = GetTickCount();
    if (g_bAllowOptimizations && now - lastCallTime < 100)  // Max ~10 calls/sec
        return cachedResult;
    lastCallTime = now;

    // 1. Try the focused element directly
    IUIAutomationElement* pFocused = nullptr;
    LARGE_INTEGER llA, llB, llF;
    QueryPerformanceCounter(&llA);
    HRESULT hrFocus = g_pUIAutomation->GetFocusedElement(&pFocused);
    QueryPerformanceCounter(&llB);
    QueryPerformanceFrequency(&llF);
    float msFocus = (float)(llB.QuadPart - llA.QuadPart) * 1000.0f / llF.QuadPart;
    if (SUCCEEDED(hrFocus) && pFocused)
    {
        // Log what element we got
        BSTR name = nullptr;
        CONTROLTYPEID ctrlType = 0;
        pFocused->get_CurrentName(&name);
        pFocused->get_CurrentControlType(&ctrlType);
        OutputDebugFormatA("UIA: focused element: type=%d name='%S'  (%.2fms)\n", ctrlType, name ? name : L"(null)", msFocus);

        // Check if VSCode needs screen reader mode toggle
        bool needsToggle = name && wcsstr(name, L"not accessible") && wcsstr(name, L"screen reader") && wcsstr(name, L"Shift+Alt+F1");
        if (name) SysFreeString(name);

        if (needsToggle)
        {
            DWORD pid = 0;
            GetWindowThreadProcessId(g_hForegroundWindow, &pid);

            bool alreadyToggled = false;
            for (int i = 0; i < s_toggledPidCount; i++)
                if (s_toggledPids[i] == pid) { alreadyToggled = true; break; }

            if (!alreadyToggled)
            {
                // First step: send Shift+Alt+F1 to toggle screen reader ON
                SendShiftAltF1();
                OutputDebugFormatA("UIA: VSCode not accessible - sending Shift+Alt+F1 (1/2, pid=%d)\n", pid);
                s_pendingTogglePid = pid;
                s_toggleState = 1;
                pFocused->Release();
                return result;  // skip this tick
            }
        }

        bool hadTextPattern = false;
        if (TryGetCaretFromElement(pFocused, result, &hadTextPattern))
        {
            pFocused->Release();
            cachedResult = result;
            s_lastResult = result;
            return result;
        }
        pFocused->Release();

        // If the focused element had a TextPattern but couldn't resolve coordinates
        // (e.g. empty line), don't fall through to the subtree search — it would
        // find a different element (like the editor) and return wrong coordinates.
        if (hadTextPattern)
        {
            OutputDebugFormatA("UIA: empty line - returning previous (%d,%d)\n", s_lastResult.x, s_lastResult.y);
            return s_lastResult;
        }
    }

#ifdef UIA_SUBTREE_SEARCH
    // 2. Try from the foreground window's root element - search entire window tree
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
            QueryPerformanceCounter(&llA);
            hr = pWindow->FindFirst(TreeScope_Descendants, pCondition, &pChild);
            QueryPerformanceCounter(&llB);
            QueryPerformanceFrequency(&llF);
            float msFind = (float)(llB.QuadPart - llA.QuadPart) * 1000.0f / llF.QuadPart;
            if (SUCCEEDED(hr) && pChild)
            {
                BSTR childName = nullptr;
                pChild->get_CurrentName(&childName);
                OutputDebugFormatA("UIA: found %s element: '%S'  (%.2fms)\n", patternNames[i], childName ? childName : L"(null)", msFind);
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
#endif // UIA_SUBTREE_SEARCH
    cachedResult = result;
    if (result.y != 0) s_lastResult = result;
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

int GetContainerMethodsForWindow(HWND hwnd)
{
#ifdef USE_PER_APP_CARET_CONFIG
    return GetMethodsForWindow(hwnd) & CONTAINER_ALL;
#else
    return CONTAINER_ALL;
#endif
}

// ---------------------------------------------------------------------------
// GetContainerRectFromUIA - use UIA ElementFromPoint to get the bounding
// rectangle of the UI element at the given screen coordinates.
// Returns a non-empty RECT on success, or {0,0,0,0} on failure.
// ---------------------------------------------------------------------------
RECT GetContainerRectFromUIA(int x, int y)
{
    RECT result = {};
    if (!g_pUIAutomation) return result;

    // Start at the caret position, then walk left to find the parent container.
    // ElementFromPoint returns the deepest element — calling it on the left edge
    // of the current rect walks up to the enclosing container.
    POINT pt = { x, y };
    RECT prevRc = {};
    const RECT nullRc = {};

    for (int iter = 0; iter < 8; iter++)  // safety limit
    {
        IUIAutomationElement* pElement = nullptr;
        LARGE_INTEGER llA, llB, llF;
        QueryPerformanceCounter(&llA);
        HRESULT hr = g_pUIAutomation->ElementFromPoint(pt, &pElement);
        QueryPerformanceCounter(&llB);
        QueryPerformanceFrequency(&llF);
        float msEFP = (float)(llB.QuadPart - llA.QuadPart) * 1000.0f / llF.QuadPart;
        if (FAILED(hr) || !pElement) break;

        RECT rc;
        hr = pElement->get_CurrentBoundingRectangle(&rc);

        BSTR bstrName = nullptr;
        pElement->get_CurrentName(&bstrName);
        BSTR bstrClass = nullptr;
        pElement->get_CurrentClassName(&bstrClass);
        pElement->Release();
        if (FAILED(hr)) { SysFreeString(bstrName); SysFreeString(bstrClass); break; }

        // Truncate name at first newline
        if (bstrName) { WCHAR* nl = wcspbrk(bstrName, L"\r\n"); if (nl) *nl = 0; }

        OutputDebugFormatA("  Container [UIA][%d]  : (%d,%d,%d,%d) %dx%d  '%ls' [%ls]  (%.2fms)\n",
                           iter, rc.left, rc.top, rc.right, rc.bottom,
                           rc.right - rc.left, rc.bottom - rc.top,
                           bstrName ? bstrName : L"",
                           bstrClass ? bstrClass : L"",
                           msEFP);
        SysFreeString(bstrName);
        SysFreeString(bstrClass);

        if( 0==memcmp(&prevRc, &nullRc, sizeof(RECT) ) )
        {
            prevRc = rc; // Must initialize prevRc...
        }
        else
        {
            // Stop if we got the same rect twice in a row
            if (rc.left == prevRc.left && rc.top == prevRc.top &&
                rc.right == prevRc.right && rc.bottom == prevRc.bottom)
                break;
        }

        if( g_bAppIsCode )
            result = rc;
        else
        {
            // By iterating this way, we are looking for a container that is SAME OR LARGER than the one we got
            // initially.
            if( (rc.right  >= prevRc.right)  &&
                (rc.left   <= prevRc.left)   &&
                (rc.top    <= prevRc.top)    &&
                (rc.bottom >= prevRc.bottom) )
                result = rc;
        }

        prevRc = rc;

        // Next iteration X: try to find a larger container starting from the left edge
        pt.x = rc.left + 0;// + 1;

        // Next iteration Y: use the upper-left corner of the current container
        pt.y = rc.top;
    }

    return result;
}

// ---------------------------------------------------------------------------
// FindSmallestChildContainingXY - find the smallest child window (by area)
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