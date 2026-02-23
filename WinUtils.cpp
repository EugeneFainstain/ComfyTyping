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

void OutputDebugFormatA(const char* format, ...)
{
    char buffer[256];
    va_list args;
    va_start(args, format);
    vsnprintf(buffer, sizeof(buffer), format, args);
    va_end(args);
    OutputDebugStringA(buffer);
}

float GetDpiScaleFactor(HWND hWnd)
{
    int dpi = GetDpiForWindow(hWnd);  // Windows 8.1+ (Use 96 DPI as base)
    float fDpiScaleFactor = dpi / 96.0f;
    return fDpiScaleFactor;
}

POINT GetCaretPosition()
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

    if (count < 4)
    {
        // Collapsed (zero-length) range — expand by 1 character and use the left edge
        SafeArrayDestroy(pRects);
        pRects = nullptr;

        IUIAutomationTextRange* pExpanded = nullptr;
        if (SUCCEEDED(pRange->Clone(&pExpanded)) && pExpanded)
        {
            int moved = 0;
            pExpanded->MoveEndpointByUnit(TextPatternRangeEndpoint_End,
                                          TextUnit_Character, 1, &moved);
            if (moved > 0)
            {
                if (SUCCEEDED(pExpanded->GetBoundingRectangles(&pRects)) && pRects)
                {
                    SafeArrayGetLBound(pRects, 1, &lBound);
                    SafeArrayGetUBound(pRects, 1, &uBound);
                    count = uBound - lBound + 1;
                }
            }
            pExpanded->Release();
        }
    }

    if (count >= 4)
    {
        double* pData = nullptr;
        if (SUCCEEDED(SafeArrayAccessData(pRects, (void**)&pData)))
        {
            // Use left edge (x) as caret position, not center of expanded char
            result.x = (LONG)(pData[0]);
            result.y = (LONG)(pData[1] + pData[3] / 2);
            OutputDebugFormatA("UIA: caret at %d, %d (rect: %.0f,%.0f,%.0f,%.0f)\n",
                               result.x, result.y, pData[0], pData[1], pData[2], pData[3]);
            SafeArrayUnaccessData(pRects);
            found = true;
        }
    }
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
                    // Check what GetBoundingRectangles gives us
                    SAFEARRAY* pRects = nullptr;
                    hr = pRange->GetBoundingRectangles(&pRects);
                    if (SUCCEEDED(hr) && pRects)
                    {
                        LONG lBound, uBound;
                        SafeArrayGetLBound(pRects, 1, &lBound);
                        SafeArrayGetUBound(pRects, 1, &uBound);
                        LONG count = uBound - lBound + 1;
                        OutputDebugFormatA("UIA: TextPattern v1 bounding rect elements: %d\n", count);
                        SafeArrayDestroy(pRects);
                    }
                    else
                    {
                        OutputDebugFormatA("UIA: TextPattern v1 GetBoundingRectangles failed hr=0x%08X\n", hr);
                    }

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
    return result;
}

POINT GetCaretPositionFromAccessibility()
{
    long x=0,y=0,w=0,h=0;

    IAccessible* pAcc = nullptr;
    if (AccessibleObjectFromWindow(g_hForegroundWindow, OBJID_CARET, IID_IAccessible, (void**)&pAcc) == S_OK)
    {
        if (pAcc)
        {
            pAcc->accLocation(&x, &y, &w, &h, {CHILDID_SELF} );
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