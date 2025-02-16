// WinUtils.cpp : helper functions for Windows programming
//
#include "framework.h"
#include "ComfyTypingDefines.h"
#include "WinUtils.h"

#include <oleacc.h>  // Include for IAccessible

#pragma comment(lib, "Oleacc.lib")  // Link the required library

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

int GetCaretY()
{
    GUITHREADINFO gti;
    gti.cbSize = sizeof(GUITHREADINFO);

    // Get caret info from the active thread
    if (GetGUIThreadInfo(0, &gti))
    if (gti.hwndCaret)
    {
        POINT caretPos = { 0, 0 };

        // Get caret position in local (client) coordinates
        caretPos.x = gti.rcCaret.left;
        caretPos.y = gti.rcCaret.top;

        // Convert client coordinates to screen coordinates
        ClientToScreen(gti.hwndCaret, &caretPos);
        
        if( caretPos.y > 2160*3 )
        { int i = 5; }

        if( caretPos.y < 0 )
        { int i = 5; }

        return caretPos.y;
    }

    return 0;
}

int GetCaretPositionFromAccessibility()
{
    long x=0,y=0,w=0,h=0;

    HWND hWnd = GetForegroundWindow();

    static HWND hPrevWnd     = nullptr;
    static IAccessible* pAcc = nullptr;

    if (hWnd != hPrevWnd)
    {
        if (pAcc)
        {
            pAcc->Release();
            pAcc = nullptr;
        }

        if (AccessibleObjectFromWindow(hWnd, OBJID_CARET, IID_IAccessible, (void**)&pAcc) == S_OK)
        {
            hPrevWnd = hWnd;
        }
    }

    VARIANT varChild = {}; // CHILDID_SELF
    if (pAcc)
    {
        // Get the caret's screen position
        if (pAcc->accLocation(&x, &y, &w, &h, varChild) == S_OK)
            { int i = 5; }
        //OutputDebugFormatA("Caret Position: X=%d Y=%d\n", x, y);
    }

//     IAccessible* pAcc = nullptr;
//     VARIANT varChild = {}; // CHILDID_SELF
//     if (AccessibleObjectFromWindow(hWnd, OBJID_CARET, IID_IAccessible, (void**)&pAcc) == S_OK) {
//         if (pAcc) {
//             // Get the caret's screen position
//             if (pAcc->accLocation(&x, &y, &w, &h, varChild) == S_OK) {
//                 //std::cout << "Caret Position: X=" << caretPos.x << " Y=" << caretPos.y << std::endl;
//                 OutputDebugFormatA("Caret Position: X=%d Y=%d\n", x, y);
//             }
//             pAcc->Release();
//         }
//     }

    return (int)y;
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

int GetFontHeight(HWND hWnd) {
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
