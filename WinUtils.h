#pragma once

#include "ComfyTypingDefines.h"
#include <cstdarg>
#include <cstdio>

#ifdef WIN_UTILS_CPP
    #define MYGLOBAL(x,y,z) x y = z
#else
    #define MYGLOBAL(x,y,z) extern x y
#endif

MYGLOBAL(HWND         , g_myWindowHandle   , nullptr);
MYGLOBAL(HWND         , g_hForegroundWindow, nullptr);
MYGLOBAL(HWINEVENTHOOK, g_hEventHook       , nullptr);
MYGLOBAL(POINT        , g_ptCaret          , {}     );
MYGLOBAL(float        , g_fDpiScaleFactor  , 1.0f   );


void OutputDebugFormatA(const char* format, ...);

float GetDpiScaleFactor(HWND hWnd);
POINT GetCaretPosition();
POINT GetCaretPositionFromAccessibility();

#ifdef USE_FONT_HEIGHT
    int GetFontHeight(HWND hWnd);
#endif

#ifdef RENDER_CURSOR
    void RenderScaledCursor(HDC hTargetDC, HWND hWnd, int width, int height);
#endif