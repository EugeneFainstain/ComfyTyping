#pragma once

#include "ComfyTyping.h"
#include <cstdarg>
#include <cstdio>

#ifdef WIN_UTILS_CPP
    #define MYGLOBAL(x,y,z) x y = z
#else
    #define MYGLOBAL(x,y,z) extern x y
#endif

MYGLOBAL(HWND         , g_myWindowHandle  , nullptr);
MYGLOBAL(float        , g_fDpiScaleFactor , 1.0f   );
MYGLOBAL(int          , g_iScreenWidth    , 100    );
MYGLOBAL(int          , g_iScreenHeight   , 100    );

MYGLOBAL(int          , g_iMyWidth        , 100    );
MYGLOBAL(int          , g_iMyHeight       , 100    );

MYGLOBAL(int          , g_iMyWindowX      , 100    );
MYGLOBAL(int          , g_iMyWindowY      , 100    );

MYGLOBAL(int          , g_iSrcX           , 0      );
MYGLOBAL(int          , g_iSrcY           , 0      );

MYGLOBAL(HWND         , g_hForegroundWindow, nullptr);

MYGLOBAL(HWND         , g_hFocusedChildWnd , nullptr);
MYGLOBAL(HWINEVENTHOOK, g_hEventHook       , nullptr);
MYGLOBAL(POINT        , g_ptCaret          , {}     );

MYGLOBAL(bool         , g_bCaretFromAccessibility, false);

MYGLOBAL(bool         , g_bFreezeSrcXandSrcY, false   );
MYGLOBAL(int          , g_iTakeCaretSnapshot, 0       );
MYGLOBAL(POINT        , g_ptCaretSnapshot   , {}      );

MYGLOBAL(bool         , g_bTemporarilyHideMyWindow, false);

MYGLOBAL(int          , g_iNewFocusedChild, 0);

MYGLOBAL(DWORD        , g_dwAppStartTime  , 0);
MYGLOBAL(RECT         , g_rcContainer     , {});

// Caret detection method flags (combinable)
#define CARET_METHOD_GUITHREADINFO   0x01  // GetGUIThreadInfo (fast Win32)
#define CARET_METHOD_IACCESSIBLE     0x02  // IAccessible/MSAA
#define CARET_METHOD_UIA             0x04  // UI Automation TextPattern
#define CARET_METHOD_ALL (CARET_METHOD_GUITHREADINFO | CARET_METHOD_IACCESSIBLE | CARET_METHOD_UIA)

// Container detection method flags (combinable, tried in order: HOOK -> ENUM -> UIA)
#define CONTAINER_METHOD_HOOK        0x01  // WinEventProc focus hook: g_hFocusedChildWnd
#define CONTAINER_METHOD_ENUM        0x02  // EnumChildWindows: smallest child containing X,Y
#define CONTAINER_METHOD_UIA         0x04  // UIA ElementFromPoint: bounding rect of element at caret
#define CONTAINER_METHOD_ALL (CONTAINER_METHOD_HOOK | CONTAINER_METHOD_ENUM | CONTAINER_METHOD_UIA)

void OutputDebugFormatA(const char* format, ...);
void DebugTraceA(const char* format, ...);

float GetDpiScaleFactor(HWND hWnd);
POINT GetCaretPosition(HWND* pCaretWnd = nullptr);
void  InitUIAutomation();
void  CleanupUIAutomation();
POINT GetCaretPositionFromUIA();
POINT GetCaretPositionFromAccessibility(HWND* pCaretWnd = nullptr);
int   GetTaskbarHeight();
int   GetCaretMethodsForWindow(HWND hwnd);
HWND  FindSmallestChildContainingXY(HWND hwndParent, int x, int y);
RECT  GetContainerRectFromUIA(int x, int y);
int   GetContainerMethodsForWindow(HWND hwnd);
void  EnableRoundedCorners(HWND hwnd);

#ifdef USE_FONT_HEIGHT
    int GetFontHeight(HWND hWnd);
#endif

#ifdef RENDER_CURSOR
    void RenderScaledCursor(HDC hTargetDC, HWND hWnd, int width, int height);
#endif