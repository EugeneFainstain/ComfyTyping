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
MYGLOBAL(bool         , g_bCaretMightHaveMoved, true); // true initially to detect on first tick

#ifdef WIN_UTILS_CPP
    char g_szAppExeName[64] = {};
#else
    extern char g_szAppExeName[64];
#endif

// Detection method flags — all share one bitmask per app.
// Caret methods (tried in order: GuiThreadInfo -> IAccessible -> UIA)
#define CARET_GTHI   0x001
#define CARET_IACC     0x002
#define CARET_UIA             0x004
#define CARET_ALL             0x007

// Container methods (tried in order: HOOK -> ENUM -> UIA)
#define CONTAINER_HOOK        0x010
#define CONTAINER_ENUM        0x020
#define CONTAINER_UIA         0x040
#define CONTAINER_ALL         0x070

#define METHOD_ALL (CARET_ALL | CONTAINER_ALL)

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
HWND  FindSmallestChildContainingXY(HWND hwndParent, int x, int y, int* pChildCount = nullptr);
RECT  GetContainerRectFromUIA(int x, int y);
int   GetContainerMethodsForWindow(HWND hwnd);
void  EnableRoundedCorners(HWND hwnd);

#ifdef USE_FONT_HEIGHT
    int GetFontHeight(HWND hWnd);
#endif

#ifdef RENDER_CURSOR
    void RenderScaledCursor(HDC hTargetDC, HWND hWnd, int width, int height);
#endif