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
MYGLOBAL(int          , g_iEffectiveWidth , 100    ); // <= g_iMyWidth, reduced when container is narrow
MYGLOBAL(int          , g_iMyHeight       , 100    );

MYGLOBAL(int          , g_iMyWindowX      , 100    );
MYGLOBAL(int          , g_iMyWindowY      , 100    );

MYGLOBAL(int          , g_iSrcX           , 0      );
MYGLOBAL(int          , g_iSrcY           , 0      );

MYGLOBAL(HWND         , g_hForegroundWindow, nullptr);

MYGLOBAL(HWND         , g_hFocusedChildWnd , nullptr);
MYGLOBAL(HWINEVENTHOOK, g_hEventHook       , nullptr);
MYGLOBAL(POINT        , g_ptCaret          , {}     );
MYGLOBAL(POINT        , g_ptLastQueriedCaret     , {}     ); // most recent detection result
MYGLOBAL(POINT        , g_ptLastSettledCaret      , {}     ); // last stable caret (after settle completed)

MYGLOBAL(bool         , g_bCaretFromAccessibility, false);

MYGLOBAL(bool         , g_bFreezeSrcXandSrcY, false   );
MYGLOBAL(int          , g_iTakeCaretSnapshot, 0       );
MYGLOBAL(POINT        , g_ptCaretSnapshot   , {}      );

MYGLOBAL(bool         , g_bTemporarilyHideMyWindow, true); // starts disabled; activated by typing 4 chars or arrow jiggle
MYGLOBAL(bool         , g_bOverlayEnabled         , false); // true while overlay is active (typing/jiggle); false at start and after ESC

MYGLOBAL(int          , g_iNewFocusedChild, 0);

MYGLOBAL(DWORD        , g_dwAppStartTime  , 0);
MYGLOBAL(RECT         , g_rcContainer     , {});
MYGLOBAL(RECT         , g_rcLastQueriedContainer  , {}); // most recent detection result
MYGLOBAL(RECT         , g_rcLastSettledContainer   , {}); // last stable container (after settle completed)
MYGLOBAL(bool         , g_bCaretMightHaveMoved, true); // true initially to detect on first tick
MYGLOBAL(bool         , g_bWaitForInputAfterToggle, false); // set after VSCode screen reader toggle; cleared by hooks
MYGLOBAL(bool         , g_bAllowOptimizations , true); // false when SCROLL_LOCK is lit
MYGLOBAL(bool         , g_bSettling           , false); // true while settle loop runs; suppresses desktop grab

// Zoom-in animation state (see block comment in ComfyTyping.cpp above AdvanceAnimation)
MYGLOBAL(int          , g_iAnimWidth  , 0    ); // current animated width  (0..g_iAnimToW)
MYGLOBAL(int          , g_iAnimHeight , 0    ); // current animated height (0..g_iAnimToH)
MYGLOBAL(LONGLONG     , g_llAnimStart , 0    ); // QPC tick when animation began
MYGLOBAL(int          , g_iAnimFromW  , 0    ); // width at animation start  (for mid-animation restart)
MYGLOBAL(int          , g_iAnimFromH  , 0    ); // height at animation start (for mid-animation restart)
MYGLOBAL(int          , g_iAnimToW    , 0    ); // target width  (= g_iEffectiveWidth)
MYGLOBAL(int          , g_iAnimToH    , 0    ); // target height (= g_iMyHeight)
MYGLOBAL(bool         , g_bFillCacheOnly, false); // when true, WM_PAINT grabs desktop to cache but skips blit
MYGLOBAL(HDC          , g_hAnimBgDC , NULL ); // captured desktop behind overlay, painted once at animation start

#ifdef WIN_UTILS_CPP
    char g_szAppExeName[64] = {};
#else
    extern char g_szAppExeName[64];
#endif

// Detection method flags - all share one bitmask per app.
// Caret methods (tried in order: GuiThreadInfo -> IAccessible -> UIA)
#define CARET_GTHI      0x001
#define CARET_IACC      0x002
#define CARET_UIA       0x004
#define CARET_ALL       0x007

// Container methods (all tried, narrowest width wins)
#define CONTAINER_HOOK  0x010
#define CONTAINER_ENUM  0x020
#define CONTAINER_UIA   0x040
#define CONTAINER_ALL   0x070

// Behavior flags
#define REUSE_CARET_ON_TYPING_LOSS      0x080  // if caret lost after typing key, reuse last settled caret
#define SKIP_CONTAINER_UPDATE_ON_TYPING 0x100  // don't update container when settle completes from a typing key

#define BEHAVIOR_ALL (REUSE_CARET_ON_TYPING_LOSS | SKIP_CONTAINER_UPDATE_ON_TYPING)

#ifdef NO_CONTAINER_FROM_HOOK
    #define METHOD_ALL ((CARET_ALL | CONTAINER_ALL | BEHAVIOR_ALL) & ~CONTAINER_HOOK)
#else
    #define METHOD_ALL (CARET_ALL | CONTAINER_ALL | BEHAVIOR_ALL)
#endif

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
bool  ShouldReuseCaretOnTypingLoss(HWND hwnd);
bool  ShouldSkipContainerUpdateOnTyping(HWND hwnd);
void  EnableRoundedCorners(HWND hwnd);

#ifdef USE_FONT_HEIGHT
    int GetFontHeight(HWND hWnd);
#endif

#ifdef RENDER_CURSOR
    void RenderScaledCursor(HDC hTargetDC, HWND hWnd, int width, int height);
#endif