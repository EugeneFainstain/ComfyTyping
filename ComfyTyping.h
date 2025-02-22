#pragma once

#define ZOOM          2
#define VERT_PORTION 27 //16 // 1/16th of screen height
#define HORZ_PORTION  8 // Portion of the width that is not occupied...

#define INVALIDATE_ON_TIMER
#define INVALIDATE_ON_HOOK
//#define RENDER_CURSOR // It seems that we don't need this for our use case...
//#define USE_FONT_HEIGHT

//#define HAS_MENU
//#define HAS_TITLE_BAR
//#define POSITION_OVER_TASKBAR
#define NO_RESIZE_NO_BORDER

#ifdef INVALIDATE_ON_TIMER
    #define INVALIDATE_TIMER_ID       1001  // Unique timer ID
    #define INVALIDATE_TIMER_INTERVAL 16    // 16 ms interval
#endif

#define CARET_TIMER_ID       1002   // Unique timer ID
#define CARET_TIMER_INTERVAL 1      //16 // 16 ms interval

#define KEY_DOWN(key)    ((key < 0) ? 0 : (GetAsyncKeyState(key) & 0x8000) !=0 )
#define KEY_TOGGLED(key) (                (GetKeyState     (key) & 0x0001) !=0 )
#define DO_WHEN_BECOMES_TRUE(condition,action) { bool kd = (condition); static bool pr_kd = kd; if (kd && !pr_kd) { action; } pr_kd = kd; }

#pragma once


