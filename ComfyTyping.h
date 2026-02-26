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
#define CARET_TIMER_INTERVAL 1000   // Background poll (not the main driver; event-driven detection handles most updates)

//#define USE_ANIMATION
//#define USE_CACHING_DC
#define ANIM_DURATION_MS     500    // Show/hide/resize animation duration

// Event-driven caret detection: hooks PostMessage instead of setting globals
#define WM_APP_DETECT_CARET  (WM_APP + 1)
#define DETECT_REASON_KEY    0      // lParam = virtual key code
#define DETECT_REASON_MOUSE  1
#define DETECT_REASON_SETTLE 2      // fired by settle timer to re-detect until caret stabilizes
#define SETTLE_TIMER_ID      1003
#define SETTLE_TIMER_MS      15

#define KEY_DOWN(key)    ((key < 0) ? 0 : (GetAsyncKeyState(key) & 0x8000) !=0 )
#define KEY_TOGGLED(key) (                (GetKeyState     (key) & 0x0001) !=0 )
#define DO_WHEN_BECOMES_TRUE(condition,action) { bool kd = (condition); static bool pr_kd = kd; if (kd && !pr_kd) { action; } pr_kd = kd; }

// ---------------------------------------------------------------------------
// Per-app caret detection configuration
// ---------------------------------------------------------------------------
// Comment out USE_PER_APP_CARET_CONFIG to use the full fallback chain for all
// apps (useful for testing compatibility with unknown apps).
#define USE_PER_APP_CARET_CONFIG

// When USE_PER_APP_CARET_CONFIG is defined, per-app detection methods are
// looked up from g_appConfigTable[] in WinUtils.cpp. Unknown apps get _ALL
// for both caret and container methods.

