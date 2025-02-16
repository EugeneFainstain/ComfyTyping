#pragma once

#define INVALIDATE_ON_TIMER
#define INVALIDATE_ON_HOOK
//#define RENDER_CURSOR // It seems that we don't need this for our use case...
//#define USE_FONT_HEIGHT



#ifdef INVALIDATE_ON_TIMER
    #define INVALIDATE_TIMER_ID       1001  // Unique timer ID
    #define INVALIDATE_TIMER_INTERVAL 16    // 16 ms interval
#endif

#define CARET_TIMER_ID       1002   // Unique timer ID
#define CARET_TIMER_INTERVAL 1      //16 // 16 ms interval
