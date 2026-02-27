// ComfyTyping.cpp : Defines the entry point for the application.
//

#define WIN_UTILS_CPP

#include "framework.h"
#include "resource.h"
#include "ComfyTyping.h"
#include "WinUtils.h"
#include <iostream>  // for std::isprint
#include <objbase.h> // for CoInitializeEx / CoUninitialize

#define MAX_LOADSTRING 100

// Global Variables:
HINSTANCE hInst;                                // current instance
WCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name

void CALLBACK WinEventProc_ForFocusedClientWnd(HWINEVENTHOOK hWinEventHook, DWORD event, HWND hwnd,
                           LONG idObject, LONG idChild, DWORD dwEventThread, DWORD dwmsEventTime);

LRESULT CALLBACK LowLevelKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK LowLevelMouseProc(int nCode, WPARAM wParam, LPARAM lParam);

// Dedicated thread for low-level hooks so UIA calls on the
// main thread can't block input delivery to other apps.
static DWORD WINAPI InputHookThreadProc(LPVOID lpParam)
{
    HHOOK hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, LowLevelKeyboardProc, (HINSTANCE)lpParam, 0);
    HHOOK hMouseHook    = SetWindowsHookEx(WH_MOUSE_LL,    LowLevelMouseProc,    (HINSTANCE)lpParam, 0);

    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    if (hKeyboardHook) UnhookWindowsHookEx(hKeyboardHook);
    if (hMouseHook)    UnhookWindowsHookEx(hMouseHook);
    return 0;
}

// Forward declarations of functions included in this code module:
ATOM                MyRegisterClass(HINSTANCE hInstance);
HWND                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    g_dwAppStartTime = GetTickCount();

    SetProcessDPIAware();  // Enable DPI awareness (Windows 7+)
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    InitUIAutomation();

    // Initialize global strings
    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_COMFYTYPING, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    // Perform application initialization:
    g_myWindowHandle = InitInstance(hInstance, nCmdShow);

    if (!g_myWindowHandle)
    {
        return FALSE;
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_COMFYTYPING));

#ifdef INVALIDATE_ON_TIMER
    SetTimer(g_myWindowHandle, INVALIDATE_TIMER_ID, INVALIDATE_TIMER_INTERVAL, NULL); // Start the timer
#endif

    SetTimer(g_myWindowHandle, CARET_TIMER_ID, CARET_TIMER_INTERVAL, NULL); // Start the timer

    HWINEVENTHOOK hWinEventHook = SetWinEventHook(
                                    EVENT_OBJECT_FOCUS,// [in] DWORD      eventMin,
                                    EVENT_OBJECT_FOCUS,// [in] DWORD      eventMax,
                                    NULL,              // [in] HMODULE    hmodWinEventProc,
                                    WinEventProc_ForFocusedClientWnd,      // [in] WINEVENTPROC pfnWinEventProc,
                                    0,                 // [in] DWORD        idProcess,
                                    0,                 // [in] DWORD        idThread,
                                    WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS | WINEVENT_SKIPOWNTHREAD );//[in] DWORD        dwFlags

    HANDLE hHookThread = CreateThread(nullptr, 0, InputHookThreadProc, hInstance, 0, nullptr);

    if( !hHookThread || !hWinEventHook )
    {
        ::MessageBoxA(0, "Error: please re-run as an Administrator.\n", "ComfyTyping", MB_ICONERROR);
        exit(-1);
    }

    MSG msg;

    // Main message loop:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    CleanupUIAutomation();
    CoUninitialize();
    return (int) msg.wParam;
}

#ifdef INVALIDATE_ON_HOOK
void CALLBACK WinEventProc(HWINEVENTHOOK hWinEventHook, DWORD event, HWND hwnd,
                           LONG idObject, LONG idChild, DWORD dwEventThread, DWORD dwmsEventTime) {
    if (hwnd != g_myWindowHandle) {  // Ignore own window
        InvalidateRect(g_myWindowHandle, NULL, FALSE);
    }
}

void StartDesktopMonitor() {
    g_hEventHook = SetWinEventHook(EVENT_OBJECT_LOCATIONCHANGE, EVENT_OBJECT_NAMECHANGE,
                                 NULL, WinEventProc, 0, 0, WINEVENT_OUTOFCONTEXT);
}

void StopDesktopMonitor() {
    if (g_hEventHook) UnhookWinEvent(g_hEventHook);
}
#endif // INVALIDATE_ON_HOOK

//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style          = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc    = WndProc;
    wcex.cbClsExtra     = 0;
    wcex.cbWndExtra     = 0;
    wcex.hInstance      = hInstance;
    wcex.hIcon          = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_COMFYTYPING));
    wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW+1);
    wcex.lpszMenuName   = MAKEINTRESOURCEW(IDC_COMFYTYPING);
    wcex.lpszClassName  = szWindowClass;
    wcex.hIconSm        = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    return RegisterClassExW(&wcex);
}

//
//   FUNCTION: InitInstance(HINSTANCE, int)
//
//   PURPOSE: Saves instance handle and creates main window
//
//   COMMENTS:
//
//        In this function, we save the instance handle in a global variable and
//        create and display the main program window.
//
HWND InitInstance(HINSTANCE hInstance, int nCmdShow)
{
    hInst = hInstance; // Store instance handle in our global variable

#ifdef NO_RESIZE_NO_BORDER
    HWND hWnd = CreateWindowExW(WS_EX_COMPOSITED, szWindowClass, szTitle, WS_POPUP | WS_MINIMIZEBOX | WS_SYSMENU, //WS_OVERLAPPEDWINDOW,
       CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);

    EnableRoundedCorners(hWnd); // Enable rounded corners for the popup window
#else
    HWND hWnd = CreateWindowExW(WS_EX_COMPOSITED, szWindowClass, szTitle, WS_OVERLAPPEDWINDOW, //WS_POPUP | WS_MINIMIZEBOX | WS_SYSMENU,
       CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);
#endif

    if (!hWnd)
    {
       return hWnd;
    }

#if !defined(HAS_MENU)
    SetMenu(hWnd, nullptr);
#endif

#if !defined(HAS_TITLE_BAR)
    LONG_PTR style = GetWindowLongPtr(hWnd, GWL_STYLE); // Retrieve the current window style
    style &= ~WS_CAPTION; // Remove the title bar (caption) from the style
    SetWindowLongPtr(hWnd, GWL_STYLE, style); // Set the new style
    //SetWindowPos(hWnd, NULL, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED); // Apply the style changes
#endif

    ShowWindow(hWnd, nCmdShow);
    UpdateWindow(hWnd);

    // Place and make topmost
//    SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
    int titleBarHeight = GetSystemMetrics(SM_CYCAPTION);
    int menuBarHeight  = GetSystemMetrics(SM_CYMENU);
    int borderHeight   = GetSystemMetrics(SM_CYSIZEFRAME);
    g_fDpiScaleFactor  = GetDpiScaleFactor(hWnd);
    g_iScreenWidth     = GetSystemMetrics(SM_CXSCREEN);
    g_iScreenHeight    = GetSystemMetrics(SM_CYSCREEN);

    g_iMyWidth         = g_iScreenWidth - g_iScreenWidth / HORZ_PORTION;
    g_iEffectiveWidth  = g_iMyWidth;
    g_iMyHeight        = g_iScreenHeight * ZOOM / VERT_PORTION;
#if !defined(NO_RESIZE_NO_BORDER)
    g_iMyHeight += borderHeight*4;
#endif
#ifdef HAS_MENU
    height += menuBarHeight * g_fDpiScaleFactor;
#endif
#ifdef HAS_TITLE_BAR
    height += (int)(titleBarHeight * g_fDpiScaleFactor);
#endif
    g_iMyWindowX  = (g_iScreenWidth - g_iMyWidth) / 2;
    g_iMyWindowY  = g_iScreenHeight - g_iMyHeight;
#if !defined(POSITION_OVER_TASKBAR)
    g_iMyWindowY -= GetTaskbarHeight();
#endif
#if !defined(NO_RESIZE_NO_BORDER)
    g_iMyWindowY += borderHeight*2;
#endif
    SetWindowPos(hWnd, HWND_TOPMOST, g_iMyWindowX, g_iMyWindowY, g_iMyWidth, g_iMyHeight, SWP_SHOWWINDOW);

#ifdef INVALIDATE_ON_HOOK
    // Start monitoring desktop for changes
    StartDesktopMonitor();
#endif

    return hWnd;
}

#ifdef USE_ANIMATION
// Advance animation: compute current width from time, derive height/position, call SetWindowPos.
// Does NOT re-evaluate the target — safe to call from the WM_PAINT loop.
// bNoRedraw: true when called from WM_PAINT (we paint immediately after, so skip DWM repaint)
//            false for steady-state positioning from the timer
static void AdvanceAnimation(HWND hWnd, bool bNoRedraw = false)
{
    // Compute animation progress: t in [0..1] using high-resolution timer
    LARGE_INTEGER now, freq;
    QueryPerformanceCounter(&now);
    QueryPerformanceFrequency(&freq);
    float elapsedMs = (float)(now.QuadPart - g_llAnimStart) * 1000.0f / freq.QuadPart;
    float t = (ANIM_DURATION_MS > 0) ? elapsedMs / ANIM_DURATION_MS : 1.0f;
    if (t > 1.0f) t = 1.0f;

    // Linear — no easing

    // Width is the sole driver — calculated from time
    g_iAnimWidth = g_iAnimFromW + (int)((g_iAnimToW - g_iAnimFromW) * t);

    // Height derived from width, preserving aspect ratio of the non-zero end
    int refW = (g_iAnimToW > 0) ? g_iAnimToW : g_iAnimFromW;
    int refH = (g_iAnimToW > 0) ? g_iAnimToH : g_iAnimFromH;
    g_iAnimHeight = (refW > 0) ? refH * g_iAnimWidth / refW : 0;

    UINT flags = SWP_SHOWWINDOW | SWP_NOZORDER | SWP_NOACTIVATE;
    if (bNoRedraw) flags |= SWP_NOREDRAW; // Suppress DWM repaint — we paint right after

    if (g_iAnimWidth < 1 || g_iAnimHeight < 1)
    {
        // Fully hidden — park off-screen
        SetWindowPos(hWnd, HWND_TOPMOST, g_iMyWindowX, 4 * g_iMyWindowY,
                     0, 0, flags | SWP_NOSIZE);
    }
    else
    {
        // Center the animated rect around the eventual full-size center
        int centerX = g_iMyWindowX + g_iEffectiveWidth / 2;
        int centerY = g_iMyWindowY + g_iMyHeight / 2;
        int animX   = centerX - g_iAnimWidth  / 2;
        int animY   = centerY - g_iAnimHeight / 2;

        SetWindowPos(hWnd, HWND_TOPMOST, animX, animY,
                     g_iAnimWidth, g_iAnimHeight, flags);
    }
}

// Set animation target (if changed) and advance one frame.
// Called from Show/HideMyWindow — the only path that evaluates g_iEffectiveWidth.
void MySetWindowPos(HWND hWnd, bool bShow)
{
    int targetW = bShow ? g_iEffectiveWidth : 0;

    // Target changed — start new animation from current width
    if (targetW != g_iAnimToW)
    {
        g_iAnimFromW  = g_iAnimWidth;
        g_iAnimFromH  = g_iAnimHeight;
        g_iAnimToW    = targetW;
        g_iAnimToH    = bShow ? g_iMyHeight : 0;
        LARGE_INTEGER now;
        QueryPerformanceCounter(&now);
        g_llAnimStart = now.QuadPart;
        InvalidateRect(hWnd, NULL, FALSE); // Kick off the WM_PAINT animation loop
    }

    // Only advance from the timer when animation is done (steady-state positioning).
    // During animation, WM_PAINT drives both AdvanceAnimation and painting in lockstep.
    if (g_iAnimWidth == g_iAnimToW && g_iAnimHeight == g_iAnimToH)
        AdvanceAnimation(hWnd);
}
#else
void MySetWindowPos(HWND hWnd, bool bShow)
{
    if (bShow)
        SetWindowPos(hWnd, HWND_TOPMOST, g_iMyWindowX, g_iMyWindowY,
                     g_iEffectiveWidth, g_iMyHeight,
                     SWP_SHOWWINDOW | SWP_NOZORDER | SWP_NOACTIVATE);
    else
        SetWindowPos(hWnd, HWND_TOPMOST, g_iMyWindowX, 4 * g_iMyWindowY,
                     0, 0, SWP_SHOWWINDOW | SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOSIZE);
}
#endif

void HideMyWindow(HWND hWnd) { MySetWindowPos(hWnd, false); g_ptCaret = {}; }
void ShowMyWindow(HWND hWnd) { MySetWindowPos(hWnd, true ); }

// ---------------------------------------------------------------------------
// Detect caret + container, update globals, optionally show/hide window.
// Called from WM_APP_DETECT_CARET (event-driven) and CARET_TIMER (periodic).
// bShowHide=false: detect only (caller will show/hide after painting).
// ---------------------------------------------------------------------------
static void DetectCaretAndUpdateWindow(HWND hWnd, bool bShowHide = true)
{
    if (!(g_hForegroundWindow = GetForegroundWindow()))
        return; // can be NULL during focus transitions

    g_fDpiScaleFactor = GetDpiScaleFactor(g_hForegroundWindow);
    g_bAllowOptimizations = !KEY_TOGGLED(VK_SCROLL);

    // Detect foreground window change
    static HWND hPrevForeground = nullptr;
    if (g_hForegroundWindow != hPrevForeground)
    {
        g_bCaretMightHaveMoved = true;
        hPrevForeground = g_hForegroundWindow;
    }

    // Only run detection when something might have changed
    if (g_bCaretMightHaveMoved || !g_bAllowOptimizations)
    {
        POINT caret = {};
        HWND  hwndCaretOwner = nullptr;
        bool bCaretFromAccessibility = false;
        int methods = GetCaretMethodsForWindow(g_hForegroundWindow);

        OutputDebugFormatA("======== APP: %s ========\n", g_szAppExeName);

        // --- Caret detection chain ---
        if (methods & CARET_GTHI)
        {
            caret = GetCaretPosition(&hwndCaretOwner);
            if (caret.y != 0)
                OutputDebugFormatA("  Caret [GuiThreadInfo]: (%d,%d)\n", caret.x, caret.y);
            else
                OutputDebugFormatA("  Caret [GuiThreadInfo]: not found\n");
        }
        else
            OutputDebugFormatA("  Caret [GuiThreadInfo]: skipped\n");

        if (methods & CARET_IACC)
        {
            if (caret.y == 0)
            {
                HWND hwndAccCaret = nullptr;
                caret = GetCaretPositionFromAccessibility(&hwndAccCaret);
                bCaretFromAccessibility = true;
                if (!hwndCaretOwner && hwndAccCaret)
                    hwndCaretOwner = hwndAccCaret;
                if (caret.y != 0)
                    OutputDebugFormatA("  Caret [IAccessible]  : (%d,%d)\n", caret.x, caret.y);
                else
                    OutputDebugFormatA("  Caret [IAccessible]  : not found\n");
            }
            else
                OutputDebugFormatA("  Caret [IAccessible]  : -\n");
        }
        else
            OutputDebugFormatA("  Caret [IAccessible]  : skipped\n");

        if (methods & CARET_UIA)
        {
            if (caret.y == 0)
            {
                caret = GetCaretPositionFromUIA();
                if (caret.y != 0)
                    OutputDebugFormatA("  Caret [UIA]          : (%d,%d)\n", caret.x, caret.y);
                else
                    OutputDebugFormatA("  Caret [UIA]          : not found\n");
            }
            else
                OutputDebugFormatA("  Caret [UIA]          : -\n");
        }
        else
            OutputDebugFormatA("  Caret [UIA]          : skipped\n");

        g_bCaretFromAccessibility = bCaretFromAccessibility;

        // --- Container detection (try all enabled, pick narrowest) ---
        if (caret.y != 0)
        {
            HWND hRoot = GetAncestor(g_hForegroundWindow, GA_ROOT);
            if (!hRoot) hRoot = g_hForegroundWindow;

            int containerMethods = GetContainerMethodsForWindow(g_hForegroundWindow);
            RECT rcBest = {};
            int  bestWidth = INT_MAX;
            const char* bestName = "fallback";

            // Method 1: HOOK
            if (containerMethods & CONTAINER_HOOK)
            {
                if (g_hFocusedChildWnd)
                {
                    RECT rc;
                    GetWindowRect(g_hFocusedChildWnd, &rc);
                    int w = rc.right - rc.left;
                    char cls[64] = {};
                    GetClassNameA(g_hFocusedChildWnd, cls, sizeof(cls));
                    OutputDebugFormatA("  Container [Hook]     : '%s' (%d,%d,%d,%d) %dx%d\n",
                                       cls, rc.left, rc.top, rc.right, rc.bottom,
                                       w, rc.bottom - rc.top);
                    if (w < bestWidth) { rcBest = rc; bestWidth = w; bestName = "Hook"; }
                }
                else
                    OutputDebugFormatA("  Container [Hook]     : no focused child\n");
            }
            else
                OutputDebugFormatA("  Container [Hook]     : skipped\n");

            // Method 2: ENUM
            if (containerMethods & CONTAINER_ENUM)
            {
                int childCount = 0;
                HWND hChild = FindSmallestChildContainingXY(hRoot, caret.x, caret.y, &childCount);
                if (hChild)
                {
                    RECT rc;
                    GetWindowRect(hChild, &rc);
                    int w = rc.right - rc.left;
                    char cls[64] = {};
                    GetClassNameA(hChild, cls, sizeof(cls));
                    OutputDebugFormatA("  Container [Enum]     : '%s' (%d,%d,%d,%d) %dx%d\n",
                                       cls, rc.left, rc.top, rc.right, rc.bottom,
                                       w, rc.bottom - rc.top);
                    if (w < bestWidth) { rcBest = rc; bestWidth = w; bestName = "Enum"; }
                }
                else
                    OutputDebugFormatA("  Container [Enum]     : no match (%d children)\n", childCount);
            }
            else
                OutputDebugFormatA("  Container [Enum]     : skipped\n");

            // Method 3: UIA ElementFromPoint
            if (containerMethods & CONTAINER_UIA)
            {
                RECT rc = GetContainerRectFromUIA(caret.x, caret.y);
                if (rc.right != 0)
                {
                    int w = rc.right - rc.left;
                    OutputDebugFormatA("  Container [UIA]      : (%d,%d,%d,%d) %dx%d\n",
                                       rc.left, rc.top, rc.right, rc.bottom,
                                       w, rc.bottom - rc.top);
                    if (w < bestWidth) { rcBest = rc; bestWidth = w; bestName = "UIA"; }
                }
                else
                    OutputDebugFormatA("  Container [UIA]      : not found\n");
            }
            else
                OutputDebugFormatA("  Container [UIA]      : skipped\n");

            // Fallback: use foreground window rect
            if (bestWidth == INT_MAX)
            {
                GetWindowRect(hRoot, &rcBest);
                OutputDebugFormatA("  Container [Fallback] : (%d,%d,%d,%d) %dx%d\n",
                                   rcBest.left, rcBest.top, rcBest.right, rcBest.bottom,
                                   rcBest.right - rcBest.left, rcBest.bottom - rcBest.top);
            }
            else
                OutputDebugFormatA("  => Winner: %s (%d wide)\n", bestName, bestWidth);

            g_rcContainer = rcBest;

            // Shrink overlay width when container is narrower than the default capture area
            int containerWidth = rcBest.right - rcBest.left;
            int newEffective = min((int)g_iMyWidth, containerWidth * ZOOM);
            if (newEffective != g_iEffectiveWidth)
            {
                g_iEffectiveWidth = newEffective;
                g_iMyWindowX = (g_iScreenWidth - g_iEffectiveWidth) / 2;
            }
        }

        if ((caret.y < 0) || (caret.y > g_iScreenHeight))
            caret.x = caret.y = 0;

        if (g_hForegroundWindow != hWnd)
            g_ptCaret = caret;

        g_bCaretMightHaveMoved = false;
    }

    if (bShowHide)
    {
        static bool bWindowAtTheBottom = false;

        if (((g_ptCaret.y != 0) && (g_bTemporarilyHideMyWindow == false)) ||
            ((g_hForegroundWindow == hWnd) && (g_bTemporarilyHideMyWindow == false)))
            ShowMyWindow(hWnd);
        else
            HideMyWindow(hWnd);
    }
}

//
//  FUNCTION: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PURPOSE: Processes messages for the main window.
//
//  WM_COMMAND  - process the application menu
//  WM_PAINT    - Paint the main window
//  WM_DESTROY  - post a quit message and return
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    bool bDestroy = false;

    // Handling the messages
    switch (message)
    {
#ifdef USE_ANIMATION
    case WM_ERASEBKGND:
        return 1; // Suppress background erase to prevent flicker during animation
#endif
    case WM_SYSCOMMAND:
        if (wParam == SC_MINIMIZE) // Blocking the "minimize" functionality...
        {
//             if( g_bWindowAtTheBottom )
//                 ShowMyWindow(hWnd); // Toggle Show/Hide instead of restore...
//             else
//                 HideMyWindow(hWnd); // Toggle Show/Hide instead of minimization...

            //g_bTemporarilyHideMyWindow = !g_bTemporarilyHideMyWindow;
            return (LPARAM)0; // Not calling DefWindowProc() here
        }
        break;
    case WM_COMMAND:
        {
            int wmId = LOWORD(wParam);
            // Parse the menu selections:
            switch (wmId)
            {
            case IDM_ABOUT:
                DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
                break;
            case IDM_EXIT:
                DestroyWindow(hWnd);
                break;
            }
        }
        break;
    case WM_TIMER:
#ifdef INVALIDATE_ON_TIMER
        if (wParam == INVALIDATE_TIMER_ID) // Ensure it's our specific timer
        {
#ifdef USE_ANIMATION
            // During animation, WM_PAINT drives itself — skip timer invalidation
            if (g_iAnimWidth == g_iAnimToW && g_iAnimHeight == g_iAnimToH)
                InvalidateRect(hWnd, NULL, FALSE);
#else
            InvalidateRect(hWnd, NULL, FALSE);
#endif
        }
#endif
        if (wParam == CARET_TIMER_ID)
        {
            if (g_ptCaret.y != 0) // Only run after event-driven detection has set a valid caret
                DetectCaretAndUpdateWindow(hWnd);
        }
        if (wParam == SETTLE_TIMER_ID)
        {
            KillTimer(hWnd, SETTLE_TIMER_ID);
            PostMessage(hWnd, WM_APP_DETECT_CARET, DETECT_REASON_SETTLE, 0);
        }
        break;
    case WM_APP_DETECT_CARET:
        {
            static POINT ptPrevSettle = {};
            static HWND  hSettleFG = nullptr;
            static int   iSettleCount = 0;
            static LARGE_INTEGER llSettlePrev = {};
            static DWORD dwLastEventTick = 0;  // for no-freeze window (rapid events skip freeze)

            // --- Process event flags (key/mouse only, not settle ticks) ---
            if (wParam == DETECT_REASON_KEY)
            {
                DWORD vk = (DWORD)lParam;
                UINT  mappedChar    = MapVirtualKey(vk, MAPVK_VK_TO_CHAR);
                bool  bPrintableKey = mappedChar && std::isprint(mappedChar);

                g_bWaitForInputAfterToggle = false;

                if (vk == VK_ESCAPE)
                    g_bTemporarilyHideMyWindow = !g_bTemporarilyHideMyWindow;
                else if (bPrintableKey || vk == VK_DELETE || vk == VK_BACK)
                    g_bTemporarilyHideMyWindow = false;
            }
            else if (wParam == DETECT_REASON_MOUSE)
            {
                g_bWaitForInputAfterToggle = false;
            }

            bool bSettling = (wParam == DETECT_REASON_SETTLE);

            if (!bSettling)
            {
                if (g_bSettling)
                {
                    // New event while settling: abort current settle
                    KillTimer(hWnd, SETTLE_TIMER_ID);
                    g_bSettling = false;
                    OutputDebugFormatA("  Settle interrupted — accepting (%d,%d)\n",
                                       g_ptCaret.x, g_ptCaret.y);
                    // Only paint+show if we have a valid caret; otherwise leave window as-is
                    if (g_ptCaret.y != 0)
                    {
                        InvalidateRect(hWnd, NULL, FALSE);
                        UpdateWindow(hWnd);
                        if (!g_bTemporarilyHideMyWindow)
                            ShowMyWindow(hWnd);
                        else
                            HideMyWindow(hWnd);
                    }
                }
                // Start (or restart) settle loop — freeze grabs only if no recent event
                DWORD dwNow = GetTickCount();
                bool bOkToFreeze = (dwNow - dwLastEventTick > 250);
                dwLastEventTick = dwNow;
                g_bSettling = bOkToFreeze ? true : false;
                hSettleFG = g_hForegroundWindow;
                ptPrevSettle = g_ptCaret;
                iSettleCount = 0;
                QueryPerformanceCounter(&llSettlePrev);
                SetTimer(hWnd, SETTLE_TIMER_ID, SETTLE_TIMER_MS, NULL);
                if (!bOkToFreeze)
                {
                    // Rapid event: grabs are live, paint immediately
                    InvalidateRect(hWnd, NULL, FALSE);
                    UpdateWindow(hWnd);
                }
            }
            else
            {
                // Settle tick: detect + paint + show
                g_bCaretMightHaveMoved = true;
                LARGE_INTEGER llDetectStart, llDetectEnd, llDetectFreq;
                QueryPerformanceCounter(&llDetectStart);
                DetectCaretAndUpdateWindow(hWnd, false);
                QueryPerformanceCounter(&llDetectEnd);
                QueryPerformanceFrequency(&llDetectFreq);
                float detectMs = (float)(llDetectEnd.QuadPart - llDetectStart.QuadPart) * 1000.0f / llDetectFreq.QuadPart;

                LARGE_INTEGER now, freq;
                QueryPerformanceCounter(&now);
                QueryPerformanceFrequency(&freq);
                float deltaMs = (float)(now.QuadPart - llSettlePrev.QuadPart) * 1000.0f / freq.QuadPart;
                llSettlePrev = now;

                // Keep going if caret moved, up to 50 iterations
                if ((g_ptCaret.x != ptPrevSettle.x || g_ptCaret.y != ptPrevSettle.y)
                    && ++iSettleCount < 50)
                {
                    OutputDebugFormatA("  Settle #%d  delta=%.1fms  detect=%.2fms: (%d,%d) -> (%d,%d)\n",
                                       iSettleCount, deltaMs, detectMs,
                                       ptPrevSettle.x, ptPrevSettle.y,
                                       g_ptCaret.x, g_ptCaret.y);
                    ptPrevSettle = g_ptCaret;
                    SetTimer(hWnd, SETTLE_TIMER_ID, SETTLE_TIMER_MS, NULL);
                }
                else
                {
                    // Settled: paint + show only with final stable coordinates
                    g_bSettling = false;
                    OutputDebugFormatA("  Settled after %d iteration(s)  delta=%.1fms  detect=%.2fms at (%d,%d)\n",
                                       iSettleCount, deltaMs, detectMs,
                                       g_ptCaret.x, g_ptCaret.y);
                    InvalidateRect(hWnd, NULL, FALSE);
                    UpdateWindow(hWnd);
                    if (((g_ptCaret.y != 0) && !g_bTemporarilyHideMyWindow) ||
                        ((g_hForegroundWindow == hWnd) && !g_bTemporarilyHideMyWindow))
                        ShowMyWindow(hWnd);
                    else
                        HideMyWindow(hWnd);
                }
            }
        }
        break;
    case WM_MOVING:
    case WM_SIZING:
        ValidateRect(hWnd, NULL);
        break;
    case WM_PAINT:
        {
#ifdef USE_ANIMATION
            // Advance animation BEFORE painting so content matches the new window size.
            // (On zoom-in, SetWindowPos grows the window; painting after ensures no blank area.)
            bool bAnimating = (g_iAnimWidth != g_iAnimToW || g_iAnimHeight != g_iAnimToH);
            if (bAnimating)
                AdvanceAnimation(hWnd, true); // true = SWP_NOREDRAW, we paint right after
#endif

            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);

            if( g_rcContainer.right > g_rcContainer.left )
            {
                int width = g_iEffectiveWidth, height = g_iScreenHeight / VERT_PORTION; // Define the region to capture
                int fullW = width, fullH = height * ZOOM;

#ifdef USE_CACHING_DC
                // --- Cached frame buffer (intermediate between desktop and window) ---
                static HDC     hCacheDC  = NULL;
                static HBITMAP hCacheBmp = NULL;
                static int     cacheW = 0, cacheH = 0;
                static bool    bCacheValid = false;

                if (!hCacheDC || cacheW != fullW || cacheH != fullH)
                {
                    if (hCacheDC)  { DeleteObject(hCacheBmp); DeleteDC(hCacheDC); }
                    hCacheDC  = CreateCompatibleDC(hdc);
                    hCacheBmp = CreateCompatibleBitmap(hdc, fullW, fullH);
                    SelectObject(hCacheDC, hCacheBmp);
                    cacheW = fullW; cacheH = fullH;
                    bCacheValid = false;
                }

                // --- During settling: skip grab, display from primary cache ---
                if (g_bSettling)
                    goto skip_grab;
#ifdef USE_ANIMATION
                // Invalidate cache when zoom-in starts (fresh grab needed).
                // Zoom-out keeps the stale cache (current desktop has no useful caret data).
                static bool bWasAnimating = false;
                if (bAnimating && !bWasAnimating && g_iAnimToW > 0)
                    bCacheValid = false;
                bWasAnimating = bAnimating;

                if (!bAnimating || !bCacheValid)
#endif // USE_ANIMATION
#endif // USE_CACHING_DC
                {

                POINT ptCaret = g_ptCaret; // Calling GetCaretPositionFromAccessibility() here is too slow...

                RECT focusedChildRect = g_rcContainer;

                int iSrcX = 0;
                int iSrcY = 0;

                if( ptCaret.y != 0 )
                {
                    #ifdef USE_FONT_HEIGHT
                        int iFontHeight = GetFontHeight(g_hForegroundWindow);
                    #endif
                    //iCaretY -= iFontHeight * 2;

                    int iSrcX_match  = ptCaret.x / ZOOM;             // Cursor always matches real cursor
                    int iSrcX_left   = focusedChildRect.left;                      // Cursor starts from the left edge
                    int iSrcX_right  = focusedChildRect.right - width / ZOOM; // Cursor starts from the right edge
                    //int iSrcX_match  = width / ZOOM / 2;             // Cursor center matches screen center
                    int iSrcX_center = ptCaret.x - width / ZOOM / 2; // Cursor always in the center of the screen

                    int iSrcX_left_edge  = ptCaret.x - 10; // Cursor always at the left edge of our window
                    int iSrcX_right_edge = ptCaret.x - width / ZOOM + 10; // Cursor always at the right edge of our window

                    static int iSrcX_center_prev = iSrcX_center;

                    // Implement the left-center-right logic
                    if( iSrcX_left >= iSrcX_right ) // This happens when the window width is small enough to fit as a whole
                        iSrcX = (iSrcX_left + iSrcX_right) / 2; //iSrcX = iSrcX_center;
                    else
                    {
                        //iSrcX = max( iSrcX_left, min( iSrcX_right, iSrcX_center) );
                        //iSrcX = max( iSrcX_left, min( iSrcX_right, g_iSrcX) );

                        if( KEY_DOWN(VK_RSHIFT) )
                            { int i = 5; }

//                         int iOffs = iSrcX_center - iSrcX_center_prev;
//                         iSrcX = g_iSrcX + iOffs;

                        // Note: iSrcX_right_edge < iSrcX_left_edge

                        {
                            int iLeftCaret   =  iSrcX_left;
                            int iCenterCaret = (iSrcX_left + iSrcX_right) / 2 + width / ZOOM / 2;
                            int iRightCaret  =               iSrcX_right      + width / ZOOM;
                            int iFirstThird  = (iLeftCaret*2 + iRightCaret*1) / 3;
                            int iSecondThird = (iLeftCaret*1 + iRightCaret*2) / 3;

//                                 if( ptCaret.x < iFirstThird )
//                                     g_iSrcX = iSrcX_left;//iSrcX_right_edge;
//                                 else
//                                 if( ptCaret.x > iSecondThird )
//                                     g_iSrcX = iSrcX_right; //iSrcX_left_edge;
//     //                             else
//     //                                 g_iSrcX = (iSrcX_left + iSrcX_right) / 2;

                                if( ptCaret.x < iCenterCaret )
                                    g_iSrcX = iSrcX_left;//iSrcX_right_edge;
                                else
    //                            if( ptCaret.x > iSecondThird )
                                    g_iSrcX = iSrcX_right; //iSrcX_left_edge;
    //                             else
    //                                 g_iSrcX = (iSrcX_left + iSrcX_right) / 2;

                        }
//                        else

                        iSrcX = g_iSrcX;

                        if( iSrcX < iSrcX_right_edge )
                            iSrcX = iSrcX_right_edge;
                        else
                        if( iSrcX > iSrcX_left_edge )
                            iSrcX = iSrcX_left_edge;
                        else
                            iSrcX = g_iSrcX;


//                         iSrcX = g_iSrcX;
//
//                         if( iSrcX < iSrcX_left )
//                             iSrcX = iSrcX_left;
//                         else
//                         if( iSrcX > iSrcX_right )
//                             iSrcX = iSrcX_right;
//                         else
//                             iSrcX = g_iSrcX;

//                             {
//
//
//     //                             int iOffCenter_min = iSrcX_center - iSrcX_right;
//     //                             int iOffCenter_max = iSrcX_center - iSrcX_left;
//     //
//     //                             iOffs = max( iOffCenter_min, min( iOffCenter_max, iOffCenter) );
//     //
//     //                             iSrcX = g_iSrcX + iOffCenter;
//                             }
                    }

                    iSrcX_center_prev = iSrcX_center;

                    //iSrcX = iSrcX_left;
                    //iSrcX = iSrcX_right;
                    //iSrcX = iSrcX_center;
                    //iSrcX = iSrcX_match;
                    //iSrcX = iSrcX_right_edge;
                    iSrcY = ptCaret.y - height/ 2;
                }

                if( g_bFreezeSrcXandSrcY || g_iTakeCaretSnapshot )
                {
                    if( g_iTakeCaretSnapshot )
                    {
                        g_ptCaretSnapshot = g_ptCaret;
                        g_iTakeCaretSnapshot--;
                    }
                    else
                    if( g_ptCaret.x != g_ptCaretSnapshot.x || g_ptCaret.y != g_ptCaretSnapshot.y )
                        g_bFreezeSrcXandSrcY = false;
                }
                else
                    g_iSrcX = iSrcX;

                g_iSrcY = iSrcY; // Y is always updated...

                #ifdef RENDER_CURSOR
                    RenderScaledCursor(hMemDC, hWnd, width, height); // Render to hMemDC
                #endif

                // Grab desktop and blit to window
                {
                    HDC hScreenDC = GetDC(NULL);
#ifdef USE_CACHING_DC
                    StretchBlt(hCacheDC, 0, 0, fullW, fullH, hScreenDC, g_iSrcX, g_iSrcY, width/ZOOM, height, SRCCOPY);
                    bCacheValid = true;
#else
                    StretchBlt(hdc, 0, 0, fullW, fullH, hScreenDC, g_iSrcX, g_iSrcY, width/ZOOM, height, SRCCOPY);
#endif
                    ReleaseDC(NULL, hScreenDC);
                }

                } // end of coordinate update + grab block
#ifdef USE_CACHING_DC
                skip_grab:
#endif

#ifdef USE_CACHING_DC
                // --- Blit from cache to window ---
                if (bCacheValid)
                {
#ifdef USE_ANIMATION
                    int dstW = (g_iAnimWidth  > 0) ? g_iAnimWidth  : fullW;
                    int dstH = (g_iAnimHeight > 0) ? g_iAnimHeight : fullH;
                    if (dstW != fullW || dstH != fullH)
                    {
                        SetStretchBltMode(hdc, HALFTONE);
                        StretchBlt(hdc, 0, 0, dstW, dstH, hCacheDC, 0, 0, fullW, fullH, SRCCOPY);
                    }
                    else
                        BitBlt(hdc, 0, 0, fullW, fullH, hCacheDC, 0, 0, SRCCOPY);
#else
                    BitBlt(hdc, 0, 0, fullW, fullH, hCacheDC, 0, 0, SRCCOPY);
#endif
                }
#endif // USE_CACHING_DC

                #ifdef RENDER_CURSOR
                    // RenderScaledCursor(hdc, hWnd, width, height); // Render directly to hdc - this also works
                #endif
            }

            EndPaint(hWnd, &ps);

#ifdef USE_ANIMATION
            // Request next animation frame (animation was already advanced before painting)
            if (g_iAnimWidth != g_iAnimToW || g_iAnimHeight != g_iAnimToH)
                InvalidateRect(hWnd, NULL, FALSE);
#endif
        }
        break;
//    case WM_LBUTTONDOWN: // Can't do anything on button down - because the input focus is on our window. Improve later.
    case WM_LBUTTONUP:
        {
            g_bTemporarilyHideMyWindow = false;

            if( g_rcContainer.right > g_rcContainer.left )
            {
                POINT ptCurrent; GetCursorPos(&ptCurrent);

                POINT ptClient = ptCurrent; ScreenToClient(hWnd, &ptClient);

                POINT ptTarget = { ptClient.x / ZOOM + g_iSrcX, ptClient.y / ZOOM + g_iSrcY };

                POINT ptTargetAbs  = { (ptTarget.x  * 65536 + g_iScreenWidth/2) / g_iScreenWidth, (ptTarget.y  * 65536 + g_iScreenHeight/2) / g_iScreenHeight };
                POINT ptCurrentAbs = { (ptCurrent.x * 65536 + g_iScreenWidth/2) / g_iScreenWidth, (ptCurrent.y * 65536 + g_iScreenHeight/2) / g_iScreenHeight };

                g_iTakeCaretSnapshot = 10;// At least 4 is needed here... O.o
                g_bFreezeSrcXandSrcY = true;

                mouse_event(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, ptTargetAbs.x, ptTargetAbs.y, 0, 0);
                mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                Sleep(1); // just in case...
                mouse_event(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, ptCurrentAbs.x, ptCurrentAbs.y, 0, 0);
            }
        }
        break;
    case WM_LBUTTONDOWN     :
    //case WM_LBUTTONUP     :
    case WM_LBUTTONDBLCLK   :
    case WM_RBUTTONDOWN     :
    case WM_RBUTTONUP       :
    case WM_RBUTTONDBLCLK   :
    case WM_MBUTTONDOWN     :
    case WM_MBUTTONUP       :
    case WM_MBUTTONDBLCLK   :
        g_bTemporarilyHideMyWindow = false;
        break;
    case WM_DESTROY:
        bDestroy = true;
        break;
    }

    // Execute the default handler
    lParam = DefWindowProc(hWnd, message, wParam, lParam);

    if (bDestroy)
    {
        #ifdef INVALIDATE_ON_TIMER
            KillTimer(hWnd, INVALIDATE_TIMER_ID); // Kill the timer
        #endif

        #ifdef INVALIDATE_ON_HOOK
            StopDesktopMonitor(); // Stop monitoring desktop for changes
        #endif

        PostQuitMessage(0);
    }

    return lParam;
}

// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

void CALLBACK WinEventProc_ForFocusedClientWnd(HWINEVENTHOOK hWinEventHook, DWORD event, HWND hwnd,
                                              LONG idObject, LONG idChild, DWORD dwEventThread, DWORD dwmsEventTime)
{
    if( event == EVENT_OBJECT_FOCUS && idObject == OBJID_CLIENT )
    {
        g_bCaretMightHaveMoved = true;

        // Don't accept the foreground window itself as a focused child -
        // the paint handler already falls back to it when child is null.
        if (hwnd == g_hForegroundWindow)
            return;

        HWND oldChild = g_hFocusedChildWnd;

        if( g_hFocusedChildWnd != NULL)
        {
            // Check if this hwnd is a parent of g_hFocusedChildWnd
            static HWND hwndPrevMatchedParent = nullptr;
            static HWND hwndPrevMatchedParent_itsChild = nullptr;

            if( hwnd == hwndPrevMatchedParent )
            {
                g_hFocusedChildWnd = hwndPrevMatchedParent_itsChild;
            }
            else
            {
                BOOL bIsParent = FALSE;
                HWND hwndParent = g_hFocusedChildWnd;

                while( hwndParent = GetParent(hwndParent) )
                if( hwnd == hwndParent )
                {
                    bIsParent = TRUE;
                    hwndPrevMatchedParent = hwndParent;
                    hwndPrevMatchedParent_itsChild = g_hFocusedChildWnd;
                    break;
                }

                if( bIsParent == FALSE )
                    g_hFocusedChildWnd = hwnd;
            }
        }
        else
        {
            g_hFocusedChildWnd = hwnd;
        }

        ///////////////////////////////////////////////////////////////
        static HWND hPrevFocusedChildWnd = nullptr;
        if (g_hFocusedChildWnd != hPrevFocusedChildWnd)
        {
            g_iNewFocusedChild   = 10;
            hPrevFocusedChildWnd = g_hFocusedChildWnd;
        }
    }
}

LRESULT CALLBACK LowLevelKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam)
{
    if (nCode == HC_ACTION && (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN))
    {
        KBDLLHOOKSTRUCT *pKB = (KBDLLHOOKSTRUCT *)lParam;
        PostMessage(g_myWindowHandle, WM_APP_DETECT_CARET,
                    DETECT_REASON_KEY, (LPARAM)pKB->vkCode);
    }
    return CallNextHookEx(NULL, nCode, wParam, lParam);
}

LRESULT CALLBACK LowLevelMouseProc(int nCode, WPARAM wParam, LPARAM lParam)
{
    if (nCode == HC_ACTION &&
        (wParam == WM_LBUTTONDOWN || wParam == WM_RBUTTONDOWN || wParam == WM_MBUTTONDOWN))
    {
        PostMessage(g_myWindowHandle, WM_APP_DETECT_CARET,
                    DETECT_REASON_MOUSE, 0);
    }
    return CallNextHookEx(NULL, nCode, wParam, lParam);
}