// ComfyTyping.cpp : Defines the entry point for the application.
//

#define WIN_UTILS_CPP

#include "framework.h"
#include "resource.h"
#include "ComfyTyping.h"
#include "WinUtils.h"
#include "IniReader.h"
#include <iostream>  // for std::isprint
#include <objbase.h> // for CoInitializeEx / CoUninitialize
#include <dwmapi.h>  // for DwmSetWindowAttribute, DwmGetWindowAttribute

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
    ReadConfiguration();

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

    if( CARET_TIMER_INTERVAL )
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
//
// Zoom-in animation overview:
//
// The window is placed at its final size immediately (no resizing during animation).
// Each WM_PAINT frame StretchBlt's the cached content at a growing scale, centered
// in the fixed-size window. This avoids the wobble caused by per-frame SetWindowPos
// (interior/exterior size mismatches, DWM border rounding).
//
// Key details:
//   - DWMWCP_DONOTROUND hides the DWM border during animation; restored to
//     DWMWCP_ROUND when complete.
//   - The desktop behind the overlay is captured before the window is shown and
//     painted onto the window surface. This makes the overlay invisible until the
//     zoomed content grows over it. Without this, the window surface would retain
//     stale content from the previous session (we show/hide by moving off-screen,
//     not by destroying the window).
//   - Both dstW and dstH must be computed directly from 't' rather than deriving
//     height from width via integer division. The overlay is very wide and short
//     (~2200 x 106), so height = toH * width / toW produces a long plateau at 0
//     followed by a jump to 1 -- causing visible flicker when combined with a
//     background fill.
//   - The animation is driven by InvalidateRect from WM_PAINT (self-sustaining
//     loop). WS_EX_COMPOSITED must remain active for adequate frame rate; without
//     it WM_PAINT is the lowest-priority message and gets starved.
//   - During animation, each frame must only paint the growing StretchBlt rect
//     (and the border) without touching the margins. Painting the full window
//     (e.g. FillRect black then StretchBlt) causes flicker because DWM can
//     composite the window surface between GDI calls within a single WM_PAINT,
//     briefly showing the intermediate all-black state.
//   - UpdateOverlayWindow's pre-show UpdateWindow fills the cache with
//     g_bFillCacheOnly=true so the desktop is grabbed while the overlay is still
//     off-screen (avoiding self-capture), but the full-size content is not
//     painted to the window surface (which would flash for one frame).
//

// Compute current animation width/height from elapsed time.
static void AdvanceAnimation()
{
    LARGE_INTEGER now, freq;
    QueryPerformanceCounter(&now);
    QueryPerformanceFrequency(&freq);
    float elapsedMs = (float)(now.QuadPart - g_llAnimStart) * 1000.0f / freq.QuadPart;
    float t = (ANIM_DURATION_MS > 0) ? elapsedMs / ANIM_DURATION_MS : 1.0f;
    if (t > 1.0f) t = 1.0f;

    g_iAnimWidth  = g_iAnimFromW + (int)((g_iAnimToW - g_iAnimFromW) * t);
    g_iAnimHeight = g_iAnimFromH + (int)((g_iAnimToH - g_iAnimFromH) * t);
}

void MySetWindowPos(HWND hWnd, bool bShow)
{
    if (bShow)
    {
        int targetW = g_iEffectiveWidth;
        if (g_bAnimateNextShow && targetW != g_iAnimToW)
        {
            g_bAnimateNextShow = false;
            // --- Start zoom-in animation ---

            // Hide DWM rounded corners / border during animation
            DWORD cornerPref = 1; // DWMWCP_DONOTROUND
            DwmSetWindowAttribute(hWnd, 33, &cornerPref, sizeof(cornerPref));

            // Capture the desktop behind the overlay (window is still off-screen)
            {
                static HDC     s_hBgDC  = NULL;
                static HBITMAP s_hBgBmp = NULL;
                HDC hScreenDC = GetDC(NULL);
                if (!s_hBgDC) s_hBgDC = CreateCompatibleDC(hScreenDC);
                if (s_hBgBmp) DeleteObject(s_hBgBmp);
                s_hBgBmp = CreateCompatibleBitmap(hScreenDC, g_iEffectiveWidth, g_iMyHeight);
                SelectObject(s_hBgDC, s_hBgBmp);
                BitBlt(s_hBgDC, 0, 0, g_iEffectiveWidth, g_iMyHeight,
                       hScreenDC, g_iMyWindowX, g_iMyWindowY, SRCCOPY);
                ReleaseDC(NULL, hScreenDC);
                g_hAnimBgDC = s_hBgDC;
            }

            // Place window at final size, then paint the captured background
            // so the overlay blends with the desktop until content grows over it
            SetWindowPos(hWnd, HWND_TOPMOST, g_iMyWindowX, g_iMyWindowY,
                         g_iEffectiveWidth, g_iMyHeight,
                         SWP_SHOWWINDOW | SWP_NOZORDER | SWP_NOACTIVATE);
            {
                HDC hdc = GetDC(hWnd);
                BitBlt(hdc, 0, 0, g_iEffectiveWidth, g_iMyHeight, g_hAnimBgDC, 0, 0, SRCCOPY);
                ReleaseDC(hWnd, hdc);
            }
            ValidateRect(hWnd, NULL);

            // Initialize animation state (zoom from zero to full size)
            g_iAnimFromW  = 0;
            g_iAnimFromH  = 0;
            g_iAnimToW    = targetW;
            g_iAnimToH    = g_iMyHeight;
            g_iAnimWidth  = 0;
            g_iAnimHeight = 0;
            LARGE_INTEGER now;
            QueryPerformanceCounter(&now);
            g_llAnimStart = now.QuadPart;
            InvalidateRect(hWnd, NULL, FALSE); // kick off WM_PAINT animation loop
        }
        else if (g_iAnimWidth == g_iAnimToW && g_iAnimHeight == g_iAnimToH)
        {
            // Steady-state: reposition at final size (no animation)
            g_iAnimToW    = targetW;
            g_iAnimToH    = g_iMyHeight;
            g_iAnimWidth  = targetW;
            g_iAnimHeight = g_iMyHeight;
            SetWindowPos(hWnd, HWND_TOPMOST, g_iMyWindowX, g_iMyWindowY,
                         g_iEffectiveWidth, g_iMyHeight,
                         SWP_SHOWWINDOW | SWP_NOZORDER | SWP_NOACTIVATE);
        }
        // else: animation in progress toward same target -- don't interfere
    }
    else
    {
        // Immediate hide (no zoom-out animation)
        if (!g_bOverlayEnabled)
        {
            // Full disable (ESC): reset animation state so next show animates
            g_iAnimWidth = 0; g_iAnimHeight = 0;
            g_iAnimToW = 0;   g_iAnimToH = 0;
        }
        // else: temp hide (no caret) -- preserve animation state so re-show is instant
        SetWindowPos(hWnd, HWND_TOPMOST, g_iMyWindowX, 4 * g_iMyWindowY,
                     0, 0, SWP_SHOWWINDOW | SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOSIZE);
    }
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
// Query caret and container from the system. Writes to g_ptLastQueriedCaret / g_rcLastQueriedContainer.
// Does NOT touch g_ptCaret, g_rcContainer, or g_iEffectiveWidth.
static void DetectCaretAndContainer(HWND hWnd)
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
        RECT rcBest = g_rcLastQueriedContainer; // keep previous if no caret
        if (caret.y != 0)
        {
            HWND hRoot = GetAncestor(g_hForegroundWindow, GA_ROOT);
            if (!hRoot) hRoot = g_hForegroundWindow;

            int containerMethods = GetContainerMethodsForWindow(g_hForegroundWindow);
            rcBest = {};
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
                    if (w > 0 && w < bestWidth) { rcBest = rc; bestWidth = w; bestName = "Hook"; }
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
                    if (w > 0 && w < bestWidth) { rcBest = rc; bestWidth = w; bestName = "Enum"; }
                }
                else
                    OutputDebugFormatA("  Container [Enum]     : no match (%d children)\n", childCount);
            }
            else
                OutputDebugFormatA("  Container [Enum]     : skipped\n");

            // Method 3: UIA ElementFromPoint (overrides Hook/Enum if valid)
            if (containerMethods & CONTAINER_UIA)
            {
                RECT rc = GetContainerRectFromUIA(caret.x, caret.y);
                if (rc.right != 0)
                {
                    int w = rc.right - rc.left;
                    OutputDebugFormatA("  Container [UIA]      : (%d,%d,%d,%d) %dx%d\n",
                                       rc.left, rc.top, rc.right, rc.bottom,
                                       w, rc.bottom - rc.top);
                    if (w > 0) { rcBest = rc; bestWidth = w; bestName = "UIA"; }
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
        }

        // Write to "queried" globals only
        if ((caret.y < 0) || (caret.y > g_iScreenHeight))
            caret.x = caret.y = 0;

        if (g_hForegroundWindow != hWnd)
            g_ptLastQueriedCaret = caret;

        g_rcLastQueriedContainer = rcBest;

        g_bCaretMightHaveMoved = false;
    }
}

// Recalculate g_iEffectiveWidth from current g_rcContainer
static void UpdateEffectiveWidth()
{
    int containerWidth = g_rcContainer.right - g_rcContainer.left;
    int newEffective = min((int)g_iMyWidth, containerWidth * ZOOM);
    if (newEffective != g_iEffectiveWidth)
    {
        g_iEffectiveWidth = newEffective;
        g_iMyWindowX = (g_iScreenWidth - g_iEffectiveWidth) / 2;
    }
}

// Paint + show/hide the overlay using current g_ptCaret / g_rcContainer
static void UpdateOverlayWindow(HWND hWnd)
{
    UpdateEffectiveWidth();
#ifdef USE_ANIMATION
    // During animation, WM_PAINT drives itself -- don't interfere.
    // When not animating, fill the cache (for the next animation) but skip the
    // blit to the window surface -- otherwise the full-size content would flash
    // for one frame when the window is shown (see g_bFillCacheOnly).
    bool bAnimInProgress = (g_iAnimWidth != g_iAnimToW || g_iAnimHeight != g_iAnimToH);
    if (!bAnimInProgress)
    {
        g_bFillCacheOnly = true;
        InvalidateRect(hWnd, NULL, FALSE);
        UpdateWindow(hWnd);
        g_bFillCacheOnly = false;
    }
#else
    InvalidateRect(hWnd, NULL, FALSE);
    UpdateWindow(hWnd);
#endif
    // Compute whether it's OK to show (caret exists and doesn't overlap overlay)
    g_bOkToShowOverlay = (g_ptCaret.y != 0) &&
                         !(g_ptCaret.y >= g_iMyWindowY && g_ptCaret.y < g_iMyWindowY + g_iMyHeight);

    if (g_bOverlayEnabled && (g_bOkToShowOverlay || g_hForegroundWindow == hWnd))
        ShowMyWindow(hWnd);
    else
        HideMyWindow(hWnd);
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

            //g_bOkToShowOverlay = !g_bOkToShowOverlay;
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
            // During animation, WM_PAINT drives itself via InvalidateRect
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
                PostMessage(hWnd, WM_APP_DETECT_CARET, DETECT_REASON_TIMER, 0);
        }
        if (wParam == SETTLE_TIMER_ID)
        {
            KillTimer(hWnd, SETTLE_TIMER_ID);
            PostMessage(hWnd, WM_APP_DETECT_CARET, DETECT_REASON_SETTLE, 0);
        }
        break;
    case WM_APP_DETECT_CARET:
        {
            static POINT ptPrevQueryCaret = {};
            static RECT  rcPrevQueryContainer = {};
            static HWND  hSettleFG = nullptr;
            static HWND  hPrevSettledFG = nullptr;
            static int   iSettleCount = 0;
            static bool  bSettleFromTyping = false;
            static DWORD dwLastEventTick = 0;  // for no-freeze window (rapid events skip freeze)

            // --- Print separator for new events (not settle ticks) ---
            if (wParam != DETECT_REASON_SETTLE)
                OutputDebugFormatA("\n\n");

            // --- Classify typing events for caret-loss fallback ---
            bool bIsTypingEvent = false;
            if (wParam == DETECT_REASON_KEY)
            {
                DWORD vk = (DWORD)lParam;
                UINT  mappedChar    = MapVirtualKey(vk, MAPVK_VK_TO_CHAR);
                bool  bPrintableKey = mappedChar && std::isprint(mappedChar);
                bIsTypingEvent = bPrintableKey || vk == VK_RETURN || vk == VK_TAB || vk == VK_BACK || vk == VK_DELETE
                                 || vk == VK_LEFT || vk == VK_RIGHT || vk == VK_UP || vk == VK_DOWN
                                /* || vk == VK_HOME */ || vk == VK_END || vk == VK_PRIOR || vk == VK_NEXT || vk == VK_INSERT; // Not VK_HOME - because we want to "refresh" the container size when it is pressed...
            }

            // --- Process event flags (key/mouse only, not settle ticks) ---
            if (wParam == DETECT_REASON_KEY)
            {
                DWORD vk = (DWORD)lParam;
                UINT  mappedChar    = MapVirtualKey(vk, MAPVK_VK_TO_CHAR);
                bool  bPrintableKey = mappedChar && std::isprint(mappedChar);
                bool  bTypingKey    = bPrintableKey || vk == VK_RETURN;

                g_bWaitForInputAfterToggle = false;

                // --- Activation tracking ---
                static int   s_iTypingCount = 0;       // consecutive typing keypresses
                static DWORD s_typingTicks[4] = {};     // circular buffer of timestamps
                static DWORD s_arrowKeys[4] = {};       // recent arrow VK codes
                static DWORD s_arrowTicks[4] = {};      // arrow timestamps
                static int   s_iArrowCount = 0;         // arrow keys recorded (0-4)

                if (vk == VK_ESCAPE)
                {
                    // Hook already blocked ESC & hid overlay. Just reset counters.
                    s_iTypingCount = 0;
                    s_iArrowCount  = 0;
                }
                else if (bTypingKey)
                {
                    s_iArrowCount = 0;
                    s_typingTicks[s_iTypingCount % 4] = GetTickCount();
                    s_iTypingCount++;
                    if (s_iTypingCount >= 4)
                    {
                        // Rolling window: check oldest of last 4 is within 1500ms
                        DWORD oldest = s_typingTicks[s_iTypingCount % 4];
                        if (GetTickCount() - oldest <= 1500 && !g_bOverlayEnabled && g_ptLastQueriedCaret.y != 0)
                        {   g_bOverlayEnabled = true;             // ACTIVATE
                            g_bAnimateNextShow = true;
                        }
                    }
                }
                else if (vk == VK_LEFT || vk == VK_RIGHT)
                {
                    s_iTypingCount = 0;
                    DWORD now = GetTickCount();
                    // Reset if gap from previous arrow > 250ms
                    if (s_iArrowCount > 0 && (now - s_arrowTicks[s_iArrowCount - 1]) > 250)
                        s_iArrowCount = 0;
                    // Append
                    if (s_iArrowCount < 4)
                    {
                        s_arrowKeys[s_iArrowCount]  = vk;
                        s_arrowTicks[s_iArrowCount] = now;
                        s_iArrowCount++;
                    }
                    // Check for alternating pattern of 4
                    if (s_iArrowCount == 4)
                    {
                        bool bAlternating = true;
                        for (int i = 1; i < 4; i++)
                            if (s_arrowKeys[i] == s_arrowKeys[i - 1])
                                bAlternating = false;
                        if (bAlternating && !g_bOverlayEnabled && g_ptLastQueriedCaret.y != 0)
                        {   g_bOverlayEnabled = true;             // ACTIVATE
                            g_bAnimateNextShow = true;
                        }
                        s_iArrowCount = 0;
                    }
                }
                else
                {
                    // Any other key (Delete, Backspace, Tab, modifiers, etc.)
                    s_iTypingCount = 0;
                    s_iArrowCount  = 0;
                }
            }
            else if (wParam == DETECT_REASON_MOUSE)
            {
                // Hit-test: skip detection for title bar / buttons / scrollbars
                int clickX = (short)LOWORD(lParam);
                int clickY = (short)HIWORD(lParam);
                LRESULT ht = SendMessage(g_hForegroundWindow, WM_NCHITTEST, 0, MAKELPARAM(clickX, clickY));
                if (ht == HTCAPTION || ht == HTMINBUTTON || ht == HTMAXBUTTON || ht == HTCLOSE ||
                    ht == HTMENU || ht == HTSYSMENU || ht == HTVSCROLL || ht == HTHSCROLL)
                    break;

                g_bWaitForInputAfterToggle = false;
            }

            // --- Handle periodic timer: skip if settling or recent user input ---
            if (wParam == DETECT_REASON_TIMER)
            {
                DWORD dwNow = GetTickCount();
                if (g_bSettling || (dwNow - dwLastEventTick < 250))
                    break; // already settling or user was active recently
            }

            // --- Query caret and container on every message ---
            g_bCaretMightHaveMoved = true;
            LARGE_INTEGER llDetectStart, llDetectEnd, llDetectFreq;
            QueryPerformanceCounter(&llDetectStart);
            DetectCaretAndContainer(hWnd);
            QueryPerformanceCounter(&llDetectEnd);
            QueryPerformanceFrequency(&llDetectFreq);
            float detectMs = (float)(llDetectEnd.QuadPart - llDetectStart.QuadPart) * 1000.0f / llDetectFreq.QuadPart;
            OutputDebugFormatA("  detect=%.2fms\n", detectMs);

            // --- Promote caret (with typing-loss fallback) ---
            // Only reuse the settled caret if the caret was lost during a typing-related
            // event (not a mouse click). Without this check, clicking on a non-text element
            // would trigger a settle whose fallback reuses the old caret, keeping the window visible.
            bool bCaretLossFallback = (g_ptLastQueriedCaret.y == 0 &&
                (bIsTypingEvent || (wParam == DETECT_REASON_SETTLE && bSettleFromTyping)) &&
                ShouldReuseCaretOnTypingLoss(g_hForegroundWindow));
            if (bCaretLossFallback)
            {
                g_ptCaret     = g_ptLastSettledCaret;
                g_rcContainer = g_rcLastSettledContainer;
                OutputDebugFormatA("  Caret lost after keypress - reusing settled caret(%d,%d) container(%d,%d,%d,%d)\n",
                                   g_ptCaret.x, g_ptCaret.y,
                                   g_rcContainer.left, g_rcContainer.top,
                                   g_rcContainer.right, g_rcContainer.bottom);
            }
            else
                g_ptCaret = g_ptLastQueriedCaret;

            if (bCaretLossFallback)
            {
                // Caret lost after typing key: cancel any running settle, paint with settled values
                if (g_bSettling)
                {
                    KillTimer(hWnd, SETTLE_TIMER_ID);
                    g_bSettling = false;
                }
                UpdateOverlayWindow(hWnd);
            }
            else if ((wParam == DETECT_REASON_KEY) || (wParam == DETECT_REASON_MOUSE) || (wParam == DETECT_REASON_ALT_UP))
            {
                // Determine if this is a context switch (long timeout needed)
                bool bLongTimeout = (g_hForegroundWindow != hPrevSettledFG) || (wParam == DETECT_REASON_ALT_UP);
                if (wParam == DETECT_REASON_MOUSE)
                {
                    int clickX = (short)LOWORD(lParam);
                    int clickY = (short)HIWORD(lParam);
                    int dx = clickX - g_ptLastSettledCaret.x;
                    int dy = clickY - g_ptLastSettledCaret.y;
                    if (dx*dx + dy*dy >= LONG_WAIT_ON_MOUSE_CLICK_DISTANCE_TH * LONG_WAIT_ON_MOUSE_CLICK_DISTANCE_TH)
                        bLongTimeout = true;
                }

                if (g_bSettling) // was in the middle of settling?...
                {
                    // New event while settling: abort current settle
                    KillTimer(hWnd, SETTLE_TIMER_ID);
                    g_bSettling = false;
                    if (!bLongTimeout)
                    {
                        // Same context: use last settled container, paint immediately
                        g_rcContainer = g_rcLastSettledContainer;
                        OutputDebugFormatA("  Settle interrupted - caret(%d,%d), last settled container(%d,%d,%d,%d)\n",
                                           g_ptCaret.x, g_ptCaret.y,
                                           g_rcContainer.left, g_rcContainer.top,
                                           g_rcContainer.right, g_rcContainer.bottom);
                        if (g_ptCaret.y != 0)
                            UpdateOverlayWindow(hWnd);
                    }
                    else
                        OutputDebugFormatA("  Settle interrupted - context switch, deferring paint\n");
                }
                // Start (or restart) settle loop — freeze grabs only if no recent event
                DWORD dwNow = GetTickCount();
                bool bOkToFreeze = (dwNow - dwLastEventTick > 250);
                dwLastEventTick = dwNow;
                g_bSettling = (bOkToFreeze || bLongTimeout) ? true : false;
                bSettleFromTyping = bIsTypingEvent;
                hSettleFG = g_hForegroundWindow;
                int firstSettleMs = bLongTimeout ? SETTLE_TIMER_ON_WINDOW_CHANGE_MS : SETTLE_TIMER_MS;
                ptPrevQueryCaret     = g_ptLastQueriedCaret;
                rcPrevQueryContainer = g_rcLastQueriedContainer;
                iSettleCount = 0;
                SetTimer(hWnd, SETTLE_TIMER_ID, firstSettleMs, NULL);
                if (!bOkToFreeze && !bLongTimeout)
                {
                    // Rapid event, same context: use last settled container, paint immediately
                    if (g_rcLastSettledContainer.right != 0)
                        g_rcContainer = g_rcLastSettledContainer;
                    else
                        g_rcContainer = g_rcLastQueriedContainer;
                    UpdateOverlayWindow(hWnd);
                }
            }
            else if (wParam == DETECT_REASON_TIMER)
            {
                // Periodic refresh: start a settle run (always freeze)
                g_bSettling = true;
                bSettleFromTyping = false;
                hSettleFG = g_hForegroundWindow;
                ptPrevQueryCaret     = g_ptLastQueriedCaret;
                rcPrevQueryContainer = g_rcLastQueriedContainer;
                iSettleCount = 0;
                SetTimer(hWnd, SETTLE_TIMER_ID, SETTLE_TIMER_MS, NULL);
            }
            else // wParam == DETECT_REASON_SETTLE
            {
                // Caret lost during settle: abort, reuse settled values
                if (bCaretLossFallback)
                {
                    KillTimer(hWnd, SETTLE_TIMER_ID);
                    g_bSettling = false;
                    OutputDebugFormatA("  Settle aborted - caret lost, reusing settled caret(%d,%d) container(%d,%d,%d,%d)\n",
                                       g_ptCaret.x, g_ptCaret.y,
                                       g_rcContainer.left, g_rcContainer.top,
                                       g_rcContainer.right, g_rcContainer.bottom);
                    UpdateOverlayWindow(hWnd);
                    break;
                }

                // Settle tick: compare queried values with previous tick
                bool bCaretMoved = (g_ptLastQueriedCaret.x != ptPrevQueryCaret.x ||
                                    g_ptLastQueriedCaret.y != ptPrevQueryCaret.y);
                bool bContainerChanged = (g_rcLastQueriedContainer.left   != rcPrevQueryContainer.left  ||
                                          g_rcLastQueriedContainer.right  != rcPrevQueryContainer.right ||
                                          g_rcLastQueriedContainer.top    != rcPrevQueryContainer.top   ||
                                          g_rcLastQueriedContainer.bottom != rcPrevQueryContainer.bottom);
                if ((bCaretMoved || bContainerChanged) && ++iSettleCount < 50)
                {
                    if (bContainerChanged && !bCaretMoved)
                        OutputDebugFormatA("  ******** Settle #%d  caret@(%d,%d), container (%d,%d,%d,%d)->(%d,%d,%d,%d), size %dx%d\n",
                                           iSettleCount,
                                           g_ptLastQueriedCaret.x, g_ptLastQueriedCaret.y,
                                           rcPrevQueryContainer.left, rcPrevQueryContainer.top,
                                           rcPrevQueryContainer.right, rcPrevQueryContainer.bottom,
                                           g_rcLastQueriedContainer.left, g_rcLastQueriedContainer.top,
                                           g_rcLastQueriedContainer.right, g_rcLastQueriedContainer.bottom,
                                           g_rcLastQueriedContainer.right - g_rcLastQueriedContainer.left,
                                           g_rcLastQueriedContainer.bottom - g_rcLastQueriedContainer.top);
                    else
                        OutputDebugFormatA("  Settle #%d  caret update (%d,%d)->(%d,%d), container@(%d,%d,%d,%d), size %dx%d\n",
                                           iSettleCount,
                                           ptPrevQueryCaret.x, ptPrevQueryCaret.y,
                                           g_ptLastQueriedCaret.x, g_ptLastQueriedCaret.y,
                                           g_rcLastQueriedContainer.left, g_rcLastQueriedContainer.top,
                                           g_rcLastQueriedContainer.right, g_rcLastQueriedContainer.bottom,
                                           g_rcLastQueriedContainer.right - g_rcLastQueriedContainer.left,
                                           g_rcLastQueriedContainer.bottom - g_rcLastQueriedContainer.top);
                    ptPrevQueryCaret     = g_ptLastQueriedCaret;
                    rcPrevQueryContainer = g_rcLastQueriedContainer;
                    SetTimer(hWnd, SETTLE_TIMER_ID, SETTLE_TIMER_MS, NULL);
                }
                else
                {
                    // Settled: promote caret; conditionally promote container
                    g_bSettling = false;
                    hPrevSettledFG       = hSettleFG;
                    g_ptLastSettledCaret = g_ptLastQueriedCaret;

                    bool bSkipContainer = bSettleFromTyping &&
                        ShouldSkipContainerUpdateOnTyping(g_hForegroundWindow);
                    if (bSkipContainer)
                    {
                        g_rcContainer = g_rcLastSettledContainer;
                        OutputDebugFormatA("  Settled:  caret@(%d,%d), container SKIPPED (typing) %dx%d ########################################\n",
                                           g_ptCaret.x, g_ptCaret.y,
                                           g_rcContainer.right - g_rcContainer.left,
                                           g_rcContainer.bottom - g_rcContainer.top);
                    }
                    else
                    {
                        g_rcLastSettledContainer = g_rcLastQueriedContainer;
                        g_rcContainer            = g_rcLastSettledContainer;
                        OutputDebugFormatA("  Settled:  caret@(%d,%d), container %dx%d ########################################\n",
                                           g_ptCaret.x, g_ptCaret.y,
                                           g_rcContainer.right - g_rcContainer.left,
                                           g_rcContainer.bottom - g_rcContainer.top);
                    }
                    UpdateOverlayWindow(hWnd);
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
            bool bAnimating = (g_iAnimWidth != g_iAnimToW || g_iAnimHeight != g_iAnimToH);
            if (bAnimating)
                AdvanceAnimation(); // updates g_iAnimWidth/Height from elapsed time
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
                // During animation, reuse the cache (filled before window was shown).
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
                    int srcW = width / ZOOM + CAPTURE_WIDTH_PADDING * 2;
                    int srcX = g_iSrcX - CAPTURE_WIDTH_PADDING;
                    StretchBlt(hCacheDC, 0, 0, fullW, fullH, hScreenDC, srcX, g_iSrcY, srcW, height, SRCCOPY);
                    bCacheValid = true;
#else
                    int srcW = width / ZOOM + CAPTURE_WIDTH_PADDING * 2;
                    int srcX = g_iSrcX - CAPTURE_WIDTH_PADDING;
                    StretchBlt(hdc, 0, 0, fullW, fullH, hScreenDC, srcX, g_iSrcY, srcW, height, SRCCOPY);
#endif
                    ReleaseDC(NULL, hScreenDC);
                }

                } // end of coordinate update + grab block
#ifdef USE_CACHING_DC
                skip_grab:
#endif

#ifdef USE_CACHING_DC
                // --- Blit from cache to window ---
                if (bCacheValid && !g_bFillCacheOnly)
                {
                    static COLORREF s_crBorder = GetSysColor(COLOR_ACTIVEBORDER);
#ifdef USE_ANIMATION
                    int dstW = g_iAnimWidth;
                    int dstH = g_iAnimHeight;
                    if (dstW > 0 && dstH > 0 && (dstW != fullW || dstH != fullH))
                    {
                        // Animating: StretchBlt the zoomed content centered in the window.
                        // The margins are left untouched -- they show the captured desktop
                        // background painted onto the window surface in MySetWindowPos.
                        int offsetX = (fullW - dstW) / 2;
                        int offsetY = (fullH - dstH) / 2;

                        SetStretchBltMode(hdc, HALFTONE);
                        StretchBlt(hdc, offsetX, offsetY, dstW, dstH,
                                   hCacheDC, 0, 0, fullW, fullH, SRCCOPY);

                        HBRUSH hBorderBrush = CreateSolidBrush(s_crBorder);
                        for (int b = 0; b < EXTRA_BORDER_THICKNESS; b++)
                        {
                            RECT rcBorder = { offsetX + b, offsetY + b, offsetX + dstW - b, offsetY + dstH - b };
                            FrameRect(hdc, &rcBorder, hBorderBrush);
                        }
                        DeleteObject(hBorderBrush);
                    }
                    else if (!bAnimating)
                    {
                        BitBlt(hdc, 0, 0, fullW, fullH, hCacheDC, 0, 0, SRCCOPY);

                        HBRUSH hBorderBrush = CreateSolidBrush(s_crBorder);
                        for (int b = 0; b < EXTRA_BORDER_THICKNESS; b++)
                        {
                            RECT rcBorder = { b, b, fullW - b, fullH - b };
                            FrameRect(hdc, &rcBorder, hBorderBrush);
                        }
                        DeleteObject(hBorderBrush);
                    }
                    // else: animation just started, dstW/dstH still 0 -- skip this frame
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
            if (g_iAnimWidth != g_iAnimToW || g_iAnimHeight != g_iAnimToH)
            {
                InvalidateRect(hWnd, NULL, FALSE); // request next animation frame
            }
            else if (bAnimating)
            {
                // Animation just finished -- restore DWM rounded corners
                DWORD cornerPref = 2; // DWMWCP_ROUND
                DwmSetWindowAttribute(hWnd, 33, &cornerPref, sizeof(cornerPref));
            }
#endif
        }
        break;
//    case WM_LBUTTONDOWN: // Can't do anything on button down - because the input focus is on our window. Improve later.
    case WM_LBUTTONUP:
        {
            if( g_rcContainer.right > g_rcContainer.left )
            {
                POINT ptCurrent; GetCursorPos(&ptCurrent);

                POINT ptClient = ptCurrent; ScreenToClient(hWnd, &ptClient);

                int srcW = g_iEffectiveWidth / ZOOM + CAPTURE_WIDTH_PADDING * 2;
                POINT ptTarget = { ptClient.x * srcW / g_iEffectiveWidth + g_iSrcX - CAPTURE_WIDTH_PADDING,
                                   ptClient.y / ZOOM + g_iSrcY };

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
    if (nCode == HC_ACTION)
    {
        KBDLLHOOKSTRUCT *pKB = (KBDLLHOOKSTRUCT *)lParam;
        if (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN)
        {
            // Block ESC when overlay is enabled: eat the keypress, disable overlay
            if (pKB->vkCode == VK_ESCAPE && g_bOverlayEnabled && g_ptCaret.y != 0)
            {
                g_bOverlayEnabled = false;
                PostMessage(g_myWindowHandle, WM_APP_DETECT_CARET,
                            DETECT_REASON_KEY, (LPARAM)VK_ESCAPE);
                return 1; // block ESC from reaching the target app
            }
            PostMessage(g_myWindowHandle, WM_APP_DETECT_CARET,
                        DETECT_REASON_KEY, (LPARAM)pKB->vkCode);
        }
        else if ((wParam == WM_KEYUP || wParam == WM_SYSKEYUP) &&
                 (pKB->vkCode == VK_MENU || pKB->vkCode == VK_LMENU || pKB->vkCode == VK_RMENU ||
                  pKB->vkCode == VK_CONTROL || pKB->vkCode == VK_LCONTROL || pKB->vkCode == VK_RCONTROL))
            PostMessage(g_myWindowHandle, WM_APP_DETECT_CARET,
                        DETECT_REASON_ALT_UP, 0);
    }
    return CallNextHookEx(NULL, nCode, wParam, lParam);
}

LRESULT CALLBACK LowLevelMouseProc(int nCode, WPARAM wParam, LPARAM lParam)
{
    LRESULT res = CallNextHookEx(NULL, nCode, wParam, lParam);
    if (nCode == HC_ACTION &&
        (wParam == WM_LBUTTONDOWN || wParam == WM_RBUTTONDOWN || wParam == WM_MBUTTONDOWN))
    {
        MSLLHOOKSTRUCT *pMS = (MSLLHOOKSTRUCT *)lParam;
        PostMessage(g_myWindowHandle, WM_APP_DETECT_CARET,
                    DETECT_REASON_MOUSE, MAKELPARAM(pMS->pt.x, pMS->pt.y));
    }
    return res;
}