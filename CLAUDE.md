# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

ComfyTyping is a Windows desktop application that displays a zoomed-in view of the area around the text cursor, improving typing visibility. It is a pure Win32 C++ application with no external frameworks or third-party dependencies.

## Build

Visual Studio 2022 (v143 toolset, Professional edition) is required. Configurations: Debug/Release for Win32 and x64. Linked libraries: `Oleacc.lib` (accessibility API), `dwmapi.lib` (DWM rounded corners).

### Building from Claude Code (bash on Windows)

Call MSBuild.exe directly by full path. Use `-p:` instead of `/p:` (bash interprets forward slashes):

```bash
"C:/Program Files/Microsoft Visual Studio/2022/Professional/MSBuild/Current/Bin/MSBuild.exe" "e:/EUGENE_PRIVATE/TypingAssist/ComfyTyping/ComfyTyping.vcxproj" -p:Configuration=Release -p:Platform=x64 -verbosity:minimal 2>&1
```

Output binary: `x64\Release\ComfyTyping.exe`

## Architecture

### Source Layout

- **ComfyTyping.cpp** — Main application: entry point (`wWinMain`), window procedure (`WndProc`), rendering via `StretchBlt`, keyboard/event hooks, window positioning logic
- **WinUtils.cpp** — Windows API utilities: DPI scaling, caret position detection (primary via `GUITHREADINFO`, fallback via `IAccessible`), taskbar height, rounded corners
- **ComfyTyping.h** — Compile-time configuration: zoom factor, window proportions, feature flags
- **WinUtils.h** — Global variable declarations using `MYGLOBAL` macro pattern, utility function prototypes

### Key Patterns

**Global state**: All globals are declared in `WinUtils.h` using `MYGLOBAL(type, name, init)`. `WinUtils.cpp` defines `WIN_UTILS_CPP` before including the header so the macro expands to definitions there and `extern` declarations everywhere else.

**Dual caret detection**: `GetCaretPosition()` uses `GetGUIThreadInfo` as the primary method; `GetCaretPositionFromAccessibility()` uses the `IAccessible` COM interface as a fallback when the primary method returns (0,0).

**Hooks**: Two hook mechanisms run simultaneously:
- `LowLevelKeyboardProc` — system-wide keyboard hook that detects printable key presses (shows window) and Escape (toggles visibility)
- `WinEventProc_ForFocusedClientWnd` — window event hook tracking focus changes to identify the correct child window for content capture

**Timer-based updates**: `INVALIDATE_TIMER_ID` (16ms) triggers repainting; `CARET_TIMER_ID` (1ms) updates caret position and manages window visibility.

**Rendering**: Captures a portion of the desktop around the caret using `StretchBlt` and displays it at `ZOOM`x magnification. The window is always-on-top (`HWND_TOPMOST`).

### Compile-Time Feature Flags (ComfyTyping.h)

Enabled: `INVALIDATE_ON_TIMER`, `INVALIDATE_ON_HOOK`, `NO_RESIZE_NO_BORDER`
Disabled (commented out): `RENDER_CURSOR`, `USE_FONT_HEIGHT`, `HAS_MENU`, `HAS_TITLE_BAR`, `POSITION_OVER_TASKBAR`

These flags guard conditional compilation blocks throughout the codebase via `#ifdef`.

### Window Positioning

The zoomed window's horizontal position follows a three-zone strategy: cursor in the left quarter of screen aligns the window left, in the right quarter aligns right, and in the middle half centers the view. The `g_bFreezeSrcXandSrcY` flag prevents view jumping when clicking on the zoomed window itself.
