#pragma once

#include "resource.h"

#include <cstdarg>
#include <cstdio>

inline void OutputDebugFormatA(const char* format, ...) {
    char buffer[256];
    va_list args;
    va_start(args, format);
    vsnprintf(buffer, sizeof(buffer), format, args);
    va_end(args);
    OutputDebugStringA(buffer);
}
