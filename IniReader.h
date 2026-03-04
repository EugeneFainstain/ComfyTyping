#pragma once

// Read (or reload) configuration from ComfyTyping.ini next to the exe.
// Populates in-memory whitelist/blacklist arrays.
void ReadConfiguration();

// In-memory lookups — no disk I/O.
bool IsAppBlacklisted(const char* exeBaseName);
bool IsAppWhitelisted(const char* exeBaseName);
bool IsOnlyWhitelistedMode();
