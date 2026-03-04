// IniReader.cpp — reads ComfyTyping.ini into memory for whitelist/blacklist lookups.
//
// Uses GetPrivateProfileSection to enumerate all key=value pairs in a section.
// The trick: it returns a buffer of null-terminated strings with a double null
// at the end.  We parse out just the keys (exe base names without .exe).

#include "framework.h"
#include "IniReader.h"

#include <cstring>

#define MAX_INI_APPS  64
#define MAX_APP_NAME  64

struct IniConfig
{
    bool onlyWhitelisted;

    char whitelisted[MAX_INI_APPS][MAX_APP_NAME];
    int  numWhitelisted;

    char blacklisted[MAX_INI_APPS][MAX_APP_NAME];
    int  numBlacklisted;
};

static IniConfig g_iniConfig = {};
static char      g_szIniPath[MAX_PATH] = {};

// ---------------------------------------------------------------------------
// Helper: parse a GetPrivateProfileSection buffer (double-null-terminated
// sequence of "key=value\0" strings) and store just the keys into outArray.
// Returns the number of entries stored.
// ---------------------------------------------------------------------------
static int ParseSectionKeys(const char* buffer, char outArray[][MAX_APP_NAME], int maxEntries)
{
    int count = 0;
    const char* p = buffer;

    while (*p && count < maxEntries)
    {
        // Each entry is "key=value\0".  Find the '=' to isolate the key.
        const char* eq = strchr(p, '=');
        int keyLen;
        if (eq)
            keyLen = (int)(eq - p);
        else
            keyLen = (int)strlen(p);   // no '=' — treat whole string as key

        // Trim trailing spaces from the key
        while (keyLen > 0 && p[keyLen - 1] == ' ')
            keyLen--;

        if (keyLen > 0 && keyLen < MAX_APP_NAME)
        {
            memcpy(outArray[count], p, keyLen);
            outArray[count][keyLen] = '\0';
            // Lowercase the key for consistent matching
            for (int j = 0; outArray[count][j]; j++)
                outArray[count][j] = (char)tolower((unsigned char)outArray[count][j]);
            count++;
        }

        // Advance past this entry's null terminator
        p += strlen(p) + 1;
    }

    return count;
}

// ---------------------------------------------------------------------------
// ReadConfiguration — (re)loads ComfyTyping.ini into g_iniConfig.
// ---------------------------------------------------------------------------
void ReadConfiguration()
{
    // Resolve INI path once (same directory as the exe)
    if (g_szIniPath[0] == '\0')
    {
        GetModuleFileNameA(NULL, g_szIniPath, MAX_PATH);
        // Replace exe filename with ini filename
        char* lastSlash = strrchr(g_szIniPath, '\\');
        if (lastSlash)
            strcpy_s(lastSlash + 1, MAX_PATH - (lastSlash + 1 - g_szIniPath), "ComfyTyping.ini");
    }

    // Check if the INI file exists
    DWORD attr = GetFileAttributesA(g_szIniPath);
    if (attr == INVALID_FILE_ATTRIBUTES)
    {
        MessageBoxA(NULL, g_szIniPath, "ComfyTyping: INI file not found", MB_OK | MB_ICONWARNING);
        return;
    }

    // Clear previous config
    memset(&g_iniConfig, 0, sizeof(g_iniConfig));

    // [default] section
    g_iniConfig.onlyWhitelisted =
        GetPrivateProfileIntA("default", "ONLY_WHITELISTED", 0, g_szIniPath) != 0;

    // [whitelisted] section — enumerate all key=value pairs
    char sectionBuf[4096] = {};
    GetPrivateProfileSectionA("whitelisted", sectionBuf, sizeof(sectionBuf), g_szIniPath);
    g_iniConfig.numWhitelisted = ParseSectionKeys(sectionBuf, g_iniConfig.whitelisted, MAX_INI_APPS);

    // [blacklisted] section
    memset(sectionBuf, 0, sizeof(sectionBuf));
    GetPrivateProfileSectionA("blacklisted", sectionBuf, sizeof(sectionBuf), g_szIniPath);
    g_iniConfig.numBlacklisted = ParseSectionKeys(sectionBuf, g_iniConfig.blacklisted, MAX_INI_APPS);
}

// ---------------------------------------------------------------------------
// Lookup functions — both INI keys and exeBaseName are lowercase.
// exeBaseName should be lowercase, without ".exe".
// ---------------------------------------------------------------------------
bool IsAppBlacklisted(const char* exeBaseName)
{
    for (int i = 0; i < g_iniConfig.numBlacklisted; i++)
    {
        if (strcmp(exeBaseName, g_iniConfig.blacklisted[i]) == 0)
            return true;
    }
    return false;
}

bool IsAppWhitelisted(const char* exeBaseName)
{
    for (int i = 0; i < g_iniConfig.numWhitelisted; i++)
    {
        if (strcmp(exeBaseName, g_iniConfig.whitelisted[i]) == 0)
            return true;
    }
    return false;
}

bool IsOnlyWhitelistedMode()
{
    return g_iniConfig.onlyWhitelisted;
}
