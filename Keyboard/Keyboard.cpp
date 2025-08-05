#include <Windows.h>

extern "C" __declspec(dllexport) void __cdecl SendKeyScan(USHORT scanCode, BOOL keyDown, BOOL extended)
{
    INPUT input = {};
    input.type = INPUT_KEYBOARD;
    input.ki.wVk = 0;
    input.ki.wScan = scanCode;
    input.ki.dwFlags = KEYEVENTF_SCANCODE;

    if (!keyDown)
        input.ki.dwFlags |= KEYEVENTF_KEYUP;

    if (extended)
        input.ki.dwFlags |= KEYEVENTF_EXTENDEDKEY;

    SendInput(1, &input, sizeof(INPUT));
}
