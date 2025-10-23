using System;
using System.Runtime.InteropServices;

namespace AiCalc;

internal static class DebugConsole
{
#if DEBUG
    private const uint ATTACH_PARENT_PROCESS = 0xFFFFFFFF;
    private static bool _initialized;
    private static bool _hasConsole;

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool AllocConsole();

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool AttachConsole(uint dwProcessId);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GetConsoleWindow();

    public static bool EnsureInitialized()
    {
        if (_initialized)
        {
            return _hasConsole;
        }

        _initialized = true;

        if (GetConsoleWindow() != IntPtr.Zero)
        {
            _hasConsole = true;
            return true;
        }

        try
        {
            if (!AttachConsole(ATTACH_PARENT_PROCESS))
            {
                _hasConsole = AllocConsole();
            }
            else
            {
                _hasConsole = true;
            }

            if (_hasConsole)
            {
                Console.OutputEncoding = System.Text.Encoding.UTF8;
                Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] Debug console attached.");
            }
        }
        catch
        {
            _hasConsole = false;
        }

        return _hasConsole;
    }
#else
    public static bool EnsureInitialized() => false;
#endif
}
