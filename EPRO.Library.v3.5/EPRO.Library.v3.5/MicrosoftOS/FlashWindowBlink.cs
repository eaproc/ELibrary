using System;
using System.Runtime.InteropServices;

namespace ELibrary.Standard.MicrosoftOS
{
    public class FlashWindowBlink
    {
        [DllImport("User32")]
        private static extern bool FlashWindowEx(ref FLASHWINFO fwInfo);

        // As defined by: http://msdn.microsoft.com/en-us/library/ms679347(v=vs.85).aspx
        public enum FlashWindowFlags : uint
        {
            // Stop flashing. The system restores the window to its original state.
            FLASHW_STOP = 0U,
            // Flash the window caption.
            FLASHW_CAPTION = 1U,
            // Flash the taskbar button.
            FLASHW_TRAY = 2U,
            // Flash both the window caption and taskbar button.
            // This is equivalent to setting the FLASHW_CAPTION | FLASHW_TRAY flags.
            FLASHW_ALL = 3U,
            // Flash continuously, until the FLASHW_STOP flag is set.
            FLASHW_TIMER = 4U,
            // Flash continuously until the window comes to the foreground.
            FLASHW_TIMERNOFG = 12U
        }

        private struct FLASHWINFO
        {
            public uint cbSize;
            public IntPtr hwnd;
            public FlashWindowFlags dwFlags;
            public uint uCount;
            public uint dwTimeout;
        }

        public static bool FlashWindow(ref IntPtr handle, bool FlashTitleBar, bool FlashTray, int FlashCount)
        {
            if (handle == default)
                return false;
            try
            {
                var fwi = new FLASHWINFO();
                {
                    var withBlock = fwi;
                    withBlock.hwnd = handle;
                    if (FlashTitleBar)
                        withBlock.dwFlags = withBlock.dwFlags | FlashWindowFlags.FLASHW_CAPTION;
                    if (FlashTray)
                        withBlock.dwFlags = withBlock.dwFlags | FlashWindowFlags.FLASHW_TRAY;
                    withBlock.uCount = (uint)FlashCount;
                    if (FlashCount == 0)
                        withBlock.dwFlags = withBlock.dwFlags | FlashWindowFlags.FLASHW_TIMERNOFG;
                    withBlock.dwTimeout = 0U; // Use the default cursor blink rate.
                    withBlock.cbSize = (uint)Marshal.SizeOf(fwi);
                }

                return FlashWindowEx(ref fwi);
            }
            catch
            {
                return false;
            }
        }

        public static bool StopFlashWindow(ref IntPtr handle)
        {
            if (handle == default)
                return false;
            try
            {
                var fwi = new FLASHWINFO();
                {
                    var withBlock = fwi;
                    withBlock.hwnd = handle;
                    withBlock.dwFlags = FlashWindowFlags.FLASHW_STOP;
                    withBlock.dwTimeout = 0U; // Use the default cursor blink rate.
                    withBlock.cbSize = (uint)Marshal.SizeOf(fwi);
                }

                return FlashWindowEx(ref fwi);
            }
            catch
            {
                return false;
            }
        }
    }
}