using Microsoft.Win32;
using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;

namespace xls2sql
{
    internal enum ThemeMode
    {
        System,
        Light,
        Dark,
    }

    /// <summary>
    /// Light/dark theme switching. Starts in <see cref="ThemeMode.System"/> (follows the Windows
    /// "app mode" setting, live) and is never persisted: every launch follows Windows again.
    /// Swaps the palette dictionary (Themes/Light.xaml or Themes/Dark.xaml) at index 0 of
    /// Application.Resources.MergedDictionaries; controls pick it up through DynamicResource.
    /// </summary>
    internal static class ThemeManager
    {
        private const string PersonalizeKey = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
        private const int PaletteIndex = 0;

        private static readonly Uri LightUri = new Uri("pack://application:,,,/Themes/Light.xaml");
        private static readonly Uri DarkUri = new Uri("pack://application:,,,/Themes/Dark.xaml");

        private static Dispatcher _dispatcher;
        private static bool? _appliedDark;

        public static ThemeMode Mode { get; private set; } = ThemeMode.System;

        /// <summary>Whether the dark palette is currently applied.</summary>
        public static bool IsDark => _appliedDark == true;

        /// <summary>Raised on the UI thread after the mode or the applied palette changes.</summary>
        public static event EventHandler ThemeChanged;

        public static void Initialize()
        {
            _dispatcher = Application.Current.Dispatcher;
            SystemEvents.UserPreferenceChanged += OnUserPreferenceChanged;
            Apply();
        }

        public static void Shutdown() => SystemEvents.UserPreferenceChanged -= OnUserPreferenceChanged;

        /// <summary>System → Light → Dark → System.</summary>
        public static void Cycle()
        {
            Mode = Mode switch
            {
                ThemeMode.System => ThemeMode.Light,
                ThemeMode.Light => ThemeMode.Dark,
                _ => ThemeMode.System,
            };
            Apply();
        }

        /// <summary>Windows "Choose your default app mode" = Dark.</summary>
        public static bool IsSystemDark()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(PersonalizeKey))
                    return key?.GetValue("AppsUseLightTheme") is int value && value == 0;
            }
            catch
            {
                return false; // Older Windows / no access: light.
            }
        }

        private static void OnUserPreferenceChanged(object sender, UserPreferenceChangedEventArgs e)
        {
            // Switching the Windows app mode raises General (ImmersiveColorSet). Raised on a SystemEvents thread.
            if (e.Category != UserPreferenceCategory.General && e.Category != UserPreferenceCategory.Color) return;
            _dispatcher?.BeginInvoke(new Action(() =>
            {
                if (Mode == ThemeMode.System) Apply();
            }));
        }

        private static void Apply()
        {
            bool dark = Mode == ThemeMode.Dark || (Mode == ThemeMode.System && IsSystemDark());
            if (_appliedDark != dark)
            {
                var dictionaries = Application.Current.Resources.MergedDictionaries;
                dictionaries[PaletteIndex] = new ResourceDictionary { Source = dark ? DarkUri : LightUri };
                _appliedDark = dark;
            }
            ThemeChanged?.Invoke(null, EventArgs.Empty);
        }

        // ---- title bar ------------------------------------------------------------------------

        private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
        private const int DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 = 19;

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int size);

        /// <summary>
        /// Dark/light native title bar (Windows 10 1809+ / 11). No-op elsewhere.
        /// Call after the window handle exists (SourceInitialized).
        /// </summary>
        public static void ApplyTitleBar(Window window)
        {
            IntPtr hwnd = new WindowInteropHelper(window).Handle;
            if (hwnd == IntPtr.Zero) return;
            int value = IsDark ? 1 : 0;
            try
            {
                if (DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ref value, sizeof(int)) != 0)
                    DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1, ref value, sizeof(int));
            }
            catch (DllNotFoundException) { }
            catch (EntryPointNotFoundException) { }
        }
    }
}
