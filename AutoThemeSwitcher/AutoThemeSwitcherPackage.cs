using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using Task = System.Threading.Tasks.Task;

namespace AutoThemeSwitcher
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [ProvideAutoLoad(UIContextGuids.NoSolution, PackageAutoLoadFlags.BackgroundLoad)]
    [Guid(AutoThemeSwitcherPackage.PackageGuidString)]
    public sealed class AutoThemeSwitcherPackage : AsyncPackage
    {
        /// <summary>
        /// AutoThemeSwitcherPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "abd06d75-bf29-4896-bebf-1604bae4218b";

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            // When initialized asynchronously, the current thread may be a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.

            await this.JoinableTaskFactory.SwitchToMainThreadAsync();

            var serviceProvider = ServiceProvider.GlobalProvider;
            Dte = (DTE)serviceProvider.GetService(typeof(DTE));
            Assumes.Present(Dte);
            var events = Dte.Events;
            SolEvents = events.SolutionEvents;
            SolEvents.Opened += SolutionEvents_Opened;
            EnableThemeSwitchAsync().FireAndForget();
        }

        private SolutionEvents SolEvents;
        private DTE Dte;
        private FileSystemWatcher FileSystemWatcher;


        /// <summary>Gets all themes installed in Visual Studio.</summary>
        /// <returns>All themes installed in Visual Studio.</returns>
        /// Source: https://github.com/frankschierle/ThemeSwitcher/blob/master/ThemeSwitcher/Logic/ThemeManager.cs
        public Dictionary<string, string> GetInstalledThemes()
        {
            string[] installedThemesKeys;

            var themes = new List<(string id, string displayname)>();
            using (var themesKey = ApplicationRegistryRoot.OpenSubKey("Themes"))
            {
                if (themesKey != null)
                {
                    installedThemesKeys = themesKey.GetSubKeyNames();

                    foreach (string key in installedThemesKeys)
                    {
                        using (RegistryKey themeKey = themesKey.OpenSubKey(key))
                        {
                            if (themeKey != null)
                            {
                                themes.Add((key, themeKey.GetValue(null)?.ToString()));
                            }
                        }
                    }
                }
            }
            return themes.ToDictionary(t => t.displayname, t => t.id);
        }

        private FileSystemWatcher WatchThemeFile(string solutionFileName)
        {
            if (!string.IsNullOrEmpty(solutionFileName))
            {
                var folder = Path.GetDirectoryName(solutionFileName);
                return new FileSystemWatcher(folder, ".vstheme");
            }
            return null;
        }

        private async Task EnableThemeSwitchAsync()
        {
            ClearFileSystemWatcher();
            await JoinableTaskFactory.SwitchToMainThreadAsync();
            string solutionFileName = Dte.Solution.FileName;
            FileSystemWatcher = WatchThemeFile(solutionFileName);
            if (FileSystemWatcher is object)
            {
                FileSystemWatcher.Created += ThemeFileChanged;
                FileSystemWatcher.Changed += ThemeFileChanged;
                FileSystemWatcher.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.CreationTime | NotifyFilters.FileName;
                FileSystemWatcher.EnableRaisingEvents = true;

                RefreshTheme(Path.Combine(FileSystemWatcher.Path, FileSystemWatcher.Filter));
            }
        }

        private void ClearFileSystemWatcher()
        {
            if (FileSystemWatcher is object)
            {
                FileSystemWatcher.Created -= ThemeFileChanged;
                FileSystemWatcher.Changed -= ThemeFileChanged;
                FileSystemWatcher.Dispose();
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                ClearFileSystemWatcher();
            }

            base.Dispose(disposing);
        }

        private void ThemeFileChanged(object sender, FileSystemEventArgs e) => RefreshTheme(e.FullPath);

        private void RefreshTheme(string themefilepath)
        {
            if (File.Exists(themefilepath))
            {
                var firstThemeName = File.ReadAllLines(themefilepath).FirstOrDefault();
                if (firstThemeName is string)
                {
                    var installedThemes = GetInstalledThemes();
                    if (installedThemes.TryGetValue(firstThemeName, out string themeID))
                    {
                        ApplyTheme(themeID);
                    }
                }
            }
        }

        private void SolutionEvents_Opened() => EnableThemeSwitchAsync().FireAndForget();

        /// <summary>Applies a given <see cref="Theme" />.</summary>
        /// <param name="theme">The theme to apply.</param>
        /// <exception cref="ArgumentNullException">Occurs if <paramref name="theme" /> is null.</exception>
        public void ApplyTheme(string themeID)
        {
            if (themeID == null)
            {
                throw new ArgumentNullException(nameof(themeID));
            }

            var key = UserRegistryRoot.OpenSubKey(@"ApplicationPrivateSettings\Microsoft\VisualStudio", true);
            if (key != null)
            {
                object oldColorTheme = key.GetValue("ColorTheme");
                object oldColorThemeNew = key.GetValue("ColorThemeNew");
                object newColorTheme = "0*System.String*" + themeID.Trim('{', '}');
                object newColorThemeNew = "0*System.String*" + themeID;
                if (oldColorTheme != newColorTheme || oldColorThemeNew != newColorThemeNew)
                {
                    key.SetValue("ColorTheme", newColorTheme);
                    key.SetValue("ColorThemeNew", newColorThemeNew);

                    NativeMethods.SendNotifyMessage(new IntPtr(NativeMethods.HWND_BROADCAST), NativeMethods.WM_SYSCOLORCHANGE, IntPtr.Zero, IntPtr.Zero);

                    RestoreColorThemeAsync(key, oldColorTheme, oldColorThemeNew).FireAndForget();
                }
            }
        }

        async Task RestoreColorThemeAsync(RegistryKey key, object oldColorTheme, object oldColorThemeNew)
        {
            await Task.Delay(60000);

            key.SetValue("ColorTheme", oldColorTheme);
            key.SetValue("ColorThemeNew", oldColorThemeNew);
            key.Dispose();
        }


        #endregion
    }
}
