using System.IO;
using System.Windows;
using System.Windows.Threading;

namespace SchoolLibraryManager
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private DatabaseService? _databaseService;
        private Mutex? _appMutex;
        private bool _ownsAppMutex;

        public static volatile bool IsRestoring;
        public static volatile bool IsMaintenance;

        public App()
        {
            AppDomain.CurrentDomain.AssemblyResolve += ResolveFromLibsFolder;
        }

        private static System.Reflection.Assembly? ResolveFromLibsFolder(object? sender, ResolveEventArgs args)
        {
            try
            {
                var name = new System.Reflection.AssemblyName(args.Name).Name;
                if (string.IsNullOrEmpty(name))
                {
                    return null;
                }

                var path = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "libs", name + ".dll");
                return System.IO.File.Exists(path) ? System.Reflection.Assembly.LoadFrom(path) : null;
            }
            catch
            {
                return null;
            }
        }


        private static bool IsAceProviderMissing(Exception? ex)
        {
            for (var current = ex; current is not null; current = current.InnerException)
            {
                var msg = current.Message ?? string.Empty;
                if (msg.Contains("Microsoft.ACE.OLEDB", StringComparison.OrdinalIgnoreCase) &&
                    (msg.Contains("not registered", StringComparison.OrdinalIgnoreCase) ||
                     msg.Contains("provider", StringComparison.OrdinalIgnoreCase)))
                {
                    return true;
                }
            }

            return false;
        }

        private static void ShowAceMissingMessage() => MessageBox.Show(
            "לא נמצא רכיב Microsoft Access Database Engine בגרסת 64 סיביות (Microsoft.ACE.OLEDB.12.0).\n\n" +
            "יש להתקין את 'Microsoft Access Database Engine 2016 Redistributable' בגרסת 64 סיביות (AccessDatabaseEngine_X64.exe) ולהפעיל את התוכנה מחדש.",
            "רכיב מסד נתונים חסר",
            MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK,
            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);

        private static bool IsDatabaseNotFound(Exception? ex)
        {
            for (var current = ex; current is not null; current = current.InnerException)
            {
                if (current is FileNotFoundException && current.Message.Contains("DB_NOT_FOUND", StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }

        protected override async void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // הגדרת פורמט תאריכים ושפה גלובלית (ישראל) לכל חלקי המערכת ופקדי ה-WPF
            var culture = new System.Globalization.CultureInfo("he-IL");
            System.Threading.Thread.CurrentThread.CurrentCulture = culture;
            System.Threading.Thread.CurrentThread.CurrentUICulture = culture;
            FrameworkElement.LanguageProperty.OverrideMetadata(typeof(FrameworkElement),
                new FrameworkPropertyMetadata(System.Windows.Markup.XmlLanguage.GetLanguage(culture.IetfLanguageTag)));

            var isRestarted = e.Args.Contains("--restarted", StringComparer.OrdinalIgnoreCase);
            _appMutex = new Mutex(true, "SchoolLibraryAppMutex", out var createdNew);
            _ownsAppMutex = createdNew;
            if (!createdNew && isRestarted)
            {
                try
                {
                    _ownsAppMutex = _appMutex.WaitOne(10000);
                }
                catch (AbandonedMutexException)
                {
                    _ownsAppMutex = true;
                }
            }

            if (!_ownsAppMutex)
            {
                MessageBox.Show("המערכת כבר פועלת.", "מערכת ספרייה", MessageBoxButton.OK, MessageBoxImage.Information);
                Shutdown();
                return;
            }

            RegisterGlobalExceptionHandlers();

            try
            {

                _databaseService = new DatabaseService();
            }
            catch (Exception ex) when (IsDatabaseNotFound(ex))
            {
                ErrorLogger.Log(ex, "Application startup (database not found)");
                MessageBox.Show(
                    "קובץ מסד הנתונים (LibraryDB.mdb) לא נמצא. אנא ודא שהקובץ נמצא באותה תיקייה של קובץ ההפעלה.",
                    "שגיאת אתחול",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error,
                    MessageBoxResult.OK,
                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                Application.Current.Shutdown();
                return;
            }
            catch (Exception ex) when (IsAceProviderMissing(ex))
            {
                ErrorLogger.Log(ex, "Application startup (ACE provider missing)");
                ShowAceMissingMessage();
                Shutdown();
                return;
            }
            catch (Exception ex)
            {
                ErrorLogger.Log(ex, "Application startup (database location)");
                MessageBox.Show($"לא ניתן להכין את תיקיית הנתונים או לאתר את מסד הנתונים.\n{ex.Message}", "שגיאת מערכת", MessageBoxButton.OK, MessageBoxImage.Error);
                Shutdown();
                return;
            }

            var splash = new SplashWindow();
            splash.Show();

            try
            {
                await Task.Delay(1500);
                _databaseService!.Initialize();

                var mainWindow = new MainWindow();
                MainWindow = mainWindow;
                ShutdownMode = ShutdownMode.OnMainWindowClose;
                mainWindow.Show();
            }
            catch (Exception ex) when (IsAceProviderMissing(ex))
            {
                ErrorLogger.Log(ex, "Application startup (ACE provider missing)");
                ShowAceMissingMessage();
                splash.Close();
                Shutdown();
                return;
            }
            catch (Exception ex)
            {
                ErrorLogger.Log(ex, "Application startup");
                MessageBox.Show($"שגיאה באתחול מסד הנתונים. האפליקציה תיסגר. בדוק את ErrorLog.txt לפרטים.\n{ex.Message}", "שגיאת מערכת", MessageBoxButton.OK, MessageBoxImage.Error);
                splash.Close();
                Shutdown();
                return;
            }

            splash.Close();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine($"=== App.OnExit REACHED (databaseService null: {_databaseService is null}) ===");

            try
            {
                if (!IsRestoring && !IsMaintenance)
                {
                    _databaseService?.CreateBackup("Shutdown");
                }
            }
            catch (Exception ex)
            {
                ErrorLogger.Log(ex, "Application shutdown backup");
            }

            try
            {
                if (_ownsAppMutex)
                {
                    _appMutex?.ReleaseMutex();
                }
            }
            catch (Exception ex)
            {
                ErrorLogger.Log(ex, "Application mutex release");
            }

            base.OnExit(e);
        }

        private void RegisterGlobalExceptionHandlers()
        {
            DispatcherUnhandledException += OnDispatcherUnhandledException;
            AppDomain.CurrentDomain.UnhandledException += OnAppDomainUnhandledException;
            TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;
        }

        private static void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            if (DatabaseService.TryNotifyDatabaseAccessError(e.Exception))
            {
                e.Handled = true;
                return;
            }

            ErrorLogger.Log(e.Exception, "UI thread unhandled exception");
            MessageBox.Show("אירעה שגיאה לא צפויה. פרטי התקלה נשמרו ביומן.", "שגיאה", MessageBoxButton.OK, MessageBoxImage.Error);
            e.Handled = true;
        }

        private static void OnAppDomainUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception exception)
            {
                ErrorLogger.Log(exception, "Non-UI unhandled exception");
                return;
            }

            ErrorLogger.Log("Unknown non-UI unhandled exception", "AppDomain exception");
        }

        private static void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
        {
            ErrorLogger.Log(e.Exception, "Unobserved task exception");
            e.SetObserved();
        }
    }

}
