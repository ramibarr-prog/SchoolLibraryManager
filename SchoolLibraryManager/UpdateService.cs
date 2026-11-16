using System.Net.Http;
using System.Reflection;
using System.Windows;
using System.Xml.Linq;
using AutoUpdaterDotNET;

namespace SchoolLibraryManager;

public sealed class UpdateService
{
    public const string UpdateXmlUrl = "https://github.com/ramibarr-prog/SchoolLibraryManager/releases/latest/download/update.xml";

    private const string AllowedDownloadPrefix = "https://github.com/ramibarr-prog/SchoolLibraryManager/releases/download/";

    private static readonly TimeSpan CheckTimeout = TimeSpan.FromSeconds(5);

    private readonly DatabaseService _databaseService;
    private readonly Window _owner;
    private bool _interactive;
    private bool _running;

    public UpdateService(DatabaseService databaseService, Window owner)
    {
        _databaseService = databaseService;
        _owner = owner;
    }

    public async Task CheckAsync(bool interactive)
    {
        if (_running || App.IsMaintenance || App.IsRestoring)
        {
            return;
        }

        _running = true;
        _interactive = interactive;
        try
        {
            var remoteVersion = await TryGetRemoteVersionAsync();
            if (remoteVersion is null)
            {
                if (interactive)
                {
                    MessageBox.Show(_owner, "לא ניתן לבדוק עדכונים כרגע (אין חיבור לאינטרנט או שאין גרסה זמינה).", "עדכוני תוכנה", MessageBoxButton.OK, MessageBoxImage.Information);
                }

                _running = false;
                return;
            }

            var current = Assembly.GetExecutingAssembly().GetName().Version ?? new Version(1, 0, 0, 0);
            if (remoteVersion <= current)
            {
                if (interactive)
                {
                    MessageBox.Show(_owner, $"התוכנה מעודכנת (גרסה {current}).", "עדכוני תוכנה", MessageBoxButton.OK, MessageBoxImage.Information);
                }

                _running = false;
                return;
            }

            AutoUpdater.CheckForUpdateEvent -= OnCheckForUpdate;
            AutoUpdater.CheckForUpdateEvent += OnCheckForUpdate;
            AutoUpdater.Mandatory = false;
            AutoUpdater.RunUpdateAsAdmin = false;
            AutoUpdater.Start(UpdateXmlUrl);
        }
        catch (Exception ex)
        {
            _running = false;
            ErrorLogger.Log(ex, "UpdateService.CheckAsync");
        }
    }

    private static async Task<Version?> TryGetRemoteVersionAsync()
    {
        try
        {
            using var client = new HttpClient { Timeout = CheckTimeout };
            var xml = await client.GetStringAsync(UpdateXmlUrl).ConfigureAwait(false);
            var text = XDocument.Parse(xml).Root?.Element("version")?.Value;
            return Version.TryParse(text?.Trim(), out var version) ? version : null;
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "UpdateService.TryGetRemoteVersionAsync");
            return null;
        }
    }

    private void OnCheckForUpdate(UpdateInfoEventArgs args)
    {
        AutoUpdater.CheckForUpdateEvent -= OnCheckForUpdate;
        try
        {
            if (args.Error is not null || !args.IsUpdateAvailable)
            {
                if (args.Error is not null)
                {
                    ErrorLogger.Log(args.Error, "UpdateService.OnCheckForUpdate");
                }

                return;
            }

            if (App.IsMaintenance || App.IsRestoring)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(args.DownloadURL)
                || !args.DownloadURL.StartsWith(AllowedDownloadPrefix, StringComparison.Ordinal))
            {
                ErrorLogger.Log($"Update blocked: unexpected download URL '{args.DownloadURL}'", "UpdateService.OnCheckForUpdate");
                if (_interactive)
                {
                    MessageBox.Show(_owner, "כתובת ההורדה של העדכון אינה מאומתת, ולכן העדכון נחסם.", "עדכוני תוכנה", MessageBoxButton.OK, MessageBoxImage.Warning);
                }

                return;
            }

            var answer = MessageBox.Show(_owner,
                $"קיימת גרסה חדשה ({args.CurrentVersion}). לפני העדכון ייווצר גיבוי של מסד הנתונים.\nלעדכן עכשיו?",
                "עדכוני תוכנה", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (answer != MessageBoxResult.Yes)
            {
                return;
            }

            if (string.IsNullOrEmpty(_databaseService.CreateBackup("PreUpdate")))
            {
                MessageBox.Show(_owner, "יצירת הגיבוי לפני העדכון נכשלה, ולכן העדכון בוטל. פרטים ביומן השגיאות.", "עדכוני תוכנה", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (AutoUpdater.DownloadUpdate(args))
            {
                Application.Current.Shutdown();
            }
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex, "UpdateService.OnCheckForUpdate");
            if (_interactive)
            {
                MessageBox.Show(_owner, "העדכון נכשל. פרטים ביומן השגיאות.", "עדכוני תוכנה", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        finally
        {
            _running = false;
        }
    }
}
