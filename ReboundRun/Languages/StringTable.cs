using System.Globalization;

namespace ReboundRun.Languages
{
    public class StringTable
    {
        public static string AppTitle;
        public static string Run;

        public StringTable()
        {
            ReadLanguage();
        }

        public static void ReadLanguage()
        {
            // Get the current culture (language) of the system
            CultureInfo currentCulture = CultureInfo.CurrentUICulture;
            if (currentCulture.Name.ToLower() == "en-us")
            {
                AppTitle = "Rebound Run";
                Run = "Run";
            }
            if (currentCulture.Name.ToLower() == "ro-ro")
            {
                AppTitle = "Executare Rebound";
                Run = "Execută";
            }
        }
    }
}
