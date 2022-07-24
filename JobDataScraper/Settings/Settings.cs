using System;
using System.Configuration;

namespace JobDataScraper.Settings
{
    public static class Settings
    {
        public static string BOT_BROWSER = ConfigurationManager.AppSettings.Get("BOT_BROWSER");

        public static bool BOT_HEADLESS = Convert.ToBoolean(ConfigurationManager.AppSettings.Get("BOT_HEADLESS"));

        public static TimeSpan BOT_RETRY_TIME = new TimeSpan(0, Convert.ToInt32(ConfigurationManager.AppSettings.Get("BOT_RETRY_TIME")), 0);

        public static int BOT_SKIP = Convert.ToInt32(ConfigurationManager.AppSettings.Get("BOT_SKIP"));

        public static int BOT_TAKE = Convert.ToInt32(ConfigurationManager.AppSettings.Get("BOT_TAKE"));
    }
}
