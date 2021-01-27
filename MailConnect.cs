using System;

namespace MailAlertService
{
    class MailConnect
    {
        public string imap { get; set; }
        public int port { get; set; }
        public string mail { get; set; }
        public string password { get; set; }

        public MailConnect()
        {
            try 
            {
                string strPath = System.IO.Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);
                INIManager manager = new INIManager(strPath + @"\settings.ini");
                imap = manager.GetPrivateString("MailSettings", "imap");
                port = int.Parse(manager.GetPrivateString("MailSettings", "port"));
                mail = manager.GetPrivateString("MailSettings", "mail");
                password = manager.GetPrivateString("MailSettings", "password");
            }
            catch (Exception ex)
            {
                Logger.Log.Error(ex.ToString());
            }
        }
    }
}
