using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.ServiceProcess;
using System.Text;
using MailKit;
using MailKit.Net.Imap;
using MimeKit;

namespace MailAlertService
{
    public partial class MailAlertService : ServiceBase
    {
        public MailAlertService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            Logger.InitLogger();

            Logger.Log.Info("Начало работы службы");
            
            System.Timers.Timer T2 = new System.Timers.Timer();
            T2.Interval = 60000;
            T2.AutoReset = true;
            T2.Enabled = true;
            T2.Start();
            T2.Elapsed += new System.Timers.ElapsedEventHandler(T2_Elapsed);
            
            //Alert();
            //MailCheck();
        }

        protected override void OnStop()
        {
            Logger.Log.Info("Окончание работы службы");
        }

        private void T2_Elapsed(object sender, EventArgs e)
        {

            try
            {
                string date = DateTime.Now.ToString("HH:mm");
                if (date == "13:00" || date == "17:00")
                {
                    MailCheck();
                    Alert();
                }

            }
            catch (Exception ex)
            {
                Logger.Log.Error(ex.ToString());
            }

        }

        private void Alert()
        {
            DBConnect conn = new DBConnect();
            SqlConnection sqlConnection = new SqlConnection(@"Data Source=" + conn.host + ";Initial Catalog=" + conn.db +
                ";" + "User ID=" + conn.user + ";Password=" + conn.password);
            List<MailSubject> mails = new List<MailSubject>();
            try
            {
                sqlConnection.Open();
                SqlCommand sqlCommand = sqlConnection.CreateCommand();
                sqlCommand.CommandText = "SELECT convert(varchar, Date_send, 104) Date_send, t0.ID, t0.Subject, t0.Send_from, t0.Send_to, t0.Comment, " +
                    "t0.Importance, t0.isRead, t1.UIN FROM Mails as t0, MyChatUsers as t1 WHERE t0.Send_to LIKE t1.MailAdress and isRead = 0 and t1.UIN != 1";
                using (DbDataReader reader = sqlCommand.ExecuteReader())
                {

                    while (reader.Read())
                    {
                        MailSubject mailsub = new MailSubject(reader.GetValue(3).ToString(), reader.GetValue(2).ToString(), reader.GetValue(4).ToString(),
                            reader.GetValue(0).ToString(), reader.GetValue(5).ToString(), reader.GetValue(6).ToString(), reader.GetValue(7).ToString());
                        mailsub.id = reader.GetValue(1).ToString();
                        mailsub.uin = reader.GetValue(8).ToString();
                        mails.Add(mailsub);
                    }

                }

                foreach (MailSubject mailsub in mails)
                {
                    SendPrivateMessage("У Вас есть важное непрочитанное письмо. От: " + mailsub.from + " Тема: " + mailsub.subject, mailsub.uin);
                }

            }
            catch (Exception ex)
            {
                Logger.Log.Error(ex.ToString());
            }
            finally
            {
                sqlConnection.Close();
                sqlConnection.Dispose();
            }
        }

        private void MailCheck()
        {
            DBConnect conn = new DBConnect();
            SqlConnection sqlConnection = new SqlConnection(@"Data Source=" + conn.host + ";Initial Catalog=" + conn.db +
                ";" + "User ID=" + conn.user + ";Password=" + conn.password);


            try
            { 
                
                sqlConnection.Open();
                SqlCommand sqlCommand = sqlConnection.CreateCommand();

                using (ImapClient client = new ImapClient())
                {
                    MailConnect mailConnect = new MailConnect();
                    client.Connect(mailConnect.imap, mailConnect.port, true);
                    client.Authenticate(mailConnect.mail, mailConnect.password);
                    
                    IMailFolder inbox = client.Inbox;
                    inbox.Open(FolderAccess.ReadWrite);
                    IList<IMessageSummary>messages = inbox.Fetch(0, -1, MessageSummaryItems.UniqueId);
                    int index = inbox.Count();

                    foreach (var item in messages)
                    {
                        MimeMessage message = inbox.GetMessage(item.UniqueId);
                        string messageText = message.GetTextBody(MimeKit.Text.TextFormat.Html);

                        sqlCommand.CommandText = $"SELECT * FROM Mails WHERE Send_from = '{message.From.ToString().Replace("'","")}' " +
                            $"and Subject = '{message.Subject.ToString().Replace("'", "")}' and Send_to = '{message.To.ToString().Replace("'", "")}'";

                        bool kostyl = false;
                        using (DbDataReader reader = sqlCommand.ExecuteReader())
                        {
                            if (!reader.HasRows)
                            {
                                kostyl = true;
                            }
                        }

                        if (kostyl)
                        {
                            SqlCommand command = sqlConnection.CreateCommand();
                            command.CommandText = "INSERT INTO Mails (Send_from, Subject, Date_send, Send_to, Importance, isRead) " +
                                $"VALUES ('{message.From.ToString().Replace("'","")}', '{message.Subject.ToString().Replace("'", "")}', " +
                                $"'{message.Date.ToString().Replace("'", "")}', '{message.To.ToString().Replace("'", "")}', 'Важное', 0)";

                            command.ExecuteNonQuery();
                            inbox.AddFlags(item.UniqueId, MessageFlags.Deleted, true);
                            inbox.Expunge();
                        }

                    }
                    
                    inbox.Close();

                }
            }
            catch (Exception ex)
            {
                Logger.Log.Error(ex.ToString());
            }
            finally
            {
                sqlConnection.Close();
                sqlConnection.Dispose();
            }
        }

        private void SendPrivateMessage(string message_, string to_)
        {
            try
            {
                string strPath = System.IO.Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);
                INIManager manager = new INIManager(strPath + @"\settings.ini");

                PrivateMessage PM = new PrivateMessage();
                PM.UserTo = to_;
                PM.UserFrom = "0";
                PM.Msg = message_;
                PM.hash = "";
                PM.APIStype = "c#";
                PM.ServerKey = "iddqd";
                string json_send_message = JsonConvert.SerializeObject(PM);

                TcpClient client = new TcpClient();
                client.Connect(manager.GetPrivateString("MyChatSettings", "host"), int.Parse(manager.GetPrivateString("MyChatSettings", "port")));

                StreamWriter sw = new StreamWriter(client.GetStream(), Encoding.GetEncoding(1251));
                sw.AutoFlush = true;
                string message = "mc5.20" + PM.CRLF;
                sw.WriteLine(message);

                StreamReader sr = new StreamReader(client.GetStream(), Encoding.GetEncoding(1251));
                sr.ReadLine();

                sw.WriteLine(PM.MagicPacket + PM.cs_integration_api + PM.iFlag + PM.MCIAPI_CS_SendPrivateMessage + json_send_message + PM.CRLF);

                client.Close();
            }
            catch (Exception ex)
            {
                Logger.Log.Error(ex.ToString());
            }
        }

        private void SendChannelMessage(string message_, string to_)
        {
            try
            {
                PrivateMessage PM = new PrivateMessage();
                PM.UID = to_;
                PM.UserFrom = "0";
                PM.Msg = message_;
                PM.hash = "";
                PM.APIStype = "c#";
                PM.ServerKey = "iddqd";
                string json_send_message = JsonConvert.SerializeObject(PM);

                TcpClient client = new TcpClient();
                client.Connect("sk-as1", 2004);

                StreamWriter sw = new StreamWriter(client.GetStream(), Encoding.GetEncoding(1251));
                sw.AutoFlush = true;
                string message = "mc5.20" + PM.CRLF;
                sw.WriteLine(message);

                StreamReader sr = new StreamReader(client.GetStream(), Encoding.GetEncoding(1251));
                sr.ReadLine();

                sw.WriteLine(PM.MagicPacket + PM.cs_integration_api + PM.iFlag + PM.MCIAPI_CS_SendChannelMessage + json_send_message + PM.CRLF);

                client.Close();
            }
            catch (Exception ex)
            {
                Logger.Log.Error(ex.ToString());
            }
        }
    }
}
