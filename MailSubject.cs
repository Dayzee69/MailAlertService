
namespace MailAlertService
{
    public class MailSubject
    {
        public string id { get; set; }
        public string uin { get; set; }
        public string from { get; set; }
        public string subject { get; set; }
        public string to { get; set; }
        public string date { get; set; }
        public string comment { get; set; }
        public string importance { get; set; }
        public string isRead { get; set; }

        public MailSubject(string from_, string subject_, string to_, string date_, string comment_, string importance_, string isRead_)
        {
            from = from_;
            subject = subject_;
            to = to_;
            date = date_;
            comment = comment_;
            importance = importance_;
            isRead = isRead_;
        }
    }
}
