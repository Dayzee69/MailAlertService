using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailAlertService
{
    class PrivateMessage
    {
        public string UserTo { get; set; }
        public string UID { get; set; }
        public string UserFrom { get; set; }
        public string Msg { get; set; }
        public string hash { get; set; }
        public string APIStype { get; set; }
        public string ServerKey { get; set; }

        public string CRLF = "\u000D\u000A";
        public string MagicPacket = "\u0017\u0006";
        public string cs_integration_api = "0077";
        public string iFlag = "30";
        public string MCIAPI_CS_SendPrivateMessage = "0002";
        public string MCIAPI_CS_SendChannelMessage = "0004";
    }
}
