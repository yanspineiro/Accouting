using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace SF_Lib
{
    class CookieAwareWebClient : WebClient
    {
        public CookieContainer cookie = new CookieContainer();
        public bool isfile;
        public int Timeout { get; set; }

        protected override WebRequest GetWebRequest(Uri address)
        {


            WebRequest request = base.GetWebRequest(address);
            if (request is HttpWebRequest)
            {
                (request as HttpWebRequest).CookieContainer = cookie;
            }
            (request as HttpWebRequest).AllowAutoRedirect = true;
            return request;
        }
    }
}
