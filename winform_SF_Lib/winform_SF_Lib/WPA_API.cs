using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Web;
using System.Collections.Specialized;
using System.Net;
using System.Net.Http;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace winform_SF_Lib
    {
    class WPA_API
        {
        public static void wpa_interface()
            {
            string url_wpa = @"https://www.enrollment123.com/api/user.cfc";
            var collection = new NameValueCollection();
            collection.Add("method", "doLookUp");
            collection.Add("corpid", "1190");
            collection.Add("ipaddress", "64.132.148.194");
            collection.Add("user_username", "");
            collection.Add("user_userpassword", "");
            collection.Add("memberid", "");
            collection.Add("phone", "");
            collection.Add("allowmultiple", "1");
            collection.Add("username", "HBO109553");
            collection.Add("password", "apiWPA109553");
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3;

            var client = new SF_Lib.CookieAwareWebClient();
            client.BaseAddress = url_wpa;
            //client.Headers.Add("authorization", "Basic " + WPA_API.base64encode("HBO109553 : apiWPA109553"));

            byte[] responseBytes = client.UploadValues("user.cfc", "POST", collection);

            string response = Encoding.UTF8.GetString(responseBytes);



            }

        static public string base64encode(string plaintext)
            {
            var plainTextByte = System.Text.Encoding.UTF8.GetBytes(plaintext);
                return System.Convert.ToBase64String(plainTextByte);
            }

        public static string postXMLData(string destinationUrl, WPA_XML _WPA_XML)
            {

            string requestXml= "";
            System.Xml.Serialization.XmlSerializer x = new System.Xml.Serialization.XmlSerializer(_WPA_XML.GetType());
            using (StringWriter sww = new StringWriter())
            using (XmlWriter writer = XmlWriter.Create(sww))
                {
                x.Serialize(writer, _WPA_XML);
                requestXml = sww.ToString(); // Your XML
                }

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(destinationUrl);
            byte[] bytes;
            bytes = System.Text.Encoding.ASCII.GetBytes(requestXml);
            request.ContentType = "text/xml; encoding='utf-8'";
            request.ContentLength = bytes.Length;
            request.Method = "POST";
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
     | SecurityProtocolType.Tls11
     | SecurityProtocolType.Tls12
     | SecurityProtocolType.Ssl3;
       
            Stream requestStream = request.GetRequestStream();
            requestStream.Write(bytes, 0, bytes.Length);
            requestStream.Close();
            HttpWebResponse response;
            response = (HttpWebResponse)request.GetResponse();
            if (response.StatusCode == HttpStatusCode.OK)
                {
                Stream responseStream = response.GetResponseStream();
                string responseStr = new StreamReader(responseStream).ReadToEnd();
                return responseStr;
                }
            return null;
            }


        }

    public class WPA_XML
        {

        public string CORPID { set; get; }
        public string USERNAME { set; get; }
        public string PASSWORD { set; get; }

        }



       
    }
