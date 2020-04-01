using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace AddressBookUsingCSOM
{
    class SharePointAndRest_WebClient : IDisposable
    {
        private readonly WebClient webClient;
        public Uri WebUri { get; private set; }
        public ICredentials credentials;

        public SharePointAndRest_WebClient(Uri webUri)
        {
            webClient = new WebClient();
            webClient.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
            webClient.Headers.Add(HttpRequestHeader.ContentType, "application/json;odata=verbose");
            webClient.Headers.Add(HttpRequestHeader.Accept, "application/json;odata=verbose");
            const string userName = "madhuri@intern2k20.onmicrosoft.com";
            const string password = "Intern1567";
            var securePassword = new SecureString();
            foreach (var c in password)
            {
                securePassword.AppendChar(c);
            }
            credentials = new SharePointOnlineCredentials(userName, securePassword);
            webClient.Credentials = credentials;
            WebUri = webUri;
        }
        
        private string GetFormDigest()
        {
            string newUri = WebUri + "/_api/contextinfo";
            var endpointUri = new Uri(newUri);
            var result = webClient.UploadString(endpointUri, "POST");
            JToken t = JToken.Parse(result);
            return t["d"]["GetContextWebInformation"]["FormDigestValue"].ToString();
        }
        
        public void UpdateItem(string listTitle, object payload, int id)
        {
            var formDigestValue = GetFormDigest();
            webClient.Headers.Add("X-RequestDigest", formDigestValue);
            //Following code is required to perform the update
            webClient.Headers.Add("X-HTTP-Method", "MERGE");
            webClient.Headers.Add("IF-MATCH", "*");
            string newUri = WebUri + string.Format("/_api/web/lists/getbytitle('{0}'))", listTitle);
            var endpointUri = new Uri(newUri);
            var payloadString = JsonConvert.SerializeObject(payload);
            webClient.Headers.Add(HttpRequestHeader.ContentType, "application/json;odata=verbose");
            Console.WriteLine(webClient.UploadString(endpointUri, "POST", payloadString));
        }
        
        public void AddItem()
        {
            var contact = new Contact();
            contact.Title = "1";
            contact.Email = "yodha@intern2k20.onmicrosoft.com";
            UpdateItem("PostDataTrialList", contact, 1);
        }

        public JToken GetListItems(string listTitle)
        {
            var endpointUri = new Uri(WebUri, string.Format("_api/web/lists/getbytitle('{0}')/items", listTitle));
            var result = webClient.DownloadString(endpointUri);
            var t = JToken.Parse(result);
            return t["d"]["results"];
        }

        public void Dispose()
        {
            webClient.Dispose();
            GC.SuppressFinalize(this);
            GC.Collect();
        }
    }
}
