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
using System.Xml;

namespace AddressBookUsingCSOM
{
    public class ContactService
    {
        private readonly WebClient webClient;
        public Uri webUri { get; private set; }
        public ICredentials credentials;

        public ContactService()
        {
            webClient = new WebClient();
            webClient.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
            webClient.Headers.Add(HttpRequestHeader.ContentType, "application/json;odata=verbose");
            webClient.Headers.Add(HttpRequestHeader.Accept, "application/json;odata=verbose");
            webUri = new Uri("https://intern2k20.sharepoint.com/sites/Technovert");
            const string userName = "madhuri@intern2k20.onmicrosoft.com";
            const string password = "Intern1567";
            var securePassword = new SecureString();
            foreach (var c in password)
            {
                securePassword.AppendChar(c);
            }
            credentials = new SharePointOnlineCredentials(userName, securePassword);
            webClient.Credentials = credentials;
        }

        public JToken GetAllContacts(string listTitle)
        {
            string newUri = webUri + string.Format("/_api/web/lists/getbytitle('{0}')/items", listTitle);
            var endpointUri = new Uri(newUri);
            var result = webClient.DownloadString(endpointUri);
            var t = JToken.Parse(result);
            return t["d"]["results"];
        }

        public JToken GetGroupLabel(int groupId,string listTitle)
        {
            string newUri = webUri + string.Format("/_api/web/lists/getbytitle('TaxonomyHiddenList')/items", listTitle) + string.Format("?$filter=Id eq {0} &$select = Id,Title",groupId);
            var endpointUri = new Uri(newUri);
            var result = webClient.DownloadString(endpointUri);
            var t = JToken.Parse(result);
            return t["d"]["results"][0];
        }

        public JToken GetContactName(int id,string listTitle)
        {
            string newUri = webUri + string.Format("/_api/web/lists/getbytitle('{0}')/items({1})?$select=ContactName/Title&$expand=ContactName/Id", listTitle,id);
            var endpointUri = new Uri(newUri);
            var result = webClient.DownloadString(endpointUri);
            var t = JToken.Parse(result);
            return t["d"]["ContactName"];
        }

        public JToken GetUserName(string email)
        {
            string siteUrl = "http://intern2k20.sharepoint.com";
            string accountName = "i: 0#.f|membership|" + email;
            string uri = siteUrl + "/_api/web/siteusers(@v)?@v='" +accountName + "'";
            var endpointUri = new Uri(uri);
            var result = webClient.DownloadString(endpointUri);
            var t = JToken.Parse(result);
            return t["d"]["results"];
        }

        public JToken GetContactById(string listTitle,int id)
        {
            string newUri = webUri + string.Format("/_api/web/lists/getbytitle('{0}')/items({1})", listTitle,id);
            var endpointUri = new Uri(newUri);
            var result = webClient.DownloadString(endpointUri);
            var t = JToken.Parse(result);
            return t["d"];
        }

        private string GetFormDigest()
        {
            string newUri = webUri + "/_api/contextinfo";
            var endpointUri = new Uri(newUri);
            var result = webClient.UploadString(endpointUri, "POST");
            JToken t = JToken.Parse(result);
            return t["d"]["GetContextWebInformation"]["FormDigestValue"].ToString();
        }

        public void AddContact(object payload,string listTitle)
        {
            var formDigestValue = GetFormDigest();
            webClient.Headers.Add("X-RequestDigest", formDigestValue);
            string newUri = webUri + string.Format("_api/web/lists/getbytitle('{0}')/items", listTitle);
            var endpointUri = new Uri(newUri);
            var payloadString = JsonConvert.SerializeObject(payload);
            webClient.UploadString(endpointUri, "POST", payloadString);
        }

        public void DeleteContact(string listTitle, int id)
        {
            var formDigestValue = GetFormDigest();
            webClient.Headers.Add("X-RequestDigest", formDigestValue);
            webClient.Headers.Add("X-HTTP-Method", "DELETE");
            webClient.Headers.Add("IF-MATCH", "*");
            string newUri = webUri + string.Format("/_api/web/lists/getbytitle('{0}')/items({1})", listTitle, id);
            var endpointUri = new Uri(newUri);
            webClient.Headers.Add(HttpRequestHeader.ContentType, "application/json;odata=verbose");
            webClient.UploadString(endpointUri, "POST", String.Empty);
        }

        public void EditContact(object payload,string listTitle,int id)
        {
            var formDigestValue = GetFormDigest();
            webClient.Headers.Add("X-RequestDigest", formDigestValue);
            webClient.Headers.Add("X-HTTP-Method", "MERGE");
            webClient.Headers.Add("IF-MATCH", "*");
            string newUri = webUri + string.Format("/_api/web/lists/getbytitle('{0}')/items({1})", listTitle, id);
            var endpointUri = new Uri(newUri);
            var payloadString = JsonConvert.SerializeObject(payload);
            webClient.Headers.Add(HttpRequestHeader.ContentType, "application/json;odata=verbose");
            Console.WriteLine(webClient.UploadString(endpointUri, "POST", payloadString));
        }
    }
}
