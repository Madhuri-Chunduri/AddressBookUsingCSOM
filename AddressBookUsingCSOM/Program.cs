using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace AddressBookUsingCSOM
{
    class Program
    {
        static void Main(string[] args)
        {
            string userName = "madhuri@intern2k20.onmicrosoft.com";
            //Console.Write("Enter your password : ");
            string password = "Intern1567";
            SecureString secureStringPassword = new SecureString();
            foreach (char c in password)
            {
                secureStringPassword.AppendChar(c);
            }
            using (var clientContext = new ClientContext("https://intern2k20.sharepoint.com/sites/Technovert"))
            {
                clientContext.Credentials = new SharePointOnlineCredentials(userName, secureStringPassword);
                Web web = clientContext.Web;
                clientContext.Load(web);
                clientContext.ExecuteQuery();

                UserActions userActions = new UserActions(clientContext);
                userActions.ShowUserActions();
            }
        }
    }
}
