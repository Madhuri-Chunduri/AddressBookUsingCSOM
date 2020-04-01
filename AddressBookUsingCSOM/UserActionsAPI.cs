using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client.Taxonomy;

namespace AddressBookUsingCSOM
{
    class UserActionsAPI
    {
        CommonMethods commonMethods = new CommonMethods();
        string listTitle = "AddressBook1";

        public void GetAllItems()
        {
            string url = "https://intern2k20.sharepoint.com/sites/Technovert/_api/web/lists";
            HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(url);
            endpointRequest.Method = "GET";
            endpointRequest.Accept = "application/json;odata=verbose";
            string password = "Intern1567";
            SecureString secureStringPassword = new SecureString();
            foreach (char c in password)
            {
                secureStringPassword.AppendChar(c);
            }
            NetworkCredential credentials = new NetworkCredential("madhuri@intern2k20.onmicrosoft.com", secureStringPassword, "intern2k20.sharepoint.com");
            //SharePointOnlineCredentials cred = new SharePointOnlineCredentials("madhuri@intern2k20.onmicrosoft.com", secureStringPassword);
            endpointRequest.Credentials = credentials;
            HttpWebResponse endpointResponse = (HttpWebResponse)endpointRequest.GetResponse();
            try
            {
                WebResponse webResponse = endpointRequest.GetResponse();
                Stream webStream = webResponse.GetResponseStream();
                StreamReader responseReader = new StreamReader(webStream);
                string response = responseReader.ReadToEnd();
                JObject jobj = JObject.Parse(response);
                JArray jarr = (JArray)jobj["d"]["results"];
                foreach (JObject j in jarr)
                {
                    Console.WriteLine(j["Title"] + " " + j["Body"]);
                }
                responseReader.Close();
                Console.ReadLine();
            }
            catch (Exception e)
            {
                Console.Out.WriteLine(e.Message); Console.ReadLine();
            }
        }

        public static JToken GetResult(string webUri, ICredentials credentials)
        {
            using (var client = new WebClient())
            {
                client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                client.Credentials = credentials;
                client.Headers.Add(HttpRequestHeader.ContentType, "application/json;odata=verbose");
                client.Headers.Add(HttpRequestHeader.Accept, "application/json;odata=verbose");
                var endpointUri = new Uri(webUri);
                var result = client.DownloadString(endpointUri);
                var t = JToken.Parse(result);
                return t["d"]["results"];
            }
        }

        public void ShowUserActions()
        {
            Console.WriteLine("Enter the action you want to perform ");
            Console.WriteLine("1. View all contacts 2. Add a contact 3.Exit");
            int choice = Int32.Parse(Console.ReadLine());
            switch (choice)
            {
                case 1:
                    ViewAllContacts();
                    break;

                case 2:
                    AddContact();
                    break;

                case 3:
                    Environment.Exit(0);
                    break;
            }
        }

        public void ViewAllContacts()
        {
            ContactService contactService = new ContactService();
            var list = contactService.GetAllContacts(listTitle);
            foreach (JObject contact in list)
            {
                ContactService service = new ContactService();
                var contactName = service.GetContactName(Int32.Parse(contact["Id"].ToString()),listTitle);
                ContactService contactService2 = new ContactService();
                var contactGroup = contact["Group"]["Label"];
                var groupValue = contactService2.GetGroupLabel(Int32.Parse(contactGroup.ToString()),listTitle);
                Console.WriteLine("------------------------------------------------------------------------------------------------------------------------");
                Console.WriteLine("Id : " + contact["Title"] + " Name : "+ contactName["Title"] + " | Email : " + contact["Email"]
                    + " \t | Mobile : " + contact["Mobile"] + " \t | Landline : " + contact["Landline"] +
                    " | Website :  " + contact["Website"] + " | Address : " + contact["Address"] + " | Group : " + groupValue["Term"]);
            }
            Console.WriteLine("------------------------------------------------------------------------------------------------------------------------");
            int choice = commonMethods.ReadInt("Enter the action that you want to perform : 1. Continue 2. View Contact 3. Exit : ");
            while(choice>3 || choice< 0)
            {
                choice = commonMethods.ReadInt("Please enter a valid choice : ");
            }
            if (choice == 1)
            {
                FurtherActions();
            }
            else if (choice == 2)
            {
                int title = commonMethods.ReadInt("Enter the id of the contact that you want to view : ");
                foreach(JToken contact in list)
                {
                    if (contact["Title"].ToString() == title.ToString())
                        ViewSelectedContact(Int32.Parse(contact["Id"].ToString()));
                }
            }
            else Environment.Exit(0);
            Console.Read();
        }

        public void ViewSelectedContact(int id)
        {
            ContactService service = new ContactService();
            var selectedContact = service.GetContactById(listTitle, id);
            ContactService contactService = new ContactService();
            var contactName = contactService.GetContactName(Int32.Parse(selectedContact["Id"].ToString()), listTitle);
            var groupObject = selectedContact["Group"];
            var group = groupObject.ToObject<TaxonomyFieldValue>();
            Console.WriteLine("Current details of the contact are : ");
            Console.WriteLine("1. Name : " + contactName["Title"] + " 2. Email : " + selectedContact["Email"]);
            Console.WriteLine("3. Mobile : " + selectedContact["Mobile"] + " 4. Landline : " + selectedContact["Landline"]);
            Console.WriteLine("5. Website : " + selectedContact["Website"] + " 6. Address : " + selectedContact["Address"]);
            Console.WriteLine("7. Group : " + group.Label);
            int choice = commonMethods.ReadInt("Enter the action you want to perform : 1. Edit Contact 2. Delete Contact 3. Continue >> ");
            while (choice > 3 || choice < 1)
            {
                choice = commonMethods.ReadInt("Please enter a valid action : 1. Edit Contact 2. Delete Contact 3. Continue >> ");
            }
            switch (choice)
            {
                case 1:
                    EditContact(selectedContact);
                    break;

                case 2:
                    DeleteContact(id);
                    break;

                case 3:
                    ShowUserActions();
                    break;
            }
        }

        public void AddContact()
        {
            string title = "3";
            string name = commonMethods.ReadString("Enter name of the contact : ");

            string email = commonMethods.ReadString("Enter email id of the contact : ");
            bool isValidEmail = commonMethods.IsValidEmail(email);
            while (!isValidEmail)
            {
                email = commonMethods.ReadString("Please enter a valid email : ");
                isValidEmail = commonMethods.IsValidEmail(email);
            }

            string mobile = commonMethods.ReadString("Enter mobile number of contact : ");
            bool isValidMobileNumber = commonMethods.IsValidPhoneNumber(mobile);
            while (!isValidMobileNumber)
            {
                mobile = commonMethods.ReadString("Please enter a valid mobile number : ");
                isValidMobileNumber = commonMethods.IsValidPhoneNumber(mobile);
            }

            Console.Write("Enter landline number of contact : ");
            string landline = Console.ReadLine();

            Console.Write("Enter website of the contact : ");
            string website = Console.ReadLine();

            Console.Write("Enter address of the contact : ");
            string address = Console.ReadLine();
            var newContact = new
            {
                Title = title,
                Email = email,
                Mobile = mobile,
                Landline = landline,
                Address = address
            };
            ContactService contactService = new ContactService();
            contactService.AddContact(newContact, listTitle);
            FurtherActions();
        }

        public void EditContact(JToken selectedContact)
        {
        }

        public void DeleteContact(int id)
        {
            int choice = commonMethods.ReadInt("Do you want to delete the contact ? 1. Yes 2. No >> ");
            choice = commonMethods.IsValidChoice(choice);
            if (choice == 1)
            {
                ContactService contactService = new ContactService();
                contactService.DeleteContact(listTitle,id);
                Console.WriteLine("Contact deleted successfully!!");
            }
            FurtherActions();
        }

        public void FurtherActions()
        {
            int choice = commonMethods.ReadInt("Do you want to continue ? 1. Continue 2. Exit >> ");
            choice = commonMethods.IsValidChoice(choice);
            if (choice == 1)
            {
                ShowUserActions();
            }
            else Environment.Exit(0);
        }
    }
}