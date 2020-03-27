using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddressBookUsingCSOM
{
    public class UserActions
    {
        ClientContext clientContext;
        CommonMethods commonMethods = new CommonMethods();

        public UserActions(ClientContext clientContext)
        {
            this.clientContext = clientContext;
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
            List list = clientContext.Web.Lists.GetByTitle("Address Book");
            CamlQuery query = new CamlQuery();
            ListItemCollection contacts = list.GetItems(query);

            clientContext.Load(list);
            clientContext.Load(contacts);
            clientContext.ExecuteQuery();
            
            foreach (ListItem contact in contacts)
            {
                Console.WriteLine("------------------------------------------------------------------------------------------------------------------------");
                Console.WriteLine("Id : "+contact["Title"] + ". | Name : " + contact["ContactName"]+ " | Email : "+contact["Email"]
                    +" \t | Mobile : "+contact["Mobile"]+" \t | Landline : "+contact["Landline"]+
                    " | Website :  "+contact["Website"]+" | Address : "+contact["Address"]);
            }

            Console.WriteLine("------------------------------------------------------------------------------------------------------------------------");
            int choice = commonMethods.ReadInt("Enter the action that you want to perform : 1. View Contact 2. Continue 3.Exit >> ");
            
            switch (choice)
            { 
                case 1: ViewContact(contacts);
                    break;

                case 2: ShowUserActions();
                    break;

                case 3: Environment.Exit(0);
                    break;
            }
        }

        public void AddContact()
        {
            List list = clientContext.Web.Lists.GetByTitle("Address Book");
            CamlQuery query = new CamlQuery();
            query.ViewXml = "<View/>";
            ListItemCollection items = list.GetItems(query);
            clientContext.Load(list);
            clientContext.Load(items);
            clientContext.ExecuteQuery();

            ListItem lastItem = items[items.Count()-1];

            ListItemCreationInformation newItem = new ListItemCreationInformation();
            ListItem newContact = list.AddItem(newItem);
            newContact["Title"] = (Int32.Parse(lastItem["Title"].ToString())+1).ToString();
            newContact["ContactName"] = commonMethods.ReadString("Enter name of the contact : ");
            string email = commonMethods.ReadString("Enter email id of the contact : ");
            bool isValidEmail = commonMethods.IsValidEmail(email);
            while (!isValidEmail)
            {
                email = commonMethods.ReadString("Please enter a valid email : ");
                isValidEmail = commonMethods.IsValidEmail(email);
            }
            newContact["Email"] = email;

            string mobile = commonMethods.ReadString("Enter new mobile number of contact : ");
            bool isValidMobileNumber = commonMethods.IsValidPhoneNumber(mobile);
            while (!isValidMobileNumber)
            {
                mobile = commonMethods.ReadString("Please enter a valid mobile number : ");
                isValidMobileNumber = commonMethods.IsValidPhoneNumber(mobile);
            }
            newContact["Mobile"] = mobile;

            Console.Write("Enter landline number of contact : ");
            string landline = Console.ReadLine();
            newContact["Landline"] = landline;

            Console.Write("Enter website of the contact : ");
            string website = Console.ReadLine();
            newContact["Website"] = website;

            Console.Write("Enter address of the contact : ");
            string address = Console.ReadLine();
            newContact["Address"] = address;

            newContact.Update();
            clientContext.ExecuteQuery();
            Console.WriteLine("Contact added successfully!!");
            FurtherActions();
        }

        public void EditContact(ListItem contact)
        {
            Console.Write("Enter the field that you want to edit : ");
            int selectedField = Int32.Parse(Console.ReadLine());

            switch (selectedField)
            {
                case 1:
                    string name = commonMethods.ReadString("Enter new name for the contact : ");
                    contact["ContactName"] = name;
                    contact.Update();
                    break;

                case 2:
                    string email = commonMethods.ReadString("Enter new email of the contact : ");
                    bool isValidEmail = commonMethods.IsValidEmail(email);
                    while (!isValidEmail)
                    {
                        email = commonMethods.ReadString("Please enter a valid email : ");
                        isValidEmail = commonMethods.IsValidEmail(email);
                    }

                    contact["Email"] = email;
                    contact.Update();
                    break;

                case 3:
                    string mobile = commonMethods.ReadString("Enter new mobile number of contact : ");
                    bool isValidMobileNumber = commonMethods.IsValidPhoneNumber(mobile);
                    while (!isValidMobileNumber)
                    {
                        mobile = commonMethods.ReadString("Please enter a valid mobile number : ");
                        isValidMobileNumber = commonMethods.IsValidPhoneNumber(mobile);
                    }

                    contact["Mobile"] = mobile;
                    contact.Update();
                    break;

                case 4:
                    string landline = commonMethods.ReadString("Enter new landline number of contact : ");
                    contact["Landline"] = landline;
                    contact.Update();
                    break;

                case 5:
                    string website = commonMethods.ReadString("Enter new website of contact : ");
                    contact["Website"] = website;
                    contact.Update();
                    break;

                case 6:
                    string address = commonMethods.ReadString("Enter new address of contact : ");
                    contact["Address"] = address;
                    contact.Update();
                    break;
            }
            clientContext.ExecuteQuery();
            Console.WriteLine("Contact edited successfully!!");
            FurtherActions();
        }

        public void DeleteContact(ListItem contact)
        {
            int choice = commonMethods.ReadInt("Do you want to delete the contact ? 1. Yes 2. No >> ");
            choice = commonMethods.IsValidChoice(choice);
            if (choice == 1)
            {
                contact.DeleteObject();
                clientContext.ExecuteQuery();
                Console.WriteLine("Contact deleted successfully!!");
            }
            FurtherActions();
        }

        public void ViewContact(ListItemCollection contacts)
        {
            Console.WriteLine("Enter the id of the contact that you want to view : ");
            int selectedId = Int32.Parse(Console.ReadLine());
            ListItem selectedContact = null;
            foreach(ListItem contact in contacts)
            {
                if (Int32.Parse(contact["Title"].ToString()) == selectedId)
                {
                    selectedContact = contact;
                }
            }
            
            Console.WriteLine("Current Details of the contact are : ");
            Console.WriteLine("1. Name : " + selectedContact["ContactName"] + " 2. Email : " + selectedContact["Email"]);
            Console.WriteLine("3. Mobile : " + selectedContact["Mobile"] + " 4. Landline : " + selectedContact["Landline"]);
            Console.WriteLine("5. Website : " + selectedContact["Website"] + " 6. Address : " + selectedContact["Address"]);

            int choice = commonMethods.ReadInt("Enter the action you want to perform : 1. Edit Contact 2. Delete Contact 3. Continue >> ");
            while(choice>3 || choice < 1)
            {
                choice = commonMethods.ReadInt("Please enter a valid action : 1. Edit Contact 2. Delete Contact 3. Continue >> ");
            }
            switch (choice)
            {
                case 1: EditContact(selectedContact);
                    break;

                case 2: DeleteContact(selectedContact);
                    break;

                case 3: ShowUserActions();
                    break;
            }
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
