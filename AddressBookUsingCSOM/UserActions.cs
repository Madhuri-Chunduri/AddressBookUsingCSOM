using Microsoft.SharePoint;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace AddressBookUsingCSOM
{
    public class UserActions
    {
        ClientContext clientContext;
        CommonMethods commonMethods = new CommonMethods();
        string listName = "AddressBook2";
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
            List list = null;
            if (!clientContext.Web.ListExists(listName))
            {
                //var listCreationInformationInfo = new ListCreationInformation
                //{
                //   Title = "AddressBook1",
                //   Description = "AddressBook list created from CSOM"
                //};
                list = CreateList();
            }
            else
            {
                list = clientContext.Web.Lists.GetByTitle(listName);
            }
            CamlQuery query = new CamlQuery();
            ListItemCollection contacts = list.GetItems(query);

            clientContext.Load(list);
            clientContext.Load(contacts);
            clientContext.ExecuteQuery();

            if (contacts.Count() > 0)
            {
                foreach (ListItem contact in contacts)
                {
                    FieldUserValue contactName = contact["ContactName"] as FieldUserValue;
                    Microsoft.SharePoint.Client.Taxonomy.TaxonomyFieldValue taxFieldValue = contact["Group"] as Microsoft.SharePoint.Client.Taxonomy.TaxonomyFieldValue;
                    Console.WriteLine("------------------------------------------------------------------------------------------------------------------------");
                    Console.WriteLine("Id : " + contact["Title"] + ". | Name : " + contactName.LookupValue.ToString() + " | Email : " + contact["Email"]
                        + " \t | Mobile : " + contact["Mobile"] + " \t | Landline : " + contact["Landline"] +
                        " | Website :  " + contact["Website"] + " | Address : " + contact["Address"]+ " | Group : "+taxFieldValue.Label);
                }

                Console.WriteLine("------------------------------------------------------------------------------------------------------------------------");
                int choice = commonMethods.ReadInt("Enter the action that you want to perform : 1. View Contact 2. Continue 3.Exit >> ");

                switch (choice)
                {
                    case 1:
                        ViewContact(contacts);
                        break;

                    case 2:
                        ShowUserActions();
                        break;

                    case 3:
                        Environment.Exit(0);
                        break;
                }
            }
            else
            {
                Console.WriteLine("There are no contacts to view!");
                int choice = commonMethods.ReadInt("Do you want to add a new contact ? 1.Add 2. Exit >> ");
                choice = commonMethods.IsValidChoice(choice);
                if (choice == 1)
                {
                    AddContact();
                }
                Environment.Exit(0);
            }
        }

        public void AddContact()
        {
            List list = null;
            if (!clientContext.Web.ListExists(listName))
            {
                list = CreateList();
            }
            else
            {
                list = clientContext.Web.GetListByTitle(listName);
            }
            CamlQuery query = new CamlQuery();
            query.ViewXml = "<View/>";
            ListItemCollection items = list.GetItems(query);

            clientContext.Load(list);
            clientContext.Load(items);
            clientContext.ExecuteQuery();
            Web web = clientContext.Web;
            clientContext.ExecuteQuery();

            var users = web.SiteUsers;
            clientContext.Load(users);
            clientContext.ExecuteQuery();
            string title;

            if (items.Count()==0)
            {
                title = "0";
            }
            else
            {
                ListItem lastItem = items[items.Count() - 1];
                title = lastItem["Title"].ToString();
            }
            

            ListItemCreationInformation newItem = new ListItemCreationInformation();
            ListItem newContact = list.AddItem(newItem);
            newContact["Title"] = (Int32.Parse(title) + 1).ToString();
            string name = commonMethods.ReadString("Enter name of the contact : ");
            //PeopleManager peopleManager = new PeopleManager(clientContext);
            //PersonProperties personProperties = peopleManager.GetPropertiesFor(name);
            //clientContext.Load(personProperties, p => p.AccountName, p => p.Email, p => p.DisplayName);
            //clientContext.ExecuteQuery();

            //SPSite spSite = new SPSite("https://intern2k20.sharepoint.com/sites/Technovert");
            //SPWeb web = spSite.OpenWeb();
            //SPUser webUser = web.EnsureUser(name);
            //SPFieldUserValue value = new SPFieldUserValue(web, webUser.ID, webUser.Name);

            newContact["ContactName"] = users.First(obj => obj.Title == name);

            string email = commonMethods.ReadString("Enter email id of the contact : ");
            bool isValidEmail = commonMethods.IsValidEmail(email);
            while (!isValidEmail)
            {
                email = commonMethods.ReadString("Please enter a valid email : ");
                isValidEmail = commonMethods.IsValidEmail(email);
            }
            newContact["Email"] = email;

            string mobile = commonMethods.ReadString("Enter mobile number of contact : ");
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
            if (UpsertGroup(newContact))
            {
                Console.WriteLine("Contact added successfully!!");
            }
            else
            {
                Console.WriteLine("Failed to assign to group");
                Console.WriteLine("Failed to insert new contact");
            }
            FurtherActions();
        }

        public void EditContact(ListItem contact)
        { 
            Console.Write("Enter the field that you want to edit : ");
            int selectedField = Int32.Parse(Console.ReadLine());
            List contacts = clientContext.Web.Lists.GetByTitle(listName);
            Web web = clientContext.Web;
            clientContext.Load(web);
            clientContext.ExecuteQuery();

            FieldCollection fields = contacts.Fields;
            var users = web.SiteUsers;
            clientContext.Load(users);
            clientContext.Load(fields);
            clientContext.ExecuteQuery();

            switch (selectedField)
            {
                case 1:
                    var nameField = fields.GetFieldByInternalName("ContactName");
                    nameField.ReadOnlyField = false;
                    nameField.Update();
                    string name = commonMethods.ReadString("Enter new name for the contact : ");
                    //var users = clientContext.LoadQuery(clientContext.Web.SiteUsers.Where(u => u.PrincipalType == PrincipalType.User && u.UserId.NameIdIssuer == "urn:federation:microsoftonline"));
                    //contact["ContactName"] = users.First(obj=> obj.Email.ToString()==contact["Email"].ToString()).Id;

                    //var user = clientContext.LoadQuery(web.SiteUsers.Where(obj => obj.LoginName == name));
                    //clientContext.ExecuteQuery();

                    //contact["ContactName"] = user;
                    //PeopleManager peopleManager = new PeopleManager(clientContext);
                    //PersonProperties personProperties = peopleManager.GetPropertiesFor(name);
                    //clientContext.Load(personProperties, p => p.AccountName, p => p.Email, p => p.DisplayName);
                    //clientContext.ExecuteQuery();
                     
                    contact["ContactName"] = users.First(obj => obj.Title == name);
                    contact.Update();

                    clientContext.ExecuteQuery();
                    nameField.ReadOnlyField = true;
                    nameField.Update();
                    
                    //web.Update();
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

                case 7:
                    //SPList list = new SPList("TemporaryList");
                    //SPListItem selectedContact = list.AddItem();
                    //selectedContact["Group"] = contact["Group"];
                    //Guid Promotypeid = selectedContact.Fields["Group"].Id;
                    //TaxonomyField taxfield = selectedContact.Fields[Promotypeid] as TaxonomyField;
                    //SPSite site = new SPSite("https://intern2k20.sharepoint.com/sites/Technovert");
                    if (!UpsertGroup(contact))
                    {
                        Console.WriteLine("Failed to assign to group!!");
                        Console.WriteLine("Failed to update contact!!");
                    }
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
            
            FieldUserValue contactName = selectedContact["ContactName"] as FieldUserValue;
            Microsoft.SharePoint.Client.Taxonomy.TaxonomyFieldValue group = selectedContact["Group"] as Microsoft.SharePoint.Client.Taxonomy.TaxonomyFieldValue;
            Console.WriteLine("Current Details of the contact are : ");
            Console.WriteLine("1. Name : " + contactName.LookupValue.ToString() + " 2. Email : " + selectedContact["Email"]);
            Console.WriteLine("3. Mobile : " + selectedContact["Mobile"] + " 4. Landline : " + selectedContact["Landline"]);
            Console.WriteLine("5. Website : " + selectedContact["Website"] + " 6. Address : " + selectedContact["Address"]);
            Console.WriteLine("7. Group : " + group.Label);
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

        public List CreateList()
        {
            List list = clientContext.Web.CreateList(ListTemplateType.GenericList, listName, false, true, string.Empty, true);
            string schemaUserField = "<Field Type='User' Name='ContactName' StaticName='ContactName' DisplayName='ContactName' />";
            Field userField = list.Fields.AddFieldAsXml(schemaUserField, true, AddFieldOptions.AddFieldInternalNameHint);

            string emailField = "<Field Type='Text' Name='Email' StaticName='Email' DisplayName='Email' />";
            Field emailTextField = list.Fields.AddFieldAsXml(emailField, true, AddFieldOptions.AddToDefaultContentType);

            string mobileField = "<Field Type='Text' Name='Mobile' StaticName='Mobile' DisplayName='Mobile' />";
            Field mobileTextField = list.Fields.AddFieldAsXml(mobileField, true, AddFieldOptions.AddToDefaultContentType);

            string landlineField = "<Field Type='Text' Name='Landline' StaticName='Landline' DisplayName='Landline' />";
            Field landlineTextField = list.Fields.AddFieldAsXml(landlineField, true, AddFieldOptions.AddToDefaultContentType);

            string websiteField = "<Field Type='Text' Name='Website' StaticName='Website' DisplayName='Website' />";
            Field websiteTextField = list.Fields.AddFieldAsXml(websiteField, true, AddFieldOptions.AddToDefaultContentType);

            string addressField = "<Field Type='Text' Name='Address' StaticName='Address' DisplayName='Address' />";
            Field addressTextField = list.Fields.AddFieldAsXml(addressField, true, AddFieldOptions.AddToDefaultContentType);
            
            string groupField = "<Field Type='TaxonomyFieldType' Name='Group' StaticName='Group' DisplayName = 'Group' /> ";
            Field groupMetadataField = list.Fields.AddFieldAsXml(groupField, true, AddFieldOptions.AddFieldInternalNameHint);
           
            Guid termStoreId = Guid.Empty;
            Guid termSetId = Guid.Empty;
            GetTaxonomyFieldInfo(clientContext, out termStoreId, out termSetId);

            TaxonomyField taxonomyField = clientContext.CastTo<TaxonomyField>(groupMetadataField);
            taxonomyField.SspId = termStoreId;
            taxonomyField.TermSetId = termSetId;
            taxonomyField.TargetTemplate = String.Empty;
            taxonomyField.AnchorId = Guid.Empty;
            taxonomyField.Update();
            clientContext.ExecuteQuery();
            return list;
        }

        public bool UpsertGroup(ListItem contact)
        {
            try
            {
                List contacts = clientContext.Web.Lists.GetByTitle(listName);
                clientContext.ExecuteQuery();
                FieldCollection fields = contacts.Fields;
                clientContext.Load(fields);
                TaxonomyField taxonomyField = fields.GetFieldByInternalName("Group") as TaxonomyField;
                TaxonomySession session = TaxonomySession.GetTaxonomySession(clientContext);
                TermStore mytermstore = session.GetDefaultKeywordsTermStore();
                TermSet termSet = mytermstore.GetTermSet(taxonomyField.TermSetId);
                TermCollection terms = termSet.Terms;
                clientContext.Load(terms);
                clientContext.ExecuteQuery();
                int count = 1;
                foreach (Term term in terms)
                {
                    Console.WriteLine(count + " " + term.Name);
                    count += 1;
                }
                int selectedChoice = commonMethods.ReadInt("Enter the group that you want to assign from above list : ");
                while (selectedChoice > count || selectedChoice < 1)
                {
                    selectedChoice = commonMethods.ReadInt("Please enter a valid choice : ");
                }
                string selectedGroup = terms[selectedChoice - 1].Name;
                Guid termGuid = Guid.Empty;
                foreach (Term term in terms)
                {
                    if (term.Name.Equals(selectedGroup, StringComparison.OrdinalIgnoreCase))
                    {
                        termGuid = term.Id;
                        break;
                    }
                }
                if (termGuid != Guid.Empty)
                {
                    Term customTerm = termSet.GetTerm(termGuid);
                    clientContext.Load(customTerm);
                    clientContext.ExecuteQuery();
                    string taxFieldInternalname = "Group";
                    contact[taxFieldInternalname] = customTerm.Name + "|" + customTerm.Id.ToString();
                    contact.Update();
                    return true;
                }
                return false;
            }
            catch(Exception exception)
            {
                Console.WriteLine(exception.StackTrace);
                return false;
            }
        }

        private void GetTaxonomyFieldInfo(ClientContext clientContext, out Guid termStoreId, out Guid termSetId)
        {
            termStoreId = Guid.Empty;
            termSetId = Guid.Empty;

            TaxonomySession session = TaxonomySession.GetTaxonomySession(clientContext);
            TermStore termStore = session.GetDefaultSiteCollectionTermStore();

            TermSetCollection termSets = termStore.GetTermSetsByName("Group", 1033);

            clientContext.Load(termSets, tsc => tsc.Include(ts => ts.Id));
            clientContext.Load(termStore, ts => ts.Id);
            clientContext.ExecuteQuery();

            termStoreId = termStore.Id;
            termSetId = termSets.FirstOrDefault().Id;
        }
    }
}
