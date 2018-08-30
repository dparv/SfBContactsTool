using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Windows.Forms;
using Microsoft.Lync.Model;
using Microsoft.Lync.Model.Group;

namespace WFP_Contacts
{
    class Program
    {
        public static LyncClient _client = LyncClient.GetClient();
        static void Main(string[] args)
        {
            Console.WriteLine("Checking client state...");
            CheckSignedIn(_client.State);
            Console.WriteLine("Running import function...");
            ImportContacts(_client);
            string filePath = Directory.GetCurrentDirectory() + "BuddyList.xml";
            Console.WriteLine("Import completed from file: " + filePath);
        }
        public static void ImportContacts(LyncClient _client)
        {
            ContactManager _contactManager = _client.ContactManager;
            XmlDocument doc = new XmlDocument();
            if (!File.Exists("BuddyList.xml"))
            {
                Console.WriteLine("BuddyList.xml does not exist.");
                Environment.Exit(0);
            }
            doc.Load("BuddyList.xml");
            XmlElement root = doc.DocumentElement;
            XmlNodeList groups = root.SelectNodes("ContactGroup");
            GroupCollection current_groups = _contactManager.Groups;
            var current_group_colelction = new List<string>();

            foreach (Group group in current_groups)
            {
                current_group_colelction.Add(group.Name);
            }

            foreach (XmlNode group in groups)
            {
                String group_name = group.ChildNodes.Item(0).FirstChild.Value;
                String group_type = group.Attributes.GetNamedItem("type").InnerText;
                if (group_type == "DistributionGroup")
                {
                    _contactManager.BeginSearch(
                        group_name,
                        (ar) =>
                    {
                        SearchResults searchResults = _contactManager.EndSearch(ar);
                        if (searchResults.Groups.Count > 0)
                        {
                            Console.WriteLine(searchResults.Groups.Count.ToString() + " found");
                            foreach (Group dg in searchResults.Groups)
                            {
                                Console.WriteLine(System.Environment.NewLine + dg.Name);
                                DistributionGroup dGroup = dg as DistributionGroup;
                                if (!current_groups.Contains(dGroup))
                                {
                                    if (group_name == dGroup.Name) {
                                        _contactManager.BeginAddGroup(dGroup, null, null);
                                    }
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("No groups found for search on " + group_name);
                        }
                    }, 
                    null);
                }
                else if (group_type == "FavoriteContacts" || group_type == "FrequentContacts")
                {
                    //skip Favourites as it always exists
                    continue;
                }
                else
                {
                    //handle custom groups and favourites normally
                    //check first if group exists or error
                    if (!current_group_colelction.Contains(group_name))
                    {
                        //continue to adding new group
                        _contactManager.BeginAddGroup(group_name, null, null);
                        Console.WriteLine(group_name + " created.");
                    }
                    else
                    {
                        //show message and do not try to create group
                        Console.WriteLine(group_name + " already exists.");
                    }
                }
            }

            //add the contacts to each group of typer Custom or Favourites
            System.Threading.Thread.Sleep(2000);
            Console.WriteLine("Importing contacts...");
            foreach (XmlNode group in groups)
            {
                if (group.Attributes.GetNamedItem("type").InnerText != "DistributionGroup")
                {
                    foreach (Group igroup in current_groups)
                    {
                        if (igroup.Name == group.ChildNodes.Item(0).FirstChild.Value)
                        {
                            Console.WriteLine(igroup.Name + " - " + igroup.Type.ToString());
                            //add contacts
                            foreach (XmlNode contact in group.ChildNodes.Item(1).ChildNodes)
                            {
                                Console.WriteLine(contact.InnerText);
                                Contact c = _contactManager.GetContactByUri(contact.InnerText);
                                if (!igroup.Contains(c)) {
                                    igroup.BeginAddContact(c, null, null);
                                }
                            }
                        }
                    }
                }
            }
            Console.WriteLine("Finished import.");
        }
        public static void CheckSignedIn(ClientState state)
        {
            if (state != ClientState.SignedIn)
            {
                Console.WriteLine("Client not Signed In");
                Environment.Exit(0);
            }
            else
            {
                Console.WriteLine("Client Found and Signed In - Continuing to Import");
            }
        }
    }
}
