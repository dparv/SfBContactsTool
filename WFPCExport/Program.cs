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
            System.Threading.Thread.Sleep(1000);
            CheckSignedIn(_client.State);
            Console.WriteLine("Running export function...");
            ExportContacts(_client);
            string filePath = Directory.GetCurrentDirectory() + "BuddyList.xml";
            Console.WriteLine("Export completed in file: "+filePath);
        }
        public static void ExportContacts(LyncClient _client)
        {
            XmlTextWriter writer = new XmlTextWriter("BuddyList.xml", System.Text.Encoding.UTF8);
            writer.WriteStartDocument(true);
            writer.WriteStartElement("LyncContacts");
            foreach (Group group in _client.ContactManager.Groups)
            {
                //Do not export groups with 0 contacts in them
                if (group.Count == 0)
                {
                    Console.WriteLine("Skipping empty group: "+group.Name);
                    continue;
                }
                Console.WriteLine("Found group: '"+group.Name+"' of type '"+group.Type+"' with "+ group.Count+ " contacts");
                writer.WriteStartElement("ContactGroup");
                writer.WriteAttributeString("type", group.Type.ToString());
                writer.WriteStartElement("Name");
                writer.WriteCData(group.Name);
                writer.WriteEndElement();
                //Opening Contatacts XML element
                writer.WriteStartElement("Contacts");
                foreach (Contact contact in group)
                {
                    Console.WriteLine(contact.Uri);
                    //Opening Single Contatact XML element
                    writer.WriteStartElement("Contact");
                    writer.WriteRaw(contact.Uri);
                    writer.WriteEndElement();
                    //Closing Single Contatact XML element
                }
                writer.WriteEndElement();
                writer.WriteEndElement();
                //Closing Contatacts XML element
            }
            writer.WriteEndElement();
            writer.WriteEndDocument();
            writer.Flush();
            writer.Close();
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
                Console.WriteLine("Client Found and Signed In - Continuing to Export");
            }
        }
    }
}
