using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Net;
using System.Security;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using System.Linq;
using System.Collections.Generic;

namespace ATC_POC_DeployGlobalNav
{
    class Program
    {
        static void Main(string[] args)
        {
            ConsoleColor defaultForeground = Console.ForegroundColor;

            // Collect information 
            string templateWebUrl = GetInput("Enter the URL of the template site: ", false, defaultForeground);
            string targetWebUrl = GetInput("Enter the URL of the target site: ", false, defaultForeground);
            string infrastructureUrl = GetInput("Enter the URL of the Infrastructure site with list: ", false, defaultForeground);
            string userName = GetInput("Enter your user name:", false, defaultForeground);
            string pwdS = GetInput("Enter your password:", true, defaultForeground);

            SecureString pwd = new SecureString();
            foreach (char c in pwdS.ToCharArray()) pwd.AppendChar(c);

            // GET the template from existing site and serialize
            // Serializing the template for later reuse is optional
            ProvisioningTemplate template = GetProvisioningTemplate(defaultForeground, templateWebUrl, userName, pwd);

            /* ----------------------------------------------------------------
             * Determine which site collections to apply global nav to based on Infrastructure site
             * list that maintains sites to be provisioned with global nav
             * 
             * ---------------------------------------------------------------------
             */
            Console.WriteLine("Determining which sites to apply global nav to now based on SP List");
            var listofSites = GetProvisioningSitesFromList(defaultForeground, infrastructureUrl, userName, pwd);

            Console.WriteLine("Go and update the Global Nav file now");
            Console.ReadLine();
            // APPLY the template to new site from 
           // ApplyProvisioningTemplate(defaultForeground, targetWebUrl, userName, pwd);

            // Pause and modify the UI to indicate that the operation is complete
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("We're done. Press Enter to continue.");
            Console.ReadLine();
        }

       
        private static List<GlobalNavSiteCollections> GetProvisioningSitesFromList(ConsoleColor defaultForeground, string infrastructureUrl, string userName, SecureString pwd)
        {
            var _sitesToDeploy = new List<GlobalNavSiteCollections>();

            using (var ctx = new ClientContext(infrastructureUrl))
            {
                
                ctx.Credentials = new SharePointOnlineCredentials(userName, pwd);
                ctx.RequestTimeout = Timeout.Infinite;

                // Assume the web has a list named "Announcements". 
                List provisioningList = ctx.Web.Lists.GetByTitle("GlobalNavSites");

                // This creates a CamlQuery that has a RowLimit of 100, and also specifies Scope="RecursiveAll" 
                // so that it grabs all list items, regardless of the folder they are in. 
                //CamlQuery query = CamlQuery.CreateAllItemsQuery(100);

                CamlQuery query = new CamlQuery();
                // Get only items with Deploy Navigation set to Yes choice
                query.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Deploy_x0020_Navigation'/><Value Type='Boolean'>1</Value></Eq></Where></Query></View>"; 
                ListItemCollection items = provisioningList.GetItems(query);

                // Retrieve all items in the ListItemCollection from List.GetItems(Query). 
                ctx.Load(items);
                ctx.ExecuteQueryRetry();
                foreach (ListItem listItem in items)
                {
                    _sitesToDeploy.Add(new GlobalNavSiteCollections()
                    {
                        SiteTitle = listItem["Title"].ToString(),
                        SiteURL = ((FieldUrlValue)(listItem["Site_x0020_Collection_x0020_URL"])).Url.ToString()
                    });

                }
                
            }

            // Return list of site collections to apply global nav to
            return _sitesToDeploy;
        }

        private static ProvisioningTemplate GetProvisioningTemplate(ConsoleColor defaultForeground, string webUrl, string userName, SecureString pwd)
        {
            using (var ctx = new ClientContext(webUrl))
            {
                // ctx.Credentials = new NetworkCredentials(userName, pwd);
                ctx.Credentials = new SharePointOnlineCredentials(userName, pwd);
                ctx.RequestTimeout = Timeout.Infinite;

                // Just to output the site details
                Web web = ctx.Web;
                ctx.Load(web, w => w.Title);
                ctx.ExecuteQueryRetry();

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("Your site title is: " + ctx.Web.Title);
                Console.ForegroundColor = defaultForeground;

                ProvisioningTemplateCreationInformation ptci
                        = new ProvisioningTemplateCreationInformation(ctx.Web);

                // Create FileSystemConnector to store a temporary copy of the template 
                ptci.FileConnector = new FileSystemConnector(@"c:\temp\pnpprovisioningdemo", "");
                ptci.PersistBrandingFiles = true;
                ptci.HandlersToProcess = Handlers.Navigation;
                ptci.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                {
                    // Only to output progress for console UI
                    Console.WriteLine("{0:00}/{1:00} - {2}", progress, total, message);
                };

                // Execute actual extraction of the template
                ProvisioningTemplate template = ctx.Web.GetProvisioningTemplate(ptci);

                // We can serialize this template to save and reuse it
                // Optional step 
                XMLTemplateProvider provider =
                        new XMLFileSystemTemplateProvider(@"c:\temp\pnpprovisioningdemo", "");
                provider.SaveAs(template, "PnPProvisioningDemo.xml");

                // Get Navigation only in Site Provisioning Template
                try
                {
                    // Load site provisioning template (Navigation only)
                    XDocument doc = XDocument.Load(@"c:\temp\pnpprovisioningdemo\PnPProvisioningDemo.xml");

                    // Get Current Navigation Nodes
                    var currNavNodes = from node in doc.Descendants(doc.Root.GetNamespaceOfPrefix("pnp") + "CurrentNavigation")
                                       select node;

                    // Remove Current Navigation
                    currNavNodes.Remove();

                    // Save new provisioning file with only Global Navigation in the Navigation Nodes
                    doc.Save(@"c:\temp\pnpprovisioningdemo\GlobalNav.xml");


                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine("There were no Navigation nodes found in the template for site: " + ctx.Web.Title);
                    Console.WriteLine("Error: " + ex.InnerException.ToString());
                    Console.ForegroundColor = defaultForeground;
                }


                return template;
            }
        }
        private static string GetInput(string label, bool isPassword, ConsoleColor defaultForeground)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("{0} : ", label);
            Console.ForegroundColor = defaultForeground;

            string value = "";

            for (ConsoleKeyInfo keyInfo = Console.ReadKey(true); keyInfo.Key != ConsoleKey.Enter; keyInfo = Console.ReadKey(true))
            {
                if (keyInfo.Key == ConsoleKey.Backspace)
                {
                    if (value.Length > 0)
                    {
                        value = value.Remove(value.Length - 1);
                        Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                        Console.Write(" ");
                        Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                    }
                }
                else if (keyInfo.Key != ConsoleKey.Enter)
                {
                    if (isPassword)
                    {
                        Console.Write("*");
                    }
                    else
                    {
                        Console.Write(keyInfo.KeyChar);
                    }
                    value += keyInfo.KeyChar;

                }

            }
            Console.WriteLine("");

            return value;
        }

    }
    public class GlobalNavSiteCollections
    {
        public string SiteTitle { get; set; }
        public string SiteURL { get; set; }
       
    }
}
