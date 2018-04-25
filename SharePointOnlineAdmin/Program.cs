using Microsoft.SharePoint.Client;
using System;
using System.Configuration;


namespace SharePointOnlineAdmin
{

    class Program
    {

        static void TestIt()
        {
            Console.WriteLine("Testing . . . ");

            string serviceUserName = ConfigurationManager.AppSettings["serviceUserName"];
            string servicePassword = ConfigurationManager.AppSettings["servicePassword"];
            string partnerSiteUrl = ConfigurationManager.AppSettings["partnerUrl"]; ;

            OfficeDevPnP.Core.AuthenticationManager authenticationMgr = new OfficeDevPnP.Core.AuthenticationManager();

            var clientContext = authenticationMgr.GetSharePointOnlineAuthenticatedContextTenant(partnerSiteUrl, serviceUserName, servicePassword);
            Web web = clientContext.Web;

            //Group aGroup = web.SiteGroups.GetByName("Bob Inc Members");

            //// give partner access to Discussion Forum w/o ability to create a new topic
            //List discussionList = web.Lists.GetByTitle(Partner.Configuration.PartnerDiscussionListName);
            //discussionList.AddPrincipalToAllFolders(aGroup);


        }

   
        static void Main(string[] args)
        {

            string command = args[0];

            string siteName;
            string libraryList;
            string CSVFilePath;
            string userLogonName;

            switch (command)
            {
                case "ShowMembers":
                    userLogonName = args[1];
                    UserGroups.RunShowMembership(userLogonName);
                    break;
                case "ECreate":
                    siteName = args[1];
                    libraryList = args[2];
                    Extranet.RunCreateSiteAndLibraries(siteName, libraryList);
                    break;
                case "EAdd":
                    siteName = args[1];
                    libraryList = args[2];
                    Extranet.RunAddLibrariesToSite(siteName, libraryList);
                    break;
                case "PAdd":
                    CSVFilePath = args[1];
                    Partner.RunAddPartners(CSVFilePath);
                    break;
                case "PDisplay":
                    CSVFilePath = args[1];
                    Partner.RunDisplayPartners(CSVFilePath);
                    break;
                case "PRemove":
                    CSVFilePath = args[1];
                    Partner.RunRemovePartners(CSVFilePath);
                    break;
                case "TestIt":
                    TestIt();
                    break;
                default:
                    Console.WriteLine("Unrecognized command!");
                    break;
            }

            Console.WriteLine("Press Enter to end");
            Console.ReadLine();
        }


    }
}
