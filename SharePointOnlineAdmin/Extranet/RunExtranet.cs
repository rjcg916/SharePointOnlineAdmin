using Microsoft.SharePoint.Client;
using System;
using System.Configuration;

namespace SharePointOnlineAdmin
{
    static public class RunExtranet
    {
        static public void AddLibrariesToSite(string siteName, string libraryList)
        {
            string serviceUserName = ConfigurationManager.AppSettings["serviceUserName"];
            string servicePassword = ConfigurationManager.AppSettings["servicePassword"];
            string extranetUrl = ConfigurationManager.AppSettings["extranetUrl"];

            string siteUrl = extranetUrl + siteName;

            Console.WriteLine($"Adding Libraries: {libraryList} to Site: {siteUrl}");

            OfficeDevPnP.Core.AuthenticationManager authenticationMgr = new OfficeDevPnP.Core.AuthenticationManager();

            ClientContext clientContext = authenticationMgr.GetSharePointOnlineAuthenticatedContextTenant(siteUrl, serviceUserName, servicePassword);
            Web web = clientContext.Web;

            ProvisionExtranet.AddLibrariesToSite(web, libraryList);


        }
        static public void CreateSiteAndLibraries(string siteName, string libraryList)
        {
            string serviceUserName = ConfigurationManager.AppSettings["serviceUserName"];
            string servicePassword = ConfigurationManager.AppSettings["servicePassword"];
            string extranetUrl = ConfigurationManager.AppSettings["extranetUrl"];

            Console.WriteLine($"Creating Site:{siteName} and Libraries:{libraryList} . . .");

            OfficeDevPnP.Core.AuthenticationManager authenticationMgr = new OfficeDevPnP.Core.AuthenticationManager();

            ClientContext clientContext = authenticationMgr.GetSharePointOnlineAuthenticatedContextTenant(extranetUrl, serviceUserName, servicePassword);
            Web web = clientContext.Web;

            ProvisionExtranet.CreateSiteAndLibraries(web, siteName, libraryList);

        }
    }
}
