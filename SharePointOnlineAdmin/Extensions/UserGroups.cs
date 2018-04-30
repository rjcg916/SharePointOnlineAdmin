using System;
using System.Configuration;
using Microsoft.SharePoint.Client;

namespace SharePointOnlineAdmin
{
    public class UserGroups
    {
        public static void RunShowMembership(string siteCollectionUrl, string userLogonName)
        {
            string serviceUserName = ConfigurationManager.AppSettings["serviceUserName"];
            string servicePassword = ConfigurationManager.AppSettings["servicePassword"];

            
            OfficeDevPnP.Core.AuthenticationManager authenticationMgr = new OfficeDevPnP.Core.AuthenticationManager();

            ClientContext clientContext = authenticationMgr.GetSharePointOnlineAuthenticatedContextTenant(siteCollectionUrl, serviceUserName, servicePassword);
            Web web = clientContext.Web;

            Console.WriteLine($"In {siteCollectionUrl}, {userLogonName} is in the following groups: ");
           
            GroupCollection collGroup = clientContext.Web.SiteGroups;
            clientContext.Load(collGroup);
            clientContext.ExecuteQuery();


            foreach (Group group in collGroup)
            {
               if ( web.IsUserInGroup(group.Title, userLogonName))
                        Console.WriteLine(group.Title);
            }


        }

    }
}
