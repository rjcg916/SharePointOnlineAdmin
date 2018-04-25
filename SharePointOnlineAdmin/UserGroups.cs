using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using Microsoft.SharePoint.Client;

namespace SharePointOnlineAdmin
{
    public class UserGroups
    {
        public static void RunShowMembership(string userLogonName)
        {
            string serviceUserName = ConfigurationManager.AppSettings["serviceUserName"];
            string servicePassword = ConfigurationManager.AppSettings["servicePassword"];

            string siteCollectionUrl = "https://anitaborgo365.sharepoint.com/";
            
            OfficeDevPnP.Core.AuthenticationManager authenticationMgr = new OfficeDevPnP.Core.AuthenticationManager();

            ClientContext clientContext = authenticationMgr.GetSharePointOnlineAuthenticatedContextTenant(siteCollectionUrl, serviceUserName, servicePassword);
            Web web = clientContext.Web;

            Console.WriteLine($"{userLogonName} is a member of the following groups: ");
           
            GroupCollection collGroup = clientContext.Web.SiteGroups;
            clientContext.Load(collGroup);
            clientContext.ExecuteQuery();


            foreach (Group group in collGroup)
            {
               if ( web.IsUserInGroup(group.Title, userLogonName))
                        Console.WriteLine(group.Title);
            }

            Console.WriteLine("press any key to end.");
            Console.ReadKey();


        }

    }
}
