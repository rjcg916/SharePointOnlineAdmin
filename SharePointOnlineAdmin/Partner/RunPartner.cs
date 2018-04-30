using System;
using System.Configuration;

namespace SharePointOnlineAdmin
{
    public static class RunPartner
    {

        static public void RemovePartners(string CSVFilePath)
        {

            Console.WriteLine(String.Format("Removing Partners found in list: {0}. . .", CSVFilePath));


            string serviceUserName = ConfigurationManager.AppSettings["serviceUserName"];
            string servicePassword = ConfigurationManager.AppSettings["servicePassword"];
            string partnerSiteUrl = ConfigurationManager.AppSettings["partnerUrl"]; ;


            OfficeDevPnP.Core.AuthenticationManager authenticationMgr = new OfficeDevPnP.Core.AuthenticationManager();

            var clientContext = authenticationMgr.GetSharePointOnlineAuthenticatedContextTenant(partnerSiteUrl, serviceUserName, servicePassword);


            ProvisionPartner.ProcessPartnerFile(clientContext.Web, CSVFilePath, ProvisionPartner.RemovePartner);

  
        }
        static public void AddPartners(string CSVFilePath)
        {

            Console.WriteLine(String.Format("Adding Partners found in list: {0}. . .", CSVFilePath));


            string serviceUserName = ConfigurationManager.AppSettings["serviceUserName"];
            string servicePassword = ConfigurationManager.AppSettings["servicePassword"];
            string partnerSiteUrl = ConfigurationManager.AppSettings["partnerUrl"]; ;


            OfficeDevPnP.Core.AuthenticationManager authenticationMgr = new OfficeDevPnP.Core.AuthenticationManager();

            var clientContext = authenticationMgr.GetSharePointOnlineAuthenticatedContextTenant(partnerSiteUrl, serviceUserName, servicePassword);

            ProvisionPartner.ProcessPartnerFile(clientContext.Web, CSVFilePath, ProvisionPartner.AddPartner);


        }

        static public void DisplayPartners(string CSVFilePath)
        {

            Console.WriteLine(String.Format("Displaying Partners found in list: {0}. . .", CSVFilePath));

            string serviceUserName = ConfigurationManager.AppSettings["serviceUserName"];
            string servicePassword = ConfigurationManager.AppSettings["servicePassword"];
            string partnerSiteUrl = ConfigurationManager.AppSettings["partnerUrl"]; ;


            OfficeDevPnP.Core.AuthenticationManager authenticationMgr = new OfficeDevPnP.Core.AuthenticationManager();

            var clientContext = authenticationMgr.GetSharePointOnlineAuthenticatedContextTenant(partnerSiteUrl, serviceUserName, servicePassword);

    
            ProvisionPartner.ProcessPartnerFile(clientContext.Web, CSVFilePath, ProvisionPartner.DisplayPartner);

        }

    }
}
