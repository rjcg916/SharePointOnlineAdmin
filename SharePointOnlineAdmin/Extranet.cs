using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using System;
using System.Linq;
using System.Security;
using System.Configuration;

namespace SharePointOnlineAdmin
{
    static public class Extranet
    {
        private class Configuration
        {
            public const string ownerGroupName = "AnitaB.org Extranet Owners";
            public const string groupSuffix = " Members";
            public const string homePageName = "Home.aspx";
            public const string webPartZoneId = "Header";
            public const string sitePageLibraryName = "Site Pages";
            public const string defaultDocumentLibraryName = "Documents";
            public const string extranetSiteTemplate = "STS#0";
        }

        static void InitalizeExtranetPage(this Web web)
        {
            List sitePages = web.Lists.GetByTitle(Configuration.sitePageLibraryName);
            web.Context.ExecuteQuery();

            sitePages.RemoveAllWebParts(Configuration.homePageName);
            sitePages.AddRecentlyChangedWebPart(Configuration.homePageName, Configuration.webPartZoneId);
        }


        static  void AddLibrariesToSite(this Web web, string libraryList)
        {

            ClientContext clientContext = (ClientContext)web.Context;

            // fetch web data           
            clientContext.Load(web.ParentWeb);
            clientContext.Load(web, w => w.Url);
            clientContext.Load(web, w => w.Title);
            clientContext.Load(web, w => w.CurrentUser.LoginName);

            clientContext.ExecuteQuery();

            string currentUserLoginName = web.CurrentUser.LoginName;


            // find parentWeb for later use to add security group
            Web parentWeb = clientContext.Site.OpenWebById(web.ParentWeb.Id);

            //for each specified library, create and apply security groups
            foreach (string ln in libraryList.Split(','))
            {
                string libraryName = ln.Trim();
                string groupName = $"{libraryName} ({web.Title}) {Configuration.groupSuffix}";

                if (!web.GroupExists(groupName))
                {
                    Console.WriteLine("Creating new library: " + libraryName);


                    Group newGroup;
                    try
                    {

                        // create group for library mgt                        
                        newGroup = web.AddGroup(groupName, "Use this group for access to an Extranet document library.", false);

                        // Set owner for group to extranet owner                       
                        var group = web.SiteGroups.GetByName(Configuration.ownerGroupName);
                        clientContext.Load(group);
                        clientContext.ExecuteQueryRetry();
                        if (group != null)
                        {
                            newGroup.Owner = group;
                            newGroup.Update();
                            clientContext.ExecuteQueryRetry();
                        }

                        // give group contributor access to library
                        List newLibrary = web.CreateDocumentLibrary(libraryName);
                        newLibrary.SetListPermission(newGroup, RoleType.Contributor);


                        // give group read access to current and parent web
                        web.AddPermissionLevelToGroup(groupName, RoleType.Reader);
                        web.RemovePermissionLevelFromUser(currentUserLoginName, RoleType.Administrator, true);

                        // clean up permissions from running program
                        newLibrary.RemovePermissionLevelFromUser(currentUserLoginName, RoleType.Administrator, true);
                        parentWeb.AddPermissionLevelToGroup(groupName, RoleType.Reader);

                        // update navigation - NOTE: no way to create navigation "link" - "header" created instead
                       // clientContext.Load(newLibrary, lib => lib.DefaultViewUrl);
                       // clientContext.ExecuteQueryRetry();
                       // Uri nodeUri = new Uri(newLibrary.DefaultViewUrl, UriKind.Relative);
                       // clientContext.Web.AddNavigationNode(newLibrary.Title, nodeUri, String.Empty, OfficeDevPnP.Core.Enums.NavigationType.QuickLaunch, false);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Source, e.Message, e.InnerException);
                    }
                }

            }
        }

        static  void CreateSiteAndLibraries(this Web web, string siteName, string libraryList)
        {

            //get current user name
            web.Context.Load(web, w => w.CurrentUser.LoginName);
            web.Context.ExecuteQueryRetry();

            string currentUserLoginName = web.CurrentUser.LoginName;

            //create site
            Console.WriteLine("Creating new site: " + siteName);

            Web newWeb;
            try
            {
                //create site

                //web.Context.RequestTimeout = 1000 * 60 * 7; //7min
                newWeb = web.CreateWeb(siteName, siteName, "Extranet room for " + siteName, Configuration.extranetSiteTemplate, 1033, false);


                //activate publishing to inherit branding but immediately remove

                var featureId = new Guid("94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb"); //publishing feature id for SPO
                var features = newWeb.Features;
                features.Add(featureId, true, FeatureDefinitionScope.None);
                web.Context.ExecuteQueryRetry();
                features.Remove(featureId, true);
                web.Context.ExecuteQueryRetry();

                // apply permission levels         
                newWeb.AddPermissionLevelToGroup(Configuration.ownerGroupName, RoleType.Administrator);

                // clean up permissions from running program
                newWeb.RemovePermissionLevelFromUser(currentUserLoginName, RoleType.Administrator, true);

                //delete default library
                string libraryTitle = Configuration.defaultDocumentLibraryName;
                if (newWeb.ListExists(libraryTitle))
                {
                    List list = newWeb.Lists.GetByTitle(libraryTitle);
                    list.DeleteObject();
                    newWeb.Context.ExecuteQuery();
                }
                //setup page
                newWeb.InitalizeExtranetPage();

                // go add libraries
                newWeb.AddLibrariesToSite(libraryList);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Source, e.Message, e.InnerException);
            }
        }
        static public void RunAddLibrariesToSite(string siteName, string libraryList)
        {
            string serviceUserName = ConfigurationManager.AppSettings["serviceUserName"];
            string servicePassword = ConfigurationManager.AppSettings["servicePassword"];
            string extranetUrl = ConfigurationManager.AppSettings["extranetUrl"];

            string siteUrl = extranetUrl + siteName;

            Console.WriteLine($"Adding Libraries: {libraryList} to Site: {siteUrl}");

            OfficeDevPnP.Core.AuthenticationManager authenticationMgr = new OfficeDevPnP.Core.AuthenticationManager();

            ClientContext clientContext = authenticationMgr.GetSharePointOnlineAuthenticatedContextTenant(siteUrl, serviceUserName, servicePassword);
            Web web = clientContext.Web;
            AddLibrariesToSite(web, libraryList);


        }
        static public void RunCreateSiteAndLibraries(string siteName, string libraryList)
        {
            string serviceUserName = ConfigurationManager.AppSettings["serviceUserName"];
            string servicePassword = ConfigurationManager.AppSettings["servicePassword"];
            string extranetUrl = ConfigurationManager.AppSettings["extranetUrl"];

            Console.WriteLine($"Creating Site:{siteName} and Libraries:{libraryList} . . .");

            OfficeDevPnP.Core.AuthenticationManager authenticationMgr = new OfficeDevPnP.Core.AuthenticationManager();

            ClientContext clientContext = authenticationMgr.GetSharePointOnlineAuthenticatedContextTenant(extranetUrl, serviceUserName, servicePassword);
            Web web = clientContext.Web;

            CreateSiteAndLibraries(web, siteName, libraryList);

        }
    }
}
