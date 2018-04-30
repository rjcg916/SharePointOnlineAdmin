using System;
using System.Linq;
using System.Configuration;
using Microsoft.SharePoint.Client;
using System.IO;
using Microsoft.VisualBasic.FileIO;

namespace SharePointOnlineAdmin
{
    public static class ProvisionPartner
    {
        public class Configuration
        {
            public const string OwnerGroupName = "AnitaB.org Partner Portal Owners";
            public const string GroupSuffix = " Members";
            public const string PartnerListName = "Partners";
            public const string PartnerDocumentLibraryName = "Partnership Folder";
            public const string PartnerLogoLibraryName = "Partner Logos";
            public const string PartnerLogoLibraryNamePath = "PartnerLogos";
            public const string PartnerContactListName = "AnitaBOrg Partnerships Contacts";
            public const string PartnerDiscussionListName = "Discussions";
        }


        public struct PartnerData
        {
            public string name;
            public string type;
            public DateTime startDate;
            public DateTime renewalDate;
            public string address;
            public string logoFileName;
            public string logoURL;
            public string bdmName;
            public string pemName;
            override public string ToString()
            {
                return "Name: " + name + Environment.NewLine + "Type: " + type + Environment.NewLine + "StartDate: " + startDate + Environment.NewLine + "RenewalDate: " + renewalDate + Environment.NewLine + "Address: " + address + Environment.NewLine + "LogoFileName: " + logoFileName + Environment.NewLine + "BDM Name: " + bdmName + Environment.NewLine + "PEM Name: " + pemName;
            }
        }


        static string GetFullLogoURL(string logoFileName)
        {
            string partnerUrl = ConfigurationManager.AppSettings["partnerUrl"];
            return partnerUrl + Configuration.PartnerLogoLibraryNamePath + logoFileName;
        }

        static bool LogoExists(Web web, string logoURL, string logoLibraryName)
        {
            List logoLibrary = web.Lists.GetByTitle(logoLibraryName);

            return logoLibrary.FindFiles(logoURL).Count > 0;
        }


        static string GetPartnerGroupName(string name)
        {
            return name + Configuration.GroupSuffix;

        }

        static ListItem AddPartnerListItem(this List partnersList, List contactList, Group group, PartnerData partnerData)
        {

            ClientContext clientContext = (ClientContext)partnersList.Context;
            Web web = clientContext.Web;

            web.Context.Load(web, w => w.CurrentUser.LoginName);
            web.Context.ExecuteQueryRetry();

            string currentUserLoginName = web.CurrentUser.LoginName;

            //add new item to partner list (name == partner name, group name)

            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem newPartner = partnersList.AddItem(itemCreateInfo);

            // update field values
            /* assumes list field names*/
            newPartner["Title"] = partnerData.name;
            newPartner["PartnerType"] = partnerData.type;
            newPartner["PartnershipStartDate"] = partnerData.startDate;
            newPartner["PartnershipRenewalDate"] = partnerData.renewalDate;
            newPartner["WorkAddress"] = partnerData.address;
            newPartner["Logo"] = partnerData.logoURL;

            newPartner.Update();
            clientContext.ExecuteQuery();

            // contacts are lookup fields
            string lookupFieldName = "FullName";
            string lookupFieldType = "Text";

            FieldLookupValue flvBDM = ListLibrary.GetLookupValue(contactList, partnerData.bdmName, lookupFieldName, lookupFieldType);
            newPartner["BusinessDevelopmentManagerContac"] = flvBDM;
            newPartner.Update();
            clientContext.ExecuteQuery();

            FieldLookupValue flvPEM = ListLibrary.GetLookupValue(contactList, partnerData.pemName, lookupFieldName, lookupFieldType);
            newPartner["ProgramEngagementManagerContact"] = flvPEM;

            newPartner.Update();
            clientContext.ExecuteQuery();
            
            return newPartner;
        }



        static public string AddPartner(Web web, PartnerData partnerData)
        {
            ClientContext clientContext = (ClientContext)web.Context;

            web.Context.Load(web, w => w.CurrentUser.LoginName);
            web.Context.Load(web, w => w.Url);
            web.Context.ExecuteQueryRetry();
            string currentUserLoginName = web.CurrentUser.LoginName;


            //don't overwrite existing entries
            if (web.GroupExists(GetPartnerGroupName(partnerData.name))) return partnerData.name + " already exists!";

            Principal owner = web.SiteGroups.GetByName(Configuration.OwnerGroupName);


            // create group for partner

            Group partnerGroup = web.AddGroup(GetPartnerGroupName(partnerData.name), "Use this group to grant people read permissions to " + partnerData.name, false);
            List partnersList = web.Lists.GetByTitle(Configuration.PartnerListName);


            // if it exists in current directory, upload logo

            string filePath = partnerData.logoFileName;
            if (System.IO.File.Exists(filePath))
            {
                var partnerLogoLibrary = web.Lists.GetByTitle(Configuration.PartnerLogoLibraryName);
                partnerLogoLibrary.UploadFile(filePath);
                ListItem newLogo = partnerLogoLibrary.GetItemByDisplayName(partnerData.logoFileName);

                if (newLogo != null)
                {
                    newLogo.BreakRoleInheritance(false, false);
                    newLogo.AddPermissionLevelToPrincipal(partnerGroup, RoleType.Reader, true);
                    newLogo.AddPermissionLevelToPrincipal(owner, RoleType.Contributor, true);
                    newLogo.RemovePermissionLevelFromUser(currentUserLoginName, RoleType.Contributor, true);
                    partnerData.logoURL = web.Url + "/" + Configuration.PartnerLogoLibraryNamePath + "/" + partnerData.logoFileName;
                }
            }
            else
                partnerData.logoURL = string.Empty;


            //create entry for partner in partner list with permissions for partner
            List contactList = clientContext.Web.Lists.GetByTitle(Configuration.PartnerContactListName);
            ListItem newPartner = partnersList.AddPartnerListItem(contactList, partnerGroup, partnerData);

            newPartner.BreakRoleInheritance(false, false);
            newPartner.AddPermissionLevelToPrincipal(partnerGroup, RoleType.Reader, true);
            newPartner.AddPermissionLevelToPrincipal(owner, RoleType.Contributor, true);
            newPartner.RemovePermissionLevelFromUser(currentUserLoginName, RoleType.Administrator, true);


            //create partner-specific folder

            List documentLibrary = web.Lists.GetByTitle(Configuration.PartnerDocumentLibraryName);


            ListItem folder = documentLibrary.RootFolder.CreateFolder(partnerData.name).ListItemAllFields;
            folder.BreakRoleInheritance(false, false);
            folder.AddPermissionLevelToPrincipal(partnerGroup, RoleType.Contributor, true);
            folder.AddPermissionLevelToPrincipal(owner, RoleType.Contributor);
            folder.RemovePermissionLevelFromUser(currentUserLoginName, RoleType.Administrator, true);


            // give partner access to Discussion Forum w/o ability to create a new topic
            List discussionList = web.Lists.GetByTitle(Configuration.PartnerDiscussionListName);
            discussionList.AddPrincipalToAllFolders(partnerGroup);
            discussionList.AddPrincipalToAllFolders(owner);

            //all partners get read access
            web.AddPermissionLevelToGroup(GetPartnerGroupName(partnerData.name), RoleType.Reader);

            return partnerData.name + " added!";
        }
        static public string DisplayPartner(Web web, PartnerData partnerData)
        {

            return partnerData.ToString();
        }


        static public string RemovePartner(Web web, PartnerData partnerData)
        {

            //remove partner logo
            List partnerLogoLibrary = web.Lists.GetByTitle(Configuration.PartnerLogoLibraryName);
            partnerLogoLibrary.RemoveItemByDisplayName(partnerData.logoFileName);

            //remove partner list entry
            List partnerList = web.Lists.GetByTitle(Configuration.PartnerListName);
            partnerList.RemoveItemByDisplayName(partnerData.name);

            //remove partner-specific folder
            List documentLibrary = web.Lists.GetByTitle(Configuration.PartnerDocumentLibraryName);
            documentLibrary.RemoveItemByDisplayName(partnerData.name);

            //remove partner-specific security group
            string theGroup = GetPartnerGroupName(partnerData.name);

            string results = string.Empty;
            string warning = string.Empty;

            if (!web.GroupExists(theGroup))
                warning = $"(Warning: Group {theGroup} not found!)";
            else
            {
                web.RemoveGroup(GetPartnerGroupName(partnerData.name));     
            }

            results = $"  {partnerData.name} completed. {warning}";

            return results;
        }


        public delegate string ProcessPartner(Web web, PartnerData partnerData);


        static string FirstNameToFullName(string name)
        {
            string newName = String.Empty;

            switch (name)
            {
                case "Misa":
                    newName = "Misa Mascovich";
                    break;
                case "Cait":
                    newName = "Cait Rogan";
                    break;
                default:
                    newName = name;
                    break;

            }
            return newName;

        }

        static public void ProcessPartnerFile(Web web, string csvName, ProcessPartner processPartner)
        {


            string csvData = System.IO.File.ReadAllText(csvName);


            // process each partner row
            foreach (string partnerData in csvData.Split('\n'))
            {

                if (!string.IsNullOrEmpty(partnerData))
                {

                    // process the current row

                    TextFieldParser parser = new TextFieldParser(new StringReader(partnerData));
                    parser.HasFieldsEnclosedInQuotes = true;
                    parser.SetDelimiters(",");

                    string[] partnerFields;

                    while (!parser.EndOfData)
                    {
                        partnerFields = parser.ReadFields();

                        PartnerData thePartner = new PartnerData();

                        for (int i = 0; i < partnerFields.Count(); i++)
                        {
                            string theValue = partnerFields[i];

                            if (String.IsNullOrEmpty(theValue))
                                continue;

                            /* this section assumes CSV column order*/

                            switch (i)
                            {
                                case 0:
                                    thePartner.name = new string(theValue.Where(x => char.IsWhiteSpace(x) || char.IsLetterOrDigit(x)).ToArray());
                                    break;
                                case 1:
                                    thePartner.type = theValue.Substring(theValue.IndexOf(' ')).Trim(); // type is in first part of data
                                    break;
                                case 2:
                                    thePartner.bdmName = theValue.Substring(theValue.IndexOf(' ')).Trim();
                                    break;
                                case 3:
                                    thePartner.pemName = FirstNameToFullName(theValue.Trim());
                                    break;
                                case 5:
                                    thePartner.startDate = DateTime.Parse(theValue);
                                    break;
                                case 4:
                                    thePartner.renewalDate = DateTime.Parse(theValue);
                                    break;
                                case 6:
                                case 7:
                                case 8:
                                case 9:
                                case 10:
                                    thePartner.address += theValue.Trim('"') + " "; //aggregate address components in a single field
                                    break;
                                case 11:
                                    thePartner.logoFileName = theValue.Trim();
                                    break;
                                default: break;

                            }
                        }


                        Console.WriteLine(processPartner(web, thePartner));
                    }

                    parser.Close();

                }

            }
        }


    }
}
