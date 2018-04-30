using Microsoft.SharePoint.Client;
using System;
using System.IO;
using System.Linq;

namespace SharePointOnlineAdmin
{
    static public class ListLibrary
    {

     
        static public void AddPrincipalToAllFolders(this List list, Principal principal)
        {

            //fetch all folders
            var clientContext = list.Context;
            var folderItems = list.GetItems(CamlQuery.CreateAllFoldersQuery());
            clientContext.Load(folderItems, icol => icol.Include(i => i.Folder, i=>i.HasUniqueRoleAssignments));
            clientContext.ExecuteQuery();

            foreach (ListItem f in folderItems) {
                if(f.HasUniqueRoleAssignments) // only add principal to folders with unique permissions
                    f.AddPermissionLevelToPrincipal(principal, RoleType.Contributor, true);
            }       
        }

        public static FieldLookupValue GetLookupValue(this List list, string value,
           string lookupFieldName, string lookupFieldType)
        {
            ClientContext clientContext = (ClientContext) list.ParentWeb.Context;     
                  
            FieldLookupValue lookupValue = null;
     
            if (list != null)
            {
                CamlQuery camlQueryForItem = new CamlQuery();
                camlQueryForItem.ViewXml = string.Format(@"<View>
                  <Query>
                      <Where>
                         <Eq>
                             <FieldRef Name='{0}'/>
                             <Value Type='{1}'>{2}</Value>
                         </Eq>
                       </Where>
                   </Query>
            </View>", lookupFieldName, lookupFieldType, value);

                ListItemCollection listItems = list.GetItems(camlQueryForItem);
                clientContext.Load(listItems, items => items.Include
                                                  (listItem => listItem["ID"],
                                                   listItem => listItem[lookupFieldName]));
                clientContext.ExecuteQuery();

                if (listItems != null)
                {
                    ListItem item = listItems[0];
                    lookupValue = new FieldLookupValue();
                    lookupValue.LookupId = int.Parse(item["ID"].ToString());                   
                }
            }

            return lookupValue;
        }

        static public string UploadFile(this List library, string fileName)
        {
            ClientContext context = (ClientContext)library.Context;

            using (var fs = new System.IO.FileStream(fileName, FileMode.Open))
            {
                var fi = new FileInfo(fileName);
                //                var list = web.Lists.GetByTitle(libraryTitle);
                context.Load(library.RootFolder);
                context.ExecuteQuery();
                var fileUrl = String.Format("{0}/{1}", library.RootFolder.ServerRelativeUrl, fi.Name);

                Microsoft.SharePoint.Client.File.SaveBinaryDirect(context, fileUrl, fs, true);

                return fileUrl;

            }

        }



        static public ListItem GetItemByDisplayName(this List list, string name)
        {

            ClientContext clientContext = (ClientContext)list.Context;

            CamlQuery camlQuery = new CamlQuery();
            ListItemCollection collListItem = list.GetItems(camlQuery);

            clientContext.Load(collListItem,
                 items => items.Include(
                    item => item.Id,
                    item => item.DisplayName,
                    item => item.HasUniqueRoleAssignments));

            clientContext.ExecuteQuery();

            // find and delete item
            foreach (ListItem theItem in collListItem)
            {
                if (String.Equals(theItem.DisplayName, name.Substring(0, theItem.DisplayName.Count())))
                {
                    ListItem oListItem = list.GetItemById(theItem.Id);
                    return oListItem;
                    
                }
            }

            return null;
        }

        static public void RemoveItemByDisplayName(this List list,  string name)
        {

            ClientContext clientContext = (ClientContext) list.Context;

            CamlQuery camlQuery = new CamlQuery();
            ListItemCollection collListItem = list.GetItems(camlQuery);

            clientContext.Load(collListItem,
                 items => items.Include(
                    item => item.Id,
                    item => item.DisplayName,
                    item => item.HasUniqueRoleAssignments));

            clientContext.ExecuteQuery();

            // find and delete item
            foreach (ListItem theItem in collListItem)
            {
                if (String.Equals(theItem.DisplayName, name))
                {
                    ListItem oListItem = list.GetItemById(theItem.Id);
                    oListItem.DeleteObject();
                    clientContext.ExecuteQuery();

                    break;
                }
            }

        }

    }

}
