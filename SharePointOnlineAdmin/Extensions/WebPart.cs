using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using System.Linq;

namespace SharePointOnlineAdmin
{
    static public class WebPart
    {
        static string recentlyChangedWebPartSchemaXml = @"<?xml version='1.0' encoding='utf-8'?>
<webParts>
  <webPart xmlns='http://schemas.microsoft.com/WebPart/v3'>
    <metaData>
      <type name='Microsoft.Office.Server.Search.WebControls.ContentBySearchWebPart, Microsoft.Office.Server.Search, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c' />
      <importErrorMessage>Cannot import this Web Part.</importErrorMessage>
    </metaData>
    <data>
      <properties>
        <property name='BypassResultTypes' type='bool'>True</property>
        <property name='ItemTemplateId' type='string'>~sitecollection/_catalogs/masterpage/Display Templates/Content Web Parts/Item_TwoLines.js</property>
        <property name='PropertyMappings' type='string' />
        <property name='ChromeState' type='chromestate'>Normal</property>
        <property name='IncludeResultTypeConstraint' type='bool'>False</property>
        <property name='StartingItemIndex' type='int'>1</property>
        <property name='ShowDefinitions' type='bool'>False</property>
        <property name='Height' type='string' />
        <property name='Hidden' type='bool'>False</property>
        <property name='HitHighlightedPropertiesJson' type='string'>['Title','Path','Author','SectionNames','SiteDescription']</property>
        <property name='ScrollToTopOnRedraw' type='bool'>False</property>
        <property name='UseSharedDataProvider' type='bool'>False</property>
        <property name='RepositionLanguageDropDown' type='bool'>False</property>
        <property name='AlwaysRenderOnServer' type='bool'>False</property>
        <property name='AllowConnect' type='bool'>True</property>
        <property name='ItemBodyTemplateId' type='string' />
        <property name='ShowAlertMe' type='bool'>True</property>
        <property name='ExportMode' type='exportmode'>All</property>
        <property name='AddSEOPropertiesFromSearch' type='bool'>False</property>
        <property name='ShowUpScopeMessage' type='bool'>False</property>
        <property name='AllowHide' type='bool'>True</property>
        <property name='AllowClose' type='bool'>True</property>
        <property name='UseSimplifiedQueryBuilder' type='bool'>True</property>
        <property name='ShouldHideControlWhenEmpty' type='bool'>True</property>
        <property name='ResultType' type='string' />
        <property name='LogAnalyticsViewEvent' type='bool'>False</property>
        <property name='MaxPagesAfterCurrent' type='int'>1</property>
        <property name='TitleUrl' type='string' />
        <property name='EmptyMessage' type='string' />
        <property name='AdvancedSearchPageAddress' type='string'>advanced.aspx</property>
        <property name='IsXGeo3SForwardingFlighted' type='bool'>False</property>
        <property name='AllowMinimize' type='bool'>True</property>
        <property name='ShowBestBets' type='bool'>False</property>
        <property name='AllowEdit' type='bool'>True</property>
        <property name='NumberOfItems' type='int'>3</property>
        <property name='HelpUrl' type='string' />
        <property name='ShowPaging' type='bool'>True</property>
        <property name='ShowViewDuplicates' type='bool'>False</property>
        <property name='SelectedPropertiesJson' type='string'>['Path','Title','FileExtension','SecondaryFileExtension','IsAllDayEvent']</property>
        <property name='TargetResultTable' type='string'>RelevantResults</property>
        <property name='HelpMode' type='helpmode'>Modeless</property>
        <property name='ShowXGeoOptions' type='bool'>False</property>
        <property name='IsXGeoFlighted' type='bool'>False</property>
        <property name='ShowPersonalFavorites' type='bool'>False</property>
        <property name='EnableXGeo3SForwarding' type='bool'>False</property>
        <property name='PreloadedItemTemplateIdsJson' type='string'>null</property>
        <property name='Description' type='string'>This Web Part will show items that have been modified recently. This can help site users track the latest activity on a site or a library. When you add it to the page, this Web Part will show items from the current site. You can change this setting to show items from another site or list by editing the Web Part and changing its search criteria.As new content is discovered by search, this Web Part will display an updated list of items each time the page is viewed.</property>
        <property name='ShowPreferencesLink' type='bool'>True</property>
        <property name='QueryGroupName' type='string'>Default</property>
        <property name='ShowResultCount' type='bool'>True</property>
        <property name='TitleIconImageUrl' type='string' />
        <property name='Direction' type='direction'>NotSet</property>
        <property name='ResultsPerPage' type='int'>3</property>
        <property name='AvailableSortsJson' type='string'>null</property>
        <property name='ShowResults' type='bool'>True</property>
        <property name='ServerIncludeScriptsJson' type='string'>null</property>
        <property name='SearchCenterXGeoLocations' type='string' />
        <property name='DataProviderJSON' type='string'>{'QueryGroupName':'Default','QueryPropertiesTemplateUrl':'sitesearch://webroot','IgnoreQueryPropertiesTemplateUrl':false,'SourceID':'ba63bbae-fa9c-42c0-b027-9a878f16557c','SourceName':'Recently changed items','SourceLevel':'Ssa','CollapseSpecification':'','QueryTemplate':null,'FallbackSort':null,'FallbackSortJson':'null','RankRules':null,'RankRulesJson':'null','AsynchronousResultRetrieval':false,'SendContentBeforeQuery':true,'BatchClientQuery':true,'FallbackLanguage':-1,'FallbackRankingModelID':'','EnableStemming':true,'EnablePhonetic':false,'EnableNicknames':false,'EnableInterleaving':false,'EnableQueryRules':true,'EnableOrderingHitHighlightedProperty':false,'HitHighlightedMultivaluePropertyLimit':-1,'IgnoreContextualScope':true,'ScopeResultsToCurrentSite':false,'TrimDuplicates':false,'Properties':{'TryCache':true,'UpdateLinksForCatalogItems':true,'EnableStacking':true,'CrossGeoQuery':'false','ListId':'44b71b83-4b39-46c2-a406-79b89a547f92','ListItemId':1,'FillIn':'false','Scope':'{Site.URL}'},'PropertiesJson':'{\'TryCache\':true,\'UpdateLinksForCatalogItems\':true,\'EnableStacking\':true,\'CrossGeoQuery\':\'false\',\'ListId\':\'44b71b83-4b39-46c2-a406-79b89a547f92\',\'ListItemId\':1,\'FillIn\':\'false\',\'Scope\':\'{Site.URL}\'}','ClientType':'ContentSearchRegular','ClientFunction':'','ClientFunctionDetails':'','UpdateAjaxNavigate':true,'SummaryLength':180,'DesiredSnippetLength':90,'PersonalizedQuery':false,'FallbackRefinementFilters':[{'n':'SPContentType','t':['\'ǂǂ446f63756d656e74\''],'o':'and','k':false,'m':null}],'IgnoreStaleServerQuery':false,'RenderTemplateId':'DefaultDataProvider','AlternateErrorMessage':null,'Title':''}</property>
        <property name='ShowAdvancedLink' type='bool'>True</property>
        <property name='ShowDidYouMean' type='bool'>False</property>
        <property name='AllowZoneChange' type='bool'>True</property>
        <property name='ChromeType' type='chrometype'>Default</property>
        <property name='GroupTemplateId' type='string'>~sitecollection/_catalogs/masterpage/Display Templates/Content Web Parts/Group_Content.js</property>
        <property name='MissingAssembly' type='string'>Cannot import this Web Part.</property>
        <property name='OverwriteResultPath' type='bool'>True</property>
        <property name='Width' type='string' />
        <property name='MaxPagesBeforeCurrent' type='int'>4</property>
        <property name='XGeoTenantsInfo' type='string' />
        <property name='ShowLanguageOptions' type='bool'>True</property>
        <property name='ResultTypeId' type='string' />
        <property name='AlternateErrorMessage' type='string' null='true' />
        <property name='Title' type='string'>Recently Changed Items</property>
        <property name='RenderTemplateId' type='string'>~sitecollection/_catalogs/masterpage/Display Templates/Content Web Parts/Control_List.js</property>
        <property name='EmitStyleReference' type='bool'>True</property>
        <property name='StatesJson' type='string'>{}</property>
        <property name='ShowSortOptions' type='bool'>False</property>
        <property name='CatalogIconImageUrl' type='string' />
      </properties>
    </data>
  </webPart>
</webParts>";


        public static void AddRecentlyChangedWebPart(this List sitePages, string pageName, string zoneId)
        {


            ClientContext context = (ClientContext)sitePages.Context;

            context.Load(sitePages.RootFolder.Files);
            context.ExecuteQuery();

            Microsoft.SharePoint.Client.File thePage = sitePages.RootFolder.Files.First(f => f.Name == pageName);
            context.Load(thePage);
            context.ExecuteQuery();

            // Gets the webparts available on the page  
            var wpm = thePage.GetLimitedWebPartManager(PersonalizationScope.Shared);
            context.Load(wpm.WebParts,
                wps => wps.Include(wp => wp.WebPart.Title));
            context.ExecuteQuery();
            var availableWebparts = wpm.WebParts;
            // Check if the current webpart already exists.  
            var filteredWebParts = from isWPAvail in availableWebparts
                                   where isWPAvail.WebPart.Title == "Recently Changed Items"
                                   select isWPAvail;
            if (filteredWebParts.Count() <= 0)
            {
                var importedWebPart = wpm.ImportWebPart(recentlyChangedWebPartSchemaXml);
                var webPart = wpm.AddWebPart(importedWebPart.WebPart, zoneId, 0);
                context.ExecuteQuery();
            }

        }
        public static void RemoveAllWebParts(this List sitePages, string pageName)
        {
            ClientContext context = (ClientContext)sitePages.Context;

            context.Load(sitePages.RootFolder.Files);
            context.ExecuteQuery();

            Microsoft.SharePoint.Client.File thePage = sitePages.RootFolder.Files.First(f => f.Name == pageName);
            context.Load(thePage);
            context.ExecuteQuery();

            LimitedWebPartManager wpm = thePage.GetLimitedWebPartManager(Microsoft.SharePoint.Client.WebParts.PersonalizationScope.Shared);
            context.Load(wpm);
            context.ExecuteQuery();

            WebPartDefinitionCollection webParts = wpm.WebParts;
            context.Load(webParts);
            context.ExecuteQuery();

            foreach (WebPartDefinition wp in webParts)
            {
                wp.DeleteWebPart();
                context.ExecuteQuery();
            }

        }

    }
}
