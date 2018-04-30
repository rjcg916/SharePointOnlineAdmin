# SharePoint Online Admin

Custom console application for Administration and Provisioning of SharePoint Online sites.  

This application provides re-usuable commands for routine, non-trivial provisioning operations. The application is designed 
for use by the SharePoint Online Administrator.

Many of the supported operations require item and folder level permissions, which can be problematic to complete manually.
This was one of the major reasons for creation of the application

This application uses the SharePoint Client Side Object Model (CSOM) and the SharePoint Online PnP libraries 
to access SharePoint Online sites.

## Running
  The program.cs file contains a switch for all commands.  The easiest means of running the commands is as follows:
   *  Select, Debug -> Properties -> Start, enter desired command and working directory
   *  Select, Debug -> Start Debugging
    

The following operations sets are supported:

### Partner (see RunPartner.cs)

The partner commands read a partner list from a specified input CSV file and process items as follows:
   * PAdd - provisioning all partner-related list items and folders for each specified partner.  In addition, a 
            unique security group is created for each partner.  After group creation, users must be added to the appropriate security group.
   * PDisplay - displays the content of a CSV file (useful for debugging)
   * PRemove - removes all partner-related list items and 
               folders for each specified partner
 
NOTE: The input processing assumes a fixed, ordered columns with no column headers (see partner.cs for specifics).

### Extranet (see RunExtranet.cs)

The extranet commands read input strings which specify Extranet sub-site and libraries for creation. The following commands are supported:
   * ECreate - create a specified sub-site and one or more document libraries
   * EAdd - create one or more document libraries within an existing sub-site

 In both of these cases, a unique security group is created to be used for each document library. After site and library creation, 
 users must be added to the appropriate security group.

## Configuration
   The App.config file contains values for standard site locations as well as credentials of the service account
   used for provisioning.  (For security reasons, this is excluded from Version Control.)
   
## Building
  This application is contained within a single Visual Studio Solution/Project.  NuGet is used to retrieve the appropriate 
  SharePoint Client and PnP libraries.
  
