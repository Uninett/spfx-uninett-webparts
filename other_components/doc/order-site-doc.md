# Order Site documentation

## Permissions

### Operational permissions

- Create a new Office 365 user account with the following settings:
  - **Name**: o365@\<yourtenant>.onmicrosoft.com
  - **Roles**: SharePoint administrator
  - Assign the following licenses:
    - Exchange Online
    - Office 365
- Give the new user account the Owner role (Full control) in your SharePoint Site Collection
- 
### Deployment permissions

## Edit files



## Deployment

### Create new Resource Group

Start by creating a new Resource Group in your Azure Subscription.

### Create Function App

Create a new Function App in the Resource Group with the following settings:
- **App name**: \<YourTenantName>SiteProvEngine (e.g. *ContosoSiteProvEngine*)
- **OS**: Windows
- **Hosting Plan**: Consumption Plan
- **Location**: same as your Resource Group
- **Runtime Stack**: .NET Core
- **Storage**: Create new -> \<yourtenantname>siteprovengine (e.g. *contosositeprovengine*)
  
### Publish Function Apps

1. Open *KDTO.sln* in Visual Studio 2019
2. Right click the Azure.Functions project -> **Build**
3. Right click the Azure.Functions project -> **Publish**
4. Pick **Azure Functions Consumption Plan** as a publish target
5. Choose "Select Existing" (check "Run from package file") -> click **Publish**  
      
   ![Publish Function App](img/figure1.png)

6. Select the *SiteProvEngine* Function App in your Resource Group -> **OK**
7. Follow the above steps for the Azure.ApplyPowershell project, but publish it to the *ApplyPowershell* Function App instead.

### Deploy Logic Apps




## Permissions

- A user needs read/write permissions on the *Bestillinger* list to be able to use the Order Site web part.
- You need a user with the Owner role of the SharePoint Site Collection you wish to deploy the solution to.  
  This user also needs read/write permissions on the *Bestillinger* list.
- In order for the provisioning engine to be able to send status emails, it needs to be authenticated with an Office 365 user account that has an Exchange Online license.  
  This is the `emailRecipient` parameter in LogicApp.parameters.json.