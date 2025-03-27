# EPPlus.Samples.SensibilityLabels
A sample project how you use EPPlus with the Microsoft Information Protection SDK to handle sensibility labels on your Excel files.

## Getting started
To use this sample, you will need to the following:
* A Microsoft Purview account
* A Microsoft 365 Business Premium subscription or similar supporting sensibility labels.
* An application registration in Azure to get access your sensitivity labelling data.

## Microsoft Purview settings
If you don't have a Microsoft Purview account, you can create one at https://purview.microsoft.com/ using your Azure account.  
This tutorial will not cover the setup of your Microsoft Purview account, but you can find guide lines here: [Microsoft Purview setup guides](https://learn.microsoft.com/en-us/purview/purview-fast-track-setup-guides)  
Add at least one sensitivity label to use in the sample under *Information Protection* - *Sensibility labels* menu.

## Workstation setup & App registration
To use the Microsoft Information Protection API, you need to register an application in the [Azure portal](https://portal.azure.com) or the [Microsoft Entra admin center](https://entra.microsoft.com).  
To setup your application registration and your workstation, please see this guide [Microsoft Information Protection (MIP) SDK setup and configuration](https://learn.microsoft.com/en-us/information-protection/develop/setup-configure-mip).  
When the setup is done, in the application registration, add a redirection Uri to: http://localhost, if you run the sample from Visual Studio.  

## Running the sample
Before you can run the sample you will need to specify a few parameters in the SetupConstants.cs file.
* _tenantId - The Directory (tenant) ID from the App Registration's Overview page.
* _clientId - The Application (client) ID from the App Registration's Overview page.
* _appName - The application display name from the app registration.

* _loginAccount - The account used to login and get the access token. This account must have access to the protected content we are working with in the samples.
* _labelSample1 - The name of a sensibility label of you choice that is set on the workbook in sample 1.
* _protectedSampleFile - The path to a protected excel file that will be read and updated by EPPlus in sample 2, to demonstrate how to work with files protected by sensibility labels.1

