# EPPlus - Sensibility Label Sample
As part of Office 365 you can apply sensibility labels to you documents. A sensibility label is a meta data tag that can be set on a document. These tags can also enforce the documents to be encrypted, apply watermarks and texts in the headers and footers. To apply sensitivity labels you use the [Microsoft Information Protection SDK](https://learn.microsoft.com/en-us/information-protection/develop/). 

### Getting started
To use this sample, you will need to the following:
* A Microsoft Purview account
* A Microsoft 365 Business Premium subscription or similar supporting sensibility labels.
* An application registration in Azure to get access to your sensitivity labelling data.

### Microsoft Purview settings
If you don't have a Microsoft Purview account, you can create one at https://purview.microsoft.com/ using your Azure account.  
This tutorial will not cover the setup of your Microsoft Purview account, but you can find guide lines here: [Microsoft Purview setup guides](https://learn.microsoft.com/en-us/purview/purview-fast-track-setup-guides)  
Add at least one sensitivity label to use in the sample under *Information Protection* - *Sensibility labels* menu.

### Workstation setup & App registration
To use the Microsoft Information Protection API, you need to register an application ("App Registrations") in the [Azure portal](https//:portal.azure.com) or the [Microsoft Entra admin center](https://entra.microsoft.com).  
To setup your application registration and your workstation, please see this guide [Microsoft Information Protection (MIP) SDK setup and configuration](https://learn.microsoft.com/en-us/information-protection/develop/setup-configure-mip).  
When the setup is done, in the application registration, add a redirection Uri to: http://localhost, if you run the sample from Visual Studio.  

### Running the sample
Before you can run the sample you will need to specify a few parameters in the SetupConstants.cs file.
* _tenantId - The Directory (tenant) ID from the App Registration's Overview page.
* _clientId - The Application (client) ID from the App Registration's Overview page.
* _appName - The application display name from the app registration.

* _loginAccount - The account used to login and get the access token. This account must have access to the protected content we are working with in the samples.
* _labelSample1 - The name of a sensibility label of you choice that is set on the workbook in sample 1.
* _protectedSampleFile - The path to a protected excel file that will be read and updated by EPPlus in sample 2, to demonstrate how to work with files protected by sensibility labels.

### Integrate the MIP SDK with EPPlus
From EPPlus 8, you can add a sensibility label handler to EPPlus to more easily handle documents with Sensibility labels. 
EPPlus will identify packages protected by sensibility labels and call the Sensitivity Label Handler to decrypt and encrypt these workbooks.
To do so EPPlus has the interface: `ISensitivityLabelHandler`  
|Method|Description|
|:-----|:------------|
|`InitAsync`|Called to initiate the handler. This method is called once when the handler is assigned to EPPlus. Here you should initiate the MIP SDK and connect to the Microsoft Entra Application used to connect to you Microsoft Purview Account|
|`DecryptPackageAsync`|Called to decrypt a protected package. When EPPlus identifies a workbook to be encrypted with a sensibility label, the stream will be passed to this function for decryption.|  
|`ApplyLabelAndSavePackageAsync`|Called from EPPlus when the package has been saved and to apply the active label. EPPlus will supply the package stream and the sensibility label to be applied using the the MIP SDK. Returns a stream containing the output from the MIPS SDK.|  
|`UpdateLabelList`|Should update the supplied list of sensibility labels with name, description and other properties not present in the Sensibility Label XML document inside the package.|
|`GetLabels`|Get all labels from the MIPS SDK|

The `DecryptPackageAsync`,`ApplyLabelAndSavePackageAsync`, `UpdateLabelList` and the `GetLabels`  takes an `id` property, that is a unique identifier supplied by EPPlus to identify packages between the calls.

