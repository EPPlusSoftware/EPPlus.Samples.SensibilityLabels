/***************************************************************************************************************************************
 * See this  https://github.com/EPPlusSoftware/EPPlus/wiki/Working-with-Sensibility-Labels
 * Setup your environment in the file: SetupConstants.cs
 *
 * Permissions for the app registration should look something like this:
 * API / Permissions name                      Type        Description                                     Admin consent required  Status
 * Azure Rights Management Services (1)        
 *     user_impersonation                      Delegated   Create and access protected content for users   No                      Granted for <your user/group>
 * Microsoft Graph (1)
 *     User.Read                               Delegated   Sign in and read user profile                   No                      Granted for <your user/group>
 * Microsoft Information Protection Sync Service (1)
 *     UnifiedPolicy.User.Read                 Delegated   Read all unified policies a user has access to. No                      Granted for 
****************************************************************************************************************************************/

using Microsoft.InformationProtection.File;
using Microsoft.InformationProtection;
using OfficeOpenXml;
using EPPlusSensibilityLabelHandler;
using Microsoft.InformationProtection.Protection;

namespace EPPlus.Samples.SensitivityLabel
{
    //https://learn.microsoft.com/en-us/information-protection/develop/
    class Program
    {
        static async Task Main(string[] args)
        {
            string outputPath;
            if (Path.IsPathRooted(SetupConstants._outputPath)) //Not a root path
            {
                outputPath = SetupConstants._outputPath;
            }
            else
            {
                outputPath = Path.Combine(AppContext.BaseDirectory, SetupConstants._outputPath); //Not a rooted path, put it under the base directory
            }

            //Set the license to Non-Commercial Personal. Licenses for commercial use can be purchased on https://epplussoftware.com/en/LicenseOverview/. 
            ExcelPackage.License.SetNonCommercialPersonal("EPPlus Sensibility Label Sample Project");

            var handler = new MySensibilityLabelHandler();
            ExcelPackage.SensibilityLabelHandler = handler;
            
            if(Directory.Exists(SetupConstants._outputPath)==false) Directory.CreateDirectory(SetupConstants._outputPath);

            var xlFilePath = $@"{SetupConstants._outputPath}\SensitivityLableEPPlus.xlsx";
            if (File.Exists(xlFilePath)) File.Delete(xlFilePath);
            await Sample1_Set_label_on_a_new_workbook(xlFilePath);
            await Sample2_Decrypt_a_protected_package(xlFilePath);
        }
        private static async Task Sample1_Set_label_on_a_new_workbook(string xlFilePath)
        {
            using var p = new ExcelPackage(xlFilePath);
            p.SensibilityLabels.SetActiveLabelByName(SetupConstants._labelSample1);
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = "This workbook is created from scratch";
            await p.SaveAsync();
        }
        private static async Task Sample2_Decrypt_a_protected_package(string xlFilePath)
        {
            var file = new FileInfo(xlFilePath);
            var newFile = new FileInfo(Path.Combine(file.Directory.FullName, "UpdatedByEPPlus" + file.Extension));
            if (file.Exists)
            {
                //Update the file with EPPlus
                using var p = new ExcelPackage(file);
                var ws = p.Workbook.Worksheets[0];
                ws.Workbook.Worksheets[0].Cells["A2"].Value = "This value is updated by EPPlus";
                await p.SaveAsAsync(newFile);
            }
        }

    }
}