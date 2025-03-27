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

using Microsoft.InformationProtection;
using Microsoft.InformationProtection.File;
using Microsoft.InformationProtection.Protection;
using OfficeOpenXml;
using SensibilityLabelHandler;

namespace EPPlus.Samples.SensitivityLabel
{
    //https://learn.microsoft.com/en-us/information-protection/develop/
    class Program
    {
        static IFileProfile _fileProfile;
        static MipContext _mipContext;
        static async Task Main(string[] args)
        {
            if (Directory.Exists(SetupConstants._outputPath) == false)
            {
                Directory.CreateDirectory(SetupConstants._outputPath);
            }
            //Set the license to Non-Commercial Personal. Licenses for commercial use can be purchased on https://epplussoftware.com/en/LicenseOverview/. 
            ExcelPackage.License.SetNonCommercialPersonal("EPPlus Sensibility Label Sample Project");

            var fileEngine = await SetupMIP();

            if (Directory.Exists(SetupConstants._outputPath) == false) Directory.CreateDirectory(SetupConstants._outputPath);
            var xlFilePath = $@"{SetupConstants._outputPath}\SensitivityLableEPPlus.xlsx";
            if(File.Exists(xlFilePath)) File.Delete(xlFilePath);

            await Sample1_Set_label_on_a_new_workbook(fileEngine, xlFilePath);
            await Sample2_Decrypt_a_protected_package(fileEngine, xlFilePath);

            fileEngine = null;
            _fileProfile = null;
            _mipContext.ShutDown();
            _mipContext = null;
        }
        private static async Task Sample1_Set_label_on_a_new_workbook(IFileEngine fileEngine, string xlFilePath)
        {
            var myLabel = fileEngine.SensitivityLabels.FirstOrDefault(x => x.Name == SetupConstants._labelSample1);
            if (myLabel == null)
            {
                throw new InvalidOperationException($"Cannot find label '{SetupConstants._labelSample1}' in the list of sensibility labels.");
            }
            using var ms = new MemoryStream();

            using var p = new ExcelPackage(ms);
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = "This workbook is created from scratch";
            p.Save();

            using var fileHandle = await fileEngine.CreateFileHandlerAsync(p.Stream, xlFilePath, true);
            var labelingOptions = new LabelingOptions()
            {
                AssignmentMethod = AssignmentMethod.Standard,
                IsDowngradeJustified = false,
            };

            var protection = new ProtectionSettings()
            {

            };

            fileHandle.SetLabel(myLabel, labelingOptions, protection);

            var fileStream = new FileStream(xlFilePath, FileMode.Create);
            await fileHandle.CommitAsync(fileStream);
            fileStream.Close();
        }

        private static async Task Sample2_Decrypt_a_protected_package(IFileEngine fileEngine, string xlFilePath)
        {
            var file = new FileInfo(xlFilePath);
            if (file.Exists)
            {
                //Read the file into a MemoryStream that we create the file handler on.
                using var stream = new MemoryStream(File.ReadAllBytes(xlFilePath));
                using var fileHandle = await fileEngine.CreateFileHandlerAsync(stream, file.FullName, true);
                IProtectionHandler protectionHandler;
                //If the file is protected, remove the protection and save it so it can be applied when we save the updated file.
                if (fileHandle.Protection != null)
                {
                    protectionHandler = fileHandle.Protection;
                    fileHandle.RemoveProtection();
                }
                else
                {
                    protectionHandler = null;
                }
                await fileHandle.CommitAsync(stream);

                //Update the file with EPPlus
                using var p = new ExcelPackage(stream);
                var ws = p.Workbook.Worksheets[0];
                ws.Workbook.Worksheets[0].Cells["A2"].Value = "This value is updated by EPPlus";
                p.Save();

                var outFile = @$"{SetupConstants._outputPath}\{file.Name}-updated{file.Extension}";
                if (protectionHandler !=null)
                {
                    using (var saveFileHandler = await fileEngine.CreateFileHandlerAsync(p.Stream, outFile, true))
                    {
                        saveFileHandler.SetProtection(protectionHandler);
                        await saveFileHandler.CommitAsync(outFile);
                    }
                }
                else
                {
                    File.WriteAllBytes(outFile, p.GetAsByteArray());
                }
            }
        }
        private static async Task<IFileEngine> SetupMIP()
        {
            // Initialize the MIP SDK
            MIP.Initialize(MipComponent.File);

            // Create ApplicationInfo, setting the clientID from Microsoft Entra App Registration as the ApplicationId.
            ApplicationInfo appInfo = new ApplicationInfo()
            {
                ApplicationId = SetupConstants._clientId,
                ApplicationName = SetupConstants._appName,
                ApplicationVersion = "1.0.0"
            };

            // Instantiate the AuthDelegateImpl object, passing in AppInfo.
            MipConfiguration mipConfiguration = new MipConfiguration(appInfo, "mip_data", LogLevel.Trace, false);

            // Create MipContext using Configuration
            _mipContext = MIP.CreateMipContext(mipConfiguration);

            // Initialize and instantiate the File Profile.
            // Create the FileProfileSettings object.
            // Initialize file profile settings to create/use local state.
            var profileSettings = new FileProfileSettings(_mipContext,
                                     CacheStorageType.OnDiskEncrypted,
                                     new ConsentDelegateImplementation());

            // Load the Profile async and wait for the result.
            _fileProfile = await MIP.LoadFileProfileAsync(profileSettings);
            AuthDelegateImplementation authDelegate = new AuthDelegateImplementation(appInfo);

            var engineSettings = new FileEngineSettings(SetupConstants._loginAccount, authDelegate, "", "en-US");
            engineSettings.Identity = new Identity(SetupConstants._loginAccount);

            return await _fileProfile.AddEngineAsync(engineSettings);
        }
    }
}
