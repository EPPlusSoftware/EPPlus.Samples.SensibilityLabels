using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SensibilityLabelHandler
{
    /// <summary>
    /// Constants used for authenticating to your Azure application, setting label names and output files and folders.
    /// </summary>
    public class SetupConstants
    {
        public const string _tenantId = "<The 'tenant id' for your Azure Application>";         //The tenent id from the Azure App Registration.
        public const string _clientId = "<The 'client id' for your Azure Application>";         //The application/client id from the Azure App Registration
        public const string _appName = "<'Application name' for your Azure application>";       //The display name from the Azure App Registration. 
        public const string _loginAccount = "your.name@yourdomain.com";                         //The account to used to login to your Azure application.

        public const string _outputPath = @"output";                                            //The output path, "the application base directory"/output by default.
        public const string _labelSample1 = "MySensibilityLabel";                               //The name of the sensibility label to apply in sample 1. The label must be present in your Microsoft Purview sensibility labels list.
        public const string _protectedSampleFile = @"Workbooks\MyProtectedExcelWorkbook.xlsx";  //The name of a sensibility label protected workbook to update in sample 2.
    }
}
