using OfficeOpenXml.Interfaces;
using OfficeOpenXml.Interfaces.SensitivityLabels;
namespace EPPlusSensibilityLabelHandler
{
    /// <summary>
    /// Implementation of the IPackageInfo interface, that passes information to EPPlus.
    /// </summary>
    public class PackageInformation : IPackageInfo
    {
        /// <summary>
        /// The package stream
        /// </summary>
        public MemoryStream PackageStream { get; set; }
        /// <summary>
        /// Protection information passed to EPPlus. 
        /// This property may hold the protection information or any other class used to hold information between the decryption operation and saving and applying the sensibility label.
        /// </summary>  
        public object ProtectionInformation { get; set; }
        /// <summary>
        /// The label id of the sensibility label to apply.
        /// </summary>
        public string ActiveLabelId { get; set; }
    }
}