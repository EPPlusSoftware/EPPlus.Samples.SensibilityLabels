using ConsoleApp2;
using Microsoft.InformationProtection;
using Microsoft.InformationProtection.File;
using Microsoft.InformationProtection.Protection;
using OfficeOpenXml.Interfaces.SensitivityLabels;
using SensibilityLabelHandler;
using System.Collections.ObjectModel;
namespace EPPlusSensibilityLabelHandler
{
    /// <summary>
    /// Example of a sensibility label handler 
    /// </summary>
    public class MySensibilityLabelHandler : ISensitivityLabelHandler
    {
        private IFileEngine _fileEngine;
        /// <summary>
        /// Initializes the MIP api and set up the authentication you want to use. 
        /// The sample below is just an example of how to authorize to your App. You should configure it according to your organization's setup and requirements.
        /// </summary>
        /// <returns></returns>
        public async Task InitAsync()
        {
            MIP.Initialize(MipComponent.File);

            //Create ApplicationInfo, setting the client ID from Microsoft Entra App Registration as the ApplicationId.
            ApplicationInfo appInfo = new ApplicationInfo()
            {
                ApplicationId = SetupConstants._clientId,
                ApplicationName = SetupConstants._appName,
                ApplicationVersion = "1.0.0"
            };

            MipConfiguration mipConfiguration = new MipConfiguration(appInfo, "mip_data", LogLevel.Trace, false);            
            MipContext mipContext = MIP.CreateMipContext(mipConfiguration);

            var profileSettings = new FileProfileSettings(mipContext,
                                     CacheStorageType.OnDiskEncrypted,
                                     new ConsentDelegateImplementation());

            
            // Load the Profile async and wait for the result.
            var fileProfile = await MIP.LoadFileProfileAsync(profileSettings);
            
            var authDelegate = new AuthDelegateImplementation(appInfo);

            var engineSettings = new FileEngineSettings(SetupConstants._loginAccount, authDelegate, "", "en-US");
            engineSettings.Identity = new Identity(SetupConstants._loginAccount);
            
            _fileEngine = await fileProfile.AddEngineAsync(engineSettings);
        }
        /// <summary>
        /// Decrypts the stream and returns it to EPPlus an unencrypted state.
        /// </summary>
        /// <param name="packageStream">The package stream to process. If the sensibility label has any type of protection, the stream is encrypted and must be decrypted before returning it to EPPlus.</param>
        /// <param name="id">The unique id for the package.</param>
        /// <returns>Returns the decrypted package and the protection information to EPPlus.</returns>
        public async Task<IPackageInfo> DecryptPackageAsync(MemoryStream packageStream, string id)
        {
            PackageInformation ret = new PackageInformation();
            var fileHandler = await _fileEngine.CreateFileHandlerAsync(packageStream, $@"myfile.xlsx", true);
            if (fileHandler.Protection != null)     //Is the stream encrypted?
            {
                //Yes, save the protection information and decrypt the stream.
                ret.ProtectionInformation = fileHandler.Protection;
                fileHandler.RemoveProtection();
                var ms = new MemoryStream();
                await fileHandler.CommitAsync(ms);
                ret.PackageStream = ms;
            }
            else
            {                
                //No, use the unencrypted stream directly.
                ret.PackageStream = packageStream;
            }
            return ret;
        }
        /// <summary>
        /// Applies a sensibility label and sets protection using the MIPS SDK.
        /// </summary>
        /// <param name="package">The package stream, protection information and the sensibility label to apply.</param>
        /// <param name="id">The unique id for the package.</param>
        /// <returns></returns>
        public async Task<MemoryStream> ApplyLabelAndSavePackageAsync(IPackageInfo package, string id)
        {
            package.PackageStream.Position = 0;
            var fileHandle = await _fileEngine.CreateFileHandlerAsync(package.PackageStream, $@"myfile.xlsx", true);            
            if (string.IsNullOrEmpty(package.ActiveLabelId) == false)
            {
                var l = _fileEngine.GetLabelById(package.ActiveLabelId);
                fileHandle.SetLabel(l, new LabelingOptions(), new ProtectionSettings() { });
            }
            if (package.ProtectionInformation is IProtectionHandler protection)
            {
                
                fileHandle.SetProtection(protection);
            }

            var ret = await fileHandle.CommitAsync(package.PackageStream);
            package.PackageStream.Position = 0;
            return package.PackageStream;
        }
        /// <summary>
        /// Returns all labels available for the package.
        /// </summary>
        /// <param name="id">The unique id for the package.</param>
        /// <returns>A collection of labels</returns>
        public IEnumerable<IExcelSensibilityLabel> GetLabels(string Id)
        {
            var list = new List<IExcelSensibilityLabel>();
            AddLabelCollectionToList(list, _fileEngine.SensitivityLabels);
            return list;
        }

        /// <summary>
        /// Updates the labels from EPPlus with properties missing in the package, such as Name, Description and Color.
        /// </summary>
        /// <param name="labels">The list of labels to update</param>
        /// <param name="Id">The unique id for the package.</param>
        public void UpdateLabelList(IEnumerable<IExcelSensibilityLabel> labels, string Id)
        {
            var lblDict = LabelsDictionary;
            foreach (var sl in labels)
            {
                UpdateItem(lblDict, sl);
            }
        }
        private void AddLabelCollectionToList(List<IExcelSensibilityLabel> list, ReadOnlyCollection<Label> sensitivityLabels)
        {
            foreach (var ss in sensitivityLabels.OrderBy(x => x.Parent))
            {
                list.Add(new SensibilityLabel()
                {
                    Id = ss.Id,
                    Name = ss.Name,
                    Description = ss.Description,
                    Tooltip = ss.Tooltip,
                    Enabled = ss.IsActive,
                    Removed = !ss.IsActive,
                    SiteId = SetupConstants._tenantId,
                    Color = ss.Color,
                    Method = ss.ActionSource == ActionSource.Automatic ? eMethod.Privileged : eMethod.Standard,
                    ContentBits = ss.Sensitivity == 0 ? 0 : eContentBits.Encryption
                });

                AddLabelCollectionToList(list, ss.Children);
            }
        }
        private static void UpdateItem(Dictionary<string, Label> lblDict, IExcelSensibilityLabel sl)
        {
            if (lblDict.TryGetValue(sl.Id, out Label? lbl))
            {
                if (sl is IExcelSensibilityLabelUpdate upd)
                {
                    IExcelSensibilityLabel parent;
                    if (lbl.Parent != null)
                    {
                        parent = GetLabel(lbl.Parent);
                    }
                    else
                    {
                        parent = null;
                    }

                    upd.Update(lbl.Name, lbl.Tooltip, lbl.Description, lbl.Color, parent);
                }
            }
        }
        private static IExcelSensibilityLabel GetLabel(Label lbl)
        {
            if(lbl==null)
            {
                return null;
            }
            return new SensibilityLabel()
            {
                Id = lbl.Id,
                Name = lbl.Name,
                Description = lbl.Description,
                Tooltip = lbl.Tooltip,
                Enabled = false,
                Removed = false,
                SiteId = SetupConstants._tenantId,
                Color = lbl.Color,
                Parent = GetLabel(lbl.Parent)
            };
        }
        Dictionary<string, Label> _labelsDictionary = null;
        private Dictionary<string, Label> LabelsDictionary
        {
            get
            {
                if (_labelsDictionary==null)
                {
                    _labelsDictionary = new Dictionary<string, Label>();
                    if (_fileEngine != null)
                    {
                        AddListToDict(_labelsDictionary, _fileEngine.SensitivityLabels);
                    }
                }

                return _labelsDictionary;
            }
        }
        private static void AddListToDict(Dictionary<string, Label> lbls, IList<Label> labels)
        {
            foreach (var l in labels)
            {
                if (lbls.ContainsKey(l.Id) == false)
                {
                    lbls.Add(l.Id, l);
                    if (l.Children != null && l.Children.Count > 0)
                    {
                        AddListToDict(lbls, l.Children);
                    }
                }
            }
        }
    }
}
