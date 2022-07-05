using System;
using System.Threading.Tasks;
using Microsoft.InformationProtection;
using Microsoft.InformationProtection.Exceptions;
using Microsoft.InformationProtection.File;
using Microsoft.InformationProtection.Protection;

namespace MIPSDKDOTNET1
{
    class Program
    {
        private const string clientId = "bbcfb3e2-6b61-4a99-882a-f62e00b1f3c0";
        private const string appName = "Test-01";

        static void Main(string[] args)
        {
            // Initialize Wrapper for File SDK operations.
            MIP.Initialize(MipComponent.File);

            // Create ApplicationInfo, setting the clientID from Azure AD App Registration as the ApplicationId.
            ApplicationInfo appInfo = new ApplicationInfo()
            {
                ApplicationId = clientId,
                ApplicationName = appName,
                ApplicationVersion = "1.0.0"
            };

            // Instantiate the AuthDelegateImpl object, passing in AppInfo.
            AuthDelegateImplementation authDelegate = new AuthDelegateImplementation(appInfo);

            // Create MipConfiguration Object
            MipConfiguration mipConfiguration = new MipConfiguration(appInfo, "mip_data", LogLevel.Trace, false);

            // Create MipContext using Configuration
            MipContext mipContext = MIP.CreateMipContext(mipConfiguration);

            // Initialize and instantiate the File Profile.
            // Create the FileProfileSettings object.
            // Initialize file profile settings to create/use local state.
            var profileSettings = new FileProfileSettings(mipContext,
                                     CacheStorageType.OnDiskEncrypted,
                                     new ConsentDelegateImplementation());

            // Load the Profile async and wait for the result.
            var fileProfile = Task.Run(async () => await MIP.LoadFileProfileAsync(profileSettings)).Result;

            // Create a FileEngineSettings object, then use that to add an engine to the profile.
            // This pattern sets the engine ID to user1@tenant.com, then sets the identity used to create the engine.
            var engineSettings = new FileEngineSettings("new.worldtestid27@gmail.com", authDelegate, "", "en-US");
            engineSettings.Identity = new Identity("new.worldtestid27@gmail.com");

            var fileEngine = Task.Run(async () => await fileProfile.AddEngineAsync(engineSettings)).Result;
            string protectedFilePath = "C:\\Test\\Test_labeled.docx"; // Originally protected file's path from previous quickstart.

//Create a fileHandler for consumption for the Protected File.
            var protectedFileHandler = Task.Run(async () =>
                            await fileEngine.CreateFileHandlerAsync(protectedFilePath,// inputFilePath
                                                                    protectedFilePath,// actualFilePath
                                                                    false, //isAuditDiscoveryEnabled
                                                                    null)).Result; // fileExecutionState

            // Store protection handler from file
            var protectionHandler = protectedFileHandler.Protection;

            //Check if the user has the 'Edit' right to the file
            if (protectionHandler.AccessCheck("Edit"))
            {
                // Decrypt file to temp path
                var tempPath = Task.Run(async () => await protectedFileHandler.GetDecryptedTemporaryFileAsync()).Result;

                /*
                    Your own application code to edit the decrypted file belongs here. 
                */

                /// Follow steps below for re-protecting the edited file. ///
                // Create a new file handler using the temporary file path.
                var republishHandler = Task.Run(async () => await fileEngine.CreateFileHandlerAsync(tempPath, tempPath, false)).Result;

                // Set protection using the ProtectionHandler from the original consumption operation.
                republishHandler.SetProtection(protectionHandler);

                // New file path to save the edited file
                string reprotectedFilePath = "C:\\Test\\Test_labeled.docx";// New file path for saving reprotected file.

    // Write changes
                var reprotectedResult = Task.Run(async () => await republishHandler.CommitAsync(reprotectedFilePath)).Result;

                var protectedLabel = protectedFileHandler.Label;
                Console.WriteLine(string.Format("Originally protected file: {0}", protectedFilePath));
                Console.WriteLine(string.Format("File LabelID: {0} \r\nProtectionOwner: {1} \r\nIsProtected: {2}",
                                    protectedLabel.Label.Id,
                                    protectedFileHandler.Protection.Owner,
                                    protectedLabel.IsProtectionAppliedFromLabel.ToString()));
                var reprotectedLabel = republishHandler.Label;
                Console.WriteLine(string.Format("Reprotected file: {0}", reprotectedFilePath));
                Console.WriteLine(string.Format("File LabelID: {0} \r\nProtectionOwner: {1} \r\nIsProtected: {2}",
                                    reprotectedLabel.Label.Id,
                                    republishHandler.Protection.Owner,
                                    reprotectedLabel.IsProtectionAppliedFromLabel.ToString()));
                Console.WriteLine("Press a key to continue.");

                protectedFileHandler = null;
                protectionHandler = null;
                Console.ReadKey();

              
            }


        }
    }
}
