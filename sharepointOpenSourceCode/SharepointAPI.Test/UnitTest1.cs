using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.SharePoint.Client;
using System.Security;
using System.Web;


namespace AutomationAnywhere.MetaBot.SharePoint
{
    [TestClass]
    public class SharePointAPIWrapperTest
    {
        ClientContext context { get; set; }

        string siteUrl  = Resource4Test.siteURL;
        string siteName = Resource4Test.siteName;

        string username = Resource4Test.username;
        string password = Resource4Test.password;

        string folderRelativePath = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/";
        string remoteFilePath       = Resource4Test.remoteFilePath;
        string remoteFolderPath     = Resource4Test.remoteFolderPath;
        string remoteWorkFilePath   = Resource4Test.remoteWorkFilePath;
        string remoteWorkFolderPath = Resource4Test.remoteWorkFolderPath;

        string localFile   = Resource4Test.localFile;
        string localFolder = Resource4Test.localFolder;

        /// [TestMethod] // if you want to check the SharePoint Client API, you can check with this methods.
        public void CheckAPIFunctionality()
        {
            var mock = new SharePointAPIWrapper().AsDynamic();
            mock.Authenticate(siteUrl, username, password);
            this.context = mock.context as ClientContext;

            var workingFolder = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/target2";
            var folder = context.Web.GetFolderByServerRelativeUrl(workingFolder);
            context.Load(folder);
            context.ExecuteQuery();

            FolderCollection folderCollection = folder.Folders;
            context.Load(folderCollection);
            context.ExecuteQuery();

            Console.WriteLine("Folders: ");
            foreach (Folder folderItem in folderCollection)
            {
                context.Load(folderItem);
                context.ExecuteQuery();
                Console.WriteLine("  {0}, {1}", folderItem.Name, folderItem.ServerRelativeUrl);
            }

            FileCollection fileCollection = folder.Files;
            context.Load(fileCollection);
            context.ExecuteQuery();

            Console.WriteLine("Files: ");
            foreach (Microsoft.SharePoint.Client.File fileItem in fileCollection)
            {
                context.Load(fileItem);
                context.ExecuteQuery();
                Console.WriteLine("  {0}, {1}", fileItem.Name, fileItem.ServerRelativeUrl);
            }
            
        }

        /// Test Methods for Authenticate
        [TestMethod, TestCategory("Authenticate"), TestCategory("Success")]
        public void Authenticate()
        {
            var api = new SharePointAPIWrapper();
            var res = api.Authenticate(siteUrl, username, password);

            Assert.AreEqual($"Connected to {siteName}", res);
        }
        [TestMethod, TestCategory("Authenticate"), TestCategory("Fail")]
        public void AuthenticateIncorrectURL()
        {
            var incorrectURL = @"https://www.google.com";
            var api = new SharePointAPIWrapper();
            var res = api.Authenticate(incorrectURL, username, password);

            CheckFailCode(res);
            Console.WriteLine(res);
        }
        [TestMethod, TestCategory("Authenticate"), TestCategory("Fail")]
        public void AuthenticateIncorrectUser()
        {
            var incorrectUsername = "dummy user";
            var api = new SharePointAPIWrapper();
            var res = api.Authenticate(siteUrl, incorrectUsername, password);

            CheckFailCode(res);
            Console.WriteLine(res);
        }
        [TestMethod, TestCategory("Authenticate"), TestCategory("Fail")]
        public void AuthenticateIncorrectPassword()
        {
            var incorrectPassword = "dummy pass";
            var api = new SharePointAPIWrapper();
            var res = api.Authenticate(siteUrl, username, incorrectPassword);

            CheckFailCode(res);
            Console.WriteLine(res);
        }
        [TestMethod, TestCategory("Authenticate"), TestCategory("Fail")]
        public void AuthenticateWithNullURL()
        {
            var api = new SharePointAPIWrapper();
            var res = api.Authenticate(null, username, password);

            CheckFailCode(res);
            Console.WriteLine(res);
        }
        [TestMethod, TestCategory("Authenticate"), TestCategory("Fail")]
        public void AuthenticateWithNullUserName()
        {
            var api = new SharePointAPIWrapper();
            var res = api.Authenticate(siteUrl, null, password);

            CheckFailCode(res);
            Console.WriteLine(res);
        }
        [TestMethod, TestCategory("Authenticate"), TestCategory("Fail")]
        public void AuthenticateWithNullPassword()
        {
            var api = new SharePointAPIWrapper();
            var res = api.Authenticate(siteUrl, username, null);

            CheckFailCode(res);
            Console.WriteLine(res);
        }

        /// Test Methods for GetSiteTitle
        [TestMethod, TestCategory("GetSiteTitle"), TestCategory("Success")]
        public void GetSiteTitle()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);
            var res = api.GetSiteTitle();

            // validate
            Assert.AreEqual(siteName, res);
        }
        [TestMethod, TestCategory("GetSiteTitle"), TestCategory("Fail")]
        public void GetSiteTitleWitoutContext()
        {
            var api = new SharePointAPIWrapper();

            var res = api.GetSiteTitle();
            CheckAuthenticationFail(res);
        }

        /// Here is an integration testing method for File operation.
        [TestMethod, TestCategory("Integrated Test"), TestCategory("Success")]
        public void FileOperation()
        {
            /// Prerequisite: The file is under {sourceFolder}
            /// Scenario
            ///   1. Copy a file, {sourceFolder}/{sourceFile}, to {targetFolder}/{targetFile}
            ///   2. Move a file, {targetFolder}/{targetFile}, to {targetFolder2}/{targetFile2}
            ///   3. Check Out a file, {targetFolder2}/{targetFile2}
            ///   4. Undo Check Out a file, {targetFolder2}/{targetFile2}
            ///   5. Check Out a file again, {targetFolder2}/{targetFile2}
            ///   6. Download a file, {targetFolder2}/{targetFile2}
            ///   7. Modify a file (Not a dll method), {targetFolder2}/{targetFile2}
            ///   8. Upload a file, {targetFolder2}/{targetFile2}
            ///   9. Check in a file, {targetFolder2}/{targetFile2}
            ///   10. Delete a file, {targetFolder2}/{targetFile2}
            /// 
            var sourceFolder = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin";
            var targetFolder = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/target";
            var targetFolder2 = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/target2";
            var sourceFile = "sample.txt";
            var targetFile = "sample.txt";
            var targetFile2 = "sample2.txt";
            var sourcePath = "";
            var destinationPath = "";
            var res = "";

            var api = this.InitializeWrapper(siteUrl, username, password);

            ///   1. Copy a file, {sourceFolder}/{sourceFile}, to {targetFolder}/{targetFile}
            sourcePath = $"{sourceFolder}/{sourceFile}";
            destinationPath = $"{targetFolder}/{targetFile}";
            res = api.CopyFile(sourcePath, destinationPath);
            Console.WriteLine(res);
            Assert.AreEqual($"Copied the file from {sourcePath} to {destinationPath}", res);

            ///   2. Move a file to xxx
            sourcePath = destinationPath;
            destinationPath = $"{targetFolder2}/{targetFile2}";
            res = api.MoveFile(sourcePath, destinationPath);
            Console.WriteLine(res);
            Assert.AreEqual($"Moved the file from {sourcePath} to {destinationPath}", res);

            ///   3. Check Out a file
            res = api.CheckOutFile(destinationPath);
            Console.WriteLine(res);
            Assert.AreEqual($"Checkout the file located on {destinationPath}", res);

            ///   4. Undo Check Out a file
            res = api.UndoCheckOutFile(destinationPath);
            Console.WriteLine(res);
            Assert.AreEqual($"Undo Checkout the File located on {destinationPath}", res);

            ///   5. Check Out a file again
            res = api.CheckOutFile(destinationPath);
            Console.WriteLine(res);
            Assert.AreEqual($"Checkout the file located on {destinationPath}", res);

            ///   6. Download a file
            string tempDir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            var tempDirInfo = Directory.CreateDirectory(tempDir);
            var localFile = Path.Combine(tempDirInfo.FullName, targetFile2);

            res = api.DownloadFile(destinationPath, localFile);
            Console.WriteLine(res);
            Assert.AreEqual($"File Downloaded from {destinationPath} to {localFile}", res);

            ///   7. Modify a file (Not a dll method)
            using (StreamWriter sw = System.IO.File.AppendText(localFile))
            {                
                sw.WriteLine(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + ": This is added from unit test");
            }

            ///   8. Upload a file
            res = api.UploadFile(localFile, destinationPath);
            Console.WriteLine(res);
            Assert.AreEqual($"File uploaded from {localFile} to {destinationPath}", res);

            ///   9. Check in a file.
            var comment = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + ":This is from unit test";
            res = api.CheckInFile(destinationPath, comment);
            Console.WriteLine(res);
            Assert.AreEqual($"Checkin the file located on {destinationPath}", res);

            ///   10. Delete a file
            res = api.DeleteFile(destinationPath);
            Console.WriteLine(res);
            Assert.AreEqual($"Deleted the file located on {destinationPath}", res);
        }

        /// Here is an integration testing method for Folder operation.
        [TestMethod, TestCategory("Integrated Test"), TestCategory("Success")]
        public void FolderOperation()
        {
            /// Prerequisite: The original folder is {sourceFolder}
            /// Scenario
            ///   1. Create a Folder
            ///   2. Copy a Folder under 1
            ///   3. Move a Folder
            ///   4. Download a Folder
            ///   5. Upload a Folder
            ///   6. Delete a Folder
            /// 
            var sourceFolder = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin";
            var res = "";

            var api = this.InitializeWrapper(siteUrl, username, password);

            ///   1. Create a Folder
            var parentFolder = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest";
            var newFolderName = "new";
            var newFolder = parentFolder + "/" + newFolderName;
            res = api.CreateFolder(parentFolder, newFolderName);
            Console.WriteLine(res);
            Assert.AreEqual($"Create the folder located in {newFolder}", res);

            ///   2. Copy a Folder under 1
            var workFolder = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/workFolder";
            res = api.CopyFolder(sourceFolder, workFolder, 3);
            Console.WriteLine(res);
            Assert.AreEqual($"Copied the folder from {sourceFolder} to {workFolder}", res);

            ///   3. Move a Folder
            var renamedFolder = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/renamed";
            res = api.MoveFolder(workFolder, renamedFolder);
            Console.WriteLine(res);
            Assert.AreEqual($"Moved the folder from {workFolder} to {renamedFolder}", res);

            ///   4. Download a Folder
            string tempDir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            var tempDirInfo = Directory.CreateDirectory(tempDir);

            res = api.DownloadFolder(renamedFolder, tempDirInfo.FullName, 3);
            Console.WriteLine(res);
            Assert.AreEqual($"Downloaded the folder from {renamedFolder} to {tempDirInfo.FullName}", res);

            ///   5. Upload a Folder
            res = api.UploadFolder(tempDirInfo.FullName, renamedFolder, 3);
            Console.WriteLine(res);
            Assert.AreEqual($"Uploaded the folder from {tempDirInfo.FullName} to {renamedFolder}", res);

            ///   6. Delete a Folder
            res = api.DeleteFolder(renamedFolder);
            Console.WriteLine(res);
            Assert.AreEqual($"Deleted the folder located on {renamedFolder}", res);
            res = api.DeleteFolder(newFolder);
            Console.WriteLine(res);
            Assert.AreEqual($"Deleted the folder located on {newFolder}", res);
        }

        /// Test Methods for Download file.
        [TestMethod, TestCategory("DownloadFile Method"), TestCategory("Success")]
        public void DownloadFile()
        {
            /// cleanup and prepare the parent folder.
            this.CleanUpLocalFile(localFile);
            this.CreateLocalParentDirectory(localFile);

            var api = this.InitializeWrapper(siteUrl, username, password);
            var res = api.DownloadFile(remoteFilePath, localFile);

            Assert.AreEqual($"File Downloaded from {remoteFilePath} to {localFile}", res);
        }
        [TestMethod, TestCategory("DownloadFile Method"), TestCategory("Success")]
        public void DownloadFileTwice()
        {
            /// cleanup and prepare the parent folder.
            this.CleanUpLocalFile(localFile);
            this.CreateLocalParentDirectory(localFile);

            var api = this.InitializeWrapper(siteUrl, username, password);
            var res = api.DownloadFile(remoteFilePath, localFile);
            Assert.AreEqual($"File Downloaded from {remoteFilePath} to {localFile}", res);
            res = api.DownloadFile(remoteFilePath, localFile);
            Assert.AreEqual($"File Downloaded from {remoteFilePath} to {localFile}", res);
        }
        [TestMethod, TestCategory("DownloadFile Method"), TestCategory("Fail")]
        public void DownloadFileWithoutContext()
        {
            var api = new SharePointAPIWrapper(); /// no authentication 

            var res = api.DownloadFile(remoteFilePath, localFile);
            CheckAuthenticationFail(res);
        }
        [TestMethod, TestCategory("DownloadFile Method"), TestCategory("Fail")]
        public void DownloadFileFileWithFolderPath()
        {
            /// working file
            var workingFile = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.DownloadFile(workingFile, localFile);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to checkout the file with error: Unknown Error
        }
        [TestMethod, TestCategory("DownloadFile Method"), TestCategory("Fail")]
        public void DownloadFileWithMissingFile()
        {
            /// Missing file
            var workingFile = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin/missing.txt";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.DownloadFile(workingFile, localFile);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to checkout the file with error: File Not Found.
        }
        [TestMethod, TestCategory("DownloadFile Method"), TestCategory("Fail")]
        public void DownloadFileWithTeamsURL() 
        {
            var workingFile = @"https://automationanywhere1.sharepoint.com/sites/SolutionArchitectTeam/Shared%20Documents/General/Working_Files_Naoya/MetabotTest/origin/sample.txt";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.DownloadFile(workingFile, localFile);
            CheckFailCode(res);
            Console.WriteLine(res);
            // FAIL:Fail to checkout the file with error: serverRelativeUrl
            // Parameter name: Specified value is not supported for the serverRelativeUrl parameter.
        }
        [TestMethod, TestCategory("DownloadFile Method"), TestCategory("Fail")]
        public void DownloadFileWithSharePointURL()
        {
            /// Getting URL from browser's address bar.
            var workingFile = @"https://automationanywhere1.sharepoint.com/sites/SolutionArchitectTeam/Shared%20Documents/Forms/AllItems.aspx?FolderCTID=0x0120000223E4C4F9AB2C4DA1E560A6277B76AB&id=%2Fsites%2FSolutionArchitectTeam%2FShared%20Documents%2FGeneral%2FWorking_Files_Naoya%2FMetabotTest%2Forigin%2Fsample%2Etxt&parent=%2Fsites%2FSolutionArchitectTeam%2FShared%20Documents%2FGeneral%2FWorking_Files_Naoya%2FMetabotTest%2Forigin";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.DownloadFile(workingFile, localFile);
            CheckFailCode(res);
            Console.WriteLine(res);
            // FAIL:Fail to checkout the file with error: serverRelativeUrl
            // Parameter name: Specified value is not supported for the serverRelativeUrl parameter.
        }
        [TestMethod, TestCategory("DownloadFile Method"), TestCategory("Fail")]
        public void DownloadFileWithSharePointLink()
        {
            /// Getting Link from SharePoint
            var workingFile = @"https://automationanywhere1.sharepoint.com/:t:/s/SolutionArchitectTeam/EYm5Qip7k2tCgKCIFWNefr8BMYVMimJgG3UHlkcQJHFdvQ?e=cqtfc8";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.DownloadFile(workingFile, localFile);
            CheckFailCode(res);
            Console.WriteLine(res);
            // FAIL:Fail to checkout the file with error: serverRelativeUrl
            // Parameter name: Specified value is not supported for the serverRelativeUrl parameter.
        }
        [TestMethod, TestCategory("DownloadFile Method"), TestCategory("Fail")]
        public void DownloadFileWithNullRemoteFilePath()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.DownloadFile(null, localFile);
            CheckFailCode(res);
            Console.WriteLine(res);
        }
        [TestMethod, TestCategory("DownloadFile Method"), TestCategory("Fail")]
        public void DownloadFileNullLocalFile()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.DownloadFile(remoteFilePath, null);
            CheckFailCode(res);
            Console.WriteLine(res);
        }
        [TestMethod, TestCategory("DownloadFile Method"), TestCategory("Fail")]
        public void DownloadFileWithMissingFolder()
        {
            var missingFolder = Path.Combine(remoteFolderPath + Path.GetRandomFileName());
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.DownloadFile(remoteFilePath, missingFolder);
            CheckFailCode(res);
            Console.WriteLine(res);
        }

        [TestMethod, TestCategory("UploadFile Method"), TestCategory("Success")]
        public void UploadFile()
        {
            /// cleanup and prepare the parent folder.
            this.CreateLocalParentDirectory(localFile);
            this.CreateLocalFile(localFile);
            this.AddTextToLocalFile(localFile);

            var api = this.InitializeWrapper(siteUrl, username, password);
            var res = api.UploadFile(localFile, remoteFilePath);
            Assert.AreEqual($"File uploaded from {localFile} to {remoteFilePath}", res);
        }
        [TestMethod, TestCategory("UploadFile Method"), TestCategory("Success")]
        public void UploadFileTwice()
        {
            /// cleanup and prepare the parent folder.
            this.CreateLocalParentDirectory(localFile);
            this.CreateLocalFile(localFile);
            this.AddTextToLocalFile(localFile);

            var api = this.InitializeWrapper(siteUrl, username, password);
            var res = api.UploadFile(localFile, remoteFilePath);
            Assert.AreEqual($"File uploaded from {localFile} to {remoteFilePath}", res);

            this.AddTextToLocalFile(localFile);
            res = api.UploadFile(localFile, remoteFilePath);
            Assert.AreEqual($"File uploaded from {localFile} to {remoteFilePath}", res);
        }
        [TestMethod, TestCategory("UploadFile Method"), TestCategory("Success")]
        public void UploadFileWithRename()
        {
            /// cleanup and prepare the parent folder.
            this.CreateLocalParentDirectory(localFile);
            this.CreateLocalFile(localFile);
            this.AddTextToLocalFile(localFile);

            /// Missing file
            var workingFile = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin/renamed_sample.txt";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.UploadFile(localFile, workingFile);
            Assert.AreEqual($"File uploaded from {localFile} to {workingFile}", res);
        }
        [TestMethod, TestCategory("UploadFile Method"), TestCategory("Fail")]
        public void UploadFileWithoutContext()
        {
            /// cleanup and prepare the parent folder.
            this.CreateLocalParentDirectory(localFile);
            this.CreateLocalFile(localFile);
            this.AddTextToLocalFile(localFile);

            var api = new SharePointAPIWrapper(); /// no authentication 

            var res = api.UploadFile(localFile, remoteFilePath);
            CheckAuthenticationFail(res);
        }
        [TestMethod, TestCategory("UploadFile Method"), TestCategory("Fail")]
        public void UploadFileWithFolderPath()
        {
            /// cleanup and prepare the parent folder.
            this.CreateLocalParentDirectory(localFile);
            this.CreateLocalFile(localFile);
            this.AddTextToLocalFile(localFile);

            /// working file
            var workingFile = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.UploadFile(localFile, workingFile);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to checkout the file with error: Unknown Error
        }
        [TestMethod, TestCategory("UploadFile Method"), TestCategory("Fail")]
        public void UploadFileWithMissingFolder()
        {
            /// cleanup and prepare the parent folder.
            this.CreateLocalParentDirectory(localFile);
            this.CreateLocalFile(localFile);
            this.AddTextToLocalFile(localFile);

            /// Missing file
            var workingFile = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin/missing/renamed_sample.txt";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.UploadFile(localFile, workingFile);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to checkout the file with error: File Not Found.
        }
        [TestMethod, TestCategory("UploadFile Method"), TestCategory("Fail")]
        public void UploadFileWithTeamsURL()
        {
            /// cleanup and prepare the parent folder.
            this.CreateLocalParentDirectory(localFile);
            this.CreateLocalFile(localFile);
            this.AddTextToLocalFile(localFile);

            var workingFile = @"https://automationanywhere1.sharepoint.com/sites/SolutionArchitectTeam/Shared%20Documents/General/Working_Files_Naoya/MetabotTest/origin/sample.txt";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.UploadFile(localFile, workingFile);
            CheckFailCode(res);
            Console.WriteLine(res);
            // FAIL:Fail to checkout the file with error: serverRelativeUrl
            // Parameter name: Specified value is not supported for the serverRelativeUrl parameter.
        }
        [TestMethod, TestCategory("UploadFile Method"), TestCategory("Fail")]
        public void UploadFileWithSharePointURL()
        {
            /// cleanup and prepare the parent folder.
            this.CreateLocalParentDirectory(localFile);
            this.CreateLocalFile(localFile);
            this.AddTextToLocalFile(localFile);

            /// Getting URL from browser's address bar.
            var workingFile = @"https://automationanywhere1.sharepoint.com/sites/SolutionArchitectTeam/Shared%20Documents/Forms/AllItems.aspx?FolderCTID=0x0120000223E4C4F9AB2C4DA1E560A6277B76AB&id=%2Fsites%2FSolutionArchitectTeam%2FShared%20Documents%2FGeneral%2FWorking_Files_Naoya%2FMetabotTest%2Forigin%2Fsample%2Etxt&parent=%2Fsites%2FSolutionArchitectTeam%2FShared%20Documents%2FGeneral%2FWorking_Files_Naoya%2FMetabotTest%2Forigin";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.UploadFile(localFile, workingFile);
            CheckFailCode(res);
            Console.WriteLine(res);
            // FAIL:Fail to checkout the file with error: serverRelativeUrl
            // Parameter name: Specified value is not supported for the serverRelativeUrl parameter.
        }
        [TestMethod, TestCategory("UploadFile Method"), TestCategory("Fail")]
        public void UploadFileWithSharePointLink()
        {
            /// cleanup and prepare the parent folder.
            this.CreateLocalParentDirectory(localFile);
            this.CreateLocalFile(localFile);
            this.AddTextToLocalFile(localFile);

            /// Getting Link from SharePoint
            var workingFile = @"https://automationanywhere1.sharepoint.com/:t:/s/SolutionArchitectTeam/EYm5Qip7k2tCgKCIFWNefr8BMYVMimJgG3UHlkcQJHFdvQ?e=cqtfc8";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.UploadFile(localFile, workingFile);
            CheckFailCode(res);
            Console.WriteLine(res);
            // FAIL:Fail to checkout the file with error: serverRelativeUrl
            // Parameter name: Specified value is not supported for the serverRelativeUrl parameter.
        }
        [TestMethod, TestCategory("UploadFile Method"), TestCategory("Fail")]
        public void UploadFileWithNullLocalFilePath()
        {
            /// cleanup and prepare the parent folder.
            this.CreateLocalParentDirectory(localFile);
            this.CreateLocalFile(localFile);
            this.AddTextToLocalFile(localFile);

            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.UploadFile(null, remoteFilePath);
            CheckFailCode(res);
            Console.WriteLine(res);
        }
        [TestMethod, TestCategory("UploadFile Method"), TestCategory("Fail")]
        public void UploadFileNullRemoteFilePath()
        {
            /// cleanup and prepare the parent folder.
            this.CreateLocalParentDirectory(localFile);
            this.CreateLocalFile(localFile);
            this.AddTextToLocalFile(localFile);

            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.UploadFile(localFile, null);
            CheckFailCode(res);
            Console.WriteLine(res);
        }

        /// Tests for CheckOutFile method
        [TestMethod, TestCategory("CheckOutFile Method"), TestCategory("Success")]
        public void CheckOutFile()
        {
            /// working file
            var workingFile = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin/sample.txt";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CheckOutFile(workingFile);
            Assert.AreEqual($"Checkout the file located on {workingFile}", res);
        }
        [TestMethod, TestCategory("CheckOutFile Method"), TestCategory("Success")]
        public void CheckOutFileTwice()
        {
            /// working file
            var workingFile = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin/sample.txt";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CheckOutFile(workingFile);
            Assert.AreEqual($"Checkout the file located on {workingFile}", res);
            res = api.CheckOutFile(workingFile);
            Assert.AreEqual($"Checkout the file located on {workingFile}", res);
        }
        [TestMethod, TestCategory("CheckOutFile Method"), TestCategory("Fail")]
        public void CheckOutFileWithoutContext()
        {
            /// working file
            var workingFile = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin/sample.txt";
            var api = new SharePointAPIWrapper(); /// no authentication 

            var res = api.CheckOutFile(workingFile);
            CheckAuthenticationFail(res);
        }
        [TestMethod, TestCategory("CheckOutFile Method"), TestCategory("Fail")]
        public void CheckOutFileWithFolderPath()
        {
            /// working file
            var workingFile = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CheckOutFile(workingFile);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to checkout the file with error: Unknown Error
        }
        [TestMethod, TestCategory("CheckOutFile Method"), TestCategory("Fail")]
        public void CheckOutFileWithMissingFile()
        {
            /// Missing file
            var workingFile = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin/missing.txt";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CheckOutFile(workingFile);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to checkout the file with error: File Not Found.
        }
        [TestMethod, TestCategory("CheckOutFile Method"), TestCategory("Fail")]
        public void CheckOutFileWithTeamsURL()
        {
            /// Missing Folder
            var workingFile = @"https://automationanywhere1.sharepoint.com/sites/SolutionArchitectTeam/Shared%20Documents/General/Working_Files_Naoya/MetabotTest/origin/sample.txt";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CheckOutFile(workingFile);
            CheckFailCode(res);
            Console.WriteLine(res);
              // FAIL:Fail to checkout the file with error: serverRelativeUrl
              // Parameter name: Specified value is not supported for the serverRelativeUrl parameter.
        }
        [TestMethod, TestCategory("CheckOutFile Method"), TestCategory("Fail")]
        public void CheckOutFileWithSharePointURL()
        {
            /// Getting URL from browser's address bar.
            var workingFile = @"https://automationanywhere1.sharepoint.com/sites/SolutionArchitectTeam/Shared%20Documents/Forms/AllItems.aspx?FolderCTID=0x0120000223E4C4F9AB2C4DA1E560A6277B76AB&id=%2Fsites%2FSolutionArchitectTeam%2FShared%20Documents%2FGeneral%2FWorking_Files_Naoya%2FMetabotTest%2Forigin%2Fsample%2Etxt&parent=%2Fsites%2FSolutionArchitectTeam%2FShared%20Documents%2FGeneral%2FWorking_Files_Naoya%2FMetabotTest%2Forigin";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CheckOutFile(workingFile);
            CheckFailCode(res);
            Console.WriteLine(res);
              // FAIL:Fail to checkout the file with error: serverRelativeUrl
              // Parameter name: Specified value is not supported for the serverRelativeUrl parameter.
        }
        [TestMethod, TestCategory("CheckOutFile Method"), TestCategory("Fail")]
        public void CheckOutFileWithSharePointLink()
        {
            /// Getting Link from SharePoint
            var workingFile = @"https://automationanywhere1.sharepoint.com/:t:/s/SolutionArchitectTeam/EYm5Qip7k2tCgKCIFWNefr8BMYVMimJgG3UHlkcQJHFdvQ?e=cqtfc8";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CheckOutFile(workingFile);
            CheckFailCode(res);
            Console.WriteLine(res); 
              // FAIL:Fail to checkout the file with error: serverRelativeUrl
              // Parameter name: Specified value is not supported for the serverRelativeUrl parameter.
        }
        [TestMethod, TestCategory("CheckOutFile Method"), TestCategory("Fail")]
        public void CheckOutFileWithNull()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CheckOutFile(null);
            CheckFailCode(res);
            Console.WriteLine(res);
        }

        /// Tests for CheckInFile method
        [TestMethod, TestCategory("CheckInFile Method"), TestCategory("Success")]
        public void CheckInFile()
        {
            /// working file
            var workingFile = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin/sample.txt";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CheckOutFile(workingFile);
            Assert.AreEqual($"Checkout the file located on {workingFile}", res);
            res = api.CheckInFile(workingFile, "checkin via unit testing");
            Assert.AreEqual($"Checkin the file located on {workingFile}", res);
        }
        [TestMethod, TestCategory("CheckInFile Method"), TestCategory("Success")]
        public void CheckInFileTwice()
        {
            /// working file
            var workingFile = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin/sample.txt";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CheckOutFile(workingFile);
            Assert.AreEqual($"Checkout the file located on {workingFile}", res);
            res = api.CheckInFile(workingFile, "checkin via unit testing");
            Assert.AreEqual($"Checkin the file located on {workingFile}", res);
            res = api.CheckInFile(workingFile, "checkin via unit testing");
            Assert.AreEqual($"Checkin the file located on {workingFile}", res);
        }
        [TestMethod, TestCategory("CheckInFile Method"), TestCategory("Success")]
        public void CheckInFileWithoutCheckOut()
        {
            /// working file
            var workingFile = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin/sample.txt";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CheckInFile(workingFile, "checkin via unit testing");
            Assert.AreEqual($"Checkin the file located on {workingFile}", res);
        }
        [TestMethod, TestCategory("CheckInFile Method"), TestCategory("Fail")]
        public void CheckInFileWithoutContext()
        {
            /// working file
            var workingFile = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin/sample.txt";
            var api = new SharePointAPIWrapper(); /// no authentication 

            var res = api.CheckInFile(workingFile, "checkin via unit testing");
            CheckAuthenticationFail(res);
        }
        [TestMethod, TestCategory("CheckInFile Method"), TestCategory("Fail")]
        public void CheckInFileWithFolderPath()
        {
            /// working file
            var workingFile = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CheckInFile(workingFile, "checkin via unit testing");
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to checkout the file with error: Unknown Error
        }
        [TestMethod, TestCategory("CheckInFile Method"), TestCategory("Fail")]
        public void CheckInFileWithMissingFile()
        {
            /// Missing file
            var workingFile = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin/missing.txt";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CheckInFile(workingFile, "checkin via unit testing");
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to checkout the file with error: File Not Found.
        }
        [TestMethod, TestCategory("CheckInFile Method"), TestCategory("Fail")]
        public void CheckInFileWithTeamsURL()
        {
            /// Missing Folder
            var workingFile = @"https://automationanywhere1.sharepoint.com/sites/SolutionArchitectTeam/Shared%20Documents/General/Working_Files_Naoya/MetabotTest/origin/sample.txt";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CheckInFile(workingFile, "checkin via unit testing");
            CheckFailCode(res);
            Console.WriteLine(res);
            // FAIL:Fail to checkout the file with error: serverRelativeUrl
            // Parameter name: Specified value is not supported for the serverRelativeUrl parameter.
        }
        [TestMethod, TestCategory("CheckInFile Method"), TestCategory("Fail")]
        public void CheckInFileWithSharePointURL()
        {
            /// Getting URL from browser's address bar.
            var workingFile = @"https://automationanywhere1.sharepoint.com/sites/SolutionArchitectTeam/Shared%20Documents/Forms/AllItems.aspx?FolderCTID=0x0120000223E4C4F9AB2C4DA1E560A6277B76AB&id=%2Fsites%2FSolutionArchitectTeam%2FShared%20Documents%2FGeneral%2FWorking_Files_Naoya%2FMetabotTest%2Forigin%2Fsample%2Etxt&parent=%2Fsites%2FSolutionArchitectTeam%2FShared%20Documents%2FGeneral%2FWorking_Files_Naoya%2FMetabotTest%2Forigin";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CheckInFile(workingFile, "checkin via unit testing");
            CheckFailCode(res);
            Console.WriteLine(res);
            // FAIL:Fail to checkout the file with error: serverRelativeUrl
            // Parameter name: Specified value is not supported for the serverRelativeUrl parameter.
        }
        [TestMethod, TestCategory("CheckInFile Method"), TestCategory("Fail")]
        public void CheckInFileWithSharePointLink()
        {
            /// Getting Link from SharePoint
            var workingFile = @"https://automationanywhere1.sharepoint.com/:t:/s/SolutionArchitectTeam/EYm5Qip7k2tCgKCIFWNefr8BMYVMimJgG3UHlkcQJHFdvQ?e=cqtfc8";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CheckInFile(workingFile, "checkin via unit testing");
            CheckFailCode(res);
            Console.WriteLine(res);
            // FAIL:Fail to checkout the file with error: serverRelativeUrl
            // Parameter name: Specified value is not supported for the serverRelativeUrl parameter.
        }
        [TestMethod, TestCategory("CheckInFile Method"), TestCategory("Fail")]
        public void CheckInFileWithNullPath()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CheckInFile(null, "checkin via unit testing");
            CheckFailCode(res);
            Console.WriteLine(res);
        }
        [TestMethod, TestCategory("CheckInFile Method"), TestCategory("Success")]
        public void CheckInFileWithNullComment()
        {
            /// working file
            var workingFile = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin/sample.txt";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CheckOutFile(workingFile);
            Assert.AreEqual($"Checkout the file located on {workingFile}", res);
            res = api.CheckInFile(workingFile, null);
            Assert.AreEqual($"Checkin the file located on {workingFile}", res);
        }

        /// Tests for UndoCheckOut method
        [TestMethod, TestCategory("UndoCheckOutFile Method"), TestCategory("Success")]
        public void UndoCheckOutFile()
        {
            /// working file
            var workingFile = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin/sample.txt";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CheckOutFile(workingFile);
            Assert.AreEqual($"Checkout the file located on {workingFile}", res);
            res = api.UndoCheckOutFile(workingFile);
            Assert.AreEqual($"Undo Checkout the File located on {workingFile}", res);
        }
        [TestMethod, TestCategory("UndoCheckOutFile Method"), TestCategory("Success")]
        public void UndoCheckOutFileTwice()
        {
            /// working file
            var workingFile = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin/sample.txt";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CheckOutFile(workingFile);
            Assert.AreEqual($"Checkout the file located on {workingFile}", res);
            res = api.UndoCheckOutFile(workingFile);
            Assert.AreEqual($"Undo Checkout the File located on {workingFile}", res);
            res = api.UndoCheckOutFile(workingFile);
            Assert.AreEqual($"Undo Checkout the File located on {workingFile}", res);
        }
        [TestMethod, TestCategory("UndoCheckOutFile Method"), TestCategory("Success")]
        public void UndoCheckOutFileWithoutCheckOut()
        {
            /// working file
            var workingFile = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin/sample.txt";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.UndoCheckOutFile(workingFile);
            Assert.AreEqual($"Undo Checkout the File located on {workingFile}", res);
        }
        [TestMethod, TestCategory("UndoCheckOutFile Method"), TestCategory("Fail")]
        public void UndoCheckOutFileFileWithoutContext()
        {
            /// working file
            var workingFile = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin/sample.txt";
            var api = new SharePointAPIWrapper(); /// no authentication 

            var res = api.UndoCheckOutFile(workingFile);
            CheckAuthenticationFail(res);
        }
        [TestMethod, TestCategory("UndoCheckOutFile Method"), TestCategory("Fail")]
        public void UndoCheckOutFileWithFolderPath()
        {
            /// working file
            var workingFile = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.UndoCheckOutFile(workingFile);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to checkout the file with error: Unknown Error
        }
        [TestMethod, TestCategory("UndoCheckOutFile Method"), TestCategory("Fail")]
        public void UndoCheckOutFileWithMissingFile()
        {
            /// Missing file
            var workingFile = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin/missing.txt";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.UndoCheckOutFile(workingFile);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to checkout the file with error: File Not Found.
        }
        [TestMethod, TestCategory("UndoCheckOutFile Method"), TestCategory("Fail")]
        public void UndoCheckOutFileWithTeamsURL()
        {
            /// Missing Folder
            var workingFile = @"https://automationanywhere1.sharepoint.com/sites/SolutionArchitectTeam/Shared%20Documents/General/Working_Files_Naoya/MetabotTest/origin/sample.txt";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.UndoCheckOutFile(workingFile);
            CheckFailCode(res);
            Console.WriteLine(res);
            // FAIL:Fail to checkout the file with error: serverRelativeUrl
            // Parameter name: Specified value is not supported for the serverRelativeUrl parameter.
        }
        [TestMethod, TestCategory("UndoCheckOutFile Method"), TestCategory("Fail")]
        public void UndoCheckOutFileWithSharePointURL()
        {
            /// Getting URL from browser's address bar.
            var workingFile = @"https://automationanywhere1.sharepoint.com/sites/SolutionArchitectTeam/Shared%20Documents/Forms/AllItems.aspx?FolderCTID=0x0120000223E4C4F9AB2C4DA1E560A6277B76AB&id=%2Fsites%2FSolutionArchitectTeam%2FShared%20Documents%2FGeneral%2FWorking_Files_Naoya%2FMetabotTest%2Forigin%2Fsample%2Etxt&parent=%2Fsites%2FSolutionArchitectTeam%2FShared%20Documents%2FGeneral%2FWorking_Files_Naoya%2FMetabotTest%2Forigin";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.UndoCheckOutFile(workingFile);
            CheckFailCode(res);
            Console.WriteLine(res);
            // FAIL:Fail to checkout the file with error: serverRelativeUrl
            // Parameter name: Specified value is not supported for the serverRelativeUrl parameter.
        }
        [TestMethod, TestCategory("UndoCheckOutFile Method"), TestCategory("Fail")]
        public void UndoCheckOutFileWithSharePointLink()
        {
            /// Getting Link from SharePoint
            var workingFile = @"https://automationanywhere1.sharepoint.com/:t:/s/SolutionArchitectTeam/EYm5Qip7k2tCgKCIFWNefr8BMYVMimJgG3UHlkcQJHFdvQ?e=cqtfc8";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.UndoCheckOutFile(workingFile);
            CheckFailCode(res);
            Console.WriteLine(res);
            // FAIL:Fail to checkout the file with error: serverRelativeUrl
            // Parameter name: Specified value is not supported for the serverRelativeUrl parameter.
        }
        [TestMethod, TestCategory("UndoCheckOutFile Method"), TestCategory("Fail")]
        public void UndoCheckOutFileWithNullPath()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.UndoCheckOutFile(null);
            CheckFailCode(res);
            Console.WriteLine(res);
        }

        /// Tests for CopyFile method
        [TestMethod, TestCategory("CopyFile Method"), TestCategory("DeleteFile Method"), TestCategory("Success")]
        public void CopyAndDeleteFile()
        {
            var targetFileName = "sample.txt";

            var api = this.InitializeWrapper(siteUrl, username, password);
            /// Cleanup the destination file
            var res = api.GetFilesURL(remoteWorkFolderPath);
            string[] list = res.Split(',');
            var isExist = this.CheckItemInList(list, targetFileName);
            if (isExist)
            {
                // If a file is already exist, then remove it.
                api.DeleteFile($"{remoteWorkFolderPath}/{targetFileName}");
            }

            res = api.CopyFile(remoteFilePath, $"{remoteWorkFolderPath}/{targetFileName}");
            Assert.AreEqual($"Copied the file from {remoteFilePath} to {remoteWorkFolderPath}/{targetFileName}", res);

            /// Cleanup the copied file.
            res = api.DeleteFile($"{remoteWorkFolderPath}/{targetFileName}");
            Assert.AreEqual($"Deleted the file located on {remoteWorkFolderPath}/{targetFileName}", res);            
        }
        [TestMethod, TestCategory("CopyFile Method"), TestCategory("Success")]
        public void CopyFileWithOverride()
        {
            var targetFileName = "sample.txt";

            var api = this.InitializeWrapper(siteUrl, username, password);
            /// Cleanup the destination file
            var res = api.GetFilesURL(remoteWorkFolderPath);
            string[] list = res.Split(',');
            var isExist = this.CheckItemInList(list, targetFileName);
            if (isExist)
            {
                // If a file is already exist, then remove it.
                api.DeleteFile($"{remoteWorkFolderPath}/{targetFileName}");
            }

            res = api.CopyFile(remoteFilePath, $"{remoteWorkFolderPath}/{targetFileName}");
            Assert.AreEqual($"Copied the file from {remoteFilePath} to {remoteWorkFolderPath}/{targetFileName}", res);
            res = api.CopyFile(remoteFilePath, $"{remoteWorkFolderPath}/{targetFileName}");
            Assert.AreEqual($"Copied the file from {remoteFilePath} to {remoteWorkFolderPath}/{targetFileName}", res);

            /// Cleanup the copied file.
            api.DeleteFile($"{remoteWorkFolderPath}/{targetFileName}");
        }
        [TestMethod, TestCategory("CopyFile Method"), TestCategory("Fail")]
        public void CopyFileWithoutContext()
        {
            var targetFileName = "sample.txt";
            var api = new SharePointAPIWrapper(); /// no authentication 

            /// Cleanup the destination file
            var res = api.GetFilesURL(remoteWorkFolderPath);
            string[] list = res.Split(',');
            var isExist = this.CheckItemInList(list, targetFileName);
            if (isExist)
            {
                // If a file is already exist, then remove it.
                api.DeleteFile($"{remoteWorkFolderPath}/{targetFileName}");
            }

            res = api.CopyFile(remoteFilePath, $"{remoteWorkFolderPath}/{targetFileName}");
            CheckAuthenticationFail(res);
        }
        [TestMethod, TestCategory("CopyFile Method"), TestCategory("Fail")]
        public void CopyFileWithFolderPathInSourcePath()
        {
            var targetFileName = "sample.txt";

            var api = this.InitializeWrapper(siteUrl, username, password);

            /// Cleanup the destination file
            var res = api.GetFilesURL(remoteWorkFolderPath);
            string[] list = res.Split(',');
            var isExist = this.CheckItemInList(list, targetFileName);
            if (isExist)
            {
                // If a file is already exist, then remove it.
                api.DeleteFile($"{remoteWorkFolderPath}/{targetFileName}");
            }

            res = api.CopyFile(remoteFolderPath, $"{remoteWorkFolderPath}/{targetFileName}");
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to checkout the file with error: Unknown Error
        }
        [TestMethod, TestCategory("CopyFile Method"), TestCategory("Fail")]
        public void CopyFileWithMissingFile()
        {
            /// Missing file
            var workingFile = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin/missing.txt";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CopyFile(workingFile, $"{remoteWorkFolderPath}/missing.txt");
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to checkout the file with error: File Not Found.
        }
        [TestMethod, TestCategory("CopyFile Method"), TestCategory("Fail")]
        public void CopyFileWithTeamsURL()
        {
            /// Missing Folder
            var workingFile = @"https://automationanywhere1.sharepoint.com/sites/SolutionArchitectTeam/Shared%20Documents/General/Working_Files_Naoya/MetabotTest/origin/sample.txt";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CopyFile(workingFile, $"{remoteWorkFolderPath}/dummy.txt");
            CheckFailCode(res);
            Console.WriteLine(res);
            // FAIL:Fail to checkout the file with error: serverRelativeUrl
            // Parameter name: Specified value is not supported for the serverRelativeUrl parameter.
        }
        [TestMethod, TestCategory("CopyFile Method"), TestCategory("Fail")]
        public void CopyFileWithSharePointURL()
        {
            /// Getting URL from browser's address bar.
            var workingFile = @"https://automationanywhere1.sharepoint.com/sites/SolutionArchitectTeam/Shared%20Documents/Forms/AllItems.aspx?FolderCTID=0x0120000223E4C4F9AB2C4DA1E560A6277B76AB&id=%2Fsites%2FSolutionArchitectTeam%2FShared%20Documents%2FGeneral%2FWorking_Files_Naoya%2FMetabotTest%2Forigin%2Fsample%2Etxt&parent=%2Fsites%2FSolutionArchitectTeam%2FShared%20Documents%2FGeneral%2FWorking_Files_Naoya%2FMetabotTest%2Forigin";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CopyFile(workingFile, $"{remoteWorkFolderPath}/dummy.txt");
            CheckFailCode(res);
            Console.WriteLine(res);
            // FAIL:Fail to checkout the file with error: serverRelativeUrl
            // Parameter name: Specified value is not supported for the serverRelativeUrl parameter.
        }
        [TestMethod, TestCategory("CopyFile Method"), TestCategory("Fail")]
        public void CopyFileWithSharePointLink()
        {
            /// Getting Link from SharePoint
            var workingFile = @"https://automationanywhere1.sharepoint.com/:t:/s/SolutionArchitectTeam/EYm5Qip7k2tCgKCIFWNefr8BMYVMimJgG3UHlkcQJHFdvQ?e=cqtfc8";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CopyFile(workingFile, $"{remoteWorkFolderPath}/dummy.txt");
            CheckFailCode(res);
            Console.WriteLine(res);
            // FAIL:Fail to checkout the file with error: serverRelativeUrl
            // Parameter name: Specified value is not supported for the serverRelativeUrl parameter.
        }
        [TestMethod, TestCategory("CopyFile Method"), TestCategory("Fail")]
        public void CopyFileWithNullSrcPath()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CopyFile(null, $"{remoteWorkFolderPath}/dummy.txt");
            CheckFailCode(res);
            Console.WriteLine(res);
        }
        [TestMethod, TestCategory("CopyFile Method"), TestCategory("Fail")]
        public void CopyFileWithNullDestPath()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CopyFile(remoteFilePath, null);
            CheckFailCode(res);
            Console.WriteLine(res);
        }

        /// Tests for DeleteFile method
        [TestMethod, TestCategory("DeleteFile Method"), TestCategory("Fail")]
        public void DeleteFileWithoutContext()
        {
            var api = new SharePointAPIWrapper(); /// no authentication 

            var res = api.DeleteFile(remoteFilePath);
            CheckAuthenticationFail(res);
        }
        [TestMethod, TestCategory("DeleteFile Method"), TestCategory("Fail")]
        public void DeleteFileWithFolderPath()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.DeleteFile(remoteFolderPath);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: Unknown Error
        }
        [TestMethod, TestCategory("DeleteFile Method"), TestCategory("Fail")]
        public void DeleteFileWithMissingFile()
        {
            var api = new SharePointAPIWrapper();
            api.Authenticate(siteUrl, username, password);

            var res = api.DeleteFile($"{remoteFolderPath}/missing.txt");
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: File Not Found.
        }
        [TestMethod, TestCategory("DeleteFile Method"), TestCategory("Fail")]
        public void DeleteFileWithTeamsURL()
        {
            /// Missing Folder
            var workingFile = @"https://automationanywhere1.sharepoint.com/sites/SolutionArchitectTeam/Shared%20Documents/General/Working_Files_Naoya/MetabotTest/origin/sample.txt";
            var api = new SharePointAPIWrapper();
            api.Authenticate(siteUrl, username, password);

            var res = api.DeleteFile(workingFile);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: File Not Found.
        }
        [TestMethod, TestCategory("DeleteFile Method"), TestCategory("Fail")]
        public void DeleteFileWithSharePointURL()
        {
            /// Getting URL from browser's address bar.
            var workingFile = @"https://automationanywhere1.sharepoint.com/sites/SolutionArchitectTeam/Shared%20Documents/Forms/AllItems.aspx?FolderCTID=0x0120000223E4C4F9AB2C4DA1E560A6277B76AB&id=%2Fsites%2FSolutionArchitectTeam%2FShared%20Documents%2FGeneral%2FWorking_Files_Naoya%2FMetabotTest%2Forigin%2Fsample%2Etxt&parent=%2Fsites%2FSolutionArchitectTeam%2FShared%20Documents%2FGeneral%2FWorking_Files_Naoya%2FMetabotTest%2Forigin";
            var api = new SharePointAPIWrapper();
            api.Authenticate(siteUrl, username, password);

            var res = api.DeleteFile(workingFile);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: Unknown Error
        }
        [TestMethod, TestCategory("DeleteFile Method"), TestCategory("Fail")]
        public void DeleteFileWithSharePointLink()
        {
            /// Getting Link from SharePoint
            var workingFile = @"https://automationanywhere1.sharepoint.com/:t:/s/SolutionArchitectTeam/EYm5Qip7k2tCgKCIFWNefr8BMYVMimJgG3UHlkcQJHFdvQ?e=cqtfc8";
            var api = new SharePointAPIWrapper();
            api.Authenticate(siteUrl, username, password);

            var res = api.DeleteFile(workingFile);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: URL is not for this web
        }
        [TestMethod, TestCategory("DeleteFile Method"), TestCategory("Fail")]
        public void DeleteFileWithNullPath()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.DeleteFile(null);
            CheckFailCode(res);
            Console.WriteLine(res);
        }

        /// Tests for MoveFile method
        [TestMethod, TestCategory("MoveFile Method"), TestCategory("Success")]
        public void MoveFileWithRename()
        {
            var targetFileName  = "copied-sample.txt";
            var renamedFileName = "renamed.txt";

            var api = this.InitializeWrapper(siteUrl, username, password);

            /// Cleanup the destination file
            var res = api.GetFilesURL(remoteFolderPath);
            string[] list = res.Split(',');
            var isExist = this.CheckItemInList(list, targetFileName);
            if (isExist)
            {
                // If a file is already exist, then remove it.
                api.DeleteFile($"{remoteFolderPath}/{targetFileName}");
            }

            res = api.CopyFile(remoteFilePath, $"{remoteFolderPath}/{targetFileName}");
            Assert.AreEqual($"Copied the file from {remoteFilePath} to {remoteFolderPath}/{targetFileName}", res);

            /// rename the file (move to the same folder, the different name.
            res = api.MoveFile($"{remoteFolderPath}/{targetFileName}", $"{remoteFolderPath}/{renamedFileName}");
            Assert.AreEqual($"Moved the file from {remoteFolderPath}/{targetFileName} to {remoteFolderPath}/{renamedFileName}", res);

            /// Cleanup the copied file.
            res = api.DeleteFile($"{remoteFolderPath}/{renamedFileName}");
            Assert.AreEqual($"Deleted the file located on {remoteFolderPath}/{renamedFileName}", res);
        }
        [TestMethod, TestCategory("MoveFile Method"), TestCategory("Success")]
        public void MoveFileWithSameName()
        {
            var targetFileName = "copied-sample.txt";

            var api = this.InitializeWrapper(siteUrl, username, password);

            /// Cleanup the destination file
            var res = api.GetFilesURL(remoteFolderPath);
            string[] list = res.Split(',');
            var isExist = this.CheckItemInList(list, targetFileName);
            if (isExist)
            {
                // If a file is already exist, then remove it.
                api.DeleteFile($"{remoteFolderPath}/{targetFileName}");
            }

            res = api.CopyFile(remoteFilePath, $"{remoteFolderPath}/{targetFileName}");
            Assert.AreEqual($"Copied the file from {remoteFilePath} to {remoteFolderPath}/{targetFileName}", res);

            /// rename the file (move to the same folder, the different name.
            res = api.MoveFile($"{remoteFolderPath}/{targetFileName}", $"{remoteFolderPath}/{targetFileName}");
            Assert.AreEqual($"Moved the file from {remoteFolderPath}/{targetFileName} to {remoteFolderPath}/{targetFileName}", res);

            /// Cleanup the copied file.
            res = api.DeleteFile($"{remoteFolderPath}/{targetFileName}");
            Assert.AreEqual($"Deleted the file located on {remoteFolderPath}/{targetFileName}", res);
        }
        [TestMethod, TestCategory("MoveFile Method"), TestCategory("Success")]
        public void MoveFileToOtherFolderWithRename()
        {
            var targetFileName = "copied-sample.txt";
            var renamedFileName = "renamed.txt";

            var api = this.InitializeWrapper(siteUrl, username, password);

            /// Cleanup the destination file
            var res = api.GetFilesURL(remoteFolderPath);
            string[] list = res.Split(',');
            var isExist = this.CheckItemInList(list, targetFileName);
            if (isExist)
            {
                // If a file is already exist, then remove it.
                api.DeleteFile($"{remoteFolderPath}/{targetFileName}");
            }

            res = api.CopyFile(remoteFilePath, $"{remoteFolderPath}/{targetFileName}");
            Assert.AreEqual($"Copied the file from {remoteFilePath} to {remoteFolderPath}/{targetFileName}", res);

            /// rename the file (move to the same folder, the different name.
            res = api.MoveFile($"{remoteFolderPath}/{targetFileName}", $"{remoteWorkFolderPath}/{renamedFileName}");
            Assert.AreEqual($"Moved the file from {remoteFolderPath}/{targetFileName} to {remoteWorkFolderPath}/{renamedFileName}", res);

            /// Cleanup the copied file.
            res = api.DeleteFile($"{remoteWorkFolderPath}/{renamedFileName}");
            Assert.AreEqual($"Deleted the file located on {remoteWorkFolderPath}/{renamedFileName}", res);
        }
        [TestMethod, TestCategory("MoveFile Method"), TestCategory("Success")]
        public void MoveFileToOtherFolderWithSameName()
        {
            var targetFileName = "copied-sample.txt";

            var api = this.InitializeWrapper(siteUrl, username, password);

            /// Cleanup the destination file
            var res = api.GetFilesURL(remoteFolderPath);
            string[] list = res.Split(',');
            var isExist = this.CheckItemInList(list, targetFileName);
            if (isExist)
            {
                // If a file is already exist, then remove it.
                api.DeleteFile($"{remoteFolderPath}/{targetFileName}");
            }

            res = api.CopyFile(remoteFilePath, $"{remoteFolderPath}/{targetFileName}");
            Assert.AreEqual($"Copied the file from {remoteFilePath} to {remoteFolderPath}/{targetFileName}", res);

            /// rename the file (move to the same folder, the different name.
            res = api.MoveFile($"{remoteFolderPath}/{targetFileName}", $"{remoteWorkFolderPath}/{targetFileName}");
            Assert.AreEqual($"Moved the file from {remoteFolderPath}/{targetFileName} to {remoteWorkFolderPath}/{targetFileName}", res);

            /// Cleanup the copied file.
            res = api.DeleteFile($"{remoteWorkFolderPath}/{targetFileName}");
            Assert.AreEqual($"Deleted the file located on {remoteWorkFolderPath}/{targetFileName}", res);
        }
        [TestMethod, TestCategory("MoveFile Method"), TestCategory("Fail")]
        public void MoveFileWithoutContext()
        {
            var api = new SharePointAPIWrapper(); /// no authentication 

            var res = api.MoveFile(remoteFilePath, remoteWorkFilePath);
            CheckAuthenticationFail(res);
        }
        [TestMethod, TestCategory("MoveFile Method"), TestCategory("Fail")]
        public void MoveFileWithFolderPath()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.MoveFile(remoteFolderPath, remoteWorkFilePath);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: Unknown Error
        }
        [TestMethod, TestCategory("MoveFile Method"), TestCategory("Fail")]
        public void MoveFileWithMissingFile()
        {
            var api = new SharePointAPIWrapper();
            api.Authenticate(siteUrl, username, password);

            var res = api.MoveFile($"{remoteFolderPath}/missing.txt", remoteWorkFilePath);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: File Not Found.
        }
        [TestMethod, TestCategory("MoveFile Method"), TestCategory("Fail")]
        public void MoveFileWithTeamsURL()
        {
            /// Missing Folder
            var workingFile = @"https://automationanywhere1.sharepoint.com/sites/SolutionArchitectTeam/Shared%20Documents/General/Working_Files_Naoya/MetabotTest/origin/sample.txt";
            var api = new SharePointAPIWrapper();
            api.Authenticate(siteUrl, username, password);

            var res = api.MoveFile(workingFile, remoteWorkFilePath);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: File Not Found.
        }
        [TestMethod, TestCategory("MoveFile Method"), TestCategory("Fail")]
        public void MoveFileWithSharePointURL()
        {
            /// Getting URL from browser's address bar.
            var workingFile = @"https://automationanywhere1.sharepoint.com/sites/SolutionArchitectTeam/Shared%20Documents/Forms/AllItems.aspx?FolderCTID=0x0120000223E4C4F9AB2C4DA1E560A6277B76AB&id=%2Fsites%2FSolutionArchitectTeam%2FShared%20Documents%2FGeneral%2FWorking_Files_Naoya%2FMetabotTest%2Forigin%2Fsample%2Etxt&parent=%2Fsites%2FSolutionArchitectTeam%2FShared%20Documents%2FGeneral%2FWorking_Files_Naoya%2FMetabotTest%2Forigin";
            var api = new SharePointAPIWrapper();
            api.Authenticate(siteUrl, username, password);

            var res = api.MoveFile(workingFile, remoteWorkFilePath);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: Unknown Error
        }
        [TestMethod, TestCategory("MoveFile Method"), TestCategory("Fail")]
        public void MoveFileWithSharePointLink()
        {
            /// Getting Link from SharePoint
            var workingFile = @"https://automationanywhere1.sharepoint.com/:t:/s/SolutionArchitectTeam/EYm5Qip7k2tCgKCIFWNefr8BMYVMimJgG3UHlkcQJHFdvQ?e=cqtfc8";
            var api = new SharePointAPIWrapper();
            api.Authenticate(siteUrl, username, password);

            var res = api.MoveFile(workingFile, remoteWorkFilePath);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: URL is not for this web
        }
        [TestMethod, TestCategory("MoveFile Method"), TestCategory("Fail")]
        public void MoveFileWithNullSrcPath()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.MoveFile(null, remoteWorkFilePath);
            CheckFailCode(res);
            Console.WriteLine(res);
        }
        [TestMethod, TestCategory("MoveFile Method"), TestCategory("Fail")]
        public void MoveFileWithNullDestPath()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.MoveFile(remoteFilePath, null);
            CheckFailCode(res);
            Console.WriteLine(res);
        }

        /// Tests for DownloadFolder method
        [TestMethod, TestCategory("DownloadFolder Method"), TestCategory("UploadFolder Method"), TestCategory("Success")]
        public void DownloadUploadFolder()
        {
            var workFolderName = "DownloadTest";
            var localWorkDirectory = Path.Combine(localFolder, workFolderName);
            this.CleanUpLocalDirectory(localWorkDirectory);

            var api = this.InitializeWrapper(siteUrl, username, password);
            var res = api.DownloadFolder(remoteFolderPath, localWorkDirectory);
            Assert.AreEqual($"Downloaded the folder from {remoteFolderPath} to {localWorkDirectory}", res);
            Console.WriteLine(res);

            string[] files = System.IO.Directory.GetFiles(localWorkDirectory);
            Assert.AreEqual(5, files.Length);

            res = api.UploadFolder(localWorkDirectory, $"{remoteWorkFolderPath}/{workFolderName}");
            Assert.AreEqual($"Uploaded the folder from {localWorkDirectory} to {remoteWorkFolderPath}/{workFolderName}", res);
            Console.WriteLine(res);

            res = api.GetFilesURL($"{remoteWorkFolderPath}/{workFolderName}");
            string[] list = res.Split(',');
            Assert.AreEqual(5, files.Length);

            /// cleanup
            /// Cleanup the copied file.
            res = api.DeleteFolder($"{remoteWorkFolderPath}/{workFolderName}");
            Assert.AreEqual($"Deleted the folder located on {remoteWorkFolderPath}/{workFolderName}", res);
            Console.WriteLine(res);
        }
        [TestMethod, TestCategory("DownloadFolder Method"), TestCategory("UploadFolder Method"), TestCategory("Success")]
        public void DownloadUploadFolderEndWithSlash()
        {
            var workFolderName = "DownloadTest";
            var localWorkDirectory = Path.Combine(localFolder, workFolderName);
            this.CleanUpLocalDirectory(localWorkDirectory);

            var api = this.InitializeWrapper(siteUrl, username, password);
            var res = api.DownloadFolder($"{remoteFolderPath}/", $"{localWorkDirectory}");
            Assert.AreEqual($"Downloaded the folder from {remoteFolderPath} to {localWorkDirectory}", res);
            Console.WriteLine(res);

            string[] files = System.IO.Directory.GetFiles(localWorkDirectory);
            Assert.AreEqual(5, files.Length);

            res = api.UploadFolder(localWorkDirectory, $"{remoteWorkFolderPath}/{workFolderName}/");
            Assert.AreEqual($"Uploaded the folder from {localWorkDirectory} to {remoteWorkFolderPath}/{workFolderName}", res);
            Console.WriteLine(res);

            res = api.GetFilesURL($"{remoteWorkFolderPath}/{workFolderName}");
            string[] list = res.Split(',');
            Assert.AreEqual(5, files.Length);

            /// cleanup
            /// Cleanup the copied file.
            res = api.DeleteFolder($"{remoteWorkFolderPath}/{workFolderName}");
            Assert.AreEqual($"Deleted the folder located on {remoteWorkFolderPath}/{workFolderName}", res);
            Console.WriteLine(res);
        }
        [TestMethod, TestCategory("DownloadFolder Method"), TestCategory("Fail")]
        public void DownloadFolderWithoutContext()
        {
            var api = new SharePointAPIWrapper(); /// no authentication 

            var res = api.DownloadFolder(remoteFolderPath, localFolder);
            CheckAuthenticationFail(res);
        }
        [TestMethod, TestCategory("DownloadFolder Method"), TestCategory("Fail")]
        public void DownloadFolderWithSrcFilePath()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.DownloadFolder(remoteFilePath, localFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: Unknown Error
        }
        [TestMethod, TestCategory("DownloadFolder Method"), TestCategory("Fail")]
        public void DownloadFolderWithMissingSrcFolder()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.DownloadFolder($"{remoteFolderPath}/MissingFolder", localFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: File Not Found.
        }
        [TestMethod, TestCategory("DownloadFolder Method"), TestCategory("Fail")]
        public void DownloadFolderWithTeamsURL()
        {
            /// Getting Link from Teams URL
            var workingFolder = @"https://teams.microsoft.com/_#/files/General?threadId=19%3A003daacb03344f4caf79f00dac7f898c%40thread.skype&ctx=channel&context=Working_Files_Naoya%252FMetabotTest";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.DownloadFolder(workingFolder, localFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: File Not Found.
        }
        [TestMethod, TestCategory("DownloadFolder Method"), TestCategory("Fail")]
        public void DownloadFolderWithSharePointURL()
        {
            /// Missing Folder
            var workingFolder = @"https://automationanywhere1.sharepoint.com/sites/SolutionArchitectTeam/Shared%20Documents/Forms/AllItems.aspx?RootFolder=%2Fsites%2FSolutionArchitectTeam%2FShared%20Documents%2FGeneral%2FWorking_Files_Naoya%2FMetabotTest&FolderCTID=0x0120000223E4C4F9AB2C4DA1E560A6277B76AB";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.DownloadFolder(workingFolder, localFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: Unknown Error
        }
        [TestMethod, TestCategory("DownloadFolder Method"), TestCategory("Fail")]
        public void DownloadFolderWithSharePointLink()
        {
            /// Missing Folder
            var workingFolder = @"https://automationanywhere1.sharepoint.com/:f:/s/SolutionArchitectTeam/EhcY8CIRG4tFrGx6EWINcCEBlZ0vvYCLSmNCU5eeEDb4tQ?e=lftPCq";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.DownloadFolder(workingFolder, localFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: URL is not for this web
        }
        [TestMethod, TestCategory("DownloadFolder Method"), TestCategory("Fail")]
        public void DownloadFolderWithNullSrcPath()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.DownloadFolder(null, localFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: URL is not for this web
        }
        [TestMethod, TestCategory("DownloadFolder Method"), TestCategory("Fail")]
        public void DownloadFolderWithNullDestPath()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.DownloadFolder(remoteFolderPath, null);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: URL is not for this web
        }

        /// Tests for UploadFolder method
        [TestMethod, TestCategory("UploadFolder Method"), TestCategory("Fail")]
        public void UploadFolderWithoutContext()
        {
            var api = new SharePointAPIWrapper(); /// no authentication 

            var res = api.UploadFolder(localFolder, remoteWorkFolderPath);
            CheckAuthenticationFail(res);
        }
        [TestMethod, TestCategory("UploadFolder Method"), TestCategory("Fail")]
        public void UploadFolderWithSrcFilePath()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.UploadFolder(localFile, remoteWorkFolderPath);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: Unknown Error
        }
        [TestMethod, TestCategory("UploadFolder Method"), TestCategory("Fail")]
        public void UploadFolderWithMissingSrcFolder()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.UploadFolder($"{localFolder}/MissingFolder", remoteWorkFolderPath);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: File Not Found.
        }
        [TestMethod, TestCategory("UploadFolder Method"), TestCategory("Fail")]
        public void UploadFolderWithTeamsURL()
        {
            /// Getting Link from Teams URL
            var workingFolder = @"https://teams.microsoft.com/_#/files/General?threadId=19%3A003daacb03344f4caf79f00dac7f898c%40thread.skype&ctx=channel&context=Working_Files_Naoya%252FMetabotTest";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.UploadFolder(localFolder, workingFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: File Not Found.
        }
        [TestMethod, TestCategory("UploadFolder Method"), TestCategory("Fail")]
        public void UploadFolderWithSharePointURL()
        {
            /// Missing Folder
            var workingFolder = @"https://automationanywhere1.sharepoint.com/sites/SolutionArchitectTeam/Shared%20Documents/Forms/AllItems.aspx?RootFolder=%2Fsites%2FSolutionArchitectTeam%2FShared%20Documents%2FGeneral%2FWorking_Files_Naoya%2FMetabotTest&FolderCTID=0x0120000223E4C4F9AB2C4DA1E560A6277B76AB";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.UploadFolder(localFolder, workingFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: Unknown Error
        }
        [TestMethod, TestCategory("UploadFolder Method"), TestCategory("Fail")]
        public void UploadFolderWithSharePointLink()
        {
            /// Missing Folder
            var workingFolder = @"https://automationanywhere1.sharepoint.com/:f:/s/SolutionArchitectTeam/EhcY8CIRG4tFrGx6EWINcCEBlZ0vvYCLSmNCU5eeEDb4tQ?e=lftPCq";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.UploadFolder(localFolder, workingFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: URL is not for this web
        }
        [TestMethod, TestCategory("UploadFolder Method"), TestCategory("Fail")]
        public void UploadFolderWithNullSrcPath()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.UploadFolder(null, remoteWorkFolderPath);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: URL is not for this web
        }
        [TestMethod, TestCategory("UploadFolder Method"), TestCategory("Fail")]
        public void UploadFolderWithNullDestPath()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.UploadFolder(localFolder, null);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: URL is not for this web
        }

        /// Tests for CopyFolder
        [TestMethod, TestCategory("CopyFolder Method"), TestCategory("Success")]
        public void CopyAndDeleteFolder()
        {
            var targetFolderName = "CopyNDelete";
            var api = this.InitializeWrapper(siteUrl, username, password);

            /// Cleanup the destination file
            var res = api.GetFoldersURL(remoteWorkFolderPath);
            string[] list = res.Split(',');
            var isExist = this.CheckItemInList(list, targetFolderName);
            if (isExist)
            {
                // If a file is already exist, then remove it.
                api.DeleteFile($"{remoteWorkFolderPath}/{targetFolderName}");
            }

            res = api.CopyFolder(remoteFolderPath, $"{remoteWorkFolderPath}/{targetFolderName}");
            Assert.AreEqual($"Copied the folder from {remoteFolderPath} to {remoteWorkFolderPath}/{targetFolderName}", res);

            /// Cleanup the copied file.
            res = api.DeleteFolder($"{remoteWorkFolderPath}/{targetFolderName}");
            Assert.AreEqual($"Deleted the folder located on {remoteWorkFolderPath}/{targetFolderName}", res);
        }
        [TestMethod, TestCategory("CopyFolder Method"), TestCategory("Success")]
        public void CopyFolderWithOverride()
        {
            var targetFolderName = "CopyNDelete";
            var api = this.InitializeWrapper(siteUrl, username, password);

            /// Cleanup the destination file
            var res = api.GetFoldersURL(remoteWorkFolderPath);
            string[] list = res.Split(',');
            var isExist = this.CheckItemInList(list, targetFolderName);
            if (isExist)
            {
                // If a file is already exist, then remove it.
                api.DeleteFile($"{remoteWorkFolderPath}/{targetFolderName}");
            }

            res = api.CopyFolder(remoteFolderPath, $"{remoteWorkFolderPath}/{targetFolderName}");
            Assert.AreEqual($"Copied the folder from {remoteFolderPath} to {remoteWorkFolderPath}/{targetFolderName}", res);
            res = api.CopyFolder(remoteFolderPath, $"{remoteWorkFolderPath}/{targetFolderName}");
            Assert.AreEqual($"Copied the folder from {remoteFolderPath} to {remoteWorkFolderPath}/{targetFolderName}", res);

            /// Cleanup the copied file.
            res = api.DeleteFolder($"{remoteWorkFolderPath}/{targetFolderName}");
            Assert.AreEqual($"Deleted the folder located on {remoteWorkFolderPath}/{targetFolderName}", res);
        }
        [TestMethod, TestCategory("CopyFolder Method"), TestCategory("Fail")]
        public void CopyFolderWithoutContext()
        {
            var api = new SharePointAPIWrapper(); /// no authentication 

            var res = api.CopyFolder(remoteFolderPath, $"{remoteWorkFolderPath}/CopyFolderTest");
            CheckAuthenticationFail(res);
        }
        [TestMethod, TestCategory("CopyFolder Method"), TestCategory("Fail")]
        public void CopyFolderWithSrcFilePath()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CopyFolder(remoteFilePath, $"{remoteWorkFolderPath}/CopyFolderTest");
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: Unknown Error
        }
        [TestMethod, TestCategory("CopyFolder Method"), TestCategory("Fail")]
        public void CopyFolderWithMissingSrcFolder()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CopyFolder($"{remoteFolderPath}/MissingFolder", $"{remoteWorkFolderPath}/CopyFolderTest");
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: File Not Found.
        }
        [TestMethod, TestCategory("CopyFolder Method"), TestCategory("Fail")]
        public void CopyFolderWithTeamsURL()
        {
            /// Getting Link from Teams URL
            var workingFolder = @"https://teams.microsoft.com/_#/files/General?threadId=19%3A003daacb03344f4caf79f00dac7f898c%40thread.skype&ctx=channel&context=Working_Files_Naoya%252FMetabotTest";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CopyFolder(workingFolder, $"{remoteWorkFolderPath}/CopyFolderTest");
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: File Not Found.
        }
        [TestMethod, TestCategory("CopyFolder Method"), TestCategory("Fail")]
        public void CopyFolderWithSharePointURL()
        {
            /// Missing Folder
            var workingFolder = @"https://automationanywhere1.sharepoint.com/sites/SolutionArchitectTeam/Shared%20Documents/Forms/AllItems.aspx?RootFolder=%2Fsites%2FSolutionArchitectTeam%2FShared%20Documents%2FGeneral%2FWorking_Files_Naoya%2FMetabotTest&FolderCTID=0x0120000223E4C4F9AB2C4DA1E560A6277B76AB";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CopyFolder(workingFolder, $"{remoteWorkFolderPath}/CopyFolderTest");
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: Unknown Error
        }
        [TestMethod, TestCategory("CopyFolder Method"), TestCategory("Fail")]
        public void CopyFolderWithSharePointLink()
        {
            /// Missing Folder
            var workingFolder = @"https://automationanywhere1.sharepoint.com/:f:/s/SolutionArchitectTeam/EhcY8CIRG4tFrGx6EWINcCEBlZ0vvYCLSmNCU5eeEDb4tQ?e=lftPCq";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CopyFolder(workingFolder, $"{remoteWorkFolderPath}/CopyFolderTest");
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: URL is not for this web
        }
        [TestMethod, TestCategory("CopyFolder Method"), TestCategory("Fail")]
        public void CopyFolderWithNullSrcPath()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CopyFolder(null, $"{remoteWorkFolderPath}/CopyFolderTest");
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: URL is not for this web
        }
        [TestMethod, TestCategory("CopyFolder Method"), TestCategory("Fail")]
        public void CopyFolderWithNullDestPath()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CopyFolder(remoteFolderPath, null);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: URL is not for this web
        }

        /// Tests for MoveFolder
        public void MoveFolderToDifferentFolder()
        {
            var targetFolderName = "MoveFolderTest";

            var api = this.InitializeWrapper(siteUrl, username, password);

            /// Cleanup the destination file
            var res = api.GetFoldersURL(remoteFolderPath);
            string[] list = res.Split(',');
            var isExist = this.CheckItemInList(list, targetFolderName);
            if (isExist)
            {
                // If a file is already exist, then remove it.
                api.DeleteFile($"{remoteFolderPath}/{targetFolderName}");
            }

            res = api.CopyFolder($"{remoteFolderPath}/temp", $"{remoteFolderPath}/{targetFolderName}");
            Console.WriteLine(res);
            Assert.AreEqual($"Copied the folder from {remoteFolderPath}/temp to {remoteFolderPath}/{targetFolderName}", res);

            /// rename the file (move to the same folder, the different name.
            res = api.MoveFolder($"{remoteFolderPath}/{targetFolderName}", $"{remoteWorkFolderPath}/{targetFolderName}");
            Console.WriteLine(res);
            Assert.AreEqual($"Moved the folder from {remoteFolderPath}/{targetFolderName} to {remoteWorkFolderPath}/{targetFolderName}", res);

            /// check a folder is listed in the folder.
            res = api.GetFoldersURL($"{remoteWorkFolderPath}");
            Console.WriteLine(res);
            list = res.Split(',');
            isExist = this.CheckItemInList(list, $"{remoteWorkFolderPath}/{targetFolderName}");
            Assert.AreEqual(isExist, true);

            /// Cleanup the copied file.
            res = api.DeleteFolder($"{remoteWorkFolderPath}/{targetFolderName}");
            Console.WriteLine(res);
            Assert.AreEqual($"Deleted the folder located on {remoteWorkFolderPath}/{targetFolderName}", res);
        }
        [TestMethod, TestCategory("MoveFolder Method"), TestCategory("Success")]
        public void MoveFolderToDifferentFolderWithRename()
        {
            var targetFolderName = "MoveFolderTest";
            var renamedFolderName = "RenamedFolderTest";

            var api = this.InitializeWrapper(siteUrl, username, password);

            /// Cleanup the destination file
            var res = api.GetFoldersURL(remoteFolderPath);
            string[] list = res.Split(',');
            var isExist = this.CheckItemInList(list, targetFolderName);
            if (isExist)
            {
                // If a file is already exist, then remove it.
                api.DeleteFile($"{remoteFolderPath}/{targetFolderName}");
            }

            res = api.CopyFolder($"{remoteFolderPath}/temp", $"{remoteFolderPath}/{targetFolderName}");
            Console.WriteLine(res);
            Assert.AreEqual($"Copied the folder from {remoteFolderPath}/temp to {remoteFolderPath}/{targetFolderName}", res);

            /// rename the file (move to the same folder, the different name.
            res = api.MoveFolder($"{remoteFolderPath}/{targetFolderName}", $"{remoteWorkFolderPath}/{renamedFolderName}");
            Console.WriteLine(res);
            Assert.AreEqual($"Moved the folder from {remoteFolderPath}/{targetFolderName} to {remoteWorkFolderPath}/{renamedFolderName}", res);

            /// check a folder is listed in the folder.
            res = api.GetFoldersURL($"{remoteWorkFolderPath}");
            Console.WriteLine(res);
            list = res.Split(',');
            isExist = this.CheckItemInList(list, $"{remoteWorkFolderPath}/{renamedFolderName}");
            Assert.AreEqual(isExist, true);

            /// Cleanup the copied file.
            res = api.DeleteFolder($"{remoteWorkFolderPath}/{renamedFolderName}");
            Console.WriteLine(res);
            Assert.AreEqual($"Deleted the folder located on {remoteWorkFolderPath}/{renamedFolderName}", res);
        }
        [TestMethod, TestCategory("MoveFolder Method"), TestCategory("Success")]
        public void MoveFolderWithRename()
        {
            var targetFolderName = "MoveFolderTest";
            var renamedFolderName = "RenamedFolderTest";

            var api = this.InitializeWrapper(siteUrl, username, password);

            /// Cleanup the destination file
            var res = api.GetFoldersURL(remoteFolderPath);
            string[] list = res.Split(',');
            var isExist = this.CheckItemInList(list, targetFolderName);
            if (isExist)
            {
                // If a file is already exist, then remove it.
                api.DeleteFile($"{remoteFolderPath}/{targetFolderName}");
            }

            res = api.CopyFolder($"{remoteFolderPath}/temp", $"{remoteWorkFolderPath}/{targetFolderName}");
            Console.WriteLine(res);
            Assert.AreEqual($"Copied the folder from {remoteFolderPath}/temp to {remoteWorkFolderPath}/{targetFolderName}", res);

            /// rename the file (move to the same folder, the different name.
            res = api.MoveFolder($"{remoteWorkFolderPath}/{targetFolderName}", $"{remoteWorkFolderPath}/{renamedFolderName}");
            Console.WriteLine(res);
            Assert.AreEqual($"Moved the folder from {remoteWorkFolderPath}/{targetFolderName} to {remoteWorkFolderPath}/{renamedFolderName}", res);

            /// check a folder is listed in the folder.
            res = api.GetFoldersURL($"{remoteWorkFolderPath}");
            Console.WriteLine(res);
            list = res.Split(',');
            isExist = this.CheckItemInList(list, $"{remoteWorkFolderPath}/{renamedFolderName}");
            Assert.AreEqual(isExist, true);

            /// Cleanup the copied file.
            res = api.DeleteFolder($"{remoteWorkFolderPath}/{renamedFolderName}");
            Console.WriteLine(res);
            Assert.AreEqual($"Deleted the folder located on {remoteWorkFolderPath}/{renamedFolderName}", res);
        }
        [TestMethod, TestCategory("MoveFolder Method"), TestCategory("Success")]
        public void MoveFolderWithSameName()
        {
            var targetFolderName = "MoveFolderTest";

            var api = this.InitializeWrapper(siteUrl, username, password);

            /// Cleanup the destination file
            var res = api.GetFoldersURL(remoteFolderPath);
            string[] list = res.Split(',');
            var isExist = this.CheckItemInList(list, targetFolderName);
            if (isExist)
            {
                // If a file is already exist, then remove it.
                api.DeleteFile($"{remoteFolderPath}/{targetFolderName}");
            }

            res = api.CopyFolder($"{remoteFolderPath}/temp", $"{remoteWorkFolderPath}/{targetFolderName}");
            Console.WriteLine(res);
            Assert.AreEqual($"Copied the folder from {remoteFolderPath}/temp to {remoteWorkFolderPath}/{targetFolderName}", res);

            /// rename the file (move to the same folder, the different name.
            res = api.MoveFolder($"{remoteWorkFolderPath}/{targetFolderName}", $"{remoteWorkFolderPath}/{targetFolderName}");
            Console.WriteLine(res);
            Assert.AreEqual($"Moved the folder from {remoteWorkFolderPath}/{targetFolderName} to {remoteWorkFolderPath}/{targetFolderName}", res);

            /// check a folder is listed in the folder.
            res = api.GetFoldersURL($"{remoteWorkFolderPath}");
            Console.WriteLine(res);
            list = res.Split(',');
            isExist = this.CheckItemInList(list, $"{remoteWorkFolderPath}/{targetFolderName}");
            Assert.AreEqual(isExist, true);

            /// Cleanup the copied file.
            res = api.DeleteFolder($"{remoteWorkFolderPath}/{targetFolderName}");
            Console.WriteLine(res);
            Assert.AreEqual($"Deleted the folder located on {remoteWorkFolderPath}/{targetFolderName}", res);
        }
        [TestMethod, TestCategory("MoveFolder Method"), TestCategory("Fail")]
        public void MoveFolderWithoutContext()
        {
            var targetFolderName = "CopyFolderTest";
            var api = new SharePointAPIWrapper(); /// no authentication 

            var res = api.MoveFolder($"{remoteFolderPath}/{targetFolderName}", $"{remoteWorkFolderPath}/{targetFolderName}");
            CheckAuthenticationFail(res);
        }
        [TestMethod, TestCategory("MoveFolder Method"), TestCategory("Fail")]
        public void MoveFolderWithSrcFilePath()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.MoveFolder(remoteFilePath, $"{remoteWorkFolderPath}/CopyFolderTest");
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: Unknown Error
        }
        [TestMethod, TestCategory("MoveFolder Method"), TestCategory("Fail")]
        public void MoveFolderWithMissingSrcFolder()
        {
            /// Missing Folder
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.MoveFolder($"{remoteFolderPath}/MissingFolder", $"{remoteWorkFolderPath}/CopyFolderTest");
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: File Not Found.
        }
        [TestMethod, TestCategory("MoveFolder Method"), TestCategory("Fail")]
        public void MoveFolderWithTeamsURL()
        {
            /// Missing Folder
            var workingFolder = @"https://teams.microsoft.com/_#/files/General?threadId=19%3A003daacb03344f4caf79f00dac7f898c%40thread.skype&ctx=channel&context=Working_Files_Naoya%252FMetabotTest";
            var api = new SharePointAPIWrapper();
            api.Authenticate(siteUrl, username, password);

            var res = api.MoveFolder(workingFolder, $"{remoteWorkFolderPath}/CopyFolderTest");
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: File Not Found.
        }
        [TestMethod, TestCategory("MoveFolder Method"), TestCategory("Fail")]
        public void MoveFolderWithSharePointURL()
        {
            /// Missing Folder
            var workingFolder = @"https://automationanywhere1.sharepoint.com/sites/SolutionArchitectTeam/Shared%20Documents/Forms/AllItems.aspx?RootFolder=%2Fsites%2FSolutionArchitectTeam%2FShared%20Documents%2FGeneral%2FWorking_Files_Naoya%2FMetabotTest&FolderCTID=0x0120000223E4C4F9AB2C4DA1E560A6277B76AB";
            var api = new SharePointAPIWrapper();
            api.Authenticate(siteUrl, username, password);

            var res = api.MoveFolder(workingFolder, $"{remoteWorkFolderPath}/CopyFolderTest");
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: Unknown Error
        }
        [TestMethod, TestCategory("MoveFolder Method"), TestCategory("Fail")]
        public void MoveFolderWithSharePointLink()
        {
            /// Missing Folder
            var workingFolder = @"https://automationanywhere1.sharepoint.com/:f:/s/SolutionArchitectTeam/EhcY8CIRG4tFrGx6EWINcCEBlZ0vvYCLSmNCU5eeEDb4tQ?e=lftPCq";
            var api = new SharePointAPIWrapper();
            api.Authenticate(siteUrl, username, password);

            var res = api.MoveFolder(workingFolder, $"{remoteWorkFolderPath}/CopyFolderTest");
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: URL is not for this web
        }
        [TestMethod, TestCategory("MoveFolder Method"), TestCategory("Fail")]
        public void MoveFolderWithNullSrcPath()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.MoveFolder(null, $"{remoteWorkFolderPath}/CopyFolderTest");
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: URL is not for this web
        }
        [TestMethod, TestCategory("MoveFolder Method"), TestCategory("Fail")]
        public void MoveFolderWithNullDestPath()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.MoveFolder(remoteFolderPath, null);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: URL is not for this web
        }
        [TestMethod, TestCategory("CreateFolder Method"), TestCategory("DeleteFolder Method"), TestCategory("Success")]
        public void CreateAndDeleteFolder()
        {
            var targetFolderName = "temp";

            var api = this.InitializeWrapper(siteUrl, username, password);

            /// create a folder
            var res = api.CreateFolder(remoteFolderPath, targetFolderName);
            Console.WriteLine(res);
            Assert.AreEqual(res, $"Create the folder located in {remoteFolderPath}/{targetFolderName}");

            /// check a folder is listed in the folder.
            res = api.GetFoldersURL($"{remoteFolderPath}");
            Console.WriteLine(res);
            string[] list = res.Split(',');
            var isExist = this.CheckItemInList(list, $"{remoteFolderPath}/{targetFolderName}");
            Assert.AreEqual(isExist, true);

            res = api.DeleteFolder($"{remoteFolderPath}/{targetFolderName}");
            Console.WriteLine(res);
            Assert.AreEqual(res, $"Deleted the folder located on {remoteFolderPath}/{targetFolderName}");
        }
        [TestMethod, TestCategory("CreateFolder Methods"), TestCategory("DeleteFolder Method"), TestCategory("Success")]
        public void CreateAndDeleteFolderEndWithSlash()
        {
            var targetFolderName = "temp";

            var api = this.InitializeWrapper(siteUrl, username, password);

            /// create a folder
            var res = api.CreateFolder(remoteFolderPath, targetFolderName);
            Console.WriteLine(res);
            Assert.AreEqual(res, $"Create the folder located in {remoteFolderPath}{targetFolderName}");

            /// check a folder is listed in the folder.
            res = api.GetFoldersURL($"{remoteFolderPath}");
            Console.WriteLine(res);
            string[] list = res.Split(',');
            var isExist = this.CheckItemInList(list, $"{remoteFolderPath}{targetFolderName}");
            Assert.AreEqual(isExist, true);

            res = api.DeleteFolder($"{remoteFolderPath}{targetFolderName}");
            Console.WriteLine(res);
            Assert.AreEqual(res, $"Deleted the folder located on {remoteFolderPath}{targetFolderName}");
        }
        [TestMethod, TestCategory("CreateFolder Method"), TestCategory("Fail")]
        public void CreateFolderWithoutContext()
        {
            /// working folder
            var workingFolder = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/origin";
            var targetFolderName = "temp";
            var api = new SharePointAPIWrapper(); /// no authentication 

            var res = api.CreateFolder(workingFolder, targetFolderName);
            CheckAuthenticationFail(res);
        }
        [TestMethod, TestCategory("CreateFolder Method"), TestCategory("Fail")]
        public void CreateFolderWithFilePath()
        {
            var targetFolderName = "temp";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CreateFolder(remoteFilePath, targetFolderName);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to create the folder with error: Unknown Error
        }
        [TestMethod, TestCategory("CreateFolder Method"), TestCategory("Fail")]
        public void CreateFolderWithMissingFolder()
        {
            /// Missing Folder
            var workingFolder = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/MissingFolder";
            var targetFolderName = "temp";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CreateFolder(workingFolder, targetFolderName);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to create the folder with error: File Not Found.
        }
        [TestMethod, TestCategory("CreateFolder Method"), TestCategory("Fail")]
        public void CreateFolderWithTeamsURL()
        {
            /// Missing Folder
            var workingFolder = @"https://teams.microsoft.com/_#/files/General?threadId=19%3A003daacb03344f4caf79f00dac7f898c%40thread.skype&ctx=channel&context=Working_Files_Naoya%252FMetabotTest";
            var targetFolderName = "temp";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CreateFolder(workingFolder, targetFolderName);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to create the folder with error: File Not Found.
        }
        [TestMethod, TestCategory("CreateFolder Method"), TestCategory("Fail")]
        public void CreateFolderWithSharePointURL()
        {
            /// Missing Folder
            var workingFolder = @"https://automationanywhere1.sharepoint.com/sites/SolutionArchitectTeam/Shared%20Documents/Forms/AllItems.aspx?RootFolder=%2Fsites%2FSolutionArchitectTeam%2FShared%20Documents%2FGeneral%2FWorking_Files_Naoya%2FMetabotTest&FolderCTID=0x0120000223E4C4F9AB2C4DA1E560A6277B76AB";
            var targetFolderName = "temp";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CreateFolder(workingFolder, targetFolderName);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to create the folder with error: Unknown Error
        }
        [TestMethod, TestCategory("CreateFolder Method"), TestCategory("Fail")]
        public void CreateFolderWithSharePointLink()
        {
            /// Missing Folder
            var workingFolder = @"https://automationanywhere1.sharepoint.com/:f:/s/SolutionArchitectTeam/EhcY8CIRG4tFrGx6EWINcCEBlZ0vvYCLSmNCU5eeEDb4tQ?e=lftPCq";
            var targetFolderName = "temp";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.CreateFolder(workingFolder, targetFolderName);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to create the folder with error: URL is not for this web
        }
        [TestMethod, TestCategory("DeleteFolder Method"), TestCategory("Fail")]
        public void DeleteFolderWithoutContext()
        {
            var api = new SharePointAPIWrapper(); /// no authentication 

            var res = api.DeleteFolder(remoteFolderPath);
            CheckAuthenticationFail(res);
        }
        [TestMethod, TestCategory("DeleteFolder Method"), TestCategory("Fail")]
        public void DeleteFolderWithFilePath()
        {
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.DeleteFolder(remoteFilePath);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: Unknown Error
        }
        [TestMethod, TestCategory("DeleteFolder Method"), TestCategory("Fail")]
        public void DeleteFolderWithMissingFolder()
        {
            /// Missing Folder
            var workingFolder = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/prigin/MissingFolder";
            var api = new SharePointAPIWrapper();
            api.Authenticate(siteUrl, username, password);

            var res = api.DeleteFolder(workingFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: File Not Found.
        }
        [TestMethod, TestCategory("DeleteFolder Method"), TestCategory("Fail")]
        public void DeleteFolderWithTeamsURL()
        {
            /// Missing Folder
            var workingFolder = @"https://teams.microsoft.com/_#/files/General?threadId=19%3A003daacb03344f4caf79f00dac7f898c%40thread.skype&ctx=channel&context=Working_Files_Naoya%252FMetabotTest";
            var api = new SharePointAPIWrapper();
            api.Authenticate(siteUrl, username, password);

            var res = api.DeleteFolder(workingFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: File Not Found.
        }
        [TestMethod, TestCategory("DeleteFolder Method"), TestCategory("Fail")]
        public void DeleteFolderWithSharePointURL()
        {
            /// Missing Folder
            var workingFolder = @"https://automationanywhere1.sharepoint.com/sites/SolutionArchitectTeam/Shared%20Documents/Forms/AllItems.aspx?RootFolder=%2Fsites%2FSolutionArchitectTeam%2FShared%20Documents%2FGeneral%2FWorking_Files_Naoya%2FMetabotTest&FolderCTID=0x0120000223E4C4F9AB2C4DA1E560A6277B76AB";
            var api = new SharePointAPIWrapper();
            api.Authenticate(siteUrl, username, password);

            var res = api.DeleteFolder(workingFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: Unknown Error
        }
        [TestMethod, TestCategory("DeleteFolder Method"), TestCategory("Fail")]
        public void DeleteFolderWithSharePointLink()
        {
            /// Missing Folder
            var workingFolder = @"https://automationanywhere1.sharepoint.com/:f:/s/SolutionArchitectTeam/EhcY8CIRG4tFrGx6EWINcCEBlZ0vvYCLSmNCU5eeEDb4tQ?e=lftPCq";
            var api = new SharePointAPIWrapper();
            api.Authenticate(siteUrl, username, password);

            var res = api.DeleteFolder(workingFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Fail to delete the folder with error: URL is not for this web
        }

        /// Test methods for GetFilesURL method.
        [TestMethod, TestCategory("GetFilesURL Method"), TestCategory("Success")]
        public void GetFilesURL()
        {
            /// Happy Path includes 4 files
            var workingFolder = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.GetFilesURL(workingFolder);
            string[] list = res.Split(',');
            this.OutputStringArray(list);

            Assert.AreEqual(list.Length, 4);
        }
        [TestMethod, TestCategory("GetFilesURL Method"), TestCategory("Success")]
        public void GetFilesURLEndWithSlash()
        {
            /// Happy Path includes 4 files
            var workingFolder = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.GetFilesURL(workingFolder);
            string[] list = res.Split(',');
            this.OutputStringArray(list);

            Assert.AreEqual(list.Length, 4);
        }
        [TestMethod, TestCategory("GetFilesURL Method"), TestCategory("Fail")]
        public void GetFilesURLWithoutContext()
        {
            var workingFolder = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest";
            var api = new SharePointAPIWrapper(); /// no authentication 

            var res = api.GetFilesURL(workingFolder);
            CheckAuthenticationFail(res);
        }
        [TestMethod, TestCategory("GetFilesURL Method"), TestCategory("Fail")]
        public void GetFilesURLWithFilePath()
        {
            /// Specified a path for file.
            var workingFolder = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/sample.txt";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.GetFilesURL(workingFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Failed to get a list of files in specified folder with error: Unknown Error
        }
        [TestMethod, TestCategory("GetFilesURL Method"), TestCategory("Fail")]
        public void GetFilesURLWithMissingFolder()
        {
            /// Missing Folder
            var workingFolder = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MissingFolder";
            var api = new SharePointAPIWrapper();
            api.Authenticate(siteUrl, username, password);

            var res = api.GetFilesURL(workingFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Failed to get a list of files in specified folder with error: File Not Found.
        }
        [TestMethod, TestCategory("GetFilesURL Method"), TestCategory("Fail")]
        public void GetFilesURLWithTeamsURL()
        {
            /// Missing Folder
            var workingFolder = @"https://teams.microsoft.com/_#/files/General?threadId=19%3A003daacb03344f4caf79f00dac7f898c%40thread.skype&ctx=channel&context=Working_Files_Naoya%252FMetabotTest";
            var api = new SharePointAPIWrapper();
            api.Authenticate(siteUrl, username, password);

            var res = api.GetFilesURL(workingFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Failed to get a list of files in specified folder with error: File Not Found.
        }
        [TestMethod, TestCategory("GetFilesURL Method"), TestCategory("Fail")]
        public void GetFilesURLWithSharePointURL()
        {
            /// Missing Folder
            var workingFolder = @"https://automationanywhere1.sharepoint.com/sites/SolutionArchitectTeam/Shared%20Documents/Forms/AllItems.aspx?RootFolder=%2Fsites%2FSolutionArchitectTeam%2FShared%20Documents%2FGeneral%2FWorking_Files_Naoya%2FMetabotTest&FolderCTID=0x0120000223E4C4F9AB2C4DA1E560A6277B76AB";
            var api = new SharePointAPIWrapper();
            api.Authenticate(siteUrl, username, password);

            var res = api.GetFilesURL(workingFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Failed to get a list of files in specified folder with error: Unknown Error
        }
        [TestMethod, TestCategory("GetFilesURL Method"), TestCategory("Fail")]
        public void GetFilesURLWithSharePointLink()
        {
            /// Missing Folder
            var workingFolder = @"https://automationanywhere1.sharepoint.com/:f:/s/SolutionArchitectTeam/EhcY8CIRG4tFrGx6EWINcCEBlZ0vvYCLSmNCU5eeEDb4tQ?e=lftPCq";
            var api = new SharePointAPIWrapper();
            api.Authenticate(siteUrl, username, password);

            var res = api.GetFilesURL(workingFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Failed to get a list of files in specified folder with error: URL is not for this web
        }

        /// Test methods for GetFoldersURL method.
        [TestMethod, TestCategory("GetFoldersURL Method"), TestCategory("Success")]
        public void GetFoldersURL()
        {
            /// Happy Path includes 4 folders
            var workingFolder = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.GetFoldersURL(workingFolder);
            string[] list = res.Split(',');
            this.OutputStringArray(list);

            Assert.AreEqual(list.Length, 4);
        }
        [TestMethod, TestCategory("GetFoldersURL Method"), TestCategory("Success")]
        public void GetFoldersURLEndWithSlash()
        {
            /// Happy Path includes 4 folders
            var workingFolder = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.GetFoldersURL(workingFolder);
            string[] list = res.Split(',');
            this.OutputStringArray(list);

            Assert.AreEqual(list.Length, 4);
        }
        [TestMethod, TestCategory("GetFoldersURL Method"), TestCategory("Fail")]
        public void GetFoldersURLWithoutContext()
        {
            /// Happy Path includes 4 folders
            var workingFolder = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/";
            var api = new SharePointAPIWrapper(); /// no authentication 

            var res = api.GetFoldersURL(workingFolder);
            CheckAuthenticationFail(res);
        }
        [TestMethod, TestCategory("GetFoldersURL Method"), TestCategory("Fail")]
        public void GetFoldersURLWithFilePath()
        {
            /// File path 
            var workingFolder = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/sample.txt";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.GetFoldersURL(workingFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Failed to get a list of folders in specified folder with error: Unknown Error
        }
        [TestMethod, TestCategory("GetFoldersURL Method"), TestCategory("Fail")]
        public void GetFoldersURLWithMissingFolder()
        {
            /// Missing Folder
            var workingFolder = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MissingFolder";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.GetFoldersURL(workingFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Failed to get a list of folders in specified folder with error: File Not Found.
        }
        [TestMethod, TestCategory("GetFolderItems Method"), TestCategory("Fail")]
        public void GetFoldersURLWithTeamsURL()
        {
            /// A URL getting from Microsoft Teams.
            var workingFolder = @"https://teams.microsoft.com/_#/files/General?threadId=19%3A003daacb03344f4caf79f00dac7f898c%40thread.skype&ctx=channel&context=Working_Files_Naoya%252FMetabotTest";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.GetFoldersURL(workingFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Failed to get a list of folders in specified folder with error: File Not Found.
        }
        [TestMethod, TestCategory("GetFolderItems Method"), TestCategory("Fail")]
        public void GetFoldersURLWithSharePointURL()
        {
            /// A URL getting from browser address bar.
            var workingFolder = @"https://automationanywhere1.sharepoint.com/sites/SolutionArchitectTeam/Shared%20Documents/Forms/AllItems.aspx?RootFolder=%2Fsites%2FSolutionArchitectTeam%2FShared%20Documents%2FGeneral%2FWorking_Files_Naoya%2FMetabotTest&FolderCTID=0x0120000223E4C4F9AB2C4DA1E560A6277B76AB";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.GetFoldersURL(workingFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Failed to get a list of folders in specified folder with error: Unknown Error
        }
        [TestMethod, TestCategory("GetFolderItems Method"), TestCategory("Fail")]
        public void GetFoldersURLWithSharePointLink()
        {
            /// A link URL getting from SharePoint 'CopyLink' functionality.
            var workingFolder = @"https://automationanywhere1.sharepoint.com/:f:/s/SolutionArchitectTeam/EhcY8CIRG4tFrGx6EWINcCEBlZ0vvYCLSmNCU5eeEDb4tQ?e=lftPCq";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.GetFoldersURL(workingFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Failed to get a list of folders in specified folder with error: URL is not for this web
        }

        /// Test methods for GetItemsURL method.
        [TestMethod, TestCategory("GetItemsURL Method"), TestCategory("Success")]
        public void GetItemsURL()
        {
            /// Happy Path includes 4 files and 4 folders
            var workingFolder = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.GetItemsURL(workingFolder);
            string[] list = res.Split(',');
            this.OutputStringArray(list);

            Assert.AreEqual(list.Length, 8);
        }
        [TestMethod, TestCategory("GetItemsURL Method"), TestCategory("Success")]
        public void GetItemsURLEndWithSlash()
        {
            /// Happy Path includes 4 files and 4 folders
            var workingFolder = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.GetItemsURL(workingFolder);
            string[] list = res.Split(',');
            this.OutputStringArray(list);

            Assert.AreEqual(list.Length, 8);
        }
        [TestMethod, TestCategory("GetItemsURL Method"), TestCategory("Fail")]
        public void GetItemsURLWithoutContext()
        {
            /// Happy Path includes 4 files and 4 folders
            var workingFolder = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/";
            var api = new SharePointAPIWrapper(); /// no authentication 

            var res = api.GetItemsURL(workingFolder);
            CheckAuthenticationFail(res);
        }
        [TestMethod, TestCategory("GetItemsURL Method"), TestCategory("Fail")]
        public void GetItemsURLWithFilePath()
        {
            /// File path 
            var workingFolder = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MetabotTest/sample.txt";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.GetItemsURL(workingFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Failed to get a list of folders in specified folder with error: Unknown Error,FAIL:Failed to get a list of files in specified folder with error: Unknown Error
        }
        [TestMethod, TestCategory("GetItemsURL Method"), TestCategory("Fail")]
        public void GetItemsURLWithMissingFolder()
        {
            /// Missing Folder
            var workingFolder = @"/sites/SolutionArchitectTeam/Shared Documents/General/Working_Files_Naoya/MissingFolder";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.GetItemsURL(workingFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Failed to get a list of folders in specified folder with error: File Not Found.,FAIL:Failed to get a list of files in specified folder with error: File Not Found.
        }
        [TestMethod, TestCategory("GetItemsURL Method"), TestCategory("Fail")]
        public void GetItemsURLWithTeamsURL()
        {
            /// A URL getting from Microsoft Teams.
            var workingFolder = @"https://teams.microsoft.com/_#/files/General?threadId=19%3A003daacb03344f4caf79f00dac7f898c%40thread.skype&ctx=channel&context=Working_Files_Naoya%252FMetabotTest";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.GetItemsURL(workingFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Failed to get a list of folders in specified folder with error: File Not Found.,FAIL:Failed to get a list of files in specified folder with error: File Not Found.
        }
        [TestMethod, TestCategory("GetItemsURL Method"), TestCategory("Fail")]
        public void GetItemsURLWithSharePointURL()
        {
            /// A URL getting from browser address bar.
            var workingFolder = @"https://automationanywhere1.sharepoint.com/sites/SolutionArchitectTeam/Shared%20Documents/Forms/AllItems.aspx?RootFolder=%2Fsites%2FSolutionArchitectTeam%2FShared%20Documents%2FGeneral%2FWorking_Files_Naoya%2FMetabotTest&FolderCTID=0x0120000223E4C4F9AB2C4DA1E560A6277B76AB";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.GetItemsURL(workingFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Failed to get a list of folders in specified folder with error: Unknown Error,FAIL:Failed to get a list of files in specified folder with error: Unknown Error
        }
        [TestMethod, TestCategory("GetItemsURL Method"), TestCategory("Fail")]
        public void GetItemsURLWithSharePointLink()
        {
            /// A link URL getting from SharePoint 'CopyLink' functionality.
            var workingFolder = @"https://automationanywhere1.sharepoint.com/:f:/s/SolutionArchitectTeam/EhcY8CIRG4tFrGx6EWINcCEBlZ0vvYCLSmNCU5eeEDb4tQ?e=lftPCq";
            var api = this.InitializeWrapper(siteUrl, username, password);

            var res = api.GetItemsURL(workingFolder);
            CheckFailCode(res);
            Console.WriteLine(res); // FAIL:Failed to get a list of folders in specified folder with error: URL is not for this web,FAIL:Failed to get a list of files in specified folder with error: URL is not for this web
        }

        /**
         * Below is utilities.
         */
        /// <summary>
        ///  This is the utility private method to authentiate the site.
        ///  This will return the SharePointAPIWrwapper instance to call other methods.
        /// </summary>
        private SharePointAPIWrapper InitializeWrapper(string url, string username, string password)
        {
            var wrapper = new SharePointAPIWrapper();
            wrapper.Authenticate(siteUrl, username, password);

            return wrapper;
        }
        /// <summary>
        ///  Output each value in the arguments.
        /// </summary>
        /// <param name="list">String array to be output</param>
        private void OutputStringArray(string[] list)
        {
            if (list.Length > 0)
            {
                foreach (var item in list)
                {
                    Console.WriteLine(item);
                }

            }
        }
        /// <summary>
        ///  Check whether specified item is in the list or not.
        /// </summary>
        /// <param name="list">String array to be output</param>
        /// <param name="target">the name to be checked.</param>
        private bool CheckItemInList(string[] list, string target)
        {
            var isExist = false;
            if (list.Length > 0)
            {
                foreach (var item in list)
                {
                    if (item.ToString() == target)
                    {
                        isExist = true;
                        break;
                    }
                }
            }
            return isExist;
        }
        /// <summary>
        ///   Check whether specified target string is in the list.
        /// </summary>
        /// <param name="result">string getting from each method.</param>
        private void CheckAuthenticationFail(string result)
        {
            Assert.AreEqual("FAIL:Authentication is needed.", result);
        }
        /// <summary>
        ///   Check whether the string starting with "FAIL:" or not.
        /// </summary>
        /// <param name="result">string getting from each method.</param>
        private void CheckFailCode(string result)
        {
            var code = result.Substring(0, 5);

            Assert.AreEqual("FAIL:", code);
        }
        /// <summary>
        ///   Delete a file if the file is already exist.
        /// </summary>
        /// <param name="filePath">Local file path</param>
        private void CleanUpLocalFile(string filePath)
        {
            // If file is there, then remove it.
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
        }
        /// <summary>
        ///  Delete a folder including subdirectories and files.
        /// </summary>
        /// <param name="folderPath">Local Folder Path</param>
        private void CleanUpLocalDirectory(string folderPath)
        {
            if (!System.IO.Directory.Exists(folderPath))
            {
                return;
            }
            System.IO.Directory.Delete(folderPath, true);
        }
        /// <summary>
        ///   Create a parent directory if that directory is NOT exist.
        /// </summary>
        /// <param name="filePath">Local file path</param>
        private void CreateLocalParentDirectory(string filePath)
        {
            // If the parent file is missing, then create that folder.
            var directoryName = Path.GetDirectoryName(filePath);
            if (!System.IO.Directory.Exists(directoryName))
            {
                System.IO.Directory.CreateDirectory(directoryName);
            }
        }
        /// <summary>
        ///   Create a local file. (If it is already exist, then remove it and create new one.)
        /// </summary>
        /// <param name="filePath">Local file path</param>
        private void CreateLocalFile(string filePath)
        {
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
        }
        /// <summary>
        ///   Add text into file. If file is not exist, it throws exception.
        /// </summary>
        /// <param name="filePath">Local file path</param>
        private void AddTextToLocalFile(string filePath)
        {
            using (StreamWriter sw = System.IO.File.AppendText(filePath))
            {
                sw.WriteLine(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + ": This is added from unit test");
            }

        }
    }
}
