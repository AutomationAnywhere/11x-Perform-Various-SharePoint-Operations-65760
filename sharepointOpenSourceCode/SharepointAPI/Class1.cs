/*
 * Version 2.0.0
 * Initial Author: Automation Anywhere
 * Current Version's Author: Naoya Murao (naoya.murao@automationanywhere.com)
 * High Level's Change Description: Add folder operation methods and related unit tests
 * Published Date: Apr 1st 2019
 */
using System;
using System.Security;
using Microsoft.SharePoint.Client;
using System.IO;
using System.Web;
using ClientOM = Microsoft.SharePoint.Client;

namespace AutomationAnywhere.MetaBot.SharePoint
{
    /// <summary>  
    ///  This class implements the methods to access Sharepoint Online.  
    /// </summary>  
    public class SharePointAPIWrapper
    {
        /// <summary>
        /// Store for the ClientContext property.
        /// </summary>
        ClientContext context { get; set; }

        /// <summary>
        /// delimiter for the list
        /// </summary>
        string delimiter = ",";

        /// <summary>
        ///   The Authenticate method to login to sharepoint with specified username and password.
        /// </summary>
        /// <param name="webUrl"> The server URL to login.</param>
        /// <param name="userName"> The user name to login.</param>
        /// <param name="password"> The password for the specified user.</param>
        /// <returns>
        ///   If authentication is successful, then this returns site title after the "Connected to " message.
        ///   Otherwise, return error message string starting with "FAIL:".
        /// </returns>
        /// <remarks>
        ///   Also see the following URL:
        ///     https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/jj164693(v%3doffice.15)
        /// </remarks>
        public string Authenticate(string webUrl, string userName, string password)
        {
            SecureString securePassword = new SecureString();

            try
            {
                foreach (char c in password)
                {
                    securePassword.AppendChar(c);
                }
                context = new ClientContext(webUrl);
                {
                    context.Credentials = new SharePointOnlineCredentials(userName, securePassword);
                    context.Load(context.Web, w => w.Title);
                    context.ExecuteQuery();

                    return $"Connected to {context.Web.Title}";
                }
            }
            catch (Exception ex)
            {
                return $"FAIL:Fail to authenticate with error: {ex.Message}";
            }
        }

        /// <summary>
        ///   The GetSiteTitle method to get the title of the sharepoint site.
        /// </summary>
        /// <returns>
        ///   This method returns site title.
        ///   If any errors happen, then return a string starting with "FAIL:".
        /// </returns>
        /// <remarks>
        ///   Also see the following URL:
        ///     https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee537040(v%3doffice.15)
        /// </remarks>
        public string GetSiteTitle()
        {
            if (!CheckClientContext())
            {
                return "FAIL:Authentication is needed.";
            }

            try
            {
                context.Load(context.Web, w => w.Title);
                context.ExecuteQuery();

                return context.Web.Title;
            }
            catch (Exception ex)
            {
                return $"FAIL:Fail to get site title with error: {ex.Message}";
            }

        }

        /// <summary>
        ///   The DownloadFile method to download a file located specified server path to specified local path including a file name.
        /// </summary>
        /// <param name="remoteFilePath"> The relative path to the file which will be downloaded.</param>
        /// <param name="localFilePath"> The absolute path to be downloaded.</param>
        /// <returns>
        ///   If any errors happen, then return a string starting with "FAIL:".
        /// </returns>
        /// <remarks>
        ///   Also see the following URL:
        ///     https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee542743(v%3doffice.15)
        /// </remarks>
        public string DownloadFile(string remoteFilePath, string localFilePath)
        {
            if (!CheckClientContext())
            {
                return "FAIL:Authentication is needed.";
            }

            try
            {
                FileInformation fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(context, remoteFilePath);

                using (var fileStream = new System.IO.FileStream(localFilePath, System.IO.FileMode.Create))
                {
                    fileInfo.Stream.CopyTo(fileStream);
                }

                return $"File Downloaded from {remoteFilePath} to {localFilePath}";
            }
            catch (Exception ex)
            {
                return $"FAIL:Fail to download the file with error:  {ex.Message}";
            }

        }

        /// <summary>
        ///   The UploadFile method to upload a local file to specified server path with specified file name.
        /// </summary>
        /// <param name="localFilePath"> The absolute path for a local file including the file name.</param>
        /// <param name="remoteFilePath"> The relative path to the folder which local file will be uploaded.</param>
        /// <returns>
        ///   TODO: will change the return value to useful value.
        /// </returns>
        /// <remarks>
        ///   Also see the following URL:
        ///     https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee542743(v%3doffice.15)
        /// </remarks>
        public string UploadFile(string localFilePath, string remoteFilePath)
        {
            if (!CheckClientContext())
            {
                return "FAIL:Authentication is needed.";
            }

            try
            {
                Console.WriteLine(remoteFilePath);

                string fileToUpload = localFilePath;
                using (FileStream fileStream =
                new FileStream(fileToUpload, FileMode.Open))
                    ClientOM.File.SaveBinaryDirect(context,
                        remoteFilePath, fileStream, true);

                return $"File uploaded from {localFilePath} to {remoteFilePath}";
            }
            catch (Exception ex)
            {
                return $"FAIL:Fail to upload the file with error: {ex.Message}";
            }

        }

        /// <summary>
        ///   The DeleteFile method to delete a file.
        /// </summary>
        /// <param name="path"> The relative path for the file.</param>
        /// <returns>
        ///   TODO: will change the return value to useful value.
        /// </returns>
        /// <remarks>
        ///   Also see the following URL:
        ///     https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee542743(v%3doffice.15)
        /// </remarks>
        public string DeleteFile(string path)
        {
            if (!CheckClientContext())
            {
                return "FAIL:Authentication is needed.";
            }

            try
            {
                var file = this.GetFileByPath(path);

                file.DeleteObject();
                context.ExecuteQuery();

                return $"Deleted the file located on {path}";
            }
            catch (Exception ex)
            {
                return $"FAIL:Fail to delete the file with error: {ex.Message}";
            }

        }

        /// <summary>
        ///   The CopyFile method to copy a file to a specified path.
        /// </summary>
        /// <param name="sourcePath"> The relative path for an original file.</param>
        /// <param name="destinationPath"> The relative path for a destination file.</param>
        /// <returns>
        ///   TODO: will change the return value to useful value.
        /// </returns>
        /// <remarks>
        ///   Also see the following URL:
        ///     https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee542743(v%3doffice.15)
        /// </remarks>
        public string CopyFile(string sourcePath, string destinationPath)
        {
            if (!CheckClientContext())
            {
                return "FAIL:Authentication is needed.";
            }

            try
            {
                var file = this.GetFileByPath(sourcePath);

                file.CopyTo(destinationPath, true);
                context.ExecuteQuery();

                return $"Copied the file from {sourcePath} to {destinationPath}";
            }
            catch (Exception ex)
            {
                return $"FAIL:Fail to copy the file with error: {ex.Message}";
            }
        }

        /// <summary>
        ///   The MoveFile method to move a file to a specified path.
        /// </summary>
        /// <param name="sourcePath"> The relative path for an original file.</param>
        /// <param name="destinationPath"> The relative path for a destination file.</param>
        /// <returns>
        ///   TODO: will change the return value to useful value.
        /// </returns>
        /// <remarks>
        ///   Also see the following URL:
        ///     https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee542743(v%3doffice.15)
        /// </remarks>
        public string MoveFile(string sourcePath, string destinationPath)
        {
            if (!CheckClientContext())
            {
                return "FAIL:Authentication is needed.";
            }

            try
            {
                var file = this.GetFileByPath(sourcePath);

                file.MoveTo(destinationPath, MoveOperations.Overwrite);
                context.ExecuteQuery();

                return $"Moved the file from {sourcePath} to {destinationPath}";
            }
            catch (Exception ex)
            {
                return $"FAIL:Fail to move the file with error: {ex.Message}";
            }
        }

        /// <summary>
        ///   The CheckOut method to check out a file which is located specified path.
        /// </summary>
        /// <param name="path"> The relative path for the file.</param>
        /// <returns>
        ///   TODO: will change the return value to useful value.
        /// </returns>
        /// <remarks>
        ///   Also see the following URL:
        ///     https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee542743(v%3doffice.15)
        /// </remarks>
        public string CheckOutFile(string path)
        {
            if (!CheckClientContext())
            {
                return "FAIL:Authentication is needed.";
            }

            try
            {
                var file = this.GetFileByPath(path);

                if (file.CheckOutType == CheckOutType.None)
                {
                    file.CheckOut();
                    context.ExecuteQuery();
                }
                return $"Checkout the file located on {path}";
            }
            catch (Exception ex)
            {
                return $"FAIL:Fail to checkout the file with error: {ex.Message}";
            }
        }

        /// <summary>
        ///   The CheckIn method to check in a file which is located specified path.
        /// </summary>
        /// <param name="path"> The relative path for the file.</param>
        /// <param name="comment"> The comment message to check in.</param>
        /// <returns>
        ///   TODO: will change the return value to useful value.
        /// </returns>
        /// <remarks>
        ///   Also see the following URL:
        ///     https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee542743(v%3doffice.15)
        ///   CheckInType is always MajorCheckIn because that option is not working with sharepoint online.
        /// </remarks>
        public string CheckInFile(string path, string comment)
        {
            if (!CheckClientContext())
            {
                return "FAIL:Authentication is needed.";
            }

            try
            {
                var file = this.GetFileByPath(path);

                var checkinType = CheckinType.MajorCheckIn;
                if (file.CheckOutType != CheckOutType.None)
                {
                    file.CheckIn(comment, checkinType);
                    context.ExecuteQuery();
                }

                return $"Checkin the file located on {path}";
            }
            catch (Exception ex)
            {
                return $"FAIL:Fail to checkin the file with error: {ex.Message}";
            }

        }

        /// <summary>
        ///   The UndoCheckOut method to undo checkout for the file which is located specified path
        /// </summary>
        /// <param name="path"> The relative path for the file.</param>
        /// <returns>
        ///   TODO: will change the return value to useful value.
        /// </returns>
        /// <remarks>
        ///   Also see the following URL:
        ///     https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee542743(v%3doffice.15)
        /// </remarks>
        public string UndoCheckOutFile(string path)
        {
            if (!CheckClientContext())
            {
                return "FAIL:Authentication is needed.";
            }

            try
            {
                var file = this.GetFileByPath(path);

                if (file.CheckOutType != CheckOutType.None)
                {
                    file.UndoCheckOut();
                    context.ExecuteQuery();
                }

                return $"Undo Checkout the File located on {path}";
            }
            catch (Exception ex)
            {
                return $"FAIL:Fail to undo checkout the file with error: {ex.Message}";
            }
        }

        /// <summary>
        ///   The CreateFolder method to create the spcified folder
        /// </summary>
        /// <param name="parentFolder"> The relative path for the folder which the new folder will be located. </param>
        /// <param name="folderName"> The new folder name. </param>
        /// <returns>
        ///   TODO: will change the return value to useful value.
        /// </returns>
        /// <remarks>
        ///   Also see the following URL:
        ///     https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee538304(v%3doffice.15)
        /// </remarks>
        public string CreateFolder(string parentFolder, string folderName)
        {
            if (!CheckClientContext())
            {
                return "FAIL:Authentication is needed.";
            }

            try
            {
                parentFolder = parentFolder.EndsWith(@"/") ? parentFolder : parentFolder + @"/";

                var folder = this.GetFolderByPath(parentFolder);

                folder.AddSubFolder(folderName);
                context.ExecuteQuery();

                return $"Create the folder located in {parentFolder}{folderName}";
            }
            catch (Exception ex)
            {
                return $"FAIL:Fail to create the folder with error: {ex.Message}";
            }
        }

        /// <summary>
        ///   The DeleteFolder method to delete the spcified folder
        /// </summary>
        /// <param name="path"> The relative path for the folder which will delete. </param>
        /// <returns>
        ///   TODO: will change the return value to useful value.
        /// </returns>
        /// <remarks>
        ///   Also see the following URL:
        ///     https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee538304(v%3doffice.15)
        /// </remarks>
        public string DeleteFolder(string path)
        {
            if (!CheckClientContext())
            {
                return "FAIL:Authentication is needed.";
            }

            try
            {
                var folder = this.GetFolderByPath(path);

                folder.DeleteObject();
                context.ExecuteQuery();

                return $"Deleted the folder located on {path}";
            }
            catch (Exception ex)
            {
                return $"FAIL:Fail to delete the folder with error: {ex.Message}";
            }
        }

        /// <summary>
        ///   The MoveFolder method to move the spcified folder to specified path
        /// </summary>
        /// <param name="sourcePath"> The relative path for the original folder which will move.</param>
        /// <param name="destinationPath"> The relative path for the destination folder.</param>
        /// <returns>
        ///   TODO: will change the return value to useful value.
        /// </returns>
        /// <remarks>
        ///   Also see the following URL:
        ///     https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee538304(v%3doffice.15)
        /// </remarks>
        public string MoveFolder(string sourcePath, string destinationPath)
        {
            if (!CheckClientContext())
            {
                return "FAIL:Authentication is needed.";
            }

            try
            {
                var folder = this.GetFolderByPath(sourcePath);

                folder.MoveTo(destinationPath);
                context.ExecuteQuery();

                return $"Moved the folder from {sourcePath} to {destinationPath}";
            }
            catch (Exception ex)
            {
                return $"FAIL:Fail to move the folder with error: {ex.Message}";
            }
        }

        /// <summary>
        ///   The MoveFolder method to move the spcified folder to specified path.
        ///   Internally, this will download the all files and folders under temporary directory and upload those to new path.
        /// </summary>
        /// <param name="sourcePath"> The relative path for the original folder which will move.</param>
        /// <param name="destinationPath"> The relative path for the destination folder.</param>
        /// <param name="timeoutMin"> A settings for timeout to operate Sharepoint.</param>
        /// <returns>
        ///   TODO: will change the return value to useful value.
        /// </returns>
        /// <remarks>
        ///   Also see the following URL:
        ///     https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee538304(v%3doffice.15)
        /// </remarks>
        public string CopyFolder(string sourcePath, string destinationPath, int timeoutMin = 3)
        {
            if (!CheckClientContext())
            {
                return "FAIL:Authentication is needed.";
            }
            
            try
            {
                string tempDir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
                var tempDirInfo = Directory.CreateDirectory(tempDir);

                var resultList = "";
                var res = this.DownloadFolder(sourcePath, tempDir, timeoutMin);
                if (res.StartsWith("FAIL:"))
                {
                    resultList = res;
                }

                res = this.UploadFolder(tempDir, destinationPath, timeoutMin);
                if (res.StartsWith("FAIL:"))
                {
                    if (resultList.Length > 0)
                    {
                        resultList = $"{resultList},{res}";
                    }
                    else
                    {
                        resultList = res;
                    }
                }

                if (resultList.Length > 0)
                {
                    return resultList;
                } else
                {
                    return $"Copied the folder from {sourcePath} to {destinationPath}";
                }
            }
            catch (Exception ex)
            {
                return $"FAIL:Failed to copy the folder with error: {ex.Message}";
            }
        }

        /// <summary>
        ///   The DownloadFolder method to download files and folders under specified folder.
        /// </summary>
        /// <param name="remotePath"> The relative path to the file which will be downloaded.</param>
        /// <param name="localPath"> The absolute path in local computer.</param>
        /// <param name="timeoutMin"> A settings for timeout to operate Sharepoint.</param>
        /// <returns>
        ///   TODO: will change the return value to useful value.
        /// </returns>
        /// <remarks>
        ///   Also see the following URL:
        ///     https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee542743(v%3doffice.15)
        /// </remarks>
        public string DownloadFolder(string remotePath, string localPath, int timeoutMin = 3)
        {
            if (!CheckClientContext())
            {
                return "FAIL:Authentication is needed.";
            }

            context.RequestTimeout = 1000 * 60 * timeoutMin;
            try
            {
                localPath = localPath.EndsWith(@"\") ? localPath.Substring(0, localPath.Length - 1) : localPath;
                remotePath = remotePath.EndsWith(@"/") ? remotePath.Substring(0, remotePath.Length - 1) : remotePath;

                var isError = false;
                var resultList = "";

                if (!Directory.Exists(localPath))
                {
                    Directory.CreateDirectory(localPath);
                }

                var folder = this.GetFolderByPath(remotePath);
                FileCollection fileCollection = folder.Files;
                context.Load(fileCollection);
                context.ExecuteQuery();

                if (fileCollection.Count > 0)
                {
                    foreach (ClientOM.File fileItem in fileCollection)
                    {
                        context.Load(fileItem);
                        context.ExecuteQuery();

                        var fileName = fileItem.Name.ToString();
                        var localFilePath = Path.Combine(localPath, fileName);
                        var res = this.DownloadFile(fileItem.ServerRelativeUrl.ToString(), localFilePath);
                        if (res.StartsWith("FAIL:"))
                        {
                            resultList = res;
                            isError = true;
                        }
                    }
                }

                FolderCollection folderCollection = folder.Folders;
                context.Load(folderCollection);
                context.ExecuteQuery();

                if (folderCollection.Count > 0)
                {
                    foreach (ClientOM.Folder folderItem in folderCollection)
                    {
                        context.Load(folderItem);
                        context.ExecuteQuery();

                        var folderName = folderItem.Name;
                        var localFolderPath = Path.Combine(localPath, folderName);
                        var res = this.DownloadFolder(folderItem.ServerRelativeUrl.ToString(), localFolderPath);
                        if (res.StartsWith("FAIL:"))
                        {
                            if (isError)
                            {
                                resultList = $"{resultList},{res}";
                            }
                            else
                            {
                                resultList = res;
                            }
                            isError = true;
                        }
                    }
                }
                if (isError)
                {
                    return resultList;
                }
                return $"Downloaded the folder from {remotePath} to {localPath}";
            }
            catch (Exception ex)
            {
                return $"FAIL:Failed to download the folder with error: {ex.Message}";
            }
        }

        /// <summary>
        ///   The UploadFolder method to upload files and folders under the specified folder.
        /// </summary>
        /// <param name="localPath"> The relative path to the file which will be downloaded.</param>
        /// <param name="remotePath"> The absolute path in local computer.</param>
        /// <param name="timeoutMin"> A settings for timeout to operate Sharepoint.</param>
        /// <returns>
        ///   TODO: will change the return value to useful value.
        /// </returns>
        /// <remarks>
        ///   Also see the following URL:
        ///     https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee542743(v%3doffice.15)
        /// </remarks>
        public string UploadFolder(string localPath, string remotePath, int timeoutMin = 3)
        {
            if (!CheckClientContext())
            {
                return "FAIL:Authentication is needed.";
            }

            context.RequestTimeout = 1000 * 60 * timeoutMin;

            try
            {
                localPath = localPath.EndsWith(@"\") ? localPath.Substring(0, localPath.Length - 1) : localPath;
                remotePath = remotePath.EndsWith(@"/") ? remotePath.Substring(0, remotePath.Length - 1) : remotePath;

                var isError = false;
                var resultList = "";
                if (Directory.Exists(localPath))
                {
                    if (!IsFolderExist(remotePath))
                    {
                        var folderName = this.GetFolderName(remotePath);
                        var parentFolderPath = this.GetParentFolderPath(remotePath);

                        this.CreateFolder(parentFolderPath, folderName);
                    }

                    string[] files = Directory.GetFiles(localPath);
                    if (files.Length > 0) {
                        foreach (string file in files)
                        {
                            var fileInfo = new FileInfo(file);
                            var fileName = fileInfo.Name;
                            var remoteFilePath = $"{remotePath}/{fileName}";
                            var res = this.UploadFile(file, remoteFilePath);
                            if (res.StartsWith("FAIL:"))
                            {
                                resultList = res;
                                isError = true;
                            }
                        }
                    }

                    string[] subdirs = Directory.GetDirectories(localPath);
                    if (subdirs.Length > 0)
                    {
                        foreach (string subdir in subdirs)
                        {
                            var dirInfo = new DirectoryInfo(subdir);
                            var folderName = dirInfo.Name;
                            var remotSubDirPath = remotePath + "/" + folderName;
                            var res = this.UploadFolder(subdir, remotSubDirPath, timeoutMin);
                            if (res.StartsWith("FAIL:"))
                            {
                                if (isError)
                                {
                                    resultList = $"{resultList},{res}";
                                }
                                else
                                {
                                    resultList = res;
                                }
                                isError = true;
                            }
                        }
                    }
                    if (isError)
                    {
                        return resultList;
                    }
                    return $"Uploaded the folder from {localPath} to {remotePath}";
                }
                else
                {
                    return $"FAIL:Failed to upload the folder with error: {localPath} not found";
                }

            }
            catch (Exception ex)
            {
                return $"FAIL:Failed to upload the folder with error: {ex.Message}";
            }
        }

        /// <summary>
        ///   The GetItemsURL method to get the url of each items (files and folders) in the specified folder.
        /// </summary>
        /// <param name="path"> The relative path for a folder.</param>
        /// <returns>
        ///   The list of URL with the comma delimiter.
        /// </returns>
        /// <remarks>
        ///   Also see the following URL:
        ///     https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee538304(v%3doffice.15)
        ///     https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee541625%28v%3doffice.15%29
        /// </remarks>
        public string GetItemsURL(string path)
        {
            if (!CheckClientContext())
            {
                return "FAIL:Authentication is needed.";
            }

            try
            {
                var builder = new System.Text.StringBuilder();

                var folders = this.GetFoldersURL(path);
                var files = this.GetFilesURL(path);

                builder.Append(folders);
                builder.Append(delimiter);
                builder.Append(files);

                return builder.ToString();
            }
            catch (Exception ex)
            {
                return $"FAIL:Failed to get a list of files and folders in specified folder with error: {ex.Message}";
            }

        }

        /// <summary>
        ///   The GetFolderItems method to get the url of each folders in the specified folder.
        /// </summary>
        /// <param name="path"> The relative path for a folder.</param>
        /// <returns>
        ///   The list of URL with the comma delimiter.
        /// </returns>
        /// <remarks>
        ///   Also see the following URL:
        ///     https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee538304(v%3doffice.15)
        ///     https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee541625%28v%3doffice.15%29
        /// </remarks>
        public string GetFoldersURL(string path)
        {
            if (!CheckClientContext())
            {
                return "FAIL:Authentication is needed.";
            }

            try
            {
                var itemsList = new System.Text.StringBuilder();

                var folder = this.GetFolderByPath(path);
                FolderCollection folderCollection = folder.Folders;
                context.Load(folderCollection);
                context.ExecuteQuery();

                /// Add Folder Item's URL to result list
                foreach (Folder folderItem in folderCollection)
                {
                    context.Load(folderItem);
                    context.ExecuteQuery();

                    if (itemsList.Length > 0)
                    {
                        itemsList.Append(delimiter);
                    }
                    itemsList.Append(folderItem.ServerRelativeUrl.ToString());
                }

                return itemsList.ToString();
            }
            catch (Exception ex)
            {
                return $"FAIL:Failed to get a list of folders in specified folder with error: {ex.Message}";
            }
        }

        /// <summary>
        ///   The GetFilesURL method to get the url of each files in the specified folder.
        /// </summary>
        /// <param name="path"> The relative path for a folder.</param>
        /// <returns>
        ///   The list of URL with the comma delimiter.
        /// </returns>
        /// <remarks>
        ///   Also see the following URL:
        ///     https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee538304(v%3doffice.15)
        ///     https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee541625%28v%3doffice.15%29
        /// </remarks>
        public string GetFilesURL(string path)
        {
            if (!CheckClientContext())
            {
                return "FAIL:Authentication is needed.";
            }

            try
            {
                var itemsList = new System.Text.StringBuilder();

                var folder = this.GetFolderByPath(path);
                FileCollection fileCollection = folder.Files;
                context.Load(fileCollection);
                context.ExecuteQuery();

                /// Add File Item's URL to result list
                foreach (Microsoft.SharePoint.Client.File fileItem in fileCollection)
                {
                    context.Load(fileItem);
                    context.ExecuteQuery();

                    if (itemsList.Length > 0)
                    {
                        itemsList.Append(delimiter);
                    }
                    itemsList.Append(fileItem.ServerRelativeUrl.ToString());
                }

                return itemsList.ToString();
            }
            catch (Exception ex)
            {
                return $"FAIL:Failed to get a list of files in specified folder with error: {ex.Message}";
            }

        }

        /// <summary>
        ///   The GetFileByPath method to generate Microsoft.Sharepoint.Client.File object from specified path.
        /// </summary>
        /// <param name="path"> The relative path for a file.</param>
        /// <returns>
        ///   Microsoft.Sharepoint.Client.File instance.
        /// </returns>
        /// <remarks>
        ///   Also see the following URL:
        ///     https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee542743(v%3doffice.15)
        ///   context.Load and context.ExecuteQuery methods are important to get the target file's properties.
        /// </remarks>
        private ClientOM.File GetFileByPath(string path)
        {
            var file = context.Web.GetFileByServerRelativeUrl(path);
            context.Load(file);
            context.ExecuteQuery();

            return file;
        }

        /// <summary>
        ///   The GetFolderByPath method to generate Microsoft.Sharepoint.Client.Folder object from specified path.
        /// </summary>
        /// <param name="path"> The relative path for a folder.</param>
        /// <returns>
        ///   Microsoft.Sharepoint.Client.Folder instance.
        /// </returns>
        /// <remarks>
        ///   Also see the following URL:
        ///     https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee538304(v%3doffice.15)
        ///   context.Load and context.ExecuteQuery methods are important to get the target folder's properties.
        /// </remarks>
        private ClientOM.Folder GetFolderByPath(string path)
        {
            var folder = context.Web.GetFolderByServerRelativeUrl(path);
            context.Load(folder);
            context.ExecuteQuery();

            return folder;
        }

        /// <summary>
        ///   The CheckClientContext method to check the Microsoft.Sharepoint.Client.Context is already generated or not.
        ///   If not, this method returns false as bool value.
        /// </summary>
        /// <returns>
        ///   true, if the context is already generated.
        /// </returns>
        private bool CheckClientContext()
        {
            if (context == null)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        ///  The IsFolderExist method is to check existence for specified folder.
        /// </summary>
        /// <param name="path">The relative path for the folder</param>
        /// <returns>
        ///   true if the folder is exist.
        /// </returns>
        private bool IsFolderExist(string path)
        {
            var folder = context.Web.GetFolderByServerRelativeUrl(path);
            context.Load(folder, f => f.Exists);

            try
            {
                context.ExecuteQuery();

                if (folder.Exists)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                // nop. Console.WriteLine($"Could not find or access to the folder: {path}, {ex.Message}");
            }
            return false;
        }

        /// <summary>
        ///  The IsFileExist method is to check existence for specified file.
        /// </summary>
        /// <param name="path">The relative path for the file</param>
        /// <returns>
        ///   true if the file is exist.
        /// </returns>
        private bool IsFileExist(string path)
        {
            var file = context.Web.GetFileByServerRelativeUrl(path);
            context.Load(file, f => f.Exists);
            try
            {
                context.ExecuteQuery();

                if (file.Exists)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                //nop. Console.WriteLine($"Could not find or access to the folder: {path}, {ex.Message}");
            }
            return false;
        }

        private string GetFolderName(string path)
        {
            var lastIdx = path.LastIndexOf("/");
            var folderName = path.Substring(lastIdx + 1);

            return folderName;
        }
        private string GetParentFolderPath(string path)
        {
            var lastIdx = path.LastIndexOf("/");
            var parentFolderPath = path.Substring(0, lastIdx);

            return parentFolderPath;
        }

    }

};