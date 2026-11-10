using Azure.Core;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Identity.Client;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using System.Web;
using static System.Net.WebRequestMethods;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using static WcfWCService.OneDriveInteraction;
using Task = System.Threading.Tasks.Task;

namespace WcfWCService
{
    public class OneDriveInteraction
    {

        /*        private const string ClientId = "78d7fc1a-b197-467f-bcf9-30dce4339bda";
                private const string TenantId = "5a5ad5c0-787a-46e4-9097-e6c41b3c97fc"; // or "common" / "organizations"
                private const string Username = "LMS@regainmaterials.com";
                private const string Password = "Nrk#2026";
                private const string ClientSecret = "dgP8Q~JEhjRJNNLT28fqAkgIvMYP0baZWBP6.cIP";
        */

        private string m_driveId;
        private string m_CADFolderId;
        private string m_AccessToken;
        private string m_MatCatId;
        private List<FolderItem> m_MatCatFolders;

        public class FolderDriveIds
        {
            public string sFolderId;
            public string sDriveId;
        }

        public class FolderItem
        {
            public string id;
            public string type;
            public string name;
            public string driveId;
        }
        public class FolderPageResult
        {
            public List<FolderItem> FolderItems { get; set; } = new List<FolderItem>();
            public string NextLink { get; set; }
        }

        public class FileItem
        {
            public string id;
            public string name;
            public string type;
            public string downloadurl;
        }
        public class FileItemExtra
        {
            public string id;
            public string name;
            public string type;
            public string extension;
            public string revision;
            public bool revfoundpdf;
            public bool revfounddxf;
            public bool revfoundstep;
            public string donwloadurl;
        }
        public class FileRequirement
        {
            public bool bResult;
            public bool bPDF;
            public bool bDXF;
            public bool bSTEP;
            public string sErrorMsg;
        }

        public class FilePageResult
        {
            public List<FileItem> FileItems { get; set; } = new List<FileItem>();
            public string NextLink { get; set; }
        }

        public class CADFileResult
        {
            public string sOverallComment = "";
            public string sPDFComment = "";
            public string sDXFComment = "";
            public string sSTEPComment = "";
            public string sLMSComment = "";
            public string sPDFRevision = "";
            public string sDXFRevision = "";
            public string sSTEPRevision = "";
            public string sOverallRevision = "";
            public string sStatus = "";
            public string sStatusComments = "";

            public void Clear()
            {
                sOverallComment = "";
                sPDFComment = "";
                sDXFComment = "";
                sSTEPComment = "";
                sLMSComment = "";
                sPDFRevision = "";
                sDXFRevision = "";
                sSTEPRevision = "";
                sOverallRevision = "";
            }

        }

        public async Task<string> SetUserAccessToken()
        {
            Environment env = new Environment();
            string accessToken = null;
            string sWCIISAppClientId = env.Get_Environment_String_Value("WCIISApplicationClientId");
            string sWCIISAppTenantId = env.Get_Environment_String_Value("WCIISApplicationTenantId");
            string sWCIISApplicationUsername = env.Get_Environment_String_Value("WCIISApplicationUsername");
            string sWCIISApplicationPassword = env.Get_Environment_String_Value("WCIISApplicationPassword");

            // 1. Force the runtime to use TLS 1.2 for all outbound calls in this process
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

            // 2. Clear out any stale service point connections 
            // (Ensures the change takes effect immediately)
            System.Net.ServicePointManager.FindServicePoint(new Uri("https://microsoftonline.com")).ConnectionLimit = 10;

            try
            {

                string[] scopes = new[] { "User.Read", "Files.Read.All" }; // "Files.Read.All", "User.Read"
                //string[] scopes = new[] { "https://graph.microsoft.com/.default" };
                var app = PublicClientApplicationBuilder.Create(sWCIISAppClientId)
                    .WithAuthority(AzureCloudInstance.AzurePublic, sWCIISAppTenantId)
                    .Build();

                var securePassword = new System.Security.SecureString();
                foreach (char c in sWCIISApplicationPassword) securePassword.AppendChar(c);

                // Acquire token via username & password (ROPC)
                AuthenticationResult result = await app.AcquireTokenByUsernamePassword(scopes, sWCIISApplicationUsername, securePassword).ExecuteAsync();

                // Acquire token using Client Credentials
                //AuthenticationResult result = await app.AcquireTokenForClient(scopes).ExecuteAsync(); 

                System.Diagnostics.Debug.WriteLine("Token acquired: " + result.AccessToken);
                accessToken = result.AccessToken;
                m_AccessToken = accessToken;

                return accessToken;
            }
            catch (System.Exception e)
            {
                System.Diagnostics.Debug.WriteLine($"Error occurred: {e.StackTrace}");
                return e.Message;
            }

        }

        public async Task<string> GetApplicationAccessToken()
        {
            Environment env = new Environment();
            string accessToken = null;
            string sWCIISAppClientId = env.Get_Environment_String_Value("WCIISApplicationClientId");
            string sWCIISAppTenantId = env.Get_Environment_String_Value("WCIISApplicationTenantId");
            string sWCIISApplicationClientSecret = env.Get_Environment_String_Value("WCIISApplicationClientSecret");
            string sWCIISApplicationPassword = env.Get_Environment_String_Value("WCIISApplicationPassword");

            // 1. Force the runtime to use TLS 1.2 for all outbound calls in this process
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

            // 2. Clear out any stale service point connections 
            // (Ensures the change takes effect immediately)
            System.Net.ServicePointManager.FindServicePoint(new Uri("https://microsoftonline.com")).ConnectionLimit = 10;

            try
            {

                //string[] scopes = new[] {"Files.Read.All" }; // "Files.Read.All", "User.Read"
                string[] scopes = new[] { "https://graph.microsoft.com/.default" };

                var app = ConfidentialClientApplicationBuilder.Create(sWCIISAppClientId)
                        .WithClientSecret(sWCIISApplicationClientSecret) // Paste your generated client secret value here
                        .WithAuthority(AzureCloudInstance.AzurePublic, sWCIISAppTenantId)
                        .Build();


                // Acquire token using Client Credentials
                AuthenticationResult result = await app.AcquireTokenForClient(scopes).ExecuteAsync();

                System.Diagnostics.Debug.WriteLine("Token acquired: " + result.AccessToken);
                accessToken = result.AccessToken;

                return accessToken;
            }
            catch (System.Exception e)
            {
                System.Diagnostics.Debug.WriteLine($"Error occurred: {e.StackTrace}");
                return e.Message;
            }

        }


        static async Task<List<FolderItem>> ReadFolderAsync(string accessToken, string sEndPoint, bool bHasRemoteElement)
        {
            List<FolderItem> folderItems = new List<FolderItem>();

            using (var httpClient = new HttpClient())
            {
                // 1. Set the authorization header using the app token
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                HttpResponseMessage response = await httpClient.GetAsync(sEndPoint);

                if (response.IsSuccessStatusCode)
                {
                    string sNextLink = "";
                    string jsonResponse = await response.Content.ReadAsStringAsync();
                    using (JsonDocument doc = JsonDocument.Parse(jsonResponse))
                    {
                        JsonElement root = doc.RootElement;

                        if (root.TryGetProperty("value", out JsonElement valueElement))
                        {
                            foreach (JsonElement item in valueElement.EnumerateArray())
                            {
                                FolderItem itemFold = new FolderItem();

                                string name = item.GetProperty("name").GetString();
                                string id = item.GetProperty("id").GetString();
                                // Check if it's a file or a nested folder
                                string type = item.TryGetProperty("folder", out _) ? "Folder" : "File";

                                // Safely traverse down the nested object
                                string sDriveId = "";
                                if (item.TryGetProperty("remoteItem", out _))
                                {
                                    if (bHasRemoteElement)
                                    {
                                        sDriveId = item.GetProperty("remoteItem")
                                                         .GetProperty("parentReference")
                                                         .GetProperty("driveId")
                                                         .GetString();
                                    }
                                }

                                System.Diagnostics.Debug.WriteLine($"[{type}] Name: {name} (ID: {id})");
                                System.Diagnostics.Trace.WriteLine($"[{type}] Name: {name} (ID: {id})");

                                itemFold.name = name;
                                itemFold.type = type;
                                itemFold.id = id;
                                itemFold.driveId = sDriveId;

                                folderItems.Add(itemFold);

                            }
                        }
                    }
                }
                else
                {
                    string errorContent = await response.Content.ReadAsStringAsync();
                    System.Diagnostics.Debug.WriteLine($"Failed to read folder. Status Code: {response.StatusCode}");
                    System.Diagnostics.Debug.WriteLine($"Error Details: {errorContent}");
                }
            }

            return folderItems;
        }

        //This reads the folder and list 200 folder items at a time. It also list the NextUrl paramenter which is
        //used to get the next 200 items. Keep looping through until no more items left.
        static async Task<FolderPageResult> ReadFolderAsync2(string accessToken, string sEndPoint)
        {
            FolderPageResult result = new FolderPageResult();

            try
            {
                using (var httpClient = new HttpClient())
                {
                    // Set standard authorization headers
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    // Use HttpCompletionOption.ResponseHeadersRead to prevent the HttpClient 
                    // from prematurely choking on a massive memory string conversion buffer.
                    HttpResponseMessage response = await httpClient.GetAsync(sEndPoint, HttpCompletionOption.ResponseHeadersRead);

                    if (response.IsSuccessStatusCode)
                    {
                        // STREAM DIRECTLY: Get the raw network stream without loading a huge string into memory
                        using (Stream responseStream = await response.Content.ReadAsStreamAsync())
                        {
                            // ParseAsync processes the network byte stream safely and sequentially
                            using (JsonDocument doc = await JsonDocument.ParseAsync(responseStream))
                            {
                                JsonElement root = doc.RootElement;

                                // Extract the next page token if it exists
                                if (root.TryGetProperty("@odata.nextLink", out JsonElement nextLinkElement))
                                {
                                    result.NextLink = nextLinkElement.GetString();
                                }

                                if (root.TryGetProperty("value", out JsonElement valueElement))
                                {
                                    foreach (JsonElement item in valueElement.EnumerateArray())
                                    {
                                        FolderItem itemFold = new FolderItem();

                                        string name = item.GetProperty("name").GetString();
                                        string id = item.GetProperty("id").GetString();
                                        string type = item.TryGetProperty("folder", out _) ? "Folder" : "File";

                                        string sDriveId = "";

                                        itemFold.name = name;
                                        itemFold.type = type;
                                        itemFold.id = id;
                                        itemFold.driveId = sDriveId;

                                        result.FolderItems.Add(itemFold);
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        // If the Graph API rejected your request, handle the error gracefully here
                        string errorContent = await response.Content.ReadAsStringAsync();
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.ToString());
                return null;
            }
        }

        //This reads the folder and list 200 file items at a time. It also list the NextUrl paramenter which is
        //used to get the next 200 items. Keep looping through until no more items left.
        static async Task<FilePageResult> ReadFileAsync2(string accessToken, string sEndPoint)
        {
            FilePageResult result = new FilePageResult();

            try
            {
                using (var httpClient = new HttpClient())
                {
                    // Set standard authorization headers
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    // Use HttpCompletionOption.ResponseHeadersRead to prevent the HttpClient 
                    // from prematurely choking on a massive memory string conversion buffer.
                    HttpResponseMessage response = await httpClient.GetAsync(sEndPoint, HttpCompletionOption.ResponseHeadersRead);

                    if (response.IsSuccessStatusCode)
                    {
                        // STREAM DIRECTLY: Get the raw network stream without loading a huge string into memory
                        using (Stream responseStream = await response.Content.ReadAsStreamAsync())
                        {
                            // ParseAsync processes the network byte stream safely and sequentially
                            using (JsonDocument doc = await JsonDocument.ParseAsync(responseStream))
                            {
                                JsonElement root = doc.RootElement;

                                // Extract the next page token if it exists
                                if (root.TryGetProperty("@odata.nextLink", out JsonElement nextLinkElement))
                                {
                                    result.NextLink = nextLinkElement.GetString();
                                }

                                if (root.TryGetProperty("value", out JsonElement valueElement))
                                {
                                    foreach (JsonElement item in valueElement.EnumerateArray())
                                    {
                                        FileItem itemFile = new FileItem();

                                        string name = item.GetProperty("name").GetString();
                                        string id = item.GetProperty("id").GetString();
                                        string type = item.TryGetProperty("folder", out _) ? "Folder" : "File";
                                        string downloadurl = item.TryGetProperty("@microsoft.graph.downloadUrl", out _) ? item.GetProperty("@microsoft.graph.downloadUrl").GetString() : "";

                                        itemFile.name = name;
                                        itemFile.type = type;
                                        itemFile.id = id;
                                        itemFile.downloadurl = downloadurl;

                                        result.FileItems.Add(itemFile);
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        // If the Graph API rejected your request, handle the error gracefully here
                        string errorContent = await response.Content.ReadAsStringAsync();
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.ToString());
                return null;
            }
        }

        private static async Task DownloadFileFromUrlAsync(string url, string outputPath)
        {
            HttpClient httpClient = new HttpClient();

            // 2. Request the file from the pre-authenticated URL. 
            // Use HttpCompletionOption.ResponseHeadersRead to stream the content efficiently without buffering it all in memory.
            using (HttpResponseMessage response = await httpClient.GetAsync(url, HttpCompletionOption.ResponseHeadersRead))
            {
                response.EnsureSuccessStatusCode();

                // 3. Stream the file directly to your local disk
                using (Stream downloadStream = await response.Content.ReadAsStreamAsync())
                using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None, 8192, true))
                {
                    await downloadStream.CopyToAsync(fileStream);
                }
            }
        }

        public async Task<FolderDriveIds> SetCADFolderAndDriveId()
        {
            string sAccessToekn = m_AccessToken;
            FolderDriveIds folderDriveIds = new FolderDriveIds();
            List<FolderItem> folderItems = new List<FolderItem>();

            string endpoint = "https://graph.microsoft.com/v1.0/me/drive/sharedWithMe";
            folderItems = await ReadFolderAsync(sAccessToekn, endpoint, true);
            foreach (FolderItem folderItem in folderItems)
            {
                if (folderItem.name.Equals("CAD") && folderItem.type.Equals("Folder"))
                {
                    folderDriveIds.sFolderId = folderItem.id;
                    folderDriveIds.sDriveId = folderItem.driveId;
                    m_CADFolderId = folderItem.id;
                    m_driveId = folderItem.driveId;
                }
            }
            return folderDriveIds;
        }

        public async Task<FolderDriveIds> SetMaterialCatalogueDriveId()
        {
            string sAccessToekn = m_AccessToken;
            FolderDriveIds folderDriveIds = new FolderDriveIds();
            List<FolderItem> folderItems = new List<FolderItem>();

            string endpoint = "https://graph.microsoft.com/v1.0/drives/" + m_driveId +
                              "/items/" + m_CADFolderId + "/children";
            folderItems = await ReadFolderAsync(sAccessToekn, endpoint, true);
            foreach (FolderItem folderItem in folderItems)
            {
                if (folderItem.name.StartsWith("316 Material Catalog") && folderItem.type.Equals("Folder"))
                {
                    folderDriveIds.sFolderId = folderItem.id;
                    folderDriveIds.sDriveId = folderItem.driveId;
                    m_MatCatId = folderItem.id;
                }
            }
            return folderDriveIds;
        }

        public async Task<List<FolderItem>> SetMaterialCatalogueFoldersObject()
        {
            string sAccessToekn = m_AccessToken;
            m_MatCatFolders = new List<FolderItem>();

            string nextUrl = "https://graph.microsoft.com/v1.0/drives/" + m_driveId +
                              "/items/" + m_MatCatId + "/children";

            // Keep looping while the API continues to feed us a nextLink address
            while (!string.IsNullOrEmpty(nextUrl))
            {
                FolderPageResult pageResult = await ReadFolderAsync2(sAccessToekn, nextUrl);

                if (pageResult?.FolderItems != null)
                {
                    foreach (FolderItem folderItem in pageResult.FolderItems)
                    {
                        if (folderItem.type.Equals("Folder"))
                        {
                            FolderItem thisItem = new FolderItem
                            {
                                name = folderItem.name,
                                type = folderItem.type,
                                id = folderItem.id
                            };
                            m_MatCatFolders.Add(thisItem);
                        }
                    }
                }

                // Re-target the loop iterator condition directly to the returned nextLink string variable.
                // It smoothly returns null on the very last page, cleanly dropping out of the while loop.
                nextUrl = pageResult?.NextLink;
            }

            return m_MatCatFolders;
        }

        public async Task<List<FolderItem>> SetMaterialCatalogueFolderItemsObject()
        {
            string sAccessToken = m_AccessToken;
            m_MatCatFolders = new List<FolderItem>();

            string nextUrl = "https://graph.microsoft.com/v1.0/drives/" + m_driveId +
                              "/items/" + m_MatCatId + "/children";

            // Keep looping while the API continues to feed us a nextLink address
            while (!string.IsNullOrEmpty(nextUrl))
            {
                FolderPageResult pageResult = await ReadFolderAsync2(sAccessToken, nextUrl);

                if (pageResult?.FolderItems != null)
                {
                    foreach (FolderItem folderItem in pageResult.FolderItems)
                    {
                        if (folderItem.type.Equals("Folder"))
                        {
                            FolderItem thisItem = new FolderItem
                            {
                                name = folderItem.name,
                                type = folderItem.type,
                                id = folderItem.id
                            };
                            m_MatCatFolders.Add(thisItem);
                        }
                    }
                }

                // Re-target the loop iterator condition directly to the returned nextLink string variable.
                // It smoothly returns null on the very last page, cleanly dropping out of the while loop.
                nextUrl = pageResult?.NextLink;
            }

            return m_MatCatFolders;
        }

        // iUpload = 0 means verify only
        // iUpload = 1 means attach to the windchill container
        public async Task<CADFileResult> FindCADFile(string sSessionId, string sUserId, string sFullname, 
                                                     string sPart, string sRevisionPDF,
                                                     string sRevisionDXF, string sRevisionSTEP,
                                                     int iUpload, int iWebAppId)
        {
            Service1 serv = new Service1();
            List<FileItemExtra> fileItems = new List<FileItemExtra>();
            FileRequirement fileReqs = new FileRequirement();
            CADFileResult result = new CADFileResult();
            List<string> allRevs = new List<string>();
            List<FileItemExtra> fileToDownload = new List<FileItemExtra>();
            string sLMSRevision = "", sLMSDocContName = "";
            string sMaxRev = "", sJobCode = "M", s2ndMaxRev = "";

            try
            {
                bool bDocExists = serv.DocExists(sPart, iWebAppId);

                bool bFolderExists = false;
                string sFolderId = "";
                if (bDocExists)
                {
                    //Get the LMS Revision
                    sLMSRevision = serv.GetDocRevision(sPart, iWebAppId);
                    allRevs.Add(sLMSRevision);
                    sJobCode = serv.GetDocJobCode(sPart, iWebAppId);
                    sLMSDocContName = serv.GetDocName(sPart, iWebAppId);

                    foreach (FolderItem folderitem in m_MatCatFolders)
                    {
                        //Only look at the Mnnnnn part and not any subpart suffix for the folder
                        if (folderitem.name.StartsWith(sPart.Substring(0, 6)))
                        {
                            bFolderExists = true;
                            sFolderId = folderitem.id;
                            break;
                        }
                    }
                }
                else
                {
                    result.sOverallComment = sPart + " document container does not exist in the LMS";
                    return result;
                }

                if (bFolderExists)
                {
                    string nextUrl = "https://graph.microsoft.com/v1.0/drives/" + m_driveId +
                                      "/items/" + sFolderId + "/children";

                    // Keep looping while the API continues to feed us a nextLink address
                    if (!string.IsNullOrEmpty(sFolderId))
                    {
                        while (!string.IsNullOrEmpty(nextUrl))
                        {
                            FilePageResult pageResult = await ReadFileAsync2(m_AccessToken, nextUrl);

                            if (pageResult?.FileItems != null)
                            {
                                foreach (FileItem fileItem in pageResult.FileItems)
                                {
                                    if (fileItem.type.Equals("File") && fileItem.name.StartsWith(sPart + " "))
                                    {
                                        string sRevision = "";
                                        int iLastDot = fileItem.name.LastIndexOf('.');
                                        if (iLastDot != -1)
                                        {
                                            int iLastV = fileItem.name.LastIndexOf("v", iLastDot, StringComparison.OrdinalIgnoreCase);
                                            if (iLastV != -1)
                                            {
                                                //We don't expect more than 5 characters for the revision
                                                //This allows up to 99.99 or ZZ.99
                                                if (iLastDot - iLastV - 1 <= 5)
                                                {
                                                    sRevision = fileItem.name.Substring(iLastV + 1, iLastDot - iLastV - 1); //The 1 in each is to ignore the v
                                                    sRevision = sRevision.Trim();
                                                }
                                                else
                                                {
                                                    //The revision could be immedaitely after the part number
                                                    //eg M06176A v1 Part Name.dxf
                                                    int iFirstSpace = fileItem.name.IndexOf(sPart + " v");
                                                    if(iFirstSpace >= 0)
                                                    {
                                                        iFirstSpace += +sPart.Length + 2;
                                                        int iNextSpace = fileItem.name.Substring(iFirstSpace).IndexOf(" ");
                                                        string sVersion = "";
                                                        if (iNextSpace > 0)
                                                        {
                                                            sVersion = fileItem.name.Substring(iFirstSpace, iNextSpace);
                                                        }

                                                        sRevision = sVersion.Trim();
                                                    }

                                                }
                                            }
                                        }

                                        string sRevisionCompare = sRevision;
                                        if(sRevision.Contains("."))
                                        {
                                            sRevisionCompare = sRevision.Substring(0, sRevision.IndexOf("."));
                                        }

                                        bool bFoundRevPDF = false;

                                        if (sRevisionPDF.Equals(sRevisionCompare) && 
                                            Path.GetExtension(fileItem.name).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
                                        {
                                            bFoundRevPDF = true;
                                        }
                                        else if(sRevisionPDF.Equals(""))
                                            bFoundRevPDF = true;

                                        bool bFoundRevDXF = false;

                                        if (sRevisionDXF.Equals(sRevisionCompare) &&
                                            Path.GetExtension(fileItem.name).Equals(".DXF", StringComparison.OrdinalIgnoreCase))
                                        {
                                            bFoundRevDXF = true;
                                        }
                                        else if (sRevisionDXF.Equals(""))
                                            bFoundRevDXF = true;

                                        bool bFoundRevSTEP = false;

                                        if (sRevisionSTEP.Equals(sRevisionCompare) &&
                                            Path.GetExtension(fileItem.name).Equals(".STEP", StringComparison.OrdinalIgnoreCase))
                                        {
                                            bFoundRevSTEP = true;
                                        }
                                        else if (sRevisionSTEP.Equals(""))
                                            bFoundRevSTEP = true;

                                        FileItemExtra thisItem = new FileItemExtra
                                        {
                                            name = fileItem.name,
                                            type = fileItem.type,
                                            id = fileItem.id,
                                            extension = Path.GetExtension(fileItem.name),
                                            revision = sRevision,
                                            revfoundpdf = bFoundRevPDF,
                                            revfounddxf = bFoundRevDXF,
                                            revfoundstep = bFoundRevSTEP,
                                            donwloadurl = fileItem.downloadurl
                                        };

                                        fileItems.Add(thisItem);
                                    }
                                }
                            }

                            // Re-target the loop iterator condition directly to the returned nextLink string variable.
                            // It smoothly returns null on the very last page, cleanly dropping out of the while loop.
                            nextUrl = pageResult?.NextLink;
                        }
                    }

                    //If we get to here we have all the items in the folder in a list
                    if (fileItems.Count > 0)
                    {
                        //Find out what types of files we need to upload
                        fileReqs = GetDrawingUploadRequirements(sPart, iWebAppId);

                        if (!fileReqs.bResult)
                        {
                            result.sOverallComment = fileReqs.sErrorMsg;
                            return result;
                        }

                        //Loop 3 times for Pdf, Dxf and Step in that order
                        for (int i = 0; i < 3; i++) //Chnage this to top of range as 3 when finished testing
                        {
                            switch (i)
                            {
                                //PDF
                                case 0:
                                    if (fileReqs.bPDF)
                                    {
                                        List<FileItemExtra> fileItemsPdf = new List<FileItemExtra>();
                                        foreach (FileItemExtra pdffileitem in fileItems)
                                        {
                                            if (pdffileitem.extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase))
                                            {
                                                if (!pdffileitem.revision.Equals("") && pdffileitem.revfoundpdf)
                                                {
                                                    FileItemExtra thisItem = new FileItemExtra
                                                    {
                                                        name = pdffileitem.name,
                                                        type = pdffileitem.type,
                                                        id = pdffileitem.id,
                                                        extension = pdffileitem.extension,
                                                        revision = pdffileitem.revision,
                                                        donwloadurl = pdffileitem.donwloadurl
                                                    };
                                                    fileItemsPdf.Add(thisItem);
                                                }

                                            }
                                        }

                                        if (fileItemsPdf.Count > 0)
                                        {
                                            List<FileItemExtra> sortedfileItemsPdf = fileItemsPdf
                                            .OrderByDescending(ii => string.IsNullOrEmpty(ii.revision) ? 2 : char.IsLetter(ii.revision[0]) ? 0 : 1)
                                            .ThenByDescending(ii => ii.revision)
                                            .ToList();

                                            result.sPDFRevision = sortedfileItemsPdf[0].revision;
                                            allRevs.Add(result.sPDFRevision);
                                            fileToDownload.Add(sortedfileItemsPdf[0]);
                                        }
                                        else
                                        {
                                            if (sRevisionPDF.Equals(""))
                                                result.sPDFComment = "Cannot find any pdf files that have any revision";
                                            else
                                                result.sPDFComment = "Cannot find any pdf files that have a revision " + sRevisionPDF;
                                        }
                                    }
                                    else
                                    {
                                        if(!sRevisionPDF.Equals(""))
                                        {
                                            result.sPDFComment = "A pdf file is not marked as required for this part";
                                        }
                                    }
                                    break;
                                //DXF
                                case 1:
                                    if (fileReqs.bDXF)
                                    {
                                        List<FileItemExtra> fileItemsDxf = new List<FileItemExtra>();
                                        foreach (FileItemExtra dxffileitem in fileItems)
                                        {
                                            if (dxffileitem.extension.Equals(".dxf", StringComparison.OrdinalIgnoreCase))
                                            {
                                                if (!dxffileitem.revision.Equals("") && dxffileitem.revfounddxf)
                                                {
                                                    FileItemExtra thisItem = new FileItemExtra
                                                    {
                                                        name = dxffileitem.name,
                                                        type = dxffileitem.type,
                                                        id = dxffileitem.id,
                                                        extension = dxffileitem.extension,
                                                        revision = dxffileitem.revision,
                                                        donwloadurl = dxffileitem.donwloadurl
                                                    };
                                                    fileItemsDxf.Add(thisItem);
                                                }

                                            }
                                        }

                                        if (fileItemsDxf.Count > 0)
                                        {
                                            List<FileItemExtra> sortedfileItemsDxf = fileItemsDxf
                                            .OrderByDescending(ii => string.IsNullOrEmpty(ii.revision) ? 2 : char.IsLetter(ii.revision[0]) ? 0 : 1)
                                            .ThenByDescending(ii => ii.revision)
                                            .ToList();

                                            result.sDXFRevision = sortedfileItemsDxf[0].revision;
                                            allRevs.Add(result.sDXFRevision);
                                            fileToDownload.Add(sortedfileItemsDxf[0]);
                                        }
                                        else
                                        {
                                            if(sRevisionDXF.Equals(""))
                                                result.sDXFComment = "Cannot find any DXF files that have any revision";
                                            else
                                                result.sDXFComment = "Cannot find any DXF files that have a revision " + sRevisionDXF;
                                        }
                                    }
                                    else
                                    {
                                        if (!sRevisionDXF.Equals(""))
                                        {
                                            result.sDXFComment = "A dxf file is not marked as required for this part";
                                        }
                                    }
                                    break;
                                //STEP
                                case 2:
                                    if (fileReqs.bSTEP)
                                    {
                                        List<FileItemExtra> fileItemsStep = new List<FileItemExtra>();
                                        foreach (FileItemExtra stepfileitem in fileItems)
                                        {
                                            if (stepfileitem.extension.Equals(".step", StringComparison.OrdinalIgnoreCase))
                                            {
                                                if (!stepfileitem.revision.Equals("") && stepfileitem.revfoundstep)
                                                {
                                                    FileItemExtra thisItem = new FileItemExtra
                                                    {
                                                        name = stepfileitem.name,
                                                        type = stepfileitem.type,
                                                        id = stepfileitem.id,
                                                        extension = stepfileitem.extension,
                                                        revision = stepfileitem.revision,
                                                        donwloadurl = stepfileitem.donwloadurl
                                                    };
                                                    fileItemsStep.Add(thisItem);
                                                }

                                            }
                                        }

                                        if (fileItemsStep.Count > 0)
                                        {
                                            List<FileItemExtra> sortedfileItemsStep = fileItemsStep
                                            .OrderByDescending(ii => string.IsNullOrEmpty(ii.revision) ? 2 : char.IsLetter(ii.revision[0]) ? 0 : 1)
                                            .ThenByDescending(ii => ii.revision)
                                            .ToList();

                                            result.sSTEPRevision = sortedfileItemsStep[0].revision;
                                            allRevs.Add(result.sSTEPRevision);
                                            fileToDownload.Add(sortedfileItemsStep[0]);
                                        }
                                        else
                                        {
                                            if (sRevisionSTEP.Equals(""))
                                                result.sSTEPComment = "Cannot find any Step files that have any revision";
                                            else
                                                result.sSTEPComment = "Cannot find any Step files that have a revision " + sRevisionSTEP;
                                        }
                                    }
                                    else
                                    {
                                        if (!sRevisionSTEP.Equals(""))
                                        {
                                            result.sSTEPComment = "A step file is not marked as required for this part";
                                        }
                                    }
                                    break;
                            }
                        }

                        // Finds the maximum value using our custom rule
                        if(allRevs.Count >0)
                            sMaxRev = allRevs
                                    .OrderByDescending(ii => string.IsNullOrEmpty(ii) ? 2 : char.IsLetter(ii[0]) ? 0 : 1)
                                    .ThenByDescending(ii => ii)
                                    .FirstOrDefault();

                        if(sMaxRev.Contains("."))
                        {
                            sMaxRev = sMaxRev.Substring(0, sMaxRev.IndexOf("."));
                        }

                        if (allRevs.Count > 1)
                            s2ndMaxRev = allRevs
                                    .OrderByDescending(ii => string.IsNullOrEmpty(ii) ? 2 : char.IsLetter(ii[0]) ? 0 : 1)
                                    .ThenByDescending(ii => ii)
                                    .Skip(1)
                                    .FirstOrDefault();

                        if(!s2ndMaxRev.Equals(sMaxRev))
                        {
                            CustomStringComparer clsComp = new CustomStringComparer();
                            int iLMSCompare = clsComp.Compare(sLMSRevision, sMaxRev);

                            if (iLMSCompare > 0)
                            {
                                result.sLMSComment = "LMS document container revision is further advanced than the highest revision of the required files and therefore will not be modified";
                            }
                        }

                        result.sOverallRevision = sMaxRev;

                        /*                        string sMaxRevision = allRevs
                                                    .OrderByDescending(s => s, new CustomStringComparer())
                                                    .FirstOrDefault() ?? string.Empty;

                                                result.sOverallRevision = sMaxRevision;
                        */
//                        return result;
                    }
                }
                else
                {
                    result.sOverallComment = "Cannot find a folder that starts with " + sPart.Substring(0,6);
                    return result;

                }
                //This download url has to be read each time from the json response because it includes the current access token
                //The json has the fully formed url in the response so just grab it and hold it in the class
                if(iUpload == 1)
                {
                    if(fileToDownload.Count >0)
                    {
                        //Here the variable sMaxRev should hold the new revision of the LMS document container
                        if(!sLMSRevision.Equals(sMaxRev) && !sMaxRev.Equals(""))
                        {
                            string sLMSRevUpdate = serv.ReviseDocumentAndRemoveAttachments(sSessionId, sUserId, sFullname, 
                                                                                           sPart, sLMSDocContName,
                                                                                           sMaxRev, "", sFullname, "", sJobCode,
                                                                                           "Revising document to " + sMaxRev + " from auto part upload process",
                                                                                           "1", iWebAppId.ToString());

                            if(!sLMSRevUpdate.StartsWith("Success"))
                            {
                                result.sStatusComments = "Could not raise revision of part document container " + sPart + " to " + sMaxRev;
                                result.sStatus = "Failure";
                                return result;
                            }
                        }

                        for (int i = 0; i < fileToDownload.Count; i++)
                        {
                            string sDownloadUrl = fileToDownload[i].donwloadurl;
                            string sThisPath = @"C:\WebRoot\Regain\Uploads\" + fileToDownload[i].name;
                            await DownloadFileFromUrlAsync(sDownloadUrl, sThisPath);
                            string sAttachResult = serv.AttachWCDoc(sSessionId, sUserId, sFullname, sPart, fileToDownload[i].name, sThisPath, "true", "Initiated from auto upload code", iWebAppId.ToString());
                            if (!sAttachResult.StartsWith("Success"))
                            {
                                result.sStatusComments += "Could not attach file " + fileToDownload[i].name + "\n";
                                result.sStatus = "Failure";
                                return result;
                            }
                        }

                        result.sStatusComments = "Uploaded";
                        result.sStatus = "Uploaded";
                    }
                }

                return result;
            }
            catch (System.Exception e)
            {
                System.Diagnostics.Debug.WriteLine($"Error occurred: {e.StackTrace}");
                result.sOverallComment = e.Message;
                return result;
            }
        }

        public FileRequirement GetDrawingUploadRequirements(string sPartNo, int iWebAppId)
        {
            Service1 serv = new Service1();
            Service1.rtnInt rtnPDF = new Service1.rtnInt();
            Service1.rtnInt rtnDXF = new Service1.rtnInt();
            Service1.rtnInt rtnSTEP = new Service1.rtnInt();
            FileRequirement rtnReq = new FileRequirement();

            rtnPDF = serv.GetPartBooleanAttribute(sPartNo, "PDF", iWebAppId);

            if (!rtnPDF.bReturnValue)
            {
                rtnReq.bResult = false;
                rtnReq.sErrorMsg = rtnPDF.sErrMessage;
                return rtnReq;
            }

            rtnReq.bPDF = rtnPDF.iReturnValue == 1 ? true : false;

            rtnDXF = serv.GetPartBooleanAttribute(sPartNo, "DXF", iWebAppId);

            if (!rtnDXF.bReturnValue)
            {
                rtnReq.bResult = false;
                rtnReq.sErrorMsg = rtnDXF.sErrMessage;
                return rtnReq;
            }

            rtnReq.bDXF = rtnDXF.iReturnValue == 1 ? true : false;

            rtnSTEP = serv.GetPartBooleanAttribute(sPartNo, "STEP", iWebAppId);

            if (!rtnSTEP.bReturnValue)
            {
                rtnReq.bResult = false;
                rtnReq.sErrorMsg = rtnSTEP.sErrMessage;
                return rtnReq;
            }

            rtnReq.bSTEP = rtnSTEP.iReturnValue == 1 ? true : false;

            rtnReq.bResult = true;
            rtnReq.sErrorMsg = "";

            return rtnReq;
        }

        public class CustomStringComparer : IComparer<string>
        {
            //Returns < 0 means x < y
            //Returns = 0 means x = y
            //Retunrs > 0 means x > y
            public int Compare(string x, string y)
            {
                // Handle nulls and exact matches
                if (x == y) return 0;
                if (x == null) return -1;
                if (y == null) return 1;

                // 1. Get the priority bucket for both strings
                int priorityX = GetTypePriority(x);
                int priorityY = GetTypePriority(y);

                // 2. If they belong to different categories, the higher priority wins
                if (priorityX != priorityY)
                {
                    return priorityX.CompareTo(priorityY);
                }

                // 3. If they are in the same category, fall back to standard ordinal sorting
                return string.Compare(x, y, StringComparison.Ordinal);
            }

            private int GetTypePriority(string str)
            {
                if (string.IsNullOrEmpty(str)) return 0; // Lowest priority

                // Check if the string is numeric (all characters are digits)
                if (str.All(char.IsDigit)) return 2;     // Highest priority

                return 1;                                // Alphas/Others middle priority
            }

        }
    }
}