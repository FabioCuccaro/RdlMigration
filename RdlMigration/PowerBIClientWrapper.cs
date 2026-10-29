// Copyright (c) 2019 Microsoft Corporation. All Rights Reserved.
// Licensed under the MIT License (MIT)

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.PowerBI.Api.V2;
using Microsoft.PowerBI.Api.V2.Models;
using Microsoft.Rest;
using Newtonsoft.Json.Linq;

using static RdlMigration.ElementNameConstants;

namespace RdlMigration
{
    /// <summary>
    /// The class of accessing the Power BI Api.
    /// </summary>
    public sealed class PowerBIClientWrapper
    {
        public string ClientId { get; set; }

        // the selected workspace. if null then means "My Workspace" is added.
        private Group workspace;

        private PowerBIClient client;
        private IImportsOperations importClient;
        private IReportsOperations reportsClient;
        private IGroupsOperations groupsClient;

        // token kept for the REST calls the SDK does not cover (folders, import into a folder)
        private string accessToken;

        // folder path relative to the workspace root -> folder id
        private Dictionary<string, string> folderCache = new Dictionary<string, string>();

        // folder id (empty for the root) -> names of the reports already in that folder
        private Dictionary<string, HashSet<string>> folderReports = new Dictionary<string, HashSet<string>>();

        public PowerBIClientWrapper(string workspaceName, string clientId)
        {
            this.ClientId = clientId;
            InitializeClients();
            GetWorkspaces(workspaceName);
        }

        public PowerBIClientWrapper(string workspaceName, string clientId,
            IImportsOperations importClient, IReportsOperations reportsClient,
            IGroupsOperations groupsClient)
        {
            this.ClientId = clientId;
            this.importClient = importClient;
            this.reportsClient = reportsClient;
            this.groupsClient = groupsClient;
            GetWorkspaces(workspaceName);
        }

        /// <summary>
        /// Pops up the log-in window for user to log in with their Microsoft account.
        /// </summary>
        public AuthenticationResult DoInteractiveSignIn()
        {
            AuthenticationContext authenticationContext = new AuthenticationContext(
                PowerBIWrapperConstants.AuthorityUri, false, new TokenCache());
            AuthenticationResult userAuthnResult = authenticationContext.AcquireTokenAsync(
                PowerBIWrapperConstants.ResourceUrl,
                this.ClientId,
                new Uri(PowerBIWrapperConstants.RedirectUrl),
                new PlatformParameters(PromptBehavior.Auto)).Result;

            return userAuthnResult;
        }

        /// <summary>
        /// Upload the report file stream to the workspace, optionally into a workspace folder.
        /// </summary>
        /// <param name="fileName">The file name displayed after uploaded, including the .rdl extension.</param>
        /// <param name="file">The file stream itself.</param>
        /// <param name="subfolderObjectId">Optional id of the target workspace folder. Null means the workspace root.</param>
        /// <returns>true if success, false if conflict.</returns>
        public bool UploadRDL(string fileName, Stream file, string subfolderObjectId = null)
        {
            if (!ExistReport(fileName, subfolderObjectId))
            {
                if (workspace == null)
                {
                    // "My Workspace" does not support folders
                    importClient.PostImportWithFile(file, fileName, ImportConflictHandlerMode.Abort);
                }
                else if (string.IsNullOrEmpty(subfolderObjectId))
                {
                    importClient.PostImportWithFileInGroup(workspace.Id, file, fileName, ImportConflictHandlerMode.Abort);
                }
                else
                {
                    // the SDK does not expose subfolderObjectId, so call the REST API directly
                    UploadRDLToFolder(fileName, file, subfolderObjectId);
                }
                GetFolderReports(subfolderObjectId).Add(Path.GetFileNameWithoutExtension(fileName));
                return true;
            }

            return false;
        }

        /// <summary>
        /// Creates the folder hierarchy in the workspace if needed and returns the id of the last folder.
        /// For example "Italy/Sales" creates "Italy", then "Sales" inside "Italy".
        /// </summary>
        /// <param name="relativePath">The folder path relative to the workspace root. Null means the root.</param>
        /// <returns>The id of the leaf folder, or null for the root or "My Workspace".</returns>
        public string EnsureFolderExists(string relativePath)
        {
            if (string.IsNullOrWhiteSpace(relativePath) || workspace == null)
            {
                return null;
            }

            if (folderCache.TryGetValue(relativePath, out string cachedId))
            {
                return cachedId;
            }

            string[] segments = relativePath.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            string parentFolderId = null;
            string cumulativePath = string.Empty;

            using (var httpClient = CreateHttpClient())
            {
                foreach (string segment in segments)
                {
                    cumulativePath = string.IsNullOrEmpty(cumulativePath) ? segment : cumulativePath + "/" + segment;

                    if (folderCache.TryGetValue(cumulativePath, out string existingId))
                    {
                        parentFolderId = existingId;
                        continue;
                    }

                    string url = $"{PowerBIWrapperConstants.FabricApiUri}/v1/workspaces/{workspace.Id}/folders";
                    var bodyObj = new JObject { ["displayName"] = segment };
                    if (parentFolderId != null)
                    {
                        bodyObj["parentFolderId"] = parentFolderId;
                    }

                    using (var content = new StringContent(bodyObj.ToString(), Encoding.UTF8, "application/json"))
                    using (var response = httpClient.PostAsync(url, content).Result)
                    {
                        string responseBody = response.Content.ReadAsStringAsync().Result;

                        if (response.IsSuccessStatusCode)
                        {
                            parentFolderId = JObject.Parse(responseBody)["id"]?.ToString();
                            Console.WriteLine($"FOLDER CREATED : '{cumulativePath}' ID: {parentFolderId}");
                        }
                        else if (response.StatusCode == System.Net.HttpStatusCode.Conflict)
                        {
                            parentFolderId = FindExistingFolderId(httpClient, segment, parentFolderId);
                            Console.WriteLine($"FOLDER EXISTS : '{cumulativePath}' ID: {parentFolderId}");
                        }
                        else
                        {
                            throw new Exception(
                                $"Failed to create folder '{cumulativePath}' at {url}: " +
                                $"{(int)response.StatusCode} {response.StatusCode} - {responseBody}");
                        }
                    }

                    if (parentFolderId == null)
                    {
                        throw new Exception($"Could not resolve the id of folder '{cumulativePath}' in the workspace");
                    }

                    folderCache[cumulativePath] = parentFolderId;
                }
            }

            return parentFolderId;
        }

        /// <summary>
        /// if the workspace root contains a specific report.
        /// </summary>
        /// <param name="reportName">the name of the report.</param>
        /// <returns>true if the report exists.</returns>
        public bool ExistReport(string reportName)
        {
            return ExistReport(reportName, null);
        }

        /// <summary>
        /// if a folder of the workspace contains a specific report.
        /// </summary>
        /// <param name="reportName">the name of the report.</param>
        /// <param name="folderId">the id of the workspace folder. Null means the workspace root.</param>
        /// <returns>true if the report exists.</returns>
        public bool ExistReport(string reportName, string folderId)
        {
            return GetFolderReports(folderId).Contains(Path.GetFileNameWithoutExtension(reportName));
        }

        private void UploadRDLToFolder(string fileName, Stream file, string subfolderObjectId)
        {
            string url = $"{PowerBIWrapperConstants.PowerBiApiUri}/v1.0/myorg/groups/{workspace.Id}/imports" +
                         $"?datasetDisplayName={Uri.EscapeDataString(fileName)}" +
                         $"&nameConflict=Abort" +
                         $"&subfolderObjectId={subfolderObjectId}";

            using (var httpClient = CreateHttpClient())
            using (var multipart = new MultipartFormDataContent())
            {
                var streamContent = new StreamContent(file);
                streamContent.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                multipart.Add(streamContent, "file", fileName);

                using (var response = httpClient.PostAsync(url, multipart).Result)
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        string body = response.Content.ReadAsStringAsync().Result;
                        throw new Exception(
                            $"Failed to upload '{fileName}' to folder {subfolderObjectId}: " +
                            $"{(int)response.StatusCode} {response.StatusCode} - {body}");
                    }
                }
            }
        }

        /// <summary>
        /// Looks up an existing folder by name and parent, following the continuation pages of the list.
        /// </summary>
        private string FindExistingFolderId(HttpClient httpClient, string folderName, string parentFolderId)
        {
            string url = $"{PowerBIWrapperConstants.FabricApiUri}/v1/workspaces/{workspace.Id}/folders";
            if (parentFolderId != null)
            {
                url += $"?rootFolderId={parentFolderId}&recursive=false";
            }
            else
            {
                url += "?recursive=false";
            }

            foreach (var folder in GetAllPages(httpClient, url))
            {
                bool nameMatch = string.Equals(folder["displayName"]?.ToString(), folderName, StringComparison.OrdinalIgnoreCase);
                var parentToken = folder["parentFolderId"];
                bool parentMatch = parentFolderId == null
                    ? parentToken == null || parentToken.Type == JTokenType.Null
                    : parentToken?.ToString() == parentFolderId;

                if (nameMatch && parentMatch)
                {
                    return folder["id"]?.ToString();
                }
            }

            return null;
        }

        /// <summary>
        /// Returns the names of the reports already in a folder. The root uses the report list read at start-up,
        /// which covers the whole workspace; folders are read on first use from the Fabric items API.
        /// </summary>
        private HashSet<string> GetFolderReports(string folderId)
        {
            string key = folderId ?? string.Empty;
            if (!folderReports.TryGetValue(key, out HashSet<string> reports))
            {
                reports = new HashSet<string>();
                if (folderId != null && workspace != null && accessToken != null)
                {
                    string url = $"{PowerBIWrapperConstants.FabricApiUri}/v1/workspaces/{workspace.Id}/items" +
                                 $"?type=PaginatedReport&recursive=false&rootFolderId={folderId}";
                    using (var httpClient = CreateHttpClient())
                    {
                        foreach (var item in GetAllPages(httpClient, url))
                        {
                            reports.Add(item["displayName"]?.ToString());
                        }
                    }
                }
                folderReports[key] = reports;
            }

            return reports;
        }

        private static IEnumerable<JToken> GetAllPages(HttpClient httpClient, string url)
        {
            while (!string.IsNullOrEmpty(url))
            {
                JObject page;
                using (var response = httpClient.GetAsync(url).Result)
                {
                    string body = response.Content.ReadAsStringAsync().Result;
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new Exception($"Request to {url} failed: {(int)response.StatusCode} {response.StatusCode} - {body}");
                    }
                    page = JObject.Parse(body);
                }

                if (page["value"] is JArray values)
                {
                    foreach (var value in values)
                    {
                        yield return value;
                    }
                }

                url = page["continuationUri"]?.Type == JTokenType.String ? page["continuationUri"].ToString() : null;
            }
        }

        private HttpClient CreateHttpClient()
        {
            var httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            return httpClient;
        }

        private void InitializeClients()
        {
            AuthenticationResult result = DoInteractiveSignIn();
            accessToken = result.AccessToken;
            client = new PowerBIClient(
                new Uri(PowerBIWrapperConstants.PowerBiApiUri),
                new TokenCredentials(result.AccessToken));

            importClient = new ImportsOperations(client);
            reportsClient = new ReportsOperations(client);
            groupsClient = new GroupsOperations(client);
        }

        private Group GetWorkspaces(string workspaceName)
        {
            if (workspaceName == PowerBIWrapperConstants.MyWorkspace)
            {
                workspace = null;
                var reportNames = reportsClient.GetReports().Value.Select(report => report.Name);
                folderReports[string.Empty] = new HashSet<string>(reportNames);
                return null;
            }
            var workspaces = groupsClient.GetGroups().Value;
            var groups = workspaces.Where(g => (g.Name == workspaceName));
            if (groups.Count() == 1)
            {
                workspace = groups.First();
                if (workspace.IsOnDedicatedCapacity == false)
                {
                    throw new Exception(
                        $"WORKSPACE {workspaceName} IS NOT A PREMIUM WORKSPACE. Only premium workspaces can upload reports");
                }
                var reportNames = reportsClient.GetReportsInGroup(workspace.Id).Value.Select(report => report.Name);
                folderReports[string.Empty] = new HashSet<string>(reportNames);
                return workspace;
            }
            else if (groups.Count() == 0)
            {
                throw new Exception(
                    $"WORKSPACE {workspaceName} NOT FOUND. Please make sure it is a valid workspace");
            }
            else
            {
                throw new Exception(
                    $"MULTIPLE PREMIUM WORKSPACE {workspaceName} FOUND. This should not happen, make sure you have valid workspaces");
            }
        }
    }
}