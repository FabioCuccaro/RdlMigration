// Copyright (c) 2019 Microsoft Corporation. All Rights Reserved.
// Licensed under the MIT License (MIT)

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.PowerBI.Api.V2;
using Microsoft.PowerBI.Api.V2.Models;
using Microsoft.Rest;
using System.Net.Http;                    // ← AGGIUNTO
using System.Net.Http.Headers;            // ← AGGIUNTO
using System.Text;                        // ← AGGIUNTO
using Newtonsoft.Json.Linq;               // ← AGGIUNTO

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

        private HashSet<string> workspaceReports = new HashSet<string>();

        private PowerBIClient client;
        private IImportsOperations importClient;
        private IReportsOperations reportsClient;
        private IGroupsOperations groupsClient;

        private string accessToken;                                          // ← AGGIUNTO
        private Dictionary<string, string> folderCache = new Dictionary<string, string>(); // ← AGGIUNTO

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

        // ═══════════════════════════════════════════════════════════════
        // ↓↓↓ METODO MODIFICATO: aggiunto parametro subfolderObjectId ↓↓↓
        // ═══════════════════════════════════════════════════════════════

        /// <summary>
        /// Upload the report file stream to workspace, optionally into a folder.
        /// </summary>
        /// <param name="fileName">The file name displayed after uploaded.</param>
        /// <param name="file">The file stream itself.</param>
        /// <param name="subfolderObjectId">Optional folder UUID. Null = root.</param>
        /// <returns>true if success, false if conflict.</returns>
        public bool UploadRDL(string fileName, Stream file, string subfolderObjectId = null)
        {
            if (!ExistReport(fileName))
            {
                if (workspace == null)
                {
                    // My Workspace - nessun supporto cartelle
                    importClient.PostImportWithFile(file, fileName, ImportConflictHandlerMode.Abort);
                }
                else if (string.IsNullOrEmpty(subfolderObjectId))
                {
                    // Workspace specifico, root
                    importClient.PostImportWithFileInGroup(
                        workspace.Id, file, fileName, ImportConflictHandlerMode.Abort);
                }
                else
                {
                    // Workspace specifico CON cartella → REST API diretta
                    UploadRDLToFolder(fileName, file, subfolderObjectId);
                }
                workspaceReports.Add(fileName);
                return true;
            }

            return false;
        }

        // ═══════════════════════════════════════════════════════
        // ↓↓↓ METODI NUOVI: upload in cartella + gestione cartelle ↓↓↓
        // ═══════════════════════════════════════════════════════

        /// <summary>
        /// Upload via REST API diretta per passare subfolderObjectId.
        /// </summary>
        private void UploadRDLToFolder(string fileName, Stream file, string subfolderObjectId)
        {
            string url = $"https://api.powerbi.com/v1.0/myorg/groups/{workspace.Id}/imports" +
                         $"?datasetDisplayName={Uri.EscapeDataString(fileName)}" +
                         $"&nameConflict=Abort" +
                         $"&subfolderObjectId={subfolderObjectId}";

            using (var httpClient = new HttpClient())
            {
                httpClient.DefaultRequestHeaders.Authorization =
                    new AuthenticationHeaderValue("Bearer", accessToken);

                var multipart = new MultipartFormDataContent();
                var streamContent = new StreamContent(file);
                streamContent.Headers.ContentType =
                    new MediaTypeHeaderValue("application/octet-stream");
                multipart.Add(streamContent, "file", fileName);

                var response = httpClient.PostAsync(url, multipart).Result;
                if (!response.IsSuccessStatusCode)
                {
                    string body = response.Content.ReadAsStringAsync().Result;
                    throw new Exception(
                        $"Failed to upload '{fileName}' to folder {subfolderObjectId}: " +
                        $"{response.StatusCode} - {body}");
                }
            }
        }

        /// <summary>
        /// Crea la gerarchia di cartelle nel workspace PBI e ritorna l'ID dell'ultima.
        /// Es: relativePath = "Italy/Sales" → crea "Italy", poi "Sales" dentro "Italy".
        /// </summary>
        /// <param name="relativePath">Path relativo (es. "Italy/Sales"). Null = root.</param>
        /// <returns>L'UUID della cartella foglia, o null se root.</returns>
        public string EnsureFolderExists(string relativePath)
        {
            if (string.IsNullOrWhiteSpace(relativePath) || workspace == null)
                return null;

            if (folderCache.ContainsKey(relativePath))
                return folderCache[relativePath];

            string[] segments = relativePath.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            string parentFolderId = null;
            string cumulativePath = "";

            using (var httpClient = new HttpClient())
            {
                httpClient.DefaultRequestHeaders.Authorization =
                    new AuthenticationHeaderValue("Bearer", accessToken);

                foreach (string segment in segments)
                {
                    cumulativePath = string.IsNullOrEmpty(cumulativePath)
                        ? segment
                        : cumulativePath + "/" + segment;

                    if (folderCache.ContainsKey(cumulativePath))
                    {
                        parentFolderId = folderCache[cumulativePath];
                        continue;
                    }

                    // Crea la cartella
                    string url = $"https://api.fabric.microsoft.com/v1/workspaces/{workspace.Id}/folders";
                    var bodyObj = new JObject { ["displayName"] = segment };
                    if (parentFolderId != null)
                        bodyObj["parentObjectId"] = parentFolderId;

                    var content = new StringContent(bodyObj.ToString(), Encoding.UTF8, "application/json");
                    var response = httpClient.PostAsync(url, content).Result;
                    string responseBody = response.Content.ReadAsStringAsync().Result;


                    if (response.IsSuccessStatusCode)
                    {
                        var folderObj = JObject.Parse(responseBody);
                        parentFolderId = folderObj["id"]?.ToString();
                        Console.WriteLine($"FOLDER CREATED: '{cumulativePath}' → ID: {parentFolderId}");
                    }
                    else if (response.StatusCode == System.Net.HttpStatusCode.Conflict)
                    {
                        Console.WriteLine($"FOLDER EXISTS: '{cumulativePath}' → recovering ID...");
                        parentFolderId = FindExistingFolderId(httpClient, segment, parentFolderId);
                        Console.WriteLine($"FOLDER FOUND: '{cumulativePath}' → ID: {parentFolderId}");
                    }
                    else
                    {
                        // ═══ LOGGING DETTAGLIATO ═══
                        Console.WriteLine($"═══════════════════════════════════════════════");
                        Console.WriteLine($"ERROR creating folder '{cumulativePath}'");
                        Console.WriteLine($"  HTTP Status : {(int)response.StatusCode} {response.StatusCode}");
                        Console.WriteLine($"  Reason      : {response.ReasonPhrase}");
                        Console.WriteLine($"  URL         : https://api.fabric.microsoft.com/v1.0/myorg/groups/{workspace.Id}/folders");
                        Console.WriteLine($"  Request Body: {bodyObj.ToString()}");
                        Console.WriteLine($"  Response    : {responseBody}");
                        Console.WriteLine($"═══════════════════════════════════════════════");
                        return null;
                    }


                    if (parentFolderId != null)
                        folderCache[cumulativePath] = parentFolderId;
                }
            }

            return parentFolderId;
        }

        /// <summary>
        /// Cerca una cartella esistente per nome e parent.
        /// </summary>
        private string FindExistingFolderId(HttpClient httpClient, string folderName, string parentFolderId)
        {
            string listUrl = $"https://api.fabric.microsoft.com/v1/workspaces/{workspace.Id}/folders";
            var listResponse = httpClient.GetAsync(listUrl).Result;
            string listBody = listResponse.Content.ReadAsStringAsync().Result;
            var parsed = JObject.Parse(listBody);
            var folders = parsed["value"] as JArray;

            if (folders != null)
            {
                foreach (var f in folders)
                {
                    bool nameMatch = f["displayName"]?.ToString() == folderName;
                    bool parentMatch = parentFolderId == null
                        ? f["parentObjectId"] == null || f["parentObjectId"].Type == JTokenType.Null
                        : f["parentObjectId"]?.ToString() == parentFolderId;

                    if (nameMatch && parentMatch)
                        return f["id"]?.ToString();
                }
            }
            return null;
        }

        // ═══════════════════════════════════════════════════════
        // ↓↓↓ FINE METODI NUOVI ↓↓↓
        // ═══════════════════════════════════════════════════════

        /// <summary>
        /// if the default workspace contains a specific report.
        /// </summary>
        public bool ExistReport(string reportName)
        {
            return workspaceReports.Contains(Path.GetFileNameWithoutExtension(reportName));
        }

        private void InitializeClients()
        {
            AuthenticationResult result = DoInteractiveSignIn();
            accessToken = result.AccessToken;  // ← AGGIUNTO: salva il token per le chiamate REST dirette
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
                workspaceReports = new HashSet<string>(reportNames);
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
                workspaceReports = new HashSet<string>(reportNames);
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