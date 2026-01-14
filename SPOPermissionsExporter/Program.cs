using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System.Text;

namespace SPOPermissionsExporter
{
    class Program
    {
        private static GraphServiceClient? _graphClient;

        static async Task Main(string[] args)
        {
            try
            {
                Console.WriteLine("SharePoint Permissions Exporter");
                Console.WriteLine("================================\n");

                // Initialize Graph client
                await InitializeGraphClient();

                // Step 1: List all SharePoint sites
                var sites = await ListSharePointSites();
                if (sites == null || sites.Count == 0)
                {
                    Console.WriteLine("No SharePoint sites found.");
                    return;
                }

                // Step 2: Select a site
                var selectedSite = SelectSite(sites);
                if (selectedSite == null)
                {
                    Console.WriteLine("No site selected.");
                    return;
                }

                // Step 3: List document libraries for selected site
                var libraries = await ListDocumentLibraries(selectedSite.Id!);
                if (libraries == null || libraries.Count == 0)
                {
                    Console.WriteLine("No document libraries found.");
                    return;
                }

                // Step 4: Select a document library
                var selectedLibrary = SelectDocumentLibrary(libraries);
                if (selectedLibrary == null)
                {
                    Console.WriteLine("No library selected.");
                    return;
                }

                // Step 5: List all folders and files with permissions
                Console.WriteLine("\nRetrieving items and permissions...");
                var items = await GetAllItemsWithPermissions(selectedSite.Id!, selectedLibrary.Id!);

                // Step 6: Export to CSV
                ExportToCSV(items, selectedSite.DisplayName!, selectedLibrary.DisplayName!);

                Console.WriteLine("\nPress any key to exit...");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nError: {ex.Message}");
                Console.WriteLine($"Details: {ex}");
            }
        }

        static async Task InitializeGraphClient()
        {
            Console.WriteLine("Initializing Microsoft Graph client...");
            Console.WriteLine("Please enter your credentials when prompted.\n");

            // Using InteractiveBrowserCredential for user authentication
            var credential = new InteractiveBrowserCredential(new InteractiveBrowserCredentialOptions
            {
                TenantId = "common", // Use "common" for multi-tenant or specify your tenant ID
                ClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e", // Microsoft Graph Explorer client ID (public)
                RedirectUri = new Uri("http://localhost")
            });

            _graphClient = new GraphServiceClient(credential, new[]
            {
                "Sites.Read.All",
                "Files.Read.All",
                "User.Read"
            });

            // Test the connection
            var user = await _graphClient.Me.GetAsync();
            Console.WriteLine($"Authenticated as: {user?.DisplayName} ({user?.UserPrincipalName})\n");
        }

        static async Task<List<Site>> ListSharePointSites()
        {
            Console.WriteLine("Retrieving SharePoint sites...\n");
            var sites = new List<Site>();

            try
            {
                var result = await _graphClient!.Sites.GetAsync((requestConfiguration) =>
                {
                    requestConfiguration.QueryParameters.Search = "*";
                    requestConfiguration.QueryParameters.Select = new[] { "id", "displayName", "name", "webUrl" };
                    requestConfiguration.QueryParameters.Top = 50;
                });

                if (result?.Value != null)
                {
                    sites.AddRange(result.Value);
                }

                Console.WriteLine($"Found {sites.Count} SharePoint site(s):\n");
                for (int i = 0; i < sites.Count; i++)
                {
                    Console.WriteLine($"{i + 1}. {sites[i].DisplayName} ({sites[i].WebUrl})");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving sites: {ex.Message}");
            }

            return sites;
        }

        static Site? SelectSite(List<Site> sites)
        {
            Console.Write("\nEnter the number of the site you want to explore: ");
            if (int.TryParse(Console.ReadLine(), out int selection) && selection > 0 && selection <= sites.Count)
            {
                var selected = sites[selection - 1];
                Console.WriteLine($"\nSelected: {selected.DisplayName}");
                return selected;
            }

            Console.WriteLine("Invalid selection.");
            return null;
        }

        static async Task<List<DriveItemInfo>> ListDocumentLibraries(string siteId)
        {
            Console.WriteLine("\nRetrieving document libraries...\n");
            var libraries = new List<DriveItemInfo>();

            try
            {
                var drives = await _graphClient!.Sites[siteId].Drives.GetAsync((requestConfiguration) =>
                {
                    requestConfiguration.QueryParameters.Select = new[] { "id", "name", "description", "webUrl" };
                });

                if (drives?.Value != null)
                {
                    foreach (var drive in drives.Value)
                    {
                        libraries.Add(new DriveItemInfo
                        {
                            Id = drive.Id,
                            DisplayName = drive.Name,
                            WebUrl = drive.WebUrl
                        });
                    }
                }

                Console.WriteLine($"Found {libraries.Count} document library(ies):\n");
                for (int i = 0; i < libraries.Count; i++)
                {
                    Console.WriteLine($"{i + 1}. {libraries[i].DisplayName}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving libraries: {ex.Message}");
            }

            return libraries;
        }

        static DriveItemInfo? SelectDocumentLibrary(List<DriveItemInfo> libraries)
        {
            Console.Write("\nEnter the number of the document library you want to explore: ");
            if (int.TryParse(Console.ReadLine(), out int selection) && selection > 0 && selection <= libraries.Count)
            {
                var selected = libraries[selection - 1];
                Console.WriteLine($"\nSelected: {selected.DisplayName}");
                return selected;
            }

            Console.WriteLine("Invalid selection.");
            return null;
        }

        static async Task<List<ItemPermissionInfo>> GetAllItemsWithPermissions(string siteId, string driveId)
        {
            var itemsWithPermissions = new List<ItemPermissionInfo>();

            try
            {
                // Get root folder items
                var rootItems = await _graphClient!.Drives[driveId].Items["root"].Children.GetAsync((requestConfiguration) =>
                {
                    requestConfiguration.QueryParameters.Select = new[] { "id", "name", "folder", "file", "webUrl", "createdBy", "lastModifiedBy" };
                });

                if (rootItems?.Value != null)
                {
                    await ProcessItems(driveId, rootItems.Value.ToList(), "", itemsWithPermissions);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving items: {ex.Message}");
            }

            return itemsWithPermissions;
        }

        static async Task ProcessItems(string driveId, List<DriveItem> items, string parentPath, List<ItemPermissionInfo> result)
        {
            foreach (var item in items)
            {
                try
                {
                    var itemPath = string.IsNullOrEmpty(parentPath) ? (item.Name ?? "Unknown") : $"{parentPath}/{item.Name ?? "Unknown"}";
                    var itemType = item.Folder != null ? "Folder" : "File";

                    Console.WriteLine($"Processing: {itemPath}");

                    // Get permissions for this item
                    var permissions = await GetItemPermissions(driveId, item.Id!);

                    var itemInfo = new ItemPermissionInfo
                    {
                        ItemName = item.Name ?? "Unknown",
                        ItemPath = itemPath,
                        ItemType = itemType,
                        WebUrl = item.WebUrl ?? "",
                        CreatedBy = item.CreatedBy?.User?.DisplayName ?? "Unknown",
                        LastModifiedBy = item.LastModifiedBy?.User?.DisplayName ?? "Unknown",
                        Permissions = permissions
                    };

                    result.Add(itemInfo);

                    // If it's a folder, process children recursively
                    if (item.Folder != null && item.Id != null)
                    {
                        var children = await _graphClient!.Drives[driveId].Items[item.Id].Children.GetAsync((requestConfiguration) =>
                        {
                            requestConfiguration.QueryParameters.Select = new[] { "id", "name", "folder", "file", "webUrl", "createdBy", "lastModifiedBy" };
                        });

                        if (children?.Value != null && children.Value.Any())
                        {
                            await ProcessItems(driveId, children.Value.ToList(), itemPath, result);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing item {item.Name ?? "Unknown"}: {ex.Message}");
                }
            }
        }

        static async Task<List<PermissionInfo>> GetItemPermissions(string driveId, string itemId)
        {
            var permissionsList = new List<PermissionInfo>();

            try
            {
                var permissions = await _graphClient!.Drives[driveId].Items[itemId].Permissions.GetAsync();

                if (permissions?.Value != null)
                {
                    foreach (var perm in permissions.Value)
                    {
                        var permInfo = new PermissionInfo
                        {
                            Id = perm.Id ?? "",
                            Roles = perm.Roles != null ? string.Join(", ", perm.Roles) : "None",
                            GrantedTo = perm.GrantedToV2?.User?.DisplayName ?? 
                                       perm.GrantedToIdentitiesV2?.FirstOrDefault()?.User?.DisplayName ?? 
                                       "Unknown",
                            PermissionType = perm.Link != null ? "Link" : 
                                           perm.GrantedToV2 != null ? "User/Group" : 
                                           "Inherited"
                        };

                        permissionsList.Add(permInfo);
                    }
                }
            }
            catch (Exception ex)
            {
                // Some items may not have accessible permissions
                permissionsList.Add(new PermissionInfo
                {
                    Id = "N/A",
                    Roles = "Error",
                    GrantedTo = $"Error: {ex.Message}",
                    PermissionType = "Error"
                });
            }

            return permissionsList;
        }

        static void ExportToCSV(List<ItemPermissionInfo> items, string siteName, string libraryName)
        {
            try
            {
                var fileName = $"Permissions_{siteName}_{libraryName}_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
                fileName = string.Join("_", fileName.Split(Path.GetInvalidFileNameChars()));

                using (var writer = new StreamWriter(fileName, false, Encoding.UTF8))
                {
                    // Write header
                    writer.WriteLine("Site,Library,Item Name,Item Path,Item Type,Web URL,Created By,Last Modified By,Permission ID,Permission Type,Roles,Granted To");

                    // Write data
                    foreach (var item in items)
                    {
                        if (item.Permissions.Count == 0)
                        {
                            // Item with no permissions
                            writer.WriteLine($"\"{siteName}\",\"{libraryName}\",\"{item.ItemName}\",\"{item.ItemPath}\",\"{item.ItemType}\",\"{item.WebUrl}\",\"{item.CreatedBy}\",\"{item.LastModifiedBy}\",\"\",\"\",\"\",\"\"");
                        }
                        else
                        {
                            foreach (var perm in item.Permissions)
                            {
                                writer.WriteLine($"\"{siteName}\",\"{libraryName}\",\"{item.ItemName}\",\"{item.ItemPath}\",\"{item.ItemType}\",\"{item.WebUrl}\",\"{item.CreatedBy}\",\"{item.LastModifiedBy}\",\"{perm.Id}\",\"{perm.PermissionType}\",\"{perm.Roles}\",\"{perm.GrantedTo}\"");
                            }
                        }
                    }
                }

                Console.WriteLine($"\nPermissions exported successfully to: {fileName}");
                Console.WriteLine($"Total items processed: {items.Count}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error exporting to CSV: {ex.Message}");
            }
        }
    }

    // Helper classes for data management
    class DriveItemInfo
    {
        public string? Id { get; set; }
        public string? DisplayName { get; set; }
        public string? WebUrl { get; set; }
    }

    class ItemPermissionInfo
    {
        public string ItemName { get; set; } = "";
        public string ItemPath { get; set; } = "";
        public string ItemType { get; set; } = "";
        public string WebUrl { get; set; } = "";
        public string CreatedBy { get; set; } = "";
        public string LastModifiedBy { get; set; } = "";
        public List<PermissionInfo> Permissions { get; set; } = new List<PermissionInfo>();
    }

    class PermissionInfo
    {
        public string Id { get; set; } = "";
        public string PermissionType { get; set; } = "";
        public string Roles { get; set; } = "";
        public string GrantedTo { get; set; } = "";
    }
}
