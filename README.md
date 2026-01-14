# SPO-GraphPermissions

A Windows .NET console application that lists all SharePoint sites, allows you to select a site and document library, and exports all folders, files, and their permissions to a CSV file using the Microsoft Graph API.

## Features

- **Authentication**: Uses interactive browser authentication with Microsoft Graph API
- **Site Listing**: Lists all accessible SharePoint sites in your tenant
- **Library Selection**: Lists all document libraries in the selected site
- **Recursive Scanning**: Recursively scans all folders and files in the selected library
- **Permission Export**: Exports comprehensive permission information to CSV format
- **Detailed Information**: Includes item names, paths, types, URLs, creators, modifiers, and permission details

## Prerequisites

- .NET 8.0 SDK or later
- Microsoft 365 tenant with SharePoint Online
- Appropriate permissions to access SharePoint sites and read permissions

## Configuration

### Using the Default Client ID (Development/Testing)

The application uses Microsoft Graph Explorer's public client ID by default, which is suitable for development and testing purposes.

### For Production Use (Recommended)

For production environments, you should register your own Azure AD application:

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations** > **New registration**
3. Configure the application:
   - **Name**: SPO Permissions Exporter (or your preferred name)
   - **Supported account types**: Choose appropriate option (e.g., "Accounts in this organizational directory only")
   - **Redirect URI**: Select "Public client/native (mobile & desktop)" and enter `http://localhost`
4. After registration, note the **Application (client) ID** and **Directory (tenant) ID**
5. Under **API permissions**, add the following Microsoft Graph delegated permissions:
   - `Sites.Read.All`
   - `Files.Read.All`
   - `User.Read`
6. Click **Grant admin consent** for your organization
7. Update the `ClientId` and `TenantId` in `Program.cs` (lines 80-81) with your values

## Installation

1. Clone the repository:
```bash
git clone https://github.com/linssenjack/SPO-GraphPermissions.git
cd SPO-GraphPermissions
```

2. Navigate to the project directory:
```bash
cd SPOPermissionsExporter
```

3. Restore dependencies:
```bash
dotnet restore
```

## Usage

1. Build the application:
```bash
dotnet build
```

2. Run the application:
```bash
dotnet run
```

3. Follow the interactive prompts:
   - **Authentication**: A browser window will open for you to sign in with your Microsoft 365 credentials
   - **Site Selection**: View the list of SharePoint sites and enter the number of the site you want to explore
   - **Library Selection**: View the list of document libraries and enter the number of the library you want to scan
   - **Processing**: The application will recursively scan all folders and files
   - **Export**: Permissions will be exported to a CSV file in the format `Permissions_<SiteName>_<LibraryName>_<Timestamp>.csv`

## CSV Output Format

The exported CSV file includes the following columns:
- **Site**: Name of the SharePoint site
- **Library**: Name of the document library
- **Item Name**: Name of the file or folder
- **Item Path**: Full path within the library
- **Item Type**: Either "File" or "Folder"
- **Web URL**: Direct link to the item
- **Created By**: User who created the item
- **Last Modified By**: User who last modified the item
- **Permission ID**: Unique identifier for the permission
- **Permission Type**: Type of permission (User/Group, Link, or Inherited)
- **Roles**: Assigned roles (e.g., read, write, owner)
- **Granted To**: User or group the permission is granted to

## Permissions Required

The application requests the following Microsoft Graph permissions:
- `Sites.Read.All` - Read all SharePoint sites
- `Files.Read.All` - Read all files
- `User.Read` - Read user profile

## Technologies Used

- **.NET 8.0**: Target framework
- **Microsoft.Graph**: Microsoft Graph SDK for .NET
- **Azure.Identity**: Azure authentication library

## Troubleshooting

### Authentication Issues
- Ensure you're using valid Microsoft 365 credentials
- Check that your account has access to SharePoint sites
- Verify that the necessary permissions are granted

### Permission Access Errors
- Some items may show permission errors if they have restricted access
- The application will continue processing other items and note the errors in the export

### No Sites Found
- Ensure your account has access to at least one SharePoint site
- Check with your administrator about site access permissions

## License

This project is provided as-is for educational and utility purposes.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.