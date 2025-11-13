# KiCad Excel Bridge

## Overview
The KiCad Excel Bridge is a .NET 9 WPF application designed to facilitate the integration of KiCad with Excel/CSV files. It provides a tray application with an HTTP API for managing and mapping fields between KiCad and Excel/CSV data. The application includes a user-friendly field-mapping UI and supports per-sheet settings for flexible configuration.

### Key Features
- **Tray Application**: Runs in the system tray for easy access.
- **HTTP API**: Exposes endpoints for querying categories, parts, and details.
- **Excel/CSV Ingestion**: Supports previewing and mapping fields from Excel/CSV files.
- **Field Mapping UI**: Allows users to map fields, add/remove custom fields, and preview data.
- **Robust ID System**: Ensures unique category and part IDs.

<img width="786" height="443" alt="image" src="https://github.com/user-attachments/assets/9c6484bb-6888-40bd-af38-fe0c5e4d810e" />
<img width="836" height="493" alt="image" src="https://github.com/user-attachments/assets/90402bbf-75ec-48bf-bb2e-f8dcb1c9e247" />

## Prerequisites
- **.NET 9 SDK**: Ensure you have the .NET 9 SDK installed.
- **PowerShell**: For running the provided query scripts.

## Getting Started

### Build and Run
1. Clone the repository:
   ```powershell
   git clone <repository-url>
   cd kicad_excel
   ```
2. Build the solution:
   ```powershell
   dotnet build KiCadExcelBridge.csproj
   ```
3. Run the application:
   ```powershell
   dotnet run --project KiCadExcelBridge.csproj
   ```

### Using the Application
- **Tray Icon**: Double-click the tray icon to open the main window.
- **Field Mapping**: Use the UI to configure field mappings for each sheet.
- **HTTP API**: The server starts automatically at `http://localhost:8088/kicad-api/v1/`.

### API Endpoints
- **Validation**: `GET /v1/` - Returns validation payload.
- **Categories**: `GET /v1/categories.json` - Lists available categories.
- **Parts**: `GET /v1/parts/category/{categoryId}.json` - Lists parts in a category.
- **Part Details**: `GET /v1/parts/{partId}.json` - Retrieves details for a specific part.

### Query Script
A PowerShell script is provided to query the API:
```powershell
.\scripts\query_kicad_api.ps1 parts-for Resistors
```

## Logs
HTTP requests are logged to `http_requests.log` in the application directory.

## Contributing
Contributions are welcome! Feel free to submit issues or pull requests.

## License
This project is licensed under the GPLv3 License.

## KiCad Setup

To integrate the HTTP library with KiCad:

1. Locate the `library.kicad_httplib` file in the project directory.
2. Open KiCad's Project Manager.
3. Navigate to `Preferences -> Manage Symbol Libraries`.
4. Click the folder icon in the bottom-left corner to add a Global Library.
5. Select the `library.kicad_httplib` file. KiCad will automatically detect it as an HTTP library.
6. Set a Nickname for the library and confirm.

Once added, the HTTP library will be available for use in KiCad.

## Installer

This project includes an [Inno Setup](https://jrsoftware.org/isinfo.php) script to create a Windows installer.

### Prerequisites
- **Inno Setup**: You must have Inno Setup installed.

### Creating the Installer
1. **Publish the application**:
   First, publish the application to a single-file executable.
   ```powershell
   dotnet publish -c Release -r win-x64 --self-contained true
   ```
   This will create the executable in `bin\Release\net9.0-windows\win-x64\publish\`.

2. **Compile the setup script**:
   - Open the `installer\setup.iss` file in the Inno Setup Compiler.
   - Click `Build -> Compile` (or press F9).
   - The installer (`KiCadExcelBridge-1.0.0-setup.exe`) will be created in the project's root directory.

The installer will:
- Install the application to the user's `Program Files` folder.
- Optionally create a desktop icon.
- Optionally configure the application to run at Windows startup.
