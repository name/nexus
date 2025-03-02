# Nexus

Automated tool for packaging and managing Intune applications. Streamlines the process of creating application packages and managing deployments through Microsoft Intune with an intuitive interface.

> Currently a work in progress.

## Features

- **Interactive UI**: User-friendly terminal interface with color-coded menus and selections
- **Application Packaging**: Create ready-to-deploy Intune application packages from MSI or EXE installers
- **Automatic Detection**: Extract product codes and version information from MSI installers
- **Script Generation**: Automatically generate installation and uninstallation scripts
- **Repackaging**: Update existing application packages with new versions
- **Intunewin Creation**: Seamlessly create .intunewin files required for Intune deployment
- **Local & Remote Sources**: Package applications from local files or direct download URLs
- **Standardized Structure**: Consistent package organization for easier management
- **Recent Packages**: Quick access to recently modified packages
- **Custom Package Directory**: Configure where packages are stored
- **Detailed Progress**: Verbose output during package creation for better visibility
- **Path Autocomplete**: Tab completion for file paths when selecting installers
- **Package Suggestions**: Common application name suggestions for faster package creation
- **Error Handling**: Comprehensive error checking and validation throughout the process

## Build

```bash
go build -ldflags="-s -w" -o dist/nexus.exe
```

## How to Use

### Creating a New Application Package

1. Run Nexus and select "New Application Package"
2. Choose your installation source (Local File or Download File)
3. Select the installer type (MSI or EXE)
4. Enter a name for your package
5. Provide the path or URL to the installer file
6. Review the package summary and confirm creation

### Repackaging an Existing Application

1. Run Nexus and select "Repackage Application"
2. Choose the application to repackage from the list (sorted by most recently modified)
3. Review the package details and confirm repackaging

### Customizing Package Location

1. Run Nexus and select "Set Packages Directory"
2. Choose "Use Default Directory" or "Set Custom Directory"
3. If setting a custom directory, enter the path with tab-completion assistance

### Package Structure

Each package created by Nexus includes:

- The original installer file (.msi or .exe)
- Install.ps1 script for installation
- Uninstall.ps1 script for removal
- .intunewin file for Intune deployment

### Customizing Installation

After package creation, you can customize the installation process:

1. Open Install.ps1 in the package directory
2. Modify installation arguments by updating the $install_args variable
3. Add custom PowerShell code before or after the install_application() function
4. Repackage the application to create a new .intunewin file with your changes

### Intune Deployment

After package creation, Nexus provides a comprehensive summary with all the information needed for Intune deployment:

```plaintext
Summary:
    • Name, version and product code (for MSI)
    • Source location

Location:
    • Installer File
    • IntuneWin File
    • Package Directory

Intune Configuration:
    • Install Script
    • Uninstall Script

Intune Detection Method:
    • MSI Product Code (for MSI installers)
    • Version Detection (for MSI installers)
    • Custom detection options (for EXE installers)

Customizing Installation:
    • Installation Arguments
    • Custom Installation Steps
```

This summary contains everything you need to configure the application in Intune, with no additional information required.

## Demo

https://github.com/user-attachments/assets/e7a5117d-af11-4e5a-8a4d-ce2a969df250
