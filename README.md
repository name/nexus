# Nexus

Automated tool for monitoring, packaging, and versioning Intune applications. Streamlines the process of detecting new software versions, repackaging applications, and managing deployments through Microsoft Intune.

> Currently a work in progress.

## Features

- **Application Packaging**: Create ready-to-deploy Intune application packages from MSI or EXE installers
- **Automatic Detection**: Extract product codes and version information from MSI installers
- **Script Generation**: Automatically generate installation and uninstallation scripts
- **Repackaging**: Update existing application packages with new versions
- **Intunewin Creation**: Seamlessly create .intunewin files required for Intune deployment
- **Local & Remote Sources**: Package applications from local files or direct download URLs
- **Standardized Structure**: Consistent package organization for easier management

## Build

```bash
go build -ldflags="-s -w" -o dist/nexus.exe
```

## How to Use

### Creating a New Application Package

1. Run Nexus and select "New Application Package"
2. Choose your installation source (Local File or Download from Internet)
3. Select the installer type (MSI or EXE)
4. Enter a name for your package
5. Provide the path or URL to the installer file
6. Review the package summary and confirm creation

### Repackaging an Existing Application

1. Run Nexus and select "Repackage Application"
2. Choose the application to repackage from the list
3. Review the package details and confirm repackaging

### Package Structure

Each package created by Nexus includes:

- The original installer file (.msi or .exe)
- Install.ps1 script for installation
- Uninstall.ps1 script for removal
- .intunewin file for Intune deployment

### Intune Deployment

After package creation, Nexus provides a comprehensive summary with all the information needed for Intune deployment:

```plaintext
Package Summary
    • Name, type, version and product code (for MSI)
    • Package location and files
    • Installation and uninstallation commands
    • Detection method details
    • Configuration options
```

This summary contains everything you need to configure the application in Intune, with no additional information required.

## Demo

![Nexus Demo](./docs/videos/nexus.mp4)
