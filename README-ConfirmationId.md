# Get-ConfirmationId PowerShell Function

This PowerShell function generates a confirmation ID from Microsoft's activation service when provided with an installation ID. It supports both Windows and Office activation scenarios.

## Overview

The `Get-ConfirmationId` function automates the process of obtaining a confirmation ID for manual activation of Windows or Office products. It contacts Microsoft's batch activation web service using the same protocol as the official activation tools.

## Features

- ✅ Supports Windows activation IDs
- ✅ Supports Office activation IDs  
- ✅ Validates installation ID format
- ✅ Checks internet connectivity
- ✅ Comprehensive error handling
- ✅ Verbose logging support
- ✅ Based on proven Microsoft Activation Scripts codebase

## Requirements

- Windows PowerShell 5.1 or PowerShell Core 6+
- Internet connection
- Valid installation ID from Windows or Office

## Installation

1. Download the `Get-ConfirmationId.ps1` file
2. Place it in your desired directory
3. Import the function:
   ```powershell
   . .\Get-ConfirmationId.ps1
   ```

## Usage

### Basic Usage

```powershell
Get-ConfirmationId -InstallationId "123456-789012-345678-901234-567890-123456-789012-345678-901234"
```

### With Custom Extended Product ID

```powershell
Get-ConfirmationId -InstallationId "123456-789012-345678-901234-567890-123456-789012-345678-901234" -ExtendedProductId "12345-67890-123-456789-04-1337-9600.0000-1234567"
```

### With Verbose Output

```powershell
Get-ConfirmationId -InstallationId "123456-789012-345678-901234-567890-123456-789012-345678-901234" -Verbose
```

## How to Get Installation ID

### For Windows:
```cmd
slmgr.vbs /dti
```

### For Office 2016/2019/2021:
```cmd
cscript //nologo "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /dinstid
```

### For Office 2013:
```cmd
cscript //nologo "C:\Program Files\Microsoft Office\Office15\OSPP.VBS" /dinstid
```

## Manual Activation Process

1. **Get Installation ID**: Use the commands above to obtain your installation ID
2. **Get Confirmation ID**: Use this PowerShell function to get the confirmation ID
3. **Apply Confirmation ID**: Use the confirmation ID to complete activation:

   **For Windows:**
   ```cmd
   slmgr.vbs /atp [ConfirmationID]
   ```
   
   **For Office:**
   ```cmd
   cscript //nologo "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /actcid:[ConfirmationID]
   ```

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `InstallationId` | String | Yes | The installation ID from Windows or Office |
| `ExtendedProductId` | String | No | Custom extended product ID (uses default if not specified) |

## Return Value

Returns a string containing the confirmation ID that can be used to complete manual activation.

## Error Handling

The function provides detailed error messages for common scenarios:

- Invalid installation ID format
- No internet connection
- Microsoft service errors
- Product key blocked
- DNS resolution issues

## Examples

### Complete Windows Activation Example

```powershell
# Step 1: Import the function
. .\Get-ConfirmationId.ps1

# Step 2: Get installation ID (run in elevated cmd)
# slmgr.vbs /dti

# Step 3: Get confirmation ID
$installationId = "123456-789012-345678-901234-567890-123456-789012-345678-901234"
$confirmationId = Get-ConfirmationId -InstallationId $installationId -Verbose

# Step 4: Apply confirmation ID (run in elevated cmd)
# slmgr.vbs /atp $confirmationId
```

### Complete Office Activation Example

```powershell
# Step 1: Import the function
. .\Get-ConfirmationId.ps1

# Step 2: Get installation ID (run in elevated cmd)
# cscript //nologo "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /dinstid

# Step 3: Get confirmation ID
$installationId = "987654-321098-765432-109876-543210-987654-321098-765432-109876"
$confirmationId = Get-ConfirmationId -InstallationId $installationId

# Step 4: Apply confirmation ID (run in elevated cmd)
# cscript //nologo "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /actcid:$confirmationId
```

## Technical Details

This function uses the same Microsoft batch activation web service protocol as the official tools:

- **Service URL**: `https://activation.sls.microsoft.com/BatchActivation/BatchActivation.asmx`
- **Protocol**: SOAP over HTTPS
- **Authentication**: HMAC-SHA256 message signing
- **Format**: Base64-encoded XML requests/responses

## Troubleshooting

### Common Issues

1. **"Installation ID format may be invalid"**
   - Ensure the installation ID follows the format: `123456-789012-345678-901234-567890-123456-789012-345678-901234`
   - Verify you copied the complete installation ID

2. **"Internet connection is required but not available"**
   - Check your internet connection
   - Verify firewall settings allow HTTPS connections
   - Try running from a different network

3. **"The installation ID is invalid"**
   - The installation ID may be from an already activated system
   - Try getting a fresh installation ID
   - Ensure the system needs activation

4. **"Product key has been blocked"**
   - The product key has exceeded activation limits
   - Contact Microsoft support or use a different key

## Related Tools

This function is based on the Microsoft Activation Scripts project:
- Homepage: https://massgrave.dev
- GitHub: https://github.com/massgravel/Microsoft-Activation-Scripts

## License

This function is derived from the Microsoft Activation Scripts project and follows the same licensing terms.