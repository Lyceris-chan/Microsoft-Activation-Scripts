@echo off

::============================================================================
::
::   Homepage: mass()grave(dot)dev
::      Email: mas.help@outlook.com
::
::============================================================================

::  Get Confirmation ID from Installation ID
::  Based on Microsoft Activation Scripts
::  This script provides a PowerShell function to get confirmation IDs

::  Set environment variables
setlocal EnableExtensions
setlocal DisableDelayedExpansion

set "PathExt=.COM;.EXE;.BAT;.CMD;.VBS;.VBE;.JS;.JSE;.WSF;.WSH;.MSC"
set "SysPath=%SystemRoot%\System32"
set "Path=%SystemRoot%\System32;%SystemRoot%;%SystemRoot%\System32\Wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\"

if exist "%SystemRoot%\Sysnative\reg.exe" (
set "SysPath=%SystemRoot%\Sysnative"
set "Path=%SystemRoot%\Sysnative;%SystemRoot%;%SystemRoot%\Sysnative\Wbem;%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\;%Path%"
)

set "ComSpec=%SysPath%\cmd.exe"
set "PSModulePath=%ProgramFiles%\WindowsPowerShell\Modules;%SysPath%\WindowsPowerShell\v1.0\Modules"

set "ps=%SysPath%\WindowsPowerShell\v1.0\powershell.exe"
set "_psc=%ps% -nop -c"
set "_err===== ERROR ===="
set _pwsh=1

if not exist %ps% set _pwsh=0
cmd /c "%_psc% "$ExecutionContext.SessionState.LanguageMode"" | find /i "FullLanguage" 1>nul || (set _pwsh=0)

if %_pwsh% equ 0 (
echo %_err%
echo Windows PowerShell is not working correctly.
echo It is required for this script to work.
goto :E_Exit
)

set "_batf=%~f0"
set "_batp=%_batf:'=''%"
setlocal EnableDelayedExpansion

cls
echo.
echo ========================================================================
echo                        Get Confirmation ID
echo ========================================================================
echo.
echo This tool generates a confirmation ID from an installation ID
echo for manual activation of Windows or Office products.
echo.
echo How to get Installation ID:
echo   Windows: slmgr.vbs /dti
echo   Office:  cscript //nologo "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /dinstid
echo.
echo ========================================================================
echo.

%_psc% "$f=[System.IO.File]::ReadAllText('!_batp!') -split ':powershell\:.*';. ([scriptblock]::Create($f[1]))"

:E_Exit
echo.
echo Press any key to exit.
pause >nul
exit /b

:powershell:

# Get-ConfirmationId PowerShell Function
function Get-ConfirmationId {
    <#
    .SYNOPSIS
        Retrieves a confirmation ID from Microsoft's activation service using an installation ID.
    
    .DESCRIPTION
        This function generates a confirmation ID by calling Microsoft's batch activation web service
        with the provided installation ID. It supports both Windows and Office activation scenarios.
    
    .PARAMETER InstallationId
        The installation ID obtained from the system that needs activation.
        This should be a string of digits separated by hyphens.
    
    .PARAMETER ExtendedProductId
        Optional. The extended product ID. If not provided, uses a default value.
    #>
    
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$InstallationId,
        
        [Parameter(Mandatory = $false)]
        [string]$ExtendedProductId = "31337-42069-123-456789-04-1337-2600.0000-2542001"
    )
    
    # Validate installation ID format
    if ($InstallationId -notmatch '^\d{6}-\d{6}-\d{6}-\d{6}-\d{6}-\d{6}-\d{6}-\d{6}-\d{6}$') {
        Write-Warning "Installation ID format may be invalid. Expected format: 123456-789012-345678-901234-567890-123456-789012-345678-901234"
    }
    
    try {
        # Check internet connectivity
        Write-Host "Checking internet connectivity..." -ForegroundColor Yellow
        $testConnection = $false
        try {
            $webClient = New-Object System.Net.WebClient
            $webClient.DownloadString("http://www.google.com") | Out-Null
            $webClient.Dispose()
            $testConnection = $true
        }
        catch {
            throw "Internet connection is required but not available."
        }
        
        Write-Host "Internet connection verified." -ForegroundColor Green
        Write-Host "Creating activation web service client..." -ForegroundColor Yellow
        
        # Define the C# code for the activation web service client
        $activationCode = @'
using System;
using System.IO;
using System.Net;
using System.Text;
using System.Xml.Linq;
using System.Security.Cryptography;
using System.Linq;

namespace ActivationWs
{
    public class ActivationHelper
    {
        private static readonly byte[] MacKey = {
            0xFE, 0x31, 0x98, 0x75, 0xFB, 0x48, 0x84, 0x86, 0x9C, 0xF3, 0xF1, 0xCE, 0x99, 0xA8, 0x90, 0x64
        };

        private static readonly XNamespace SoapEnvelopeNs = "http://schemas.xmlsoap.org/soap/envelope/";
        private static readonly XNamespace XmlSchemaNs = "http://www.w3.org/2001/XMLSchema";
        private static readonly XNamespace BatchActivationServiceNs = "http://www.microsoft.com/BatchActivationService";
        private static readonly XNamespace BatchActivationRequestNs = "http://www.microsoft.com/DRM/SL/BatchActivationRequest/1.0";
        private static readonly XNamespace BatchActivationResponseNs = "http://www.microsoft.com/DRM/SL/BatchActivationResponse/1.0";

        public static string CallWebService(int requestType, string installationId, string extendedProductId)
        {
            XDocument soapRequest = CreateSoapRequest(requestType, installationId, extendedProductId);
            HttpWebRequest webRequest = CreateWebRequest(soapRequest);

            try
            {
                using (WebResponse webResponse = webRequest.GetResponse())
                using (StreamReader streamReader = new StreamReader(webResponse.GetResponseStream()))
                {
                    XDocument soapResponse = XDocument.Parse(streamReader.ReadToEnd());
                    return ParseSoapResponse(soapResponse);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to get confirmation ID: " + ex.Message);
            }
        }

        private static XDocument CreateSoapRequest(int requestType, string installationId, string extendedProductId)
        {
            XElement activationRequest = new XElement(BatchActivationRequestNs + "ActivationRequest",
                new XElement(BatchActivationRequestNs + "VersionNumber", "2.0"),
                new XElement(BatchActivationRequestNs + "RequestType", requestType),
                new XElement(BatchActivationRequestNs + "Requests",
                    new XElement(BatchActivationRequestNs + "Request",
                        new XElement(BatchActivationRequestNs + "PID", extendedProductId),
                        requestType == 1 ? new XElement(BatchActivationRequestNs + "IID", installationId) : null)
                )
            );

            byte[] bytes = Encoding.Unicode.GetBytes(activationRequest.ToString());
            string requestXml = Convert.ToBase64String(bytes);

            using (HMACSHA256 hMACSHA = new HMACSHA256(MacKey))
            {
                string digest = Convert.ToBase64String(hMACSHA.ComputeHash(bytes));

                return new XDocument(
                    new XDeclaration("1.0", "UTF-8", "no"),
                    new XElement(SoapEnvelopeNs + "Envelope",
                        new XAttribute(XNamespace.Xmlns + "soap", SoapEnvelopeNs),
                        new XAttribute(XNamespace.Xmlns + "xsi", XmlSchemaNs + "-instance"),
                        new XAttribute(XNamespace.Xmlns + "xsd", XmlSchemaNs),
                        new XElement(SoapEnvelopeNs + "Body",
                            new XElement(BatchActivationServiceNs + "BatchActivate",
                                new XElement(BatchActivationServiceNs + "request",
                                    new XElement(BatchActivationServiceNs + "Digest", digest),
                                    new XElement(BatchActivationServiceNs + "RequestXml", requestXml)
                                )
                            )
                        )
                    )
                );
            }
        }

        private static HttpWebRequest CreateWebRequest(XDocument soapRequest)
        {
            HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create("https://activation.sls.microsoft.com/BatchActivation/BatchActivation.asmx");
            webRequest.Headers.Add("SOAPAction", "http://www.microsoft.com/BatchActivationService/BatchActivate");
            webRequest.ContentType = "text/xml; charset=utf-8";
            webRequest.Accept = "text/xml";
            webRequest.Method = "POST";
            webRequest.UserAgent = "BatchActivationClient";
            webRequest.Timeout = 30000;

            byte[] soapBytes = Encoding.UTF8.GetBytes(soapRequest.ToString());
            webRequest.ContentLength = soapBytes.Length;

            using (Stream requestStream = webRequest.GetRequestStream())
            {
                requestStream.Write(soapBytes, 0, soapBytes.Length);
            }

            return webRequest;
        }

        private static string ParseSoapResponse(XDocument soapResponse)
        {
            XElement responseXmlElement = soapResponse.Descendants(BatchActivationServiceNs + "ResponseXml").FirstOrDefault();
            
            if (responseXmlElement == null)
            {
                throw new Exception("Invalid response format from activation service");
            }

            byte[] responseBytes = Convert.FromBase64String(responseXmlElement.Value);
            string responseXmlString = Encoding.Unicode.GetString(responseBytes);
            XDocument responseXml = XDocument.Parse(responseXmlString);

            if (responseXml.Descendants(BatchActivationResponseNs + "ErrorCode").Any())
            {
                string errorCode = responseXml.Descendants(BatchActivationResponseNs + "ErrorCode").First().Value;
                string errorDescription = responseXml.Descendants(BatchActivationResponseNs + "ErrorDescription").FirstOrDefault()?.Value ?? "Unknown error";
                
                switch (errorCode)
                {
                    case "0x8007000D":
                        throw new Exception("The installation ID is invalid. Please check the Installation ID and try again");
                    default:
                        throw new Exception($"Activation service error {errorCode}: {errorDescription}");
                }
            }
            else if (responseXml.Descendants(BatchActivationResponseNs + "ResponseType").Any())
            {
                string responseType = responseXml.Descendants(BatchActivationResponseNs + "ResponseType").First().Value;

                switch (responseType)
                {
                    case "1":
                        return responseXml.Descendants(BatchActivationResponseNs + "CID").First().Value;
                    case "2":
                        string activationsRemaining = responseXml.Descendants(BatchActivationResponseNs + "ActivationRemaining").First().Value;
                        throw new Exception($"Product key blocked. Activations remaining: {activationsRemaining}");
                    default:
                        throw new Exception("Unrecognized response type from activation service");
                }
            }
            else
            {
                throw new Exception("Unrecognized response format from activation service");
            }
        }
    }
}
'@

        Add-Type -TypeDefinition $activationCode -Language CSharp -ErrorAction Stop
        
        Write-Host "Contacting Microsoft activation service..." -ForegroundColor Yellow
        $confirmationId = [ActivationWs.ActivationHelper]::CallWebService(1, $InstallationId, $ExtendedProductId)
        
        return $confirmationId
    }
    catch {
        Write-Error "Failed to get confirmation ID: $($_.Exception.Message)"
        throw
    }
}

# Interactive mode
Write-Host "Enter your Installation ID (or press Enter to exit):" -ForegroundColor Cyan
$installationId = Read-Host

if ([string]::IsNullOrWhiteSpace($installationId)) {
    Write-Host "Exiting..." -ForegroundColor Yellow
    return
}

try {
    Write-Host ""
    Write-Host "Processing Installation ID: $installationId" -ForegroundColor White
    Write-Host ""
    
    $confirmationId = Get-ConfirmationId -InstallationId $installationId
    
    Write-Host ""
    Write-Host "SUCCESS!" -ForegroundColor Green
    Write-Host "========" -ForegroundColor Green
    Write-Host "Confirmation ID: $confirmationId" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "To complete activation:" -ForegroundColor Cyan
    Write-Host "  Windows: slmgr.vbs /atp $confirmationId" -ForegroundColor White
    Write-Host "  Office:  cscript //nologo \"C:\Program Files\Microsoft Office\Office16\OSPP.VBS\" /actcid:$confirmationId" -ForegroundColor White
}
catch {
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Common issues:" -ForegroundColor Yellow
    Write-Host "  - Installation ID format is incorrect" -ForegroundColor White
    Write-Host "  - System is already activated" -ForegroundColor White
    Write-Host "  - Internet connection problems" -ForegroundColor White
    Write-Host "  - Product key has been blocked" -ForegroundColor White
}