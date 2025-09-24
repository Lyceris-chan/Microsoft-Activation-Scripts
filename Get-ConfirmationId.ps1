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
    
    .EXAMPLE
        Get-ConfirmationId -InstallationId "123456-789012-345678-901234-567890-123456-789012-345678-901234"
        
        Returns the confirmation ID for the specified installation ID.
    
    .NOTES
        Based on Microsoft Activation Scripts by massgravel.dev
        Requires internet connection to contact Microsoft's activation service.
    #>
    
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$InstallationId,
        
        [Parameter(Mandatory = $false)]
        [string]$ExtendedProductId = "31337-42069-123-456789-04-1337-2600.0000-2542001"
    )
    
    # Validate installation ID format (basic check)
    if ($InstallationId -notmatch '^\d{6}-\d{6}-\d{6}-\d{6}-\d{6}-\d{6}-\d{6}-\d{6}-\d{6}$') {
        Write-Warning "Installation ID format may be invalid. Expected format: 123456-789012-345678-901234-567890-123456-789012-345678-901234"
    }
    
    try {
        # Check internet connectivity
        Write-Verbose "Checking internet connectivity..."
        $testConnection = Test-Connection -ComputerName "google.com" -Count 1 -Quiet -ErrorAction Continue
        if (-not $testConnection) {
            try {
                # Alternative connectivity test using .NET
                $webClient = New-Object System.Net.WebClient
                $webClient.DownloadString("http://www.google.com") | Out-Null
                $webClient.Dispose()
            }
            catch {
                throw "Internet connection is required but not available."
            }
        }
        
        Write-Verbose "Creating activation web service client..."
        
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
            XDocument soapResponse = new XDocument();

            try
            {
                IAsyncResult asyncResult = webRequest.BeginGetResponse(null, null);
                asyncResult.AsyncWaitHandle.WaitOne();

                using (WebResponse webResponse = webRequest.EndGetResponse(asyncResult))
                using (StreamReader streamReader = new StreamReader(webResponse.GetResponseStream()))
                {
                    soapResponse = XDocument.Parse(streamReader.ReadToEnd());
                }

                return ParseSoapResponse(soapResponse);
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to get confirmation ID from activation service: " + ex.Message, ex);
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

            XDocument soapRequest = new XDocument();

            using (HMACSHA256 hMACSHA = new HMACSHA256(MacKey))
            {
                string digest = Convert.ToBase64String(hMACSHA.ComputeHash(bytes));

                soapRequest = new XDocument(
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

            return soapRequest;
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
            try
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
                        case "0x8007232B":
                            throw new Exception("DNS name does not exist");
                        case "0x8007232A":
                            throw new Exception("DNS request timed out");
                        case "0x80072EE7":
                            throw new Exception("The server name or address could not be resolved");
                        default:
                            throw new Exception($"Activation service returned error {errorCode}: {errorDescription}");
                    }
                }
                else if (responseXml.Descendants(BatchActivationResponseNs + "ResponseType").Any())
                {
                    string responseType = responseXml.Descendants(BatchActivationResponseNs + "ResponseType").First().Value;

                    switch (responseType)
                    {
                        case "1":
                            string confirmationId = responseXml.Descendants(BatchActivationResponseNs + "CID").First().Value;
                            return confirmationId;

                        case "2":
                            string activationsRemaining = responseXml.Descendants(BatchActivationResponseNs + "ActivationRemaining").First().Value;
                            throw new Exception($"Product key has been blocked. Activations remaining: {activationsRemaining}");

                        default:
                            throw new Exception("The remote server returned an unrecognized response type");
                    }
                }
                else
                {
                    throw new Exception("The remote server returned an unrecognized response format");
                }
            }
            catch (Exception ex) when (!(ex.Message.Contains("activation service") || ex.Message.Contains("installation ID") || ex.Message.Contains("Product key")))
            {
                throw new Exception("Failed to parse activation service response: " + ex.Message, ex);
            }
        }
    }
}
'@

        # Add the C# type to PowerShell session
        Write-Verbose "Compiling activation service client..."
        Add-Type -TypeDefinition $activationCode -Language CSharp -ErrorAction Stop
        
        Write-Verbose "Calling Microsoft activation web service..."
        Write-Verbose "Installation ID: $InstallationId"
        Write-Verbose "Extended Product ID: $ExtendedProductId"
        
        # Call the web service (requestType 1 = confirmation ID request)
        $confirmationId = [ActivationWs.ActivationHelper]::CallWebService(1, $InstallationId, $ExtendedProductId)
        
        if ([string]::IsNullOrEmpty($confirmationId)) {
            throw "Received empty confirmation ID from activation service"
        }
        
        Write-Verbose "Successfully retrieved confirmation ID"
        return $confirmationId
        
    }
    catch {
        Write-Error "Failed to get confirmation ID: $($_.Exception.Message)"
        throw
    }
}

# Export the function if this script is imported as a module
if ($MyInvocation.InvocationName -ne '.') {
    Export-ModuleMember -Function Get-ConfirmationId
}