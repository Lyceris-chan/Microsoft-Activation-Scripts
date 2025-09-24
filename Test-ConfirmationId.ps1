# Test script for Get-ConfirmationId function
# This script demonstrates how to use the Get-ConfirmationId function

# Import the function
. "$PSScriptRoot\Get-ConfirmationId.ps1"

Write-Host "=== Get-ConfirmationId Function Test ===" -ForegroundColor Cyan
Write-Host ""

# Example usage with verbose output
Write-Host "Example usage:" -ForegroundColor Green
Write-Host "Get-ConfirmationId -InstallationId `"123456-789012-345678-901234-567890-123456-789012-345678-901234`" -Verbose" -ForegroundColor Yellow
Write-Host ""

Write-Host "Note: This function requires a valid installation ID from a Windows or Office system" -ForegroundColor Yellow
Write-Host "that needs activation. You can obtain an installation ID using:" -ForegroundColor Yellow
Write-Host ""
Write-Host "For Windows:" -ForegroundColor White
Write-Host "  slmgr.vbs /dti" -ForegroundColor Gray
Write-Host ""
Write-Host "For Office:" -ForegroundColor White
Write-Host "  cscript //nologo `"C:\Program Files\Microsoft Office\Office16\OSPP.VBS`" /dinstid" -ForegroundColor Gray
Write-Host ""

Write-Host "The function will:" -ForegroundColor Green
Write-Host "  1. Validate the installation ID format" -ForegroundColor White
Write-Host "  2. Check internet connectivity" -ForegroundColor White
Write-Host "  3. Contact Microsoft's activation service" -ForegroundColor White
Write-Host "  4. Return the confirmation ID for manual activation" -ForegroundColor White
Write-Host ""

Write-Host "Example with a test installation ID (this will fail as it's not real):" -ForegroundColor Yellow
try {
    $testId = "123456-789012-345678-901234-567890-123456-789012-345678-901234"
    Write-Host "Testing with: $testId" -ForegroundColor Gray
    $result = Get-ConfirmationId -InstallationId $testId -Verbose -ErrorAction Stop
    Write-Host "Confirmation ID: $result" -ForegroundColor Green
}
catch {
    Write-Host "Expected error (test ID is not valid): $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "To use with a real installation ID, replace the test ID above with your actual installation ID." -ForegroundColor Cyan