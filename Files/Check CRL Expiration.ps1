# --------------------------------------------------------------------------------------------------------------------------
# "CheckCRLExpiration" script was designed to check CRL files exposed by HTTP and validate their availability and
# validity period based on a predefined threshold set by the user.
# Written by Omer Eldan, Customer Engineer, Microsoft.
# --------------------------------------------------------------------------------------------------------------------------

# --------------------------------------------------------------------------------------------------------------------------
# Verify-CRLFile Function - Retrieve CRL information from a URL and verify expiration time based on a predefined threshold
# --------------------------------------------------------------------------------------------------------------------------
Function Verify-CRLFile ($CRLServer,$CRLFileName,$CRLExpirationThresholdInHours,$OutputDirectory,$SendAlertsByEmail,$SendAlertsToEventViewer) {
    # Configure required variables for email notificaiton, if required.
    If ($SendAlertsByEmail) {
        $MailMessage = New-Object System.Net.Mail.MailMessage 
        $MailMessage.From = "PKI@domain.local"  
        $MailMessage.To.Add("Administator@domain.local") 
        $MailMessage.IsBodyHtml = $True 
        $MailMessage.Subject = "PKI - Check CRL Expiration Script"
        $SMTPServer = "Exchange.domain.local" 
        $SMTPClient = New-Object Net.Mail.SmtpClient($SMTPServer)  
    }

     # Check URL response status code to make sure the CRL file is available.
    $CRLURLPath = "http://$CRLServer/CertEnroll/$CRLFileName"
    $OutputFile = "$OutputDirectory\$CRLServer_$CRLFileName"
    $WebRequestResponse = Invoke-WebRequest -Uri $CRLURLPath
    If (($WebRequestResponse.StatusCode) -ne "200") {
        $Message = "There might be a problem with CRL availability. The CRL file '$CRLFileName' on server '$CRLServer' is not accessible."
        Write-Host $Message -ForegroundColor Red
        if ($SendAlertsToEventViewer) {
            Write-EventLog -LogName Application -Source "CRL Monitoring Script" -Message $Message -EntryType Error -EventId 1003
        }
        if ($SendAlertsByEmail) {
            $MailMessage.Body = $Message 
            $SMTPClient.Send($MailMessage)
        }
        Return   
    }

    # Check CRL expiration times and validate it against the predefined threshold.
    Invoke-WebRequest -Uri $CRLURLPath -OutFile $OutputFile
    $CRLData = Get-CertificateRevocationList $OutputFile
    $TimeSpan = New-TimeSpan -End $CRLData.NextUpdate
    if (($TimeSpan.Days*24+$TimeSpan.Hours) -le $CRLExpirationThresholdInHours) {
        $Message = "There might be a problem with CRL updating process. The CRL file '$CRLFileName' on server '$CRLServer' will expire within $($TimeSpan.Days) days and $($TimeSpan.Hours) hours."
        Write-Host $Message -ForegroundColor Red
        if ($SendAlertsToEventViewer) {
            Write-EventLog -LogName Application -Source "CRL Monitoring Script" -Message $Message -EntryType Warning -EventId 1002
        }
        if ($SendAlertsByEmail) {
            $MailMessage.Body = $Message 
            $SMTPClient.Send($MailMessage)
        }
    }
    else {
        $Message = "CRL file on server '$CRLServer' is updated, and is within the valid range of expiration ($CRLExpirationThresholdInHours hours). The CRL file '$CRLFileName' will expire within $($TimeSpan.Days) days and $($TimeSpan.Hours) hours."
        Write-Host $Message -ForegroundColor Green
        if ($SendAlertsToEventViewer) {
            Write-EventLog -LogName Application -Source "CRL Monitoring Script" -Message $Message -EntryType Information -EventId 1001
        }
        if ($SendAlertsByEmail) {
            $MailMessage.Body = $Message 
            $SMTPClient.Send($MailMessage)
        }
    }
}

# --------------------------------------------------------------------------------------------------------------------------
# Verify-CRLAvailabilityUsingNLB Function - Validate CRL availability when accessing the CRL using a NLB address
# --------------------------------------------------------------------------------------------------------------------------
Function Verify-CRLAvailabilityUsingNLB ($CRLNLBName,$CRLFileName,$SendAlertsByEmail,$SendAlertsToEventViewer) {
    $CRLURLPath = "http://$CRLNLBName/CertEnroll/$CRLFileName"
    $WebRequestResponse = Invoke-WebRequest -Uri $CRLURLPath
    If (($WebRequestResponse.StatusCode) -eq "200") {
        $Message = "CRL file named '$CRLFileName' is accessible and valid using the network load balancer address ($CRLURLPath)."
        Write-EventLog -LogName Application -Source "CRL Monitoring Script" -Message $Message -EntryType Information -EventId 2001
        Write-Host $Message -ForegroundColor Green   
    }
    else {
        $Message = "There was a problem to access the CRL file named '$CRLFileName using the network load balancer address ($CRLURLPath)."
        Write-EventLog -LogName Application -Source "CRL Monitoring Script" -Message $Message -EntryType Error -EventId 2002
        Write-Host $Message -ForegroundColor Red
    }
}

# --------------------------------------------------------------------------------------------------------------------------
# Install and import required PowerShell modules
# --------------------------------------------------------------------------------------------------------------------------
<# Install PSPKI PowerShell module and its prerequisites for offline environments. Should be pefromed once:
#  Download Nuget from an online, connected workstation, and copy the Nuget folder to "C:\Program Files\PackageManagement\ProviderAssemblies".
#  Download PSPKI PowerShell module, and copy the whole folder to "C:\Program Files\WindowsPowerShell\Modules".
#>

Import-PackageProvider -Name NuGet -RequiredVersion 2.8.5.208
Import-Module PSPKI
 
# --------------------------------------------------------------------------------------------------------------------------
# Define general variables
# --------------------------------------------------------------------------------------------------------------------------
$SendAlertsByEmail = $false
$SendAlertsToEventViewer = $true
$OutputDirectory = "C:\Scripts\PKI Monitoring"
$CRLServers = "CRL01.domain.local","CRL02.domain.local" ## Change accroding to your CRL server addresses
$CRLNLBs = "CRLVIP01" ## Change accroding to your CRL network load balancer addresses

# --------------------------------------------------------------------------------------------------------------------------
# Validate prerequisites
# --------------------------------------------------------------------------------------------------------------------------

#Checks if the output folder is exist, and if not creates it.
If ( -Not (Test-Path -Path $OutputDirectory) ) {
    New-Item -ItemType Directory -Path $OutputDirectory
}

#Checks if Event Viewer is used by the script, and if so creates the event source if it's not already existed
If ($SendAlertsToEventViewer) {
    $EventLogName = "Application"
    $EventSource = "CRL Monitoring Script"
    if ([System.Diagnostics.EventLog]::SourceExists($EventSource) -eq $false) {
        [System.Diagnostics.EventLog]::CreateEventSource($EventSource,$EventLogName)
    }    
}

# --------------------------------------------------------------------------------------------------------------------------
# Main - Call relevant functions with provided parameters
# --------------------------------------------------------------------------------------------------------------------------
foreach ($CRLServer in $CRLServers) {
    Verify-CRLFile -CRLServer $CRLServer -CRLFileName "SubCA01.CRL" -CRLExpirationThresholdInHours 48 -OutputDirectory $OutputDirectory -SendAlertsByEmail $false -SendAlertsToEventViewer $true
}