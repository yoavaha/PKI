Function Create-HTMLReport ($CertificatesObjectsArray) {
    $Date = Get-Date -Format "yyyy-MM-dd"
    $ReportsDir = "C:\Reports"
    If( -Not (Test-Path -Path $ReportsDir ) )
    {
        New-Item -ItemType Directory -Path $ReportsDir
    }
    $ReportFileName = "ADCSExpiringCerts_$Date.html"
    $Body = "<H1>AD CS - Expiring Certificates Report</H1>
    <br>
    <table>
        <tr>
            <th>Request ID</th>
            <th>Common Name</th>
            <th>Expired On</th>
            <th>Requester Name</th>
            <th>Certificate Template</th>
        </tr>"
        foreach ($CertificateObject in $CertificatesObjectsArray) {
            $Body += "<tr>
                <td>$($CertificateObject.'Request ID')</td>
                <td>$($CertificateObject.'Common Name')</td>
                <td>$($CertificateObject.'Expired On')</td>
                <td>$($CertificateObject.'Requester Name')</td>
                <td>$($CertificateObject.'Certificate Template')</td>
            </tr>"   
        }
    $ConvertParams = @{
    head = @"
    <Title>AD CS - Expiring Certificates Report</Title>
    <Style>
        body { 
            background-color:#E5E4E2;
            font-family:Arial, sans-serif;
            font-size:10pt; }
        table {
            border-collapse: collapse;
            width: 80%;
            margin-left:auto; 
            margin-right:auto; }
        th, td {
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #ddd; }
        th {
            color:white;
            background-color:#525251; }
        tr {background-color:#F0F0F0;}
        tr:hover {background-color:#D4D4D4;}
        H1 {
            font-family:Tahoma;
            color:#525251;
            text-align:center; }
        H@ {
            font-family:Tahoma;
            color:#6D7B8D; }
    </style>
"@
 body = $Body
}
    ConvertTo-Html @ConvertParams | Out-File "$ReportsDir\$ReportFileName"
    If(Test-Path -Path "$ReportsDir\$ReportFileName") {
        Write-Host "The report has been created successfully in the following location: $ReportsDir\$ReportFileName" -ForegroundColor Green
    }
    Else {
        Write-Host "Failed to create the report because of an error." -ForegroundColor Red
    }
}

#Install-Module -Name PSPKI
Import-Module -Name PSPKI

#Initialize Parameters 
$CertificateTemplates = @()
$CertificateTemplatesOID = @()
$CertificatesObjectsArray = @()
$CertificateTemplatesFilter = "Web Server - SHA256","Contoso-CodeSigning","Contoso-SmartcardLogon" #Change to match organization certificate templates
$CAName = "Contoso Subordinate" #Change to match organization CA Common Name
$ExpirationTime = 72 #In Months

#Select the Enterprise CA to review
$EnterpriseCA = Get-CertificationAuthority -Name $CAName

#Create an array of Certificate Templates objects based on the Certificate Template Filter configured in 'CertificateTemplatesFilter'
foreach ($CertificateTemplateFilter in $CertificateTemplatesFilter) {
    $CertificateTemplates += Get-CertificateTemplate -Name $CertificateTemplateFilter
}

#Create an array of Certificate Templates OIDs based on the Certificate Template objects array
foreach ($CertificateTemplateObject in $CertificateTemplates) {
    $CertificateTemplatesOID += $CertificateTemplateObject.OID.Value
}

#Configure a filter for certificates retrieval CmdLet (Get-IssuedRequest) 
$IssuedRequestFilter = "NotAfter -ge $(Get-Date)", "NotAfter -le $((Get-Date).AddMonths($ExpirationTime))"

#Get issued requests based on the filter created
$ExpiringCertificates = Get-IssuedRequest -CertificationAuthority $EnterpriseCA -Filter $IssuedRequestFilter | Select-Object RequestId, Request.RequesterName, CommonName, NotAfter,CertificateTemplate

#Filter retrieved certificates based on the their certificate template
foreach ($Certificate in $ExpiringCertificates) {
    if ($CertificateTemplatesOID.Contains($Certificate.CertificateTemplate)) {
        $CertificateObject = New-Object -TypeName PSObject
        $CertificateTemplate = Get-CertificateTemplate -OID $Certificate.CertificateTemplate
        Add-Member -InputObject $CertificateObject -MemberType 'NoteProperty' -Name 'Request ID' -Value $($Certificate.RequestID)
        Add-Member -InputObject $CertificateObject -MemberType 'NoteProperty' -Name 'Common Name' -Value $($Certificate.CommonName)
        Add-Member -InputObject $CertificateObject -MemberType 'NoteProperty' -Name 'Expired On' -Value $($Certificate.NotAfter)
        Add-Member -InputObject $CertificateObject -MemberType 'NoteProperty' -Name 'Requester Name' -Value $($Certificate.'Request.RequesterName')
        Add-Member -InputObject $CertificateObject -MemberType 'NoteProperty' -Name 'Certificate Template' -Value $($CertificateTemplate.DisplayName)
        $CertificatesObjectsArray += $CertificateObject
    }
}

Create-HTMLReport($CertificatesObjectsArray)

# ---------------------------------------------------------------------------------------------------------------
# Send Mail - Define Variables
# ----------------------------------------------------------------------------------------------------------------

$date = (get-date -Format dd.MM.yyyy) 
$fromaddress = "PKI@contoso.local" 
$toaddress = "Administator@contoso.local" 
$Subject = "PKI Expiring Certificates Report" 
$Date = Get-Date -Format "yyyy-MM-dd"
$ReportPath = "C:\Reports\ADCSExpiringCerts_$Date.html"
$body = Get-Content $ReportPath
#$attachment = "c:\Reports\Attachment.txt" 
$smtpserver = "Exchange.contoso.local" 
 
# ---------------------------------------------------------------------------------------------------------------
# Send Mail - Create And Send The Mail
# ----------------------------------------------------------------------------------------------------------------
 
$message = New-Object System.Net.Mail.MailMessage 
$message.From = $fromaddress 
$message.To.Add($toaddress) 
$message.IsBodyHtml = $True 
$message.Subject = $Subject 
#$attach = New-Object Net.Mail.Attachment($attachment) 
#$message.Attachments.Add($attach) 
$message.body = $body 
$smtp = New-Object Net.Mail.SmtpClient($smtpserver) 
$smtp.Send($message)