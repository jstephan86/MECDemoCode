#Install modules first before import
#https://www.powershellgallery.com/packages/MSAL.PS/4.37.0.0
#https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation?view=graph-powershell-1.0
#https://learn.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps

#Import the MSAL PS.module first, then Graph modules, then EXO Powershell last - issues can arise with token fetch if EXO imported before token module
Import-Module MSAL.PS
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Users.Actions
Import-Module Microsoft.Graph.Mail
Import-Module Microsoft.Graph.Calendar
Import-Module ExchangeOnlineManagement

$ConnectParameters = 
@{
# Azure AD Application Id for Authentication
ClientId = ""
# The Id or Name of the tenant to authenticate against
TenantId = ""
# Thumbprint of the certificate to use for authentication
CertificateThumbprint = ""
}

function SetEWSConnection
	{
		param (
			[switch]$Refresh
		)
		
		if (!$Refresh)
		{
			#Load EWS Binary
			try
			{
				#https://www.nuget.org/packages/Exchange.WebServices.Managed.Api/
				$strEWSPackagePath = Get-Package -Name Exchange.WebServices.Managed.Api | ForEach-Object{ $_.Source.Substring(0, $_.Source.LastIndexOf('\')) }
				$strEWSDllPath = (Get-ChildItem $strEWSPackagePath -Recurse | Where-Object{ $_.Name -eq "Microsoft.Exchange.WebServices.dll" }).FullName
				Add-Type -Path $strEWSDllPath
			}
			catch
			{
				if ($global:blEnableLogging -eq $true) { Function_LogMessage -strMessage "Error loading EWS Binaries.`r`n $($_.Exception.Message)" }
				return "Error"
			}
			
			#Load Microsoft.Identity.Client Binary
			try
			{
				Add-Type -Path "C:\Program Files\WindowsPowerShell\Modules\MSAL.PS\4.37.0.0\Microsoft.Identity.Client.4.37.0\net45\Microsoft.Identity.Client.dll"
			}
			catch
			{
				if ($global:blEnableLogging -eq $true) { Function_LogMessage -strMessage "Error loading Microsoft.Identity.Client Binaries.`r`n $($_.Exception.Message)" }
				return "Error"
			}
		}
		
		#Obtain MSAL Token
		try
		{
			$objCertificate = (Get-ChildItem Cert:\CurrentUser\My) | Where-Object { $_.Thumbprint -match $ConnectParameters.CertificateThumbprint }
			$strScope = "https://outlook.office365.com/.default"
			if ($Refresh)
			{
				$objMSALToken = Get-MsalToken -ClientId $ConnectParameters.ClientId -TenantId $ConnectParameters.TenantID -ClientCertificate $objCertificate -Scopes $strScope -ForceRefresh
			}
			else
			{
				$objMSALToken = Get-MsalToken -ClientId $ConnectParameters.ClientId -TenantId $ConnectParameters.TenantID -ClientCertificate $objCertificate -Scopes $strScope
			}
			
		}
		catch
		{
			if ($global:blEnableLogging -eq $true) { Function_LogMessage -strMessage "Error getting MSAL token.`r`n $($_.Exception.Message)" }
			return "Error"
		}
		
		#Set EWS Credentials
		try
		{
			$objEWSService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016)
			$objEWSService.Url = "https://outlook.office365.com/EWS/Exchange.asmx"
			$objEWSService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $objMSALToken.AccessToken
		}
		catch
		{
			if ($global:blEnableLogging -eq $true) { Function_LogMessage -strMessage "Error connecting to EWS.`r`n $($_.Exception.Message)" }
			return "Error"
		}
		
		#Add Impersonation properties
		try
		{
			$objEWSService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, "")
			$objEWSService.HttpHeaders.Add("X-AnchorMailbox", "")
		}
		catch
		{
			if ($global:blEnableLogging -eq $true) { Function_LogMessage -strMessage "Error adding EWS Impersonation properties.`r`n $($_.Exception.Message)" }
			return "Error"
		}
		
		#Start timer for token refresh
		if ($global:objStopwatch.IsRunning -eq $true)
		{
			$global:objStopwatch.Restart()
		}
		else
		{
			$global:objStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
		}
		
		return $objEWSService
		
	} #End Function_SetEWSConnection

function GraphSendMail
{
	<#
    Ref:

    https://mikecrowley.us/2021/10/27/sending-email-with-send-mgusermail-microsoft-graph-powershell
    https://docs.microsoft.com/en-us/graph/api/user-sendmail
    https://docs.microsoft.com/en-us/powershell/module/microsoft.graph.users.actions/send-mgusermail

    #>

	$emailRecipients   = @(
        'julian.stephan@qtjustdev.onmicrosoft.com'
        
    )
    $emailSender  = 'julian.stephan@qtjustdev.onmicrosoft.com'

    $emailSubject = "Sample Email | " + (Get-Date -UFormat %e%b%Y)

	Function ConvertTo-IMicrosoftGraphRecipient {
        [cmdletbinding()]
        Param(
            [array]$SmtpAddresses        
        )
        foreach ($address in $SmtpAddresses) {
            @{
                emailAddress = @{address = $address}
            }    
        }    
    }

    Function ConvertTo-IMicrosoftGraphAttachment {
        [cmdletbinding()]
        Param(
            [string]$UploadDirectory        
        )
        $directoryContents = Get-ChildItem $UploadDirectory -Attributes !Directory -Recurse
        foreach ($file in $directoryContents) {
            $encodedAttachment = [convert]::ToBase64String((Get-Content $file.FullName -Encoding byte))
            @{
                "@odata.type"= "#microsoft.graph.fileAttachment"
                name = ($File.FullName -split '\\')[-1]
                contentBytes = $encodedAttachment
            }   
        }    
    }

	[array]$toRecipients = ConvertTo-IMicrosoftGraphRecipient -SmtpAddresses $emailRecipients 

    $attachments = ConvertTo-IMicrosoftGraphAttachment -UploadDirectory C:\tmp

    $emailBody  = @{
        ContentType = 'html'
        Content = Get-Content 'C:\tmp\Email-Framework-master\boilerplate.html'    
    }

	$body += @{subject      = $emailSubject}
    $body += @{toRecipients = $toRecipients}    
    $body += @{attachments  = $attachments}
    $body += @{body         = $emailBody}

    $bodyParameter += @{'message'         = $body}
    $bodyParameter += @{'saveToSentItems' = $false}

    Send-MgUserMail -UserId $emailSender -BodyParameter $bodyParameter



}

function GraphCreateFolder
{
	$params = @{
	  DisplayName = "Graph Test Folder"
	  IsHidden = $false
	}

	New-MgUserMailFolder -UserId julian.stephan@qtjustdev.onmicrosoft.com -BodyParameter $params
}

function GraphCreateOnlineMeeting
{
	
$params = @{
	Subject = "Attend MEC Airlift"
	Body = @{
		ContentType = "HTML"
		Content = "September 13-14, 2022"
	}
	Start = @{
		DateTime = "2022-09-13T10:00:00"
		TimeZone = "Eastern Standard Time"
	}
	End = @{
		DateTime = "2022-09-14T15:00:00"
		TimeZone = "Eastern Standard Time"
	}
	Location = @{
		DisplayName = "Teams Virtual Sessions"
	}
	Attendees = @(
		@{
			EmailAddress = @{
				Address = "adelev@qtjusdev.onmicrosoft.com"
				Name = "Adele Vance"
			}
			Type = "required"
		}
	)
	IsOnlineMeeting = $true
	OnlineMeetingProvider = "teamsForBusiness"
}

# A UPN can also be used as -UserId.
$CalendarID = (Get-MgUserDefaultCalendar -UserId julian.stephan@qtjustdev.onmicrosoft.com).Id
New-MgUserCalendarEvent -UserId julian.stephan@qtjustdev.onmicrosoft.com -CalendarId $CalendarID -BodyParameter $params
}

function CreateAppAccessPolicy
{
	Connect-ExchangeOnline
	New-ApplicationAccessPolicy -Description "Allow EWS Access" -AppId $ConnectParameters.ClientId -AccessRight RestrictAccess -PolicyScopeGroupID 'ewsappaccesstest@qtjustdev.onmicrosoft.com'
}

function TestAppAccessPolicy
{
	{
		param (
			[switch]$Refresh
		)
		
		if (!$Refresh)
		{
			#Load EWS Binary
			try
			{
				$strEWSPackagePath = Get-Package -Name Exchange.WebServices.Managed.Api | ForEach-Object{ $_.Source.Substring(0, $_.Source.LastIndexOf('\')) }
				$strEWSDllPath = (Get-ChildItem $strEWSPackagePath -Recurse | Where-Object{ $_.Name -eq "Microsoft.Exchange.WebServices.dll" }).FullName
				Add-Type -Path $strEWSDllPath
			}
			catch
			{
				if ($global:blEnableLogging -eq $true) { Function_LogMessage -strMessage "Error loading EWS Binaries.`r`n $($_.Exception.Message)" }
				return "Error"
			}
			
			#Load Microsoft.Identity.Client Binary
			try
			{
				Add-Type -Path "C:\Program Files\WindowsPowerShell\Modules\MSAL.PS\4.37.0.0\Microsoft.Identity.Client.4.37.0\net45\Microsoft.Identity.Client.dll"
			}
			catch
			{
				if ($global:blEnableLogging -eq $true) { Function_LogMessage -strMessage "Error loading Microsoft.Identity.Client Binaries.`r`n $($_.Exception.Message)" }
				return "Error"
			}
		}
		
		#Obtain MSAL Token
		try
		{
			$objCertificate = (Get-ChildItem Cert:\CurrentUser\My) | Where-Object { $_.Thumbprint -match $ConnectParameters.CertificateThumbprint }
			$strScope = "https://outlook.office365.com/.default"
			if ($Refresh)
			{
				$objMSALToken = Get-MsalToken -ClientId $ConnectParameters.ClientId -TenantId $ConnectParameters.TenantID -ClientCertificate $objCertificate -Scopes $strScope -ForceRefresh
			}
			else
			{
				$objMSALToken = Get-MsalToken -ClientId $ConnectParameters.ClientId -TenantId $ConnectParameters.TenantID -ClientCertificate $objCertificate -Scopes $strScope
			}
			
		}
		catch
		{
			if ($global:blEnableLogging -eq $true) { Function_LogMessage -strMessage "Error getting MSAL token.`r`n $($_.Exception.Message)" }
			return "Error"
		}
		
		#Set EWS Credentials
		try
		{
			$objEWSService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016)
			$objEWSService.Url = "https://outlook.office365.com/EWS/Exchange.asmx"
			$objEWSService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $objMSALToken.AccessToken
		}
		catch
		{
			if ($global:blEnableLogging -eq $true) { Function_LogMessage -strMessage "Error connecting to EWS.`r`n $($_.Exception.Message)" }
			return "Error"
		}
		
		#Add Impersonation properties
		try
		{
			$objEWSService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, "")
			$objEWSService.HttpHeaders.Add("X-AnchorMailbox", "")
		}
		catch
		{
			if ($global:blEnableLogging -eq $true) { Function_LogMessage -strMessage "Error adding EWS Impersonation properties.`r`n $($_.Exception.Message)" }
			return "Error"
		}
		
		#Start timer for token refresh
		if ($global:objStopwatch.IsRunning -eq $true)
		{
			$global:objStopwatch.Restart()
		}
		else
		{
			$global:objStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
		}
		#Add Impersonation properties, test app access policy true
		try
		{
			$objEWSService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, "julian.stephan@qtjustdev.onmicrosoft.com")
			$objEWSService.GetDelegates("julian.stephan@qtjustdev.onmicrosoft.com",$true)

		}
		catch
		{
			if ($global:blEnableLogging -eq $true) { Function_LogMessage -strMessage "Error adding EWS Impersonation properties.`r`n $($_.Exception.Message)" }
			return "Error"
		}
		#Add Impersonation properties, test app access policy false
		try
		{
			$objEWSService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, "diegos@qtjustdev.onmicrosoft.com")
			$objEWSService.GetDelegates("diegos@qtjustdev.onmicrosoft.com",$true)
		}
		catch
		{
			if ($global:blEnableLogging -eq $true) { Function_LogMessage -strMessage "Error adding EWS Impersonation properties.`r`n $($_.Exception.Message)" }
			return "Error"
		}
		
}
}

function GraphSendMailAppAccess
{
	<#
    Ref:

    https://mikecrowley.us/2021/10/27/sending-email-with-send-mgusermail-microsoft-graph-powershell
    https://docs.microsoft.com/en-us/graph/api/user-sendmail
    https://docs.microsoft.com/en-us/powershell/module/microsoft.graph.users.actions/send-mgusermail

    #>

	$emailRecipients   = @(
        'diegos@qtjustdev.onmicrosoft.com'
        
    )
    $emailSender  = 'diegos@qtjustdev.onmicrosoft.com'

    $emailSubject = "Sample Email | " + (Get-Date -UFormat %e%b%Y)

	Function ConvertTo-IMicrosoftGraphRecipient {
        [cmdletbinding()]
        Param(
            [array]$SmtpAddresses        
        )
        foreach ($address in $SmtpAddresses) {
            @{
                emailAddress = @{address = $address}
            }    
        }    
    }

    Function ConvertTo-IMicrosoftGraphAttachment {
        [cmdletbinding()]
        Param(
            [string]$UploadDirectory        
        )
        $directoryContents = Get-ChildItem $UploadDirectory -Attributes !Directory -Recurse
        foreach ($file in $directoryContents) {
            $encodedAttachment = [convert]::ToBase64String((Get-Content $file.FullName -Encoding byte))
            @{
                "@odata.type"= "#microsoft.graph.fileAttachment"
                name = ($File.FullName -split '\\')[-1]
                contentBytes = $encodedAttachment
            }   
        }    
    }

	[array]$toRecipients = ConvertTo-IMicrosoftGraphRecipient -SmtpAddresses $emailRecipients 

    $attachments = ConvertTo-IMicrosoftGraphAttachment -UploadDirectory C:\tmp

    $emailBody  = @{
        ContentType = 'html'
        Content = Get-Content 'C:\tmp\Email-Framework-master\boilerplate.html'    
    }

	$body += @{subject      = $emailSubject}
    $body += @{toRecipients = $toRecipients}    
    $body += @{attachments  = $attachments}
    $body += @{body         = $emailBody}

    $bodyParameter += @{'message'         = $body}
    $bodyParameter += @{'saveToSentItems' = $false}

    Send-MgUserMail -UserId $emailSender -BodyParameter $bodyParameter



}


SetEWSConnection
Connect-MgGraph @ConnectParameters

