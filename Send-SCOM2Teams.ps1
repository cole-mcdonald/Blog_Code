﻿####################################
# Connecting SCOM to Microsoft Teams
####################################
# Cole McDonald - Beyond Impact
# Cole.McDonald@BeyondImpactLLC.com
#
# Portions by Tao Yang as indicated
#
# replace param defaults with your
# environmental values
####################################

param (
	[string]$AlertID = "<guid-for-alert-goes-here>",
	[string]$SubscriptionID = "<guid-for-subscription-goes-here>",
	[string]$SCOM_HOOK = "<incoming webhook connector URL from teams channel>",
	[string]$SCOM_MS = "scom_ms.domain.com",
	[string]$SCOM_URL = "https://scom.domain.com"
)

# Import the Operations Manager Module
Import-Module OperationsManager

# Grab the alert and subscription delay
$alert = Get-SCOMAlert -ComputerName $SCOM_MS | Where-Object -Property ID -eq $AlertID
$thisSub = Get-SCOMNotificationSubscription -ComputerName $SCOM_MS | Where-Object { $_.Id -eq $SubscriptionID }
$delay = $thisSub.Configuration.IdleMinutes

## The knowledge parts and toHTML functions are pulled from Tyson Paul's 2.3 version of Tao Yang's Enhanced Email 2.0
## ( https://blogs.msdn.microsoft.com/tysonpaul/2014/08/04/scom-enhanced-email-notification-script-version-2-1/ )

# Company Knowledge
# Functions to parse and sanitize the content
function trim-braces ($inString) {
	$instring = $inString.trimstart("{")
	$instring = $inString.trimend("}")
	return $inString
}
function fnMamlToHTML ($MAMLText) {
    $HTMLText = ""
    $HTMLText = $MAMLText -replace ('xmlns:maml="http://schemas.microsoft.com/maml/2004/10"')
    $HTMLText = $HTMLText -replace ("<maml:para>", "`n")
    $HTMLText = $HTMLText -replace ("maml:")
    $HTMLText = $HTMLText -replace ("</section>")
    $HTMLText = $HTMLText -replace ("<section>")
    $HTMLText = $HTMLText -replace ("<section >")
    $HTMLText = $HTMLText -replace ("<title>")
    $HTMLText = $HTMLText -replace ("</title>")
    $HTMLText = $HTMLText -replace ("<listitem>")
    $HTMLText = $HTMLText -replace ("</listitem>")
    $HTMLText = $HTMLText -replace ("`"", " ")
    return $HTMLText
}
function fnTrimHTML ($HTMLText) {
    $TrimmedText = ""
    $TrimmedText = $HTMLText -replace ("&lt;", "<")
    $TrimmedText = $TrimmedText -replace ("&gt;", ">")
    $TrimmedText = $TrimmedText -replace ("<html>")
    $TrimmedText = $TrimmedText -replace ("<HTML>")
    $TrimmedText = $TrimmedText -replace ("</html>")
    $TrimmedText = $TrimmedText -replace ("</HTML>")
    $TrimmedText = $TrimmedText -replace ("<body>")
    $TrimmedText = $TrimmedText -replace ("<BODY>")
    $TrimmedText = $TrimmedText -replace ("</body>")
    $TrimmedText = $TrimmedText -replace ("</BODY>")
    $TrimmedText = $TrimmedText -replace ("<h1>", "<h3>")
    $TrimmedText = $TrimmedText -replace ("</h1>", "</h3>")
    $TrimmedText = $TrimmedText -replace ("<h2>", "<h3>")
    $TrimmedText = $TrimmedText -replace ("</h2>", "</h3>")
    $TrimmedText = $TrimmedText -replace ("<H1>", "<h3>")
    $TrimmedText = $TrimmedText -replace ("</H1>", "</h3>")
    $TrimmedText = $TrimmedText -replace ("<H2>", "<h3>")
    $TrimmedText = $TrimmedText -replace ("</H2>", "</h3>")
    $TrimmedText = $TrimmedText -replace ("`"", " ")
    return $TrimmedText
}

# Strip the curly braces off of the incoming values
$alertID = trim-braces -inString $alertID
$SubscriptionID = trim-braces -inString $SubscriptionID

# Load the SDK
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.EnterpriseManagement.OperationsManager.Common") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.EnterpriseManagement.OperationsManager") | Out-Null

# Connect to Management Server
$MS_CONNECTION = New-Object Microsoft.EnterpriseManagement.ManagementGroupConnectionSettings($SCOM_MS)
$MGMT_GROUP = New-Object Microsoft.EnterpriseManagement.ManagementGroup($MS_CONNECTION)

#knowledge article / company knowledge
$KnowledgeArticles = $null
$KnowledgeArticles = $MGMT_GROUP.GetMonitoringKnowledgeArticles([string]$alert.ruleid)

if ($KnowledgeArticles -eq $null) {
    $message += "No resolutions were found for this alert.`n"
} else {
    #Convert Knowledge articles
    Foreach ($article in $KnowledgeArticles)
    {
     If ($article.Visible)
     {
      #Retrieve and format article content
      $MamlText = $null
      $HtmlText = $null
   
      if ($article.MamlContent -ne $null)
      {
       $MamlText = $article.MamlContent
       $knowledgeText = "$(fnMamlToHtml($MamlText))`n"
      }
   
      if ($article.HtmlContent -ne $null)
      {
       $HtmlText = $article.HtmlContent
       $knowledgeText = "$(fnTrimHTML($HtmlText))`n"
      }
     }
    }
}

# Empty message variable
$message = ""
# Title
$title = "SCOM Dispatch`n"
# Time and Path
$message += "$($alert.TimeRaised) - $($alert.MonitoringObjectPath)`n"
# Alert Name
$message += "$($alert.name)`n`n"
# Description
$message += "$($alert.Description)`n`n"
# Add the knowledge to the output
$message += $knowledgeText
# Alert ID
$message += "Alert: $($alert.id)`n"
# Monitor ID
$message += "Monitor: $($alert.ruleid)`n"
# Subscription Info
$message += "Subscription ID: $SubscriptionID - Delay: $delay`n"
# Add action Link
$message += "[ $SCOM_URL ]"

# Build content for the post
$content_JSON = @"
{
    "title" : "$title",
    "text" : "$message"
}
"@

# Send message to Teams Channel
Invoke-WebRequest -Uri $SCOM_HOOK -Method Post -Body $content_JSON -ContentType application/json