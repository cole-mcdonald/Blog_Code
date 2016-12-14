####################################
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
    $AlertID = "<guid-for-alert-goes-here>",
    $SubscriptionID = "<guid-for-subscription-goes-here>",
    $SCOM_HOOK = "<incoming webhook connector URL from teams channel>",
    $SCOM_MS = "scom_ms.domain.com",
    $SCOM_URL = "https://scom.domain.com"
)

# Import the Operations Manager Module
Import-Module OperationsManager

# Grab the alert and subscription delay
$alert = Get-SCOMAlert -ComputerName $SCOM_MS | Where-Object -Property ID -eq $AlertID
$thisSub = Get-SCOMNotificationSubscription -ComputerName $SCOM_MS | Where-Object { $_.Id -eq $SubscriptionID }
$delay = $thisSub.Configuration.IdleMinutes

## The knowledge parts are pulled from Tyson Paul's 2.3 version of Tao Yang's Enhanced Email 2.0
## ( https://blogs.msdn.microsoft.com/tysonpaul/2014/08/04/scom-enhanced-email-notification-script-version-2-1/ )

# Company Knowledge
# Functions to parse the content
# Needed to remove quotes to get the JSON to pass correctly as well.
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
$title = "SCOM Dispatch *** TEST ***`n"
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