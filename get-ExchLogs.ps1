<#
.SYNOPSIS
  Script to execute get-MessageTrackingLog cmdlet.

.DESCRIPTION
  Executes get-MessageTrackingLog cmdlet with specified parameters and credentials. Most parameters from get-MessageTrackingLog cmdlet.

.PARAMETER EventId
 The EventId parameter filters the message tracking log entries by the value of the EventId field. The EventId value classifies each message event. Example values include DSN, Defer, Deliver, Send, or Receive.

.PARAMETER InternalMessageId
 The InternalMessageId parameter filters the message tracking log entries by the value of the InternalMessageId field. The InternalMessageId value is a message identifier that's assigned by the Exchange server that's currently processing the message.

.PARAMETER MessageId
 The MessageId parameter filters the message tracking log entries by the value of the MessageId field. The value of MessageId corresponds to the value of the Message-Id: header field in the message. If the Message-ID header field is blank or doesn't exist, an arbitrary value is assigned.

.PARAMETER MessageSubject
 The MessageSubject parameter filters the message tracking log entries by the value of the message subject. The value of the MessageSubject parameter automatically supports partial matches without using wildcards or special characters. For example, if you specify the MessageSubject value sea, the results include messages with Seattle in the subject. By default, message subjects are stored in the message tracking logs.

.PARAMETER Recipients
 The Recipients parameter filters the message tracking log entries by the SMTP email address of the message recipients. Multiple recipients in a single message are logged in a single message tracking log entry. Unexpanded distribution group recipients are logged by using the group's SMTP email address.

.PARAMETER Reference
 The Reference parameter filters the message tracking log entries by the value of the Reference field. The Reference field contains additional information for specific types of events. For example, the Reference field value for a DSN message tracking entry contains the InternalMessageId value of the message that caused the DSN. For many types of events, the value of Reference is blank.

.PARAMETER ResultSize
 The ResultSize parameter specifies the maximum number of results to return. If you want to return all requests that match the query, use unlimited for the value of this parameter. The default value is 1000.

.PARAMETER Sender
 The Sender parameter filters the message tracking log entries by the sender's SMTP email address.

.PARAMETER Server
 The Server parameter specifies the Exchange server where you want to run this command. You can use any value that uniquely identifies the server.

.PARAMETER NetworkMessageId
 The NetworkMessageId parameter filters the message tracking log entries by the value of the NetworkMessageId field. This field contains a unique message ID value that persists across copies of the message that may be created due to bifurcation or distribution group expansion. 

.PARAMETER Source
 The Source parameter filters the message tracking log entries by the value of the Source field. These values indicate the transport component that's responsible for the message tracking event.
 
.PARAMETER TransportTrafficType
 The TransportTrafficType parameter filters the message tracking log entries by the value of the TransportTrafficType field. However, this field isn't interesting for on-premises Exchange organizations.

.PARAMETER Start
 The Start parameter specifies the start date and time of the date range

.PARAMETER End
 The End parameter specifies the end date and time of the date range

.PARAMETER HideHealthMessages
 The HideHealthMessages specifies that messages where sender match 'maildeliveryprobe' or 'HealthMailbox', or recipient match 'HealthMailbox', or messages with missed messsageid will be filtered from output.

.PARAMETER Powershellpath
 Path to connect to powershell wirtual directory on exchange server.

.PARAMETER Username
 Username credential parameter.

.PARAMETER Password
 Password credential parameter.

.OUTPUTS
  Outputs JSON object in UTF8 encoding.

.NOTES
  Version:        1.0
  Author:         Admin
  Creation Date:  12.04.2019
  Purpose/Change: Initial script development
  
.EXAMPLE
  get-ExchLogs.ps1 -Powershellpath 'https://<EXCHANGE>/powershell' -start 11.04.2019 -end 12.04.2019 
#>


param (
[string]$EventId,
[string]$InternalMessageId,
[string]$MessageId,
[string]$MessageSubject,
[string]$Recipients,
[string]$Reference,
[string]$ResultSize,
[string]$Sender,
[string]$Server,
[string]$NetworkMessageId,
[string]$Source,
[string]$TransportTrafficType,
[datetime]$Start,
[datetime]$End,
[string]$HideHealthMessages,
[string]$Powershellpath,
[string]$Username,
[string]$Password
)

# ENCODING (force script output to UTF8 encoding)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# FUNCTION TO CONVERT CP866
function cp866-to-utf8 ([string]$str){
$byte = [System.Text.Encoding]::GetEncoding("cp866").getbytes($str)
$utf8 = [System.Text.Encoding]::UTF8.GetString($byte)
return $utf8
}

# FUNCTION TO ADD QUOTES
function add-quotes ([string]$str){
$strwithquotes = "'"+$str+"'"
return $strwithquotes
}

# PASS SCRIPT PARAMETERS WITH $PSBoundParameters
$ScriptParams =[ordered]@{}
$PSBoundParameters.Keys | foreach { 
switch ($_) {
{$_ -eq "HideHealthMessages"}{return} # exclude unnecessary pararmeters
{$_ -eq "Powershellpath"}{return} # exclude unnecessary pararmeters
{$_ -eq "Username"}{return} # exclude unnecessary pararmeters
{$_ -eq "Password"}{return} # exclude unnecessary pararmeters       
{$_ -eq "Start"} {$ScriptParams.Add($_,(add-quotes($PSBoundParameters.item($_))));return} #add quotes to date string
{$_ -eq "End"}   {$ScriptParams.Add($_,(add-quotes($PSBoundParameters.item($_))));return} #add quotes to date string
{$_ -eq "MessageSubject"} {$ScriptParams.Add($_,(add-quotes(cp866-to-utf8($PSBoundParameters.item($_)))));return} #add quotes to string and convert to utf8
default {$ScriptParams.Add($_,$PSBoundParameters.item($_))}
}
}

$ScriptParams.Add('WarningAction','silentlyContinue') #supress warnings

# CREDENTIALS
if ($Username -and $Password) {
$Credentials = [System.Management.Automation.PSCredential]::new($Username, ($Password | ConvertTo-SecureString -asPlainText -Force))
}

# SKIP CERTIFICATE CHECK SESSION OPTIONS
$Sessionoption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck

# TRY TO MAKE NEW PSSESSION
try{
if ($Credentials) {
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $Powershellpath -Credential $Credentials -Authentication Basic -SessionOption $Sessionoption -ErrorAction Stop
}
else {
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $Powershellpath -SessionOption $Sessionoption -ErrorAction Stop
}
}
catch{
throw $_
break
}

# CREATING SCRIPTBLOCK WITH $Command AND $ScriptParams
$Command = 'Get-MessageTrackingLog'
$ScriptBlock = [Scriptblock]::Create("$Command $(&{$args} @ScriptParams)")

# EXECUTING SCRIPTBLOCK IN REMOTE SESSION
try {
$Messages = Invoke-Command -Session $Session -ScriptBlock $ScriptBlock
}
catch{
throw $_
break
}

# DESTROY REMOTE SEESION
Remove-PSSession $Session

# FORMATING OUTPUT OBJECT
$Output = $Messages | Select-Object @{Name = 'MessageId'; Expression = {
if ($_.MessageId) {$_.MessageId.trim().Substring(1,($_.MessageId.Length-2))} else {"Missed"}}},Timestamp, MessageSubject, ClientHostname, ServerHostname, Sender, 
@{Name = 'Recipients'; Expression = {if($_.Recipients) {[string]($_.Recipients)}}},Directionality,Source,EventId,@{Name='Size'; Expression ={
if ($_.TotalBytes -gt 1GB) {“$([math]::Round(($_.TotalBytes / 1GB),0)) GB”;return}
if ($_.TotalBytes -gt 1MB) {“$([math]::Round(($_.TotalBytes / 1MB),0)) MB”;return}
if ($_.TotalBytes -gt 1KB) {“$([math]::Round(($_.TotalBytes / 1KB),0)) KB”;return}
else {“$([math]::Round(($_.TotalBytes / 1KB),0)) Byte”}}} #converting bytes

# FILTER SERVICE MESSAGES
If ($HideHealthMessages -eq 'Yes') {
$Output = $Output | Where-Object {($_.Sender -notmatch 'maildeliveryprobe') -and ($_.Sender -notmatch 'HealthMailbox') -and ($_.Recipients -notmatch 'HealthMailbox') -and ($_.MessageId -ne 'Missed')}
}

# OUTPUT TO JSON
$Output = $Output | ConvertTo-Json

# RETURN
$Output