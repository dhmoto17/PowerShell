<#
.SYNOPSIS

Script to update safe senders list with focus on OWA users specifically

.DESCRIPTION

Use this command to create a report of safe senders and update them if necessary. Use comments if you want to report, test and not action

.EXAMPLE
.\Safe_Sender_Check_Update.ps1
Runs with defaults

.EXAMPLE
.\Safe_Sender_Check_Update.ps1 ALL "news@dom1","games@dom2" C:\Data\rep.txt C:\Data\reperr.txt C:\Data\repact.txt
Runs with appending "ALL" to files, for safe senders "news@dom1" & "games@dom2", reports to rep.txt, errors to reperr.txt, action report to repact.txt

.EXAMPLE
.\Safe_Sender_Check_Update.ps1 -safesenders "news@dom1","games@dom2"
Runs with defaults for safe senders "news@dom1" & "games@dom2"
#>

# define parameters
[CmdletBinding()]
Param(
  [Parameter(Mandatory=$False,Position=1,HelpMessage="Set some chars to tag the file name with")]
  [string]$denom,
  [Parameter(Mandatory=$False,Position=2,HelpMessage="Set some safe senders, in speech marks coma seperated if more than one")]
  [string[]]$safesenders,
  [Parameter(Mandatory=$False,Position=3,HelpMessage="Set a report file path and name")]
  [string]$outfilearr,
  [Parameter(Mandatory=$False,Position=4,HelpMessage="Set a report error file path and name")]
  [string]$erroutfile,
  [Parameter(Mandatory=$False,Position=5,HelpMessage="Set an action performed file path and name")]
  [string]$actoutfile
)

#-- Time script --
$StopWatch = [system.diagnostics.stopwatch]::startNew()

# -- edit this block to suit --
# $denom is just a label to append to the file name, see file paths further down
If (!$denom) { $denom = "ALLMBX" }
# $safesenders can be multiple entriess, comma seperated eg "news@dom1","games@dom2" - it scans for the safe senders
If (!$safesenders) { $safesenders = "news@rewardgateway.co.uk" }
# these two are the file paths for data and errors
If (!$outfilearr) {$outfilearr = "\\$fs\Software\PowerShell\Exchange\Results\safesenders_$((get-date).tostring("yyyyMMddHHmmss"))_$($denom).txt"}
If (!$erroutfile) {$erroutfile = "\\$fs\Software\PowerShell\Exchange\Results\safesenders_$((get-date).tostring("yyyyMMddHHmmss"))_$($denom)_err.txt"}
If (!$actoutfile) {$actoutfile = "\\$fs\Software\PowerShell\Exchange\Results\safesenders_$((get-date).tostring("yyyyMMddHHmmss"))_$($denom)_act.txt"}
# --
write-host "denom: $denom"
write-host "Sites: $safesenders"
write-host "Report file: $outfilearr"
write-host "Error file: $erroutfile"
write-host "Action file: $actoutfile"

# -- script ini DON'T EDIT THIS BIT --
$list = New-Object System.Collections.Generic.List[System.String]
$varerr = $null
$varaction = $null
$Error.clear()
$quote = [char]34

# start work

# this line will set action against all mailboxes in an organisation, adjust to taste if that is not required
$allmbxs = Get-Mailbox -resultsize unlimited | select name,alias
#$allmbxs = Get-Mailbox "100016" -resultsize unlimited | select name,alias

foreach($mbx in $allmbxs) {
  $safelist = Get-Mailbox $($mbx).alias | Get-MailboxJunkEmailConfiguration | select Identity,TrustedSendersAndDomains
  
  Foreach($safesender in $safesenders){
    $Error.clear()
    # $safelist.TrustedSendersAndDomains -like $safesender
    if($safelist.TrustedSendersAndDomains -ccontains $safesender) {
      $list.Add("$($quote)$($mbx.name)$($quote),$($quote)$($mbx.alias)$($quote),$($quote)$($safesender)$($quote),$($quote)Safesender exists - would not be added again$($quote)")
      }
    else {
      $list.Add("$($quote)$($mbx.name)$($quote),$($quote)$($mbx.alias)$($quote),$($quote)$($safesender)$($quote),$($quote)Safesender does not exist - would be added$($quote)")
      try {
        $varaction += "Setting safesender $($quote)$($safesender)$($quote) for user $($quote)$($mbx.name)$($quote)" + "`r`n"
        Get-Mailbox $($mbx).alias | Set-MailboxJunkEmailConfiguration -TrustedSendersAndDomains @{Add=$safesender} -ErrorAction Stop
        }
      Catch {
        # Run this if a terminating error occurred in the Try block
        $varerr += ($Error[0]).exception.message + "`r`n"
        $varaction += ($Error[0]).exception.message + "`r`n"
        }
      $varaction += "`r`n"
      }
    }
  }

$array = $list.ToArray()
$array | out-file $outfilearr
if($varerr) { $varerr | out-file $erroutfile }
if($varaction) { $varaction | out-file $actoutfile }

#-- Stop clock --
$StopWatch.Stop()
write-host "The script took $($StopWatch.elapsed.minutes) minutes $($StopWatch.elapsed.seconds) seconds"

