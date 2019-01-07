
## The Ultimate add everyone to everyone's calendar script                   ##
## ========================================================================= ##
## Last Update: 1/7/2019                                                     ##
## Richard Loucks                                                            ##
## ========================================================================= ##
##                                                                           ##
## Requires Exchange connection via powershell loaded w/ Admin Priv.         ##
## 1) Script will connect to the exchange server                             ##
## 2) Script will load required session script (for the get commands)        ##
##                                                                           ##
##                                                                           ##
##                                                                           ##
##                                                                           ##
## ========================================================================= ##

## ===[ Tell the errors they shall not pass ]================================ ##
#$ErrorActionPreference = "SilentlyContinue"

## ===[ See if the user even wants to connect to exchange ]================================ ##
$confirmation = Read-Host "Do you want to connect to your Exchange Server? (y/N)"
if ($confirmation -eq 'y') {

## ===[ Connecting to Exchange ]============================================= ##
Write-Host "|===================[ Connecting to Exchange - Enter username / password for admin user ]=====================|" -ForegroundColor yellow -BackgroundColor red
$UserCredential = Get-Credential
#Connect-MsolService -Credential $UserCredential
Set-Executionpolicy -ExecutionPolicy Unrestricted

## ===[ Create the session and download the commands ]======================= ##
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session
}

## ===[ Resetting the confirmation element for reuse ]================================ ##
$confirmation = ""

## ===[ Tell everyone whats up ]============================================= ##
Write-Host "|===================[ Connected Sucessfully! :) ]=====================|" -ForegroundColor yellow -BackgroundColor red

## ===[ Start a loop of all the mailboxes ]================================== ##
$mbxs = Get-Mailbox -ResultSize Unlimited |

## ===[ Get the primary SMTP for the element in the loop ]==================== ##
Select-Object PrimarySmtpAddress | 

## ===[ Clean up the email address to its useable ]=========================== ##
ForEach-Object {$_ -replace '@{PrimarySmtpAddress=' -replace '}' }

## ===[ Take our element from the last group and start a new loop ]=========== ##
foreach ($mbx in $mbxs) {

## ===[ Tell everyone whats up ]============================================= ##
Write-Host "Working on " $mbx -ForegroundColor yellow -BackgroundColor red
$confirmation = Read-Host "Do you want to continue with $mbx (y/N)"
if ($confirmation -eq 'y') {

## ===[ Get the current element in the new loop ]============================== ##
(Get-Mailbox).identity |


## ===[ Tell everyone whats up ]============================================= ##
foreach { echo "______________________________________________________________________" "$mbx being added to $_" "______________________________________________________________________"; 

## ===[ Setting a try statement to the add permission to manage the error ]================================ ##
    try { Add-MailboxFolderPermission $_":\calendar" -User $mbx -AccessRights reviewer -ErrorAction Stop |

## ===[ Tell the world of the good you have done ]================================ ##
    Write-Host "|==============[ Added Reviewer access to the calendar ]===============|`n" -ForegroundColor white -BackgroundColor green } 

## ===[ or you can confess your sins back to the screens ]================================ ##
    catch { Write-Host "|==============[ Skipped - Already has Access! ]===============|" -ForegroundColor yellow -BackgroundColor red }
    
 } ## ===[ Closing the foreach #2 ]================================ ##
  
} ## ===[ Closing the foreach #1 ]================================ ##
} ## ===[ Closing the confirmation ]================================ ##
