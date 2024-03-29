cls
# Very initial variables
# Define the output file path and name, with the date and time
$OutputFile = "MailboxSizes_" + (Get-Date -Format "yyyy-MM-dd_HH-mm-ss") + ".csv"
# Get the current user's documents folder, and store it in a variable
$DocumentsFolder = [Environment]::GetFolderPath("MyDocuments")

#Define array variable that will store all the mailbox objects with the size information
$MailboxSizeCollection = @()

# Get all your databases
$Databases = Get-MailboxDatabase
$msg = 'V2Ugd2lsbCBwYXNzIHRocm91Z2ggbWFueSBkYXRhYmFzZXMgISBIYW5nIG91dCwgTWlrZSAhIDotKQ==';$msg = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($msg));Write-Host "$msg" -ForegroundColor Yellow
Write-Host "Nb databases: $($Databases.count)`n`n" -ForegroundColor Yellow
# Initialize the counter for the progress bar
$CounterDB = 0
# Loop through each database
ForEach ($Database in $Databases) {
    $CounterDB++
    Write-Host "**** Database counter: $CounterDB ****" -ForegroundColor Green
    Write-Host "Database name: $Database" -ForegroundColor Red -BackgroundColor Blue
    
    $percentCompleteDB = ($CounterDB / $Databases.Count) * 100
    Write-Progress -Id 1 -Activity "Calculating Mailbox Sizes" -Status "Calculating Mailbox Sizes for Database: $($Database.Name)" -PercentComplete $percentCompleteDB

    # Get all mailboxes in the database
    # NOTE: I am initializing an array and using += to get mailboxes to cover the cases where we have just one mailbox on the database.
    # Otherwise if you are sure you have no cases where you just have 1 mailbox in the database 
    # you can just use $Mailboxes = Get-Mailbox -Database $Database.Identity  | Select Identity | Get-MailboxStatistics | Select DisplayName, PrimarySMTPADdress, TotalItemSize, TotalDeletedItemSize
    # NOTE2: for the cases where you have just 1 mailbox, you can also treat that case in a separate IF statement.
    # NOTE3: this is because if you have just 1 mailbox, the $Mailboxes variable is not an array by default. So we "force" it to be an array at the first place, and it will be a 1 item array in case
    # we have just 1 mailbox returned by the Get-Mailbox statement!
    $Mailboxes = @()
    $Mailboxes += Get-Mailbox -Database $Database.Identity -Filter {Name -notlike "*DiscoverySearchMailbox*"} -ResultSize Unlimited -ErrorAction SilentlyContinue | Select Identity, PrimarySMTPAddress 
    
    Write-Host "Nb Mailboxes: $($Mailboxes.count)" -ForegroundColor Red
    If ($Mailboxes.Count -gt 0){
        # Loop through each mailbox
        # Initialize the counter for the mailboxes progress bar
        $CounterMB = 0
        ForEach ($Mailbox in $Mailboxes) {
            $CounterMB++
            $percentCompleteMB = ($CounterMB / $Mailboxes.Count) * 100
            Write-Progress -ParentId 1 -Activity "Calculating Mailbox Sizes" -Status "Calculating Mailbox Sizes for Mailbox: $($Mailbox.DisplayName)" -PercentComplete $percentCompleteMB
            
            $MailboxStats = Get-MailboxStatistics -Identity $Mailbox.Identity | Select DisplayName, TotalItemSize, TotalDeletedItemSize

            $TotalItemSizeInKB = $MailboxStats.TotalItemSize.Value.ToKB() | Measure-Object -Sum
            $TotalItemSizeInMB = $MailboxStats.TotalItemSize.Value.ToMB() | Measure-object -sum
            $TotalItemSizeInGB = $MailboxStats.TotalItemSize.Value.ToGB() | Measure-Object -Sum
            $TotalDeletedItemSizeInKB = $MailboxStats.TotalDeletedItemSize.Value.ToKB() | Measure-Object -Sum
            $TotalDeletedItemSizeInMB = $MailboxStats.TotalDeletedItemSize.Value.ToMB() | Measure-object -sum
            $TotalDeletedItemSizeInGB = $MailboxStats.TotalDeletedItemSize.Value.ToGB() | Measure-Object -sum

            $MailboxTotalKB = $TotalItemSizeInKB.sum + $TotalDeletedItemSizeInKB.sum
            $MailboxTotalMB = $TotalItemSizeInMB.sum + $TotalDeletedItemSizeInMB.sum
            $MailboxTotalGB = $TotalItemSizeInGB.sum + $TotalDeletedItemSizeInGB.sum
    
            #Build the Array
            $Object = New-Object PSObject
            $Object | Add-Member NoteProperty -Name "DisplayName" -Value $Mailbox.DisplayName
            $Object | Add-Member NoteProperty -Name "PrimarySMTPAddress" -Value $Mailbox.PrimarySMTPAddress
            $Object | Add-Member NoteProperty -Name "MbxSize(In KB)" -Value $MailboxTotalKB
            $Object | Add-Member NoteProperty -Name "MbxSize(In MB)" -Value $MailboxTotalMB
            $Object | Add-Member NoteProperty -Name "MbxSize(In GB)" -Value $MailboxTotalGB
            $MailboxSizeCollection += $Object    
        }
    }
}

$MailboxSizeCollection | Export-Csv -Path "$DocumentsFolder\$OutputFile" -NoTypeInformation

$NumberOfMailboxes = $MailboxSizeCollection.Count
$TotalSizeOfMAilboxesKB = ($MailboxSizeCollection | Measure-Object -Property "MbxSize(In KB)" -Sum).Sum
$TotalSizeOfMailboxesMB = ($MailboxSizeCollection | Measure-Object -Property "MbxSize(In MB)" -Sum).Sum
$TotalSizeOfMailboxesGB = ($MailboxSizeCollection | Measure-Object -Property "MbxSize(In GB)" -Sum).Sum

Write-Host "`n`nTotal Number of Mailboxes: $NumberOfMailboxes`n`n" -ForegroundColor Yellow -BackgroundColor DarkBlue

Write-Host "Total Size of Mailboxes in KB: $TotalSizeOfMailboxesKB KB" -ForegroundColor Green -BackgroundColor Blue
Write-Host "Total Size of Mailboxes in MB: $TotalSizeOfMailboxesMB MB" -ForegroundColor Yellow -BackgroundColor Blue
Write-Host "Total Size of Mailboxes in GB: $TotalSizeOfMailboxesGB GB" -ForegroundColor White -BackgroundColor Blue

Write-Host "`n`nMailbox Sizes gathered successfully and saved to $DocumentsFolder\$OutputFile!" -ForegroundColor White -BackgroundColor DarkBlue


$msg = 'WW91IGRpZCB3ZWxsLCBNaWtlICEgOi0p';$msg = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($msg));Write-Host "`n $msg" -ForegroundColor Yellow
