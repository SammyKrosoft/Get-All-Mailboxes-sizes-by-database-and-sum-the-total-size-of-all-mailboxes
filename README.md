# Script to get All  Mailboxes sizes by server or database, and sum the total size of all mailboxes
What I changed:

-	Added the “-DomainController” mandatory parameter (below example my Lab domain controller is “GC01”)
-	Added “-ServerNames” parameter (optional but default is E2016-01, E2019-01), add as many servers as you want, from 1 to 16 servers

```powershell
.\ ForMike-GetMailboxSizesAndSumTotalByServer.ps1 -DomainController GC01 -ServerNames E2016-01, E2019-01 
```

-	The script spits out 3 files:

o	Error log (if there are errors)

 ![image](https://github.com/SammyKrosoft/Get-All-Mailboxes-sizes-by-database-and-sum-the-total-size-of-all-mailboxes/assets/33433229/26bef11c-5c20-4439-b66c-0e2aee8b5845)


o	Mailbox sizes file

 ![image](https://github.com/SammyKrosoft/Get-All-Mailboxes-sizes-by-database-and-sum-the-total-size-of-all-mailboxes/assets/33433229/ed45f72d-99ca-4911-bc2e-7894eee58066)


o	Summary file with the total size for the last script run

 ![image](https://github.com/SammyKrosoft/Get-All-Mailboxes-sizes-by-database-and-sum-the-total-size-of-all-mailboxes/assets/33433229/f4e44588-9e50-4f7f-8d37-71f7fd362661)



The “package” has the exact same date_time suffix, and consists of 2 txt, and 1 csv:

 ![image](https://github.com/SammyKrosoft/Get-All-Mailboxes-sizes-by-database-and-sum-the-total-size-of-all-mailboxes/assets/33433229/f00c49d9-7e0b-4d6f-bb5c-96b01f7803d1)


-	Also, to avoid accuracy issues, I removed the KB/MB/GB calculation for each mailbox, and I put the size in Bytes only on the mailbox size output file, and only at the very end of the script after calculating the sum of all mailboxes, I convert the final result only to MB and GB (in the console output + in the summary file)

-	Finally, I put the code lines to calculate the mailbox sums for all the CSV files in the user’s DOCUMENTS directory:

```powershell
Write-Host "`n`n******************** Displaying statistics for mailboxes information on ALL CSVs in the current directorty ********************"

$import = get-childitem "$($env:Userprofile)\Documents\*.csv" | % {Import-Csv $_}

$TotalSumInBytes = ($import | Measure-Object -Property "MbxSize(In Bytes)" -Sum).sum
$TotalMBXCSV = $import.count
$TotalSumInB = [math]::Round($TotalSumInBytes, 0)
$TotalSumInMB = [math]::Round($TotalSumInBytes / 1MB, 2)
$TotalSumInGB = [math]::Round($TotalSumInBytes / 1GB, 2)


write-Host "`n`nTotal number of mailboxes on all CSVs: $TotalMBXCSV mailboxes" -BackgroundColor Yellow -ForegroundColor blue
Write-Host "`n`nTotal Sum of all Mailboxes in Bytes: $($TotalSumInB) Bytes" -ForegroundColor Green -BackgroundColor Blue
Write-Host "Total Sum of all Mailboxes in MB: $TotalSumInMB MB" -ForegroundColor Yellow -BackgroundColor Blue
Write-Host "Total Sum of all Mailboxes in GB: $TotalSumInGB GB" -ForegroundColor White -BackgroundColor Blue

Write-Host "`n`n*******************************************************************************************************************************"
```
