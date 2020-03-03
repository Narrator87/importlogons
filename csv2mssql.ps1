<#
    .SYNOPSIS
    ImportLogons
    Version: 0.04 04.08.2017
    
    © Anton Kosenko mail:Anton.Kosenko@gmail.com
    Licensed under the Apache License, Version 2.0

    .DESCRIPTION
    This script import users logins from csv to MSSQL DB
#>

# requires -version 3

# Declare Variable
    $LogFile = "./info.log"
    $Global:PrefferedExchangeServer = "ip or domain name"
    $Global:Now = (Get-Date)
    $SQLSRV="ip or domain name"
    $SQLDB="name DB"
    $SQLID="sql user"
    $SQLPASS="pass"
    $SQLTable = "table"
    $DstFolder = "path to folder"
    $CntFilesDomain=0
    $CntFiles=0
    $nf=0
    $nfa=0
    $CntDataDomain=0
    $CntData=0
    $CntSqlAft=0
    $CntSqlAll=0
    $CntError=0
    $CntDataAll=0
    $Data=$null
    $Mail_TextBody = ""
# Start writing log
    Start-Transcript -path $LogFile -append
# Create SQL-connection
    $sqlconn=New-Object System.Data.SqlClient.SqlConnection
    $sqlconn.ConnectionString = "Data Source=$SQLSRV;Initial Catalog=$SQLDB;User Id=$SQLID;Password=$SQLPASS;"
# Function mail sending
    function MailError
    {
        $PSEmailServer = "ip or domain name"
        $Mail_HTMLBody = "<head><style>table {border-collapse: collapse; padding: 2px;}table, td, th {border: 1px solid #ffffff;}</style></head>"
        $Mail_HTMLBody += "<body style='background:#ffffff'><font face='Courier New'; size='2' color=#cc0000>"		
        $Mail_Subject = "Import Logons Error Report $Now"
        $Mail_HTMLBody += "<center><h2>Import Logons Error Report</h2></center>"
        $Mail_HTMLBody += $Mail_TextBody
        $Mail_HTMLBody += "</font></body>"
        Send-MailMessage -From "service user" -To "admin mail" -Subject $Mail_Subject -Body $Mail_HTMLBody -BodyAsHtml -Encoding UTF8
    }
# Create mask for chacking date
    $Date=(Get-Date).ToString("yyyyMMdd")
# Check destination folder
    $ChkImPath=Test-Path $DstFolder
    if ($ChkImPath -match "True")
    {
        $ImPath=Get-ChildItem $DstFolder
        Write-host "`tConnect folder" -ForegroundColor Green
    }
    else {
        Write-host "Cannot connect to folder." -foregroundcolor Red
        $Mail_TextBody += "<center>Cannot connect to folder.</center>`n"
        MailError
        Stop-Transcript
        Exit
        }
# List all files in folder
    $CntFilesDomain=$ImPath.Count
# Run cycle processing files in destination folder with check to empty strings
    Foreach ($file in $ImPath)
    {
        if (($null -eq $file) -or ($file -eq "")) { continue }
    $nfa=$nfa+1
# Check files having current date
        $FileName=$file.Name
        if ($FileName -match $Date) {continue}
# Select fullpath on file
        [string]$FileCsvDomain=$file.FullName 
# Import from file and add data to variable
        [array]$DataDomain=import-csv $FileCsvDomain
        $Data=$Data+$DataDomain
# Calculate count string
        $CntDataDomain=$CntDataDomain+$dataDomain.count
        $nf=$nf+1
    }
# Check processing all files
    $CntFiles=$CntFilesDomain
    $CntData=$CntDataDomain
    if ($CntFiles -eq $nfa) 
    {
        Write-host "All files have been processed." -ForegroundColor Green
    }
    else {
        Write-host "Not all files have been processed." -foregroundcolor Red
        $Mail_TextBody += "<center>Not all files have been processed.</center>`n"
        MailError
        Stop-Transcript
        Exit
        }
# Check availability sql-server
    $sqlconn.open()
    if ($sqlconn.State -eq 'Close')
    {
        Write-host "Could not connect DB" -foregroundcolor Red
        $Mail_TextBody += "<center>Could not connect DB.</center>`n"
        MailError
        Stop-Transcript
        Exit
     }
# Count strings in table
    $SqlCmd=$sqlconn.CreateCommand()
    $SqlCmd.CommandText="SELECT COUNT(*) FROM $SQLTable"
        $objReader=$SqlCmd.ExecuteReader()
        while ($objReader.read()) {
            $CntSqlBef=$objReader.GetValue(0)
        }
         $objReader.close()
# Run line by line cycle processing array with check empty strings
    Foreach ($i in $Data)
    {
    if (($null -eq $i) -or ($i -eq "")) { continue }
# Get values from array
    $Timestamp=$i.Timestamp
    $Computer=$i.Computer
    $User=$i.User
    $TSClientName=$i.TSClientName
# Send import request in SQL table
    try {
    $SqlCmd.CommandText="INSERT INTO $SQLTable (LoginTime,ComputerName,UserName,DTUName) VALUES ('$Timestamp','$Computer','$User','$TSClientName')"
    $SqlCmd.ExecuteNonQuery() | Out-Null
    }
    catch {
        $CntError=$CntError+1
        $Mail_TextBody += "<center>Bad Row </br> $i </center>`n"
        Write-host "Bad Row" $i -ForegroundColor Magenta} 
    }
    Write-Host "`tImport done"  -foregroundcolor green
# Count importing strings in table
    $SqlCmd=$sqlconn.CreateCommand()
    $SqlCmd.CommandText="SELECT COUNT(*) FROM $SQLTable"
        $objReader=$SqlCmd.ExecuteReader()
        while ($objReader.read()) {
            $CntSqlAft=$objReader.GetValue(0)
        }
        $objReader.close()
    $CntSqlAll=$CntSqlAft-$CntSqlBef
# Show info about bad strings
    if ($CntError -eq 0)
        {   
            $CntDataAll=$CntData 
            Write-Host "All rows are correct"
        }
    else {
        Write-Host "`t$CntError bad rows found" -ForegroundColor Magenta
        $Mail_TextBody += "<center>$CntError bad rows found</center>`n"
        $CntDataAll=$Data.count-$CntError
    }
# Compare quantity strings before and after import
    if ($CntDataAll -eq $CntSqlAll)
    {
        Write-Host "`tImport are correct" -foregroundcolor green
    }
    else {
        Write-Host "`tImport are NOT correct" -foregroundcolor Red
        $Mail_TextBody += "<center>Import are NOT correct</center>`n"
        MailError
        Stop-Transcript
        Exit
          }
#  Run line by line cycle processing files from folder with check empty strings
    Foreach ($file in $ImPath)
    {
    if (($null -eq $file) -or ($file -eq "")) { continue }
# Check files having current date
    $FileName=$file.Name
    if ($FileName -match $Date) {continue}
# Select fullpath on files and delete them
    [string]$FileCsvDomain=$file.FullName 
    Remove-Item $FileCsvDomain
    }
# End work
    $sqlconn.Close()
# If $Mail_TextBody not empty - send mail
    If ( $Mail_TextBody.length -gt 0 )
            {
                MailError
            }
    Write-Host "######## End ###################" -foregroundcolor green
    Stop-Transcript