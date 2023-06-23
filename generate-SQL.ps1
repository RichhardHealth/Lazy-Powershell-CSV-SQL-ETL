# ==========================================================================================
# Description: A script to look in a folder and for each CSV import in to databse as a table
# Created: 13/05/2017 
# Created by: Richard Holley
# Modifications:
# 13/05/2016 - Version 1 - RH - Original script
# ==========================================================================================

# User Parameters (change these)
# ==========================================================================================
# Set the folder path where the CSV files are located
$folderPath = "C:\Path\To\CSV\Folder" #No trailing \ for folder path

# Set the SQL server and database connection details
$serverName = "WSSQL0193\SQL01"
$databaseName = "Test"
$tableNamePrefix = "CSVTable_"
$tableNamePostfix =""



# System Parameters (change these)
# ==========================================================================================

$logPath="c:\Scripts\Test\Log\1.txt"






#Functions
# ==========================================================================================
#Write messages to a log file
function Log($message)
{
	$date = Get-Date
    $lineMessage =$date.ToString("yyyy-MM-dd HH:mm:ss")+": "+$message
	Add-Content -Path $logPath -Value $lineMessage
}



#function formats time from ms in to human readable time
function FormatElapsedTime($ts) 
{
    $elapsedTime = ""
    if($ts.Hours -gt 0)
	{
		$elapsedTime = [string]::Format( "{0:00} hr{1:00} min. {2:00}.{3:00} sec.",$ts.Hours, $ts.Minutes, $ts.Seconds, $ts.Milliseconds / 10 );
	}
	elseif($ts.Minutes -gt 0)
	{
		$elapsedTime = [string]::Format( "{0:00} min. {1:00}.{2:00} sec.", $ts.Minutes, $ts.Seconds, $ts.Milliseconds / 10 );
	}	
    elseif ( $ts.Seconds -gt 0 )
    {
         $elapsedTime = [string]::Format( "{0:00}.{1:00} sec.", $ts.Seconds, $ts.Milliseconds / 10 );
    }
    elseif ($ts.Milliseconds -gt 0 )
    {
        $elapsedTime = [string]::Format("{0:00} ms.", $ts.Milliseconds);
    }
    else
    {
        $elapsedTime = [string]::Format("{0} ms", $ts.TotalMilliseconds);
    }
    return $elapsedTime

}




#Start of script actually doing things
# ==========================================================================================
cls
#Get start time
$date = Get-Date
Write-Host "Started script:" $date.ToString("yyyy-MM-dd HH:mm:ss")". Searching for CSV in folder: "$Input_Folder

#Get start time for elapsed time
$sw = [Diagnostics.Stopwatch]::StartNew()



# Create a SQL connection string
$connectionString = "Server=$serverName;Database=$databaseName;Integrated Security=True"

# Get the list of CSV files in the folder
$csvFiles = Get-ChildItem -Path $folderPath -Filter "*.csv"

# Loop through each CSV file
foreach ($csvFile in $csvFiles) {
    # Create a table name based on the CSV file name
    $tableName = $tableNamePrefix + $csvFile.BaseName + $tableNamePostfix

    # Read the CSV file and get the column names
    $csvData = Import-Csv -Path $csvFile.FullName
    $columnNames = $csvData | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name

    # Create the SQL CREATE TABLE statement
    $createTableQuery = "IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = '$tableName')
    BEGIN  
    CREATE TABLE $tableName ("

    foreach ($columnName in $columnNames) {
        $createTableQuery += "`n$columnName VARCHAR(MAX),"
    }

    $createTableQuery = $createTableQuery.TrimEnd(",") + ")END"

    # Create the SQL INSERT INTO statement
    $insertIntoQuery = "INSERT INTO $tableName VALUES "

    foreach ($row in $csvData) {
        $values = $columnNames | ForEach-Object { "'" + $row.$_ + "'" }
        $insertIntoQuery += "`n(" + $values -join "," + "),"
    }

    $insertIntoQuery = $insertIntoQuery.TrimEnd(",") + ";"

    # Create a new SQL connection and execute the queries
    $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    $connection.Open()

    $command = $connection.CreateCommand()
    $command.CommandText = $createTableQuery
    $command.ExecuteNonQuery()

    $command.CommandText = $insertIntoQuery
    $command.ExecuteNonQuery()

    $connection.Close()

    Write-Host "Imported data from $($csvFile.Name) into table $tableName"
}


#Get stop time for elapsed time
$sw.Stop()
$time = $sw.Elapsed
$formatTime = FormatElapsedTime $time

#Get stop time1
$date = Get-Date
Write-Host "Finished script:" $date.ToString("yyyy-MM-dd HH:mm:ss") ". Script took: "$formatTime    

#End of Script
#===========================================================

