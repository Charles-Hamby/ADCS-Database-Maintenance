#REQUIRES -Version 2.0

<#  
.SYNOPSIS  
Delete Failed Requests and Expired Certificates from the CA Database.

.DESCRIPTION
Delete Failed Requests and Expired Certificates from database by using the -Failed and/or -ExpiredDays/ExpiredDate parameter.

.NOTES
Version        : 1.6
File Name      : delete-failed-requests-V1.6.ps1
Author         : Charles Hamby
Prerequisite   : PowerShell V2 over Vista and upper.

.LINK  

.EXAMPLE  
	.\delete-failed-requests-V1.6.ps1
	
	Provides information of CA Database, if CA is running and calculates the total number of failed requests.
	
.EXAMPLE
	.\delete-failed-requests-V1.6.ps1 -Defrag
	
	Defrags the CA Database.

.EXAMPLE
	.\delete-failed-requests-V1.6.ps1 -Failed
	
	Will delete all failed requests from CA Database.

.EXAMPLE    
	.\delete-failed-requests-V1.6.ps1 -ExpiredDate 06/20/2020
	
	Will delete all expired certificates before 06/20/2020.

.EXAMPLE
	.\delete-failed-requests-V1.6.ps1 -ExpiredDate 06/20/2020 -Failed
	
	Will delete all expired certificates before 06/20/2020 and all failed requests.

.EXAMPLE    
	.\delete-failed-requests-V1.6.ps1 -ExpiredDays 90
	
	Will delete all expired certificates greater than 90 days before current date.

.EXAMPLE
	.\delete-failed-requests-V1.6.ps1 -ExpiredDays 90 -Failed
	
	Will delete all expired certificates greater than 90 days before current date and all failed requests.

#>
	
	
	##############################################################################
	#
	# Author: Charles Hamby
	# Date: 6/01/2019 (v1.1) First release
	# Date: 10/22/2019 (v1.2) Added -Failed and -Expired parameter
	# Date: 11/1/2019 (V1.3) Clean up
	# Date: 11/6/2019 (V1.4) Added Defrag parameter and Logging
	# Date: 10/06/2021 (V1.5) Added -ExpiredDays parameter to delete expired following certain amount of days.
    	# Date: 11/29/2021 (V1.6) Added Event Logging in EventVwr & Log each expired cert that is going to be deleted at runtime
    	# Current Version: 1.6
	#
	# Usage: Removing all failed and expired requests and defragmentation of the Database
	#
	##############################################################################

	# Commandline parameters
	[CmdletBinding()]
	param (

		[Parameter(Mandatory=$false)][String]$ExpiredDate="",
        [Parameter(Mandatory=$false)][String]$ExpiredDays="",
		[Switch]$Failed=$false,
		[Switch]$Defrag=$false,
		[Parameter(Mandatory=$false )][String]$LogDir="$PWD\CA-DB-Maintenance",
		[Parameter(Mandatory=$false )][String]$LogFile="CA-DB-Maintenance.log"
	);
	
    	#Preload Required Modules
    	Import-Module Logging

	# Clear screen
	Clear-Host;
	
	$certutil = "$ENV:SystemRoot\System32\certutil.exe";
	
	# Start stopwatch
	$totalTime = New-Object -TypeName System.Diagnostics.Stopwatch;
	$totalTime.Start();

	# Suppress warnings
	$WarningPreference = "SilentlyContinue";

	# Suppress errors
	$ErrorActionPreference = "SilentlyContinue";
	
	# Function LogWrite
	Function LogWrite {
	
	Param ([string]$LogString);
		
	# Check if CA log folder exists	
	if(!(Test-Path -Path $LogDir)) {
	
		# Create folder if not exist
		New-Item -ItemType directory -Path $LogDir | Out-Null;
		
	}; # End of CA log
		
	#$Date = Get-Date -Format u;
	$Date = Get-Date -Format o | foreach {$_ -replace ":", "."};
	Add-content "$LogDir\$LogFile" -Value "$Date| $LogString";
	
	}; # End Function LogWrite

    	#Start Logging
    	Start-Log -EventId 4000 -EventLog "CA Database Maintenance" -ScriptName "delete-failed-requests-V1.6.ps1"

	# Checking the CA database status
	Write-Host "";
	Write-Host "=== CA Database Maintenance ===" -ForegroundColor "Green";
	LogWrite "=== Start CA Database Maintenance ===";
    	Write-Log -Level Information -EventId 4001 -Message @{
        Message = "Start CA Database Maintenance"
    
    	};
	Write-Host "";
	$srvName = "Active Directory Certificate Services";
	Write-Host "Checking the $srvName status..." -ForegroundColor "Yellow";
	LogWrite "Checking the $srvName status...";
    	Write-Log -Level Information -EventId 4002 -Message @{
        Message = "Checking the $srvName status..."
    	};

	# 5 second wait
	Start-Sleep -s 5;

	$servicePrior = Get-Service $srvName;
	$servicePrior = $servicePrior.Status;
	Write-Host "";
	Write-Host "$srvName is" $servicePrior -ForegroundColor "Yellow";
	LogWrite "$srvName is $servicePrior";

	# Check is Service is stopped.
	if ($servicePrior.status -eq "Stopped") {

	# Starting the CA
	Start-Service certsvc;
	Write-Host "Starting the $srvName and wait for 60 seconds to make sure it has been started." -ForegroundColor "Yellow";
	LogWrite "Starting the $srvName and wait for 60 seconds to make sure it has been started.";
        Write-Log -Level Information -EventId 4003 -Message @{
        Message = "Starting the $srvName and wait for 60 seconds to make sure it has been started."
        };
	Write-Host "";
		
	# 15 second wait
	Start-Sleep -s 60;
		
	# Check status
	$srvName = "Active Directory Certificate Services";
	$servicePrior = Get-Service $srvName;
	$dbstatus = $servicePrior.status;
	Write-Host "";
	Write-Host "$srvName is now $dbstatus." -ForegroundColor "Yellow";
	LogWrite "$srvName is now $dbstatus";
        Write-Log -Level Information -EventId 4004 -Message @{
        Message = "$srvName is now $dbstatus"
        };
	Write-Host "";
		
	};
	
	# Get CAname
	$CAName = Get-ItemProperty -Path HKLM:\System\CurrentControlSet\Services\CertSvc\Configuration -Name Active | Select-Object -ExpandProperty Active;
	
	# Get CA database location
	$CADatabaseLocation = Get-ItemProperty -Path HKLM:\System\CurrentControlSet\Services\CertSvc\Configuration -Name DBDirectory | Select-Object -ExpandProperty DBDirectory;
	Write-Host "";
	Write-Host "";
	Write-Host "# CA Database Information #" -ForegroundColor "Green";
	Write-Host "";
	Write-Host "Database location: $CADataBaseLocation" -ForegroundColor "Yellow";
	LogWrite "Database location: $CADataBaseLocation";
	
	# Database name
	$CADataBaseName = "$CAName.edb";
	Write-Host "Database name: $CADataBaseName" -ForegroundColor "Yellow";
	LogWrite "Database name: $CADataBaseName";
	
	# CA Database size before
	$CADatabaseSize = (Get-ChildItem -Path "$CADataBaseLocation" -Recurse -Filter "$CADataBaseName").length/1MB;
	Write-Host "Pre-maintenance Database size: $CADatabaseSize MB" -ForegroundColor "Yellow";
	LogWrite "Pre-maintenance Database size: $CADatabaseSize MB";
    	Write-Log -Level Information -EventId 4005 -Message @{
        Message = "Pre-maintenance Database size: $CADatabaseSize MB"
        };

	# 2 second wait
	Start-Sleep -s 2;

	# Checking for failed request and count them
	Write-Host "";
	Write-Host "# Checking for failed requests... #" -ForegroundColor "Green";
	LogWrite "Checking for failed requests...";
    	Write-Log -Level Information -EventId 4010 -Message @{
        Message = "Checking for failed requests..."
        };
	Write-Host "";
	
	# Retrieving Failed Requests from the database 
	$FailedRequest = .$certutil '-silent' '-view' '-out' ""RequestID"" 'LogFail';

	# Select the line which starts with "Issued"
	$FailedRequest = $FailedRequest | select-string "Issued";
	
	# Retrieving failed requests from the database 
	$Count = .$certutil '-view' '-out' "RequestID" "LogFail" 2>&1;

	# Select the line which starts with "Maximum Row Index"
	$Count = $Count | Select-String "Maximum Row Index";
	$Count = ($Count.tostring()).Split(":");
	$Count = ($Count[1]).Split("Maximum Row Index");
	Write-Host "Total of Failed Requests:$Count" -ForegroundColor "Yellow";
	LogWrite "Total of Failed Requests:$Count";
    	Write-Log -Level Information -EventId 4011 -Message @{
        Message = "Total of Failed Requests:$Count"
        };
	Write-Host "";

	# Check if $Failed is true, delete failed requests
	if ($Failed -eq $true) {
		
		Write-Host "";
		Write-Host "# Deleting failed requests #" -ForegroundColor "Green";
		Write-Log -Level Information -EventId 4012 -Message @{
		Message = "Deleting failed requests"
		};
		Write-Host "";

		# Loop to delete failed requests
		foreach ($f in $FailedRequest) {

			# Put every line in to a String and split at ":"
			$j=($f.tostring()).Split(":");

			# Split the second line from the output of the split and split at "("
			$i=($j[1]).Split("(");

			# Remove any empty spaces from the first line
			$k=$i[0].Trim();

			# Convert Hex to Decimal
			$dec=[Convert]::ToInt32($k,16);

			# Delete the rows and surpress the output
			.$certutil '-deleterow' $dec | Out-Null;

			# Create a custom std output
			if ($i -match "0x") {

				Write-Host "Failed request with RequestID $dec deleted" -ForegroundColor "Yellow";

			} 
			else {

				Write-Host "";
				Write-Host "There are no failed requests found in the database." -ForegroundColor "Red";
				LogWrite "There are no failed requests found in the database.";
				Write-Log -Level Information -EventId 4013 -Message @{
				    Message = "There are no failed requests found in the database."
				};
			
			Write-Host "";

			};

		};

	}; # End loop to delete failed requests

	# Check if $ExpiredDate has a value or not
	if ($ExpiredDate -ne "" ) {
	
		$DelExpired = .$certutil '-deleterow' $ExpiredDate Cert;
		
		if ($DelExpired -match "Rows deleted: 0" ) {
		
			Write-Host "";
			Write-Host "No expired certificates found." -ForegroundColor "Yellow";
			LogWrite "No expired certificates found.";
            		Write-Log -Level Information -EventId 4020 -Message @{
                	Message = "No expired certificates found."
           		 };
			Write-Host "";
			
		} else {
			
			$DelExpired = $DelExpired[0].Split(":")[1].Trim();
			Write-Host "";
			Write-Host "# Deleting expired certificates #" -ForegroundColor "Green";
            		Write-Log -Level Information -EventId 4021 -Message @{
                		Message = "Deleting expired certificates"
            		};
			Write-Host "";
			Write-Host "Total of expired certificates deleted: $DelExpired" -ForegroundColor "Yellow";
			LogWrite "Total of expired certificates deleted: $DelExpired";
            		Write-Log -Level Information -EventId 4022 -Message @{
                		Message = "Total of expired certificates deleted: $DelExpired"
            		};
			Write-Host "";
			
		};
		
	}; # End of check $ExpiredDate


	# Check if $ExpiredDays has a value or not
	if ($ExpiredDays -ne "" ) {
	    
        $ExpiredDays = (Get-Date).AddDays(-$ExpiredDays)
        $ExpiredDays = Get-Date $ExpiredDays -Format "MM/dd/yyyy"

        $ExpiredCerts = @(certutil -view -restrict "NotAfter<=$ExpiredDays,Disposition=20" -out "CommonName,SerialNumber,NotAfter,CertificateTemplate" csv)
		
        if ($ExpiredCerts.count -eq 1) {
		
		Write-Host "";
		Write-Host "No expired certificates found." -ForegroundColor "Yellow";
		LogWrite "No expired certificates found.";
            	Write-Log -Level Information -EventId 4020 -Message @{
                	Message = "No expired certificates found."
            	};
	
        }Else { 
        
            Foreach ($c in $ExpiredCerts) {

                If ($c -match '"Issued Common Name","Serial Number","Certificate Expiration Date","Certificate Template"') {
                    
                     Continue;

                }else {
                    
                    $ccommonname = $c.Split(",")[0];
                    $cserialnumber = $c.Split(",")[1];
                    $cexpirationdate = $c.Split(",")[2];
                    $ccertificatetemplate = $c.Split(",")[3];
                    LogWrite "Deleting expired certificate - $cserialnumber,$ccommonname,$cexpirationdate,$ccertificatetemplate";
                    Write-Log -Level Information -EventId 4023 -Message @{
                        Message = "Deleting expired certificate - $cserialnumber"
                        Add = "'$ccommonname','$cserialnumber','$cexpirationdate','$ccertificatetemplate'"
                    };
                    
                };
           
            
            };

            $DelExpired = .$certutil '-deleterow' $ExpiredDays Cert;
		
		if ($DelExpired -match "Rows deleted: 0" ) {
		
			Write-Host "";
			Write-Host "No expired certificates found." -ForegroundColor "Yellow";
			LogWrite "No expired certificates found.";
                    	Write-Log -Level Information -EventId 4020 -Message @{
                       		Message = "No expired certificates found."
                    	};
			Write-Host "";

                } else {
			
			$DelExpired = $DelExpired[0].Split(":")[1].Trim();
			Write-Host "";
			Write-Host "# Deleting expired certificates #" -ForegroundColor "Green";
                    	Write-Log -Level Information -EventId 4021 -Message @{
                        	Message = "Deleting expired certificates"
                    	};
			
			Write-Host "";
			Write-Host "Total of expired certificates deleted: $DelExpired" -ForegroundColor "Yellow";
		        LogWrite "Total of expired certificates deleted: $DelExpired";
                    	Write-Log -Level Information -EventId 4022 -Message @{
                        	Message = "Total of expired certificates deleted: $DelExpired"
                    	};
			Write-Host "";
			
		};

        	}
		
	}; # End of check $ExpiredDays

	# Get database location
	$dblocation = .$certutil -databaselocations;
	$dblocation = $dblocation | select-string "edb";

	$i=($dblocation.tostring()).Split(":");
	$dblocation = ($i[1]).Replace("44","").Trim();
	
	# Defragmenting the database
	if ($Defrag -eq $true) {

		# Stopping the CA
		Write-Host "";
		Write-Host "Stopping the $srvName in order to defragment the database..." -ForegroundColor "Yellow";
		LogWrite "Stopping the $srvName in order to defragment the database...";
        	Write-Log -Level Information -EventId 4030 -Message @{
                	Message = "Stopping the $srvName in order to defragment the database..."
            	};
		Stop-Service certsvc;

		# Check CA status
		$srvName = "Active Directory Certificate Services";
		$servicePrior = Get-Service $srvName;
		$Servicestatus = $servicePrior.status;
		Write-Host "";
		Write-Host "$srvName is $Servicestatus" -ForegroundColor "Yellow";
		LogWrite "$srvName is $Servicestatus";
        	Write-Log -Level Information -EventId 4031 -Message @{
            		Message = "Defrag - $srvName is $Servicestatus"
        	};
		Write-Host "";

		# Change directory
		cd "D:\CAScripts\DBTool";

		# Defragmenting the database
		.\eseutil /d "$dblocation";
        	Write-Log -Level Information -EventId 4032 -Message @{
            		Message = "Defrag - Defragmenting the database"
        	};

		cd ..;
	
		# CA Database size after
		$CADatabaseSize = (Get-ChildItem -Path "$CADataBaseLocation" -Recurse -Filter "$CADataBaseName").length/1MB;
		Write-Host "Database size after defragmentation: $CADatabaseSize MB" -ForegroundColor "Yellow";
		LogWrite "Database size after defragmentation: $CADatabaseSize MB";
        	Write-Log -Level Information -EventId 4033 -Message @{
            		Message = "Database size after defragmentation: $CADatabaseSize MB"
        	};
		Write-Host "";
		
		# Starting the CA
		Start-Service certsvc;
		Write-Host "Starting the $srvName and wait for 60 seconds to make sure it has been started." -ForegroundColor "Yellow";
		LogWrite "Starting the CA";
        	Write-Log -Level Information -EventId 4034 -Message @{
            		Message = "Defrag - Starting the CA"
        	};
		Write-Host "";
		
		# 60 second wait
		Start-Sleep -s 60;

		# Check service status
		$servicePrior = Get-Service $srvName;
		$Servicestatus = $servicePrior.status;
		Write-Host "";
		Write-Host "$srvName is now $Servicestatus" -ForegroundColor "Green";
		LogWrite "$srvName is now $Servicestatus";
        	Write-Log -Level Information -EventId 4035 -Message @{
            		Message = "Defrag - $srvName is now $Servicestatus"
        	};
		Write-Host "";
        	Write-Log -Level Information -EventId 4036 -Message @{
            		Message = "Defragmentation process is now complete"
        	};
		
	}; # End of Defrag
    
    	# CA Database size after
	$CADatabaseSize = (Get-ChildItem -Path "$CADataBaseLocation" -Recurse -Filter "$CADataBaseName").length/1MB;
	Write-Host "Database size after CA Maintenance: $CADatabaseSize MB" -ForegroundColor "Yellow";
	LogWrite "Database size after CA Maintenance: $CADatabaseSize MB";
        Write-Log -Level Information -EventId 4006 -Message @{
            Message = "Database size after CA Maintenance: $CADatabaseSize MB"
        };
	Write-Host "";

	# Stop stopwatch
	$totalTime.Stop();
	$ts = $totalTime.Elapsed;
	$totalTime = [system.String]::Format("{0:00}:{1:00}:{2:00}",$ts.Hours, $ts.Minutes, $ts.Seconds);
	Write-Host "";
	Write-Host "Total process time: $totalTime" -ForegroundColor "Yellow";
	LogWrite "Total process time: $totalTime";
    	Write-Log -Level Information -EventId 4007 -Message @{
    		Message = "Total process time: $totalTime"
    	};
	Write-Host "";
	cd \scripts;
	Write-Host "=== Done ===" -ForegroundColor "Green";
	LogWrite "=== Done ===";
    	Write-Log -Level Information -EventId 4008 -Message @{
   		Message = "CA DB Maintenance Script Done"
    	};
	Write-Host "";
	
	##############################################################################
	# End of Script
	##############################################################################
