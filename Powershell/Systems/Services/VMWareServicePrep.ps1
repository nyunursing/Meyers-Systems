<#
Starting and Stopping Windows Services
--------------------------------------

command:  Get-Service
Examples:
    Get-Service -Name "SQL*"

    Status   Name               DisplayName                           
    ------   ----               -----------                           
    Running  SQLBrowser         SQL Server Browser                    
    Running  SQLSERVERAGENT     SQL Server Agent (MSSQLSERVER)        
    Running  SQLTELEMETRY       SQL Server CEIP service (MSSQLSERVER) 
    Running  SQLWriter          SQL Server VSS Writer 

    VS......

    Get-Service -DisplayName "SQL*"
    Status   Name               DisplayName                           
    ------   ----               -----------                           
    Running  MsDtsServer130     SQL Server Integration Services 13.0  
    Running  MSSQLSERVER        SQL Server (MSSQLSERVER)              
    Running  MSSQLServerOLAP... SQL Server Analysis Services (MSSQL...
    Running  ReportServer       SQL Server Reporting Services (MSSQ...
    Running  SQLBrowser         SQL Server Browser                    
    Running  SQLSERVERAGENT     SQL Server Agent (MSSQLSERVER)        
    Running  SQLTELEMETRY       SQL Server CEIP service (MSSQLSERVER) 
    Running  SQLWriter          SQL Server VSS Writer                 
    Running  SSASTELEMETRY      SQL Server Analysis Services CEIP (...
    Running  SSISTELEMETRY130   SQL Server Integration Services CEI...


How to Start/Stop a Server
    Stop-Servic -Name [Insert Name Of Service]


How to Change Startup State
    Set-Service -Name [Insert Name Of Service] -StartupType [Automatic, Boot, Disabled, Manual, System]

#>

New-Variable -Scope global -name CRLF -Value "`r`n"
    
#Error Config
#------------
#Preference
    #$ErrorActionPreference = "STOP"
    $ErrorActionPreference = "Continue"
    #$ErrorActionPreference = "SilentlyContinue"
New-Variable -Scope global -Name gbolChange -Value $False
New-Variable -Scope global -Name gbolError -Value $False
New-Variable -Scope global -Name LogMessage -Value ""

New-Variable -Scope global -Name gbolDebug -Value $True

##Dev or Prod
New-Variable -Scope global -Name gbolProd -Value $False
If($gbolProd -eq $False){
    #Dev
}
Else{
    #Prod
}

#Event Log
$strLogName = "Application"
$strSource = "Start-Cluster.ps1"


#Services Variables
enum enumServiceState{
    Start = 1
    Stop = 2
}

#enum enumServiceGroup{
#    IIS = 1
#    SQL = 2
#}

Function UpdateLogMessage(){
    Param(
        [Parameter(Mandatory=$true)]
        [String]$strMsg
    )

    If ($global:LogMessage.Length -eq 0){$global:LogMessage =(Get-Date -Format o).ToString() + " " + $strMsg}
    Else {$global:LogMessage =$global:LogMessage.trim() + $CRLF + (Get-Date -Format o).ToString() + $strMsg}

    If ($Error.Count -ne 0){
        $global:gbolError = $True        
        #Record Errors
        $global:LogMessage = $global:LogMessage.trim() + $CRLF + $CRLF + "Error Messages:" + $CRLF
    
        For ($i = 0;$i -le ($Error.Count -1);$i++){
            #$global:LogMessage = $global:LogMessage.trim() + $CRLF + "Error " + ($i+1)
            $global:LogMessage = $global:LogMessage.trim() + $CRLF + $Error[$i].ToString()
            $global:LogMessage = $global:LogMessage.trim() + $CRLF + "------------------------------"
        }
    
        #Clear Error Count
        $Error.Clear()
    }
}

Function ModifyService(){
    Param(
        [Parameter(Mandatory=$true)]
        [enumServiceState]$ServiceState,
        [Parameter(Mandatory=$true)]
        [String]$strServiceName
    )

    #Debug: Default Values
    If($Global:gbolDebug -eq $true){
        Write-Host "Modify Service Default Values"
        Write-Host $ServiceState
        Write-Host $strServiceName
        Write-Host "---------------------------"

    }


    If($ServiceState -eq [enumServiceState]::Start){
        #Start
        If ((Get-Service -Name $strServiceName).Status -ne "Running"){
            #Debug
            If($Global:gbolDebug -eq $true){Write-Host "Start Service '" + $strServiceName+ "'"}
            Start-Service -Name $strServiceName

        }
        #Debug
        If($Global:gbolDebug -eq $true){Write-Host "Change Startup to Automatic for '" + $strServiceName + "'"}
        Set-service -Name $strServiceName -StartupType Automatic
    }
    Else{
        #Stop
        #Debug
        If($Global:gbolDebug -eq $true){Write-Host "Stop Service '" + $strServiceName+ "'"}
        Stop-Service -Force -Name $strServiceName

        #Debug
        If($Global:gbolDebug -eq $true){Write-Host "Change Startup to Manual for '" + $strServiceName + "'"}
        set-service -Name $strServiceName -StartupType Manual
    }
}


Function VMWareMove(){
    Param(
        [Parameter(Mandatory=$true)]
        [enumServiceState]$ServiceState
    )



    #EMS
    $objServices = Get-Service -DisplayName "EMS*"

    If ($objServices.Count -gt 0){
        #Debug
        If($Global:gbolDebug -eq $true){Write-Host "EMS Section" + $CRLF + "---------------------------------------"}

        #Process Services
        For ($i=0; $i -le ($objServices.Count -1);$i++){
            #Debug
            If($Global:gbolDebug -eq $true){Write-Host "Processing Count : " ($i+1) + ", Service Name: " + $objServices[$i].Name}
        
            #Services
            If($ServiceState -eq [enumServiceState]::Start){ModifyService -ServiceState Start -strServiceName $objServices[$i].Name.ToString()}
            Else{ModifyService -ServiceState Stop -strServiceName $objServices[$i].Name.ToString()}
        }
    }

    #IIS
    $objServices = Get-Service -DisplayName "IIS*"

    If ($objServices.Count -gt 0){
        #Debug
        If($Global:gbolDebug -eq $true){Write-Host "IIS Section" + $CRLF + "---------------------------------------"}

        #Process Services
        For ($i=0; $i -le ($objServices.Count -1);$i++){
            #Debug
            If($Global:gbolDebug -eq $true){Write-Host "Processing Count : " ($i+1) + ", Service Name: " + $objServices[$i].Name}
        
            #Services
            If($ServiceState -eq [enumServiceState]::Start){iisreset /start}
            Else{iisreset /stop}
        }
    }

    #SQL Server
    $objServices = Get-Service -DisplayName "SQL*"

    If ($objServices.Count -gt 0){
        #Debug
        If($Global:gbolDebug -eq $true){Write-Host "SQL Server Section" + $CRLF + "---------------------------------------"}


        #Process Services
        For ($i=0; $i -le ($objServices.Count -1);$i++){
            #Debug
            If($Global:gbolDebug -eq $true){Write-Host "Processing Count : " ($i+1) + ", Service Name: " + $objServices[$i].Name}
        
            #Services
            If($ServiceState -eq [enumServiceState]::Start){ModifyService -ServiceState Start -strServiceName $objServices[$i].Name.ToString()}
            Else{ModifyService -ServiceState Stop -strServiceName $objServices[$i].Name.ToString()}
        }
    }

    Write-Host ""
    Write-Host ""
    Write-Host ""
    Write-Host ""
    Write-Host ""
    Write-Host ""
    Write-Host "------------------------------"
    Write-Host "TASK COMPLETED!"
    Write-Host "------------------------------"
}



Write-Host "To Change the EMS Services run one of the following commands.."
Write-Host "To STOP services:   VMWareMove -ServiceState Stop"
Write-Host "To START services:  VMWareMove -ServiceState Start"
Write-Host "---------------------------------------------------------------"