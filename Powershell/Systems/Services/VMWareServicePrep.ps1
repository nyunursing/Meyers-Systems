<#
*** You will need to change the execution policy for Powershell if running this via the Powershell ISE
    *Must Have Admin Access to Do this
    
    #Will show current Policy
    Get-ExecutionPolicy   

    #Will Allow you to Run Script
    Set-ExecutionPolicy -ExecutionPolicy Unrestricted   

    #Will Set the Execution Policy Back to the Default Settings
    Set-ExecutionPolicy -ExecutionPolicy Default        


Resources:
    About special Characters:  https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_special_characters?view=powershell-7.1

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

#Default Variables
#----------------
#Text Editing
#------------
New-Variable -Scope global -name CRLF -Value "`r`n"
New-Variable -Scope global -Name TAB -Value "`t"

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

#Email Settings
New-Variable -Scope global -Name gstrEmailFrom -Value ($env:COMPUTERNAME + ".NoReply@NYU.EDU").ToString()
New-Variable -Scope global -Name gstrEmailSubject -Value ((Get-Date -Format "yyyyMMdd-hh:mm").ToString() + " VMWare Service Prep")


##Dev or Prod
New-Variable -Scope global -Name gbolProd -Value $False
If($gbolProd -eq $False){
    #Dev
    #Email Settings
    New-Variable -Scope global -Name gstrEmailTo -Value "km193@nyu.edu"
}
Else{
    #Prod
    #-----
    #Email Settings
    New-Variable -Scope global -Name gstrEmailTo -Value "l5b4z8w8i9w1x4h5@nyumeyers.slack.com"
}



#Event Log
$strLogName = "Application"
$strSource = "Start-Cluster.ps1"


#Services Variables
enum enumServiceState{
    Start = 1
    Stop = 2
}


Function UpdateString(){
    Param(
        [Parameter(Mandatory=$true)]
        [String]$strInfo = "",
        [Parameter(Mandatory=$true)]
        [String]$strUpdate = "",
        [Parameter(Mandatory=$false)]
        [Boolean]$bolNewLine = $true
    )

    <#
        $strInfo is not set to Mandatory since an error will occur if that variable is blank/Null
    #>

    If ($strInfo.Trim().Length -gt 0){
        If($bolNewLine -eq $true){Return ($strInfo.Trim() + $CRLF +$strUpdate)}
        Else{Return ($strInfo.Trim() + $strUpdate)}
    }
    Else{Return $strUpdate}
}


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


Function SendEmail(){
    Param(
        [Parameter(Mandatory=$true)]
        [String]$strEmailTo = $gstrEmailTo,
        [Parameter(Mandatory=$true)]
        [String]$strEmailFrom= $gstrEmailFrom,
        [Parameter(Mandatory=$False)]
        [String]$strEmailSubject = $gstrEmailSubject,
        [Parameter(Mandatory=$true)]
        [String]$strEmailBody,
        [Parameter(Mandatory=$false)]
        [String]$strEmailMsg="",
        [Parameter(Mandatory=$false)]
        [String]$strEmailAttachementPath=""

    )
    
    #SMTP server name
    $SMTPServer = "smtp.nyu.edu"
    #Create a Mail object
    $SMTPMessage = New-Object Net.Mail.MailMessage($strEmailFrom, $strEmailTo, $strEmailSubject, $strEmailBody)
    #Create an Attachment Objct
    If($strEmailAttachementPath.Length -gt 0){
        If((Test-Path -Path $strEmailAttachementPath.Trim()) -eq $true){
            $SMTPMessage.Attachments.Add($Attachment)
            $Attachment = New-Object Net.Mail.Attachment($strEmailAttachementPath)
        }
        Else{$strEmailBody = $CRLF + "WARNING - Check Attachment Path: " + $strEmailAttachementPath + $CRLF}
    }

    #Create SMTP server object
    $SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer) 
    #Send
    $SMTPClient.Send($SMTPMessage)    

}


Function VMWareMove(){
    Param(
        [Parameter(Mandatory=$true)]
        [enumServiceState]$ServiceState
    )

    <#
        IIS Service needs to be handled differently as we are not stopping the IISAdmin Service.
        We are using a different method to start and stop the Web Sites
        --->  iisreset /Start
        --->  iisreset /Stop
    #>


    $arryServiceSearch = @('Cluster*','EMS*', 'SQL*')
    [Object]$arryTmp
    [Object]$arryServices
    [String]$strSettingChanges = $env:COMPUTERNAME + $CRLF + "---------------------------------------" + $CRLF 

    #Determine if IIS is installed on the Server
    #IIS
    $objServices = Get-Service -DisplayName "IIS*"

    If ($objServices.Count -gt 0){
        #Debug
        If($Global:gbolDebug -eq $true){Write-Host "IIS Section" + $CRLF + "---------------------------------------"}

        $strSettingChanges = UpdateString -strInfo $strSEttingChanges -strUpdate ("IIS Section" + $CRLF + "----------" + $CRLF )

        #Start/Stop Website(s)
        If($ServiceState -eq [enumServiceState]::Start){
            iisreset /start
            $strSettingChanges = UpdateString -strInfo $strSEttingChanges -strUpdate ($TAB + "IIS Started" + $CRLF + "----------" + $CRLF )
        }
        Else{
            iisreset /stop
            $strSettingChanges = UpdateString -strInfo $strSEttingChanges -strUpdate ($TAB + "IIS Stopped" + $CRLF + "--------------------" + $CRLF )
        }

    }
    
    #Determine What other Services Are Installed
    $arryServiceSearch | ForEach-Object {
        $objServices = Get-Service -DisplayName $_
        
        If ($objServices.Count -gt 0){
            $strSettingChanges = UpdateString -strInfo $strSettingChanges -strUpdate ($_ + " Section" + $CRLF + "----------" + $CRLF )
            
            #Modify Services
            For ($i=0; $i -le ($objServices.Count -1);$i++){
                #Debug
                If($Global:gbolDebug -eq $true){Write-Host $objServices[$i].Name + " Section" + $CRLF + "---------------------------------------"}
                $strSettingChanges = UpdateString -strInfo $strSettingChanges -strUpdate ($TAB + $objServices[$I].Name +  $CRLF)
                
                If($ServiceState -eq [enumServiceState]::Start){
                    #Start Service
                    If($objServices[$i].Status -ne "Running" ){
                        $strSettingChanges = UpdateString -strInfo $strSettingChanges -strUpdate ($TAB + $TAB + "Status: " + $objServices[$I].Status +  " ---> Started" + $CRLF )
                        Start-Service -Name ($objServices[$i].Name)
                    }
                    Else{$strSettingChanges = UpdateString -strInfo $strSettingChanges -strUpdate ($TAB + $TAB + "Status: " + $objServices[$I].Status +  " ---> NO CHANGE" + $CRLF )}
                    
                    #Change Startup Option
                    If($objServices[$i].StartType -ne "Automatic"){
                        $strSettingChanges = UpdateString -strInfo $strSettingChanges -strUpdate ($TAB + $TAB + "StartUp: " + $objServices[$I].StartType +  " ---> Automatic" + $CRLF )
                        Set-service -Name ($objServices[$i].Name) -StartupType Automatic
                    }
                    Else{$strSettingChanges = UpdateString -strInfo $strSettingChanges -strUpdate ($TAB + $TAB + "StartUp: " + $objServices[$I].StartType +  " ---> NO CHANGE" + $CRLF )}
                }
                Else{
                    #Stop Service
                    If($objServices[$i].Status -ne "Stopped" ){
                        $strSettingChanges = UpdateString -strInfo $strSettingChanges -strUpdate ($TAB + $TAB + "Status: " + $objServices[$I].Status +  " ---> Stopped" + $CRLF )
                        Stop-Service -Force -Name ($objServices[$i].Name)
                    }
                    Else{$strSettingChanges = UpdateString -strInfo $strSettingChanges -strUpdate ($TAB + $TAB + "Status: " + $objServices[$I].Status +  " ---> NO CHANGE" + $CRLF )}
                    
                    #Change Startup Option
                    If($objServices[$i].StartType -ne "Manual"){
                        $strSettingChanges = UpdateString -strInfo $strSettingChanges -strUpdate ($TAB + $TAB + "StartUp: " + $objServices[$I].StartType +  " ---> Manual" + $CRLF )
                        Set-service -Name ($objServices[$i].Name) -StartupType Manual
                    }
                    Else{$strSettingChanges = UpdateString -strInfo $strSettingChanges -strUpdate ($TAB + $TAB + "StartUp: " + $objServices[$I].StartType +  " ---> NO CHANGE" + $CRLF )}
                }
            }
        }
    }



    Write-Host $strSettingChanges

    #Sent Email
    SendEMail -strEmailTo $gstrEmailTo -strEmailFrom $gstrEmailFrom -strEmailSubject $gstrEmailSubject -strEmailBody $strSettingChanges

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