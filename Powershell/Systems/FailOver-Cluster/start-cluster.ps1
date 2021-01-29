<#

Resources
----------------------------------------------
Powershell
    Comparison Operators:  http://ss64.com/ps/syntax-compare.html
    Data Types:  http://ss64.com/ps/syntax-datatypes.html
    It's a Trap, Passing variables to functions:  http://stackoverflow.com/questions/957707/parameters-and-powershell-functions
    Variable Scope:  https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_scopes?view=powershell-7.1
        $private:test = 1 The variable exists only in the current scope. It cannot be accessed in any other scope. 
        $local:test = 1 Variables will be created only in the local scope. That's the default for variables that are specified without a scope. Local variables can be read from scopes originating from the current scope, but they cannot be modified. 
        $test = 1 This scope represents the top-level scope in a script. All functions and parts of a script can share variables by addressing this scope. 
        $test = 1 This scope represents the scope of the PowerShell console. So if a variable is defined in this scope, it will still exist even when the script that is defining it is no longer running. 
    Creating Custom Objects:  http://social.technet.microsoft.com/wiki/contents/articles/7804.powershell-creating-custom-objects.aspx
    PowerShell Book:  http://powershell.com/cs/blogs/ebookv2/default.aspx
    Microsoft TechNet:  https://technet.microsoft.com/en-us/library/jj159398.aspx


KMC Updated Script 20210127  (ver 1.4)
--------------------------------------
Updated the script so that the SQL Jobs know which jobs should be enabled dependent on which server is the Primary Server
    1)  Added Function SQL-AG-AlwaysOnMonitor.  Runs that Job on all the SQL Servers


KMC Updated Script 20210125  (ver 1.3)
--------------------------------------
Updated the script to Write Information to the APPLICATION Event Log When it actually does something
    1) Must Register the SOURCE using New-Eventlog (One Time Deal)
    2) From then on you can write event logs using that source

Event ID's:
    100     |  Event Log Creation.
    1000    |  Script completed with No Errors.
    3000    |  Errors detected, not all processes/services restarted.


KMC Updated Script 20210123  (ver 1.2)
--------------------------------------
Specifying the Cluster Name creates various errors when the Cluster Service is down.
    EXAMPLE
    If ($FileShareWitness.State -eq "Offline"){Start-ClusterResource -Cluster $strClusterName -Name "File Share Witness"}
    Get-ClusterResource : Check the spelling of the cluster name. Otherwise, there might be a problem with your network. Make sure the cluster nodes are turned on and 
    connected to the network or contact your network administrator.
        The RPC server is unavailable
    At line:6 char:21
    + ... reWitness = Get-ClusterResource -Cluster $strClusterName -Name "File  ...
    +                 ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        + CategoryInfo          : ConnectionError: (:) [Get-ClusterResource], ClusterCmdletException
        + FullyQualifiedErrorId : ClusterRpcConnection,Microsoft.FailoverClusters.PowerShell.GetResourceCommand


Need to run this script on the Cluster Itself to avoid errors like this

    WARNING: If you are running Windows PowerShell remotely, note that some failover clustering cmdlets do not work remotely. When possible, run the cmdlet locally and
     specify a remote computer as the target. To run the cmdlet remotely, try using the Credential Security Service Provider (CredSSP). All additional errors or warnin
    gs from this cmdlet might be caused by running it remotely.


KMC Updated Script 20210121.  (ver 1.1)
--------------------------------------
#Problem:
----------
1)  If the Cluster Fails, the command above does not start the cluster resource due to the following error:

    WARNING: If you are running Windows PowerShell remotely, note that some failover 
    clustering cmdlets do not work remotely. When possible, run the cmdlet locally and
    specify a remote computer as the target. To run the cmdlet remotely, try using the
    Credential Security Service Provider (CredSSP). All additional errors or warnings 
    from this cmdlet might be caused by running it remotely.


2)  Updated the Order


Neima Ullah Version 1.0 (Original Code)
--------------------------------------
    #Start-Cluster -Name CONEMSSQLCLSTR
    #Start-ClusterGroup -Cluster CONEMSSQLCLSTR -Name CONEMSSQLAG
    #Start-ClusterResource -Cluster CONEMSSQLCLSTR -Name "File Share Witness"
#>


#Default Variables
#----------------
#Text Editing
#------------
New-Variable -Scope global -name CRLF -Value "`r`n"
    
#Error Config
#------------
#Preference
    #$ErrorActionPreference = "STOP"
    $ErrorActionPreference = "Continue"
    #$ErrorActionPreference = "SilentlyContinue"


#Mail Settings
$EmailTo = "l5b4z8w8i9w1x4h5@nyumeyers.slack.com"
$EmailBody = ""
#$EmailAttachment = "Full File Path"
New-Variable -Scope global -Name gbolProd -Value $true


##Dev or Prod
If($gbolProd -eq $False){
    #Dev
    $strRoleName = "CON-SQL1-AG"
    $strClusterName = "CON-SQL1-CLSTR"
    $EmailFrom		= "CON940@NoReply.com"
    $EmailSubject		= "Dev " + $strClusterName
}
Else{
    #Prod
    $strRoleName = "CONEMSSQLAG"
    $strClusterName = "CONEMSSQLCLSTR"
    $EmailFrom		= "CON928@NoReply.com"
    $EmailSubject		= "Prod " + $strClusterName
}

#Event Log
$strLogName = "Application"
$strSource = "Start-Cluster.ps1"

New-Variable -Scope global -Name gbolChange -Value $False
New-Variable -Scope global -Name gbolError -Value $False
New-Variable -Scope global -Name LogMessage -Value ""


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


Function SQL-AG-AlwaysOnMonitor{
    param( 
	    [Parameter(Mandatory=$true)]
        [string]$ServerName,

        [Parameter(Mandatory=$true)]	    
        [string]$JobName,

        [Parameter(Mandatory=$true)]
	    [string]$StepName 
    )

    $bolRunJob = $False
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | out-null
    $srv = New-Object Microsoft.SqlServer.Management.SMO.Server("$ServerName")
    $job = $srv.jobserver.jobs["$JobName"] 

    If ($job){	
	    #Start Job
        If($StepName -ne ''){$job.Start($StepName)}
	    Else {$job.Start()}

        $bolRunJob = $true
    }


    If($bolRunJob -eq $true){UpdateLogMessage -strMsg ($ServerName + ": AG Jobs Updated.  Job Status = " + $job.LastRunOutcome + " On " + $job.LastRunDate)}
    Else{
        #Job Should Exist on All SQL Servers
        $global:gbolError = $true
        UpdateLogMessage -strMsg ("SQL Job: " + $JobName + ", could not be found on Server: " + $ServerName)
    }
}


<#
Registering the Source if it doesn't exists
    Notes:
    - One Time Thing, Need to run with Admin Rights
    - Try/Catch will only work on TERMINATING Execptions, Get-EventLog when it fails is not a Terminating Exception!
#>
Try{Get-EventLog -LogName $strLogName -Source $strSource -ErrorAction Stop}
Catch{
    #Source Doesn't Exist so Create it
    Write-Host "Creating Event Log Source"
    New-EventLog -LogName $strLogName -Source $strSource
    Write-EventLog -LogName $strLogName -Source $strSource -EventId 100 -EntryType Information -Message ("Created Event Log '" +$strLogName + "', Source '" + $strSource + "'.")
    
    #Clear Error Log
    $Error.clear()
}




#Cluster Service
$ClusterGroup = Get-ClusterGroup -Name "Cluster Group"
If ($ClusterGroup.State -ne "Online"){
<#
    Start-Cluster -Name $strClusterName   #Cannot Start Cluster This way :(
    
    WARNING: If you are running Windows PowerShell remotely, note that some failover 
    clustering cmdlets do not work remotely. When possible, run the cmdlet locally and
    specify a remote computer as the target. To run the cmdlet remotely, try using the
    Credential Security Service Provider (CredSSP). All additional errors or warnings 
    from this cmdlet might be caused by running it remotely.
#>
    Start-ClusterResource -Name "Cluster Name"
    
    $gbolChange = $True
    UpdateLogMessage -strMsg ("Restarted: Cluster Group - " + $strClusterName)
    
    #if ($LogMessage.Length -eq 0){$LogMessage =(Get-Date -Format o).ToString() +  " Restarted: Cluster Group - " + $strClusterName}
    #Else {$LogMessage =$LogMessage.trim() + $CRLF + (Get-Date -Format o).ToString() +  " Restarted: Cluster Group - " + $strClusterName}
}


#File Share Witness
$FileShareWitness = Get-ClusterResource -Name "File Share Witness"
If ($FileShareWitness.State -eq "Offline"){
    Start-ClusterResource -Name "File Share Witness"
    
    $gbolChange = $True
    UpdateLogMessage -strMsg ("Restarted: File Share Witness")
#    if ($LogMessage.Length -eq 0){$LogMessage =(Get-Date -Format o).ToString() +  " Restarted: File Share Witness"}
#    Else {$LogMessage =$LogMessage.trim() + $CRLF + (Get-Date -Format o).ToString() +  " Restarted: File Share Witness"}
}


#Role
$Role = Get-ClusterResource -Name $strRoleName
If ($Role.State -ne "Online"){
    Start-ClusterResource -Name $strRoleName

    $gbolChange = $True
    UpdateLogMessage -strMsg ("Restarted: Role - " + $strRoleName)
    #if ($LogMessage.Length -eq 0){$LogMessage =(Get-Date -Format o).ToString() +  " Restarted: Role - " + $strRoleName}
    #Else {$LogMessage =$LogMessage.trim() + $CRLF + (Get-Date -Format o).ToString() +  " Restarted: Role - " + $strRoleName}
}

#Ensure Primary Node is started 1st, then start all other nodes.
$PrimaryNode = Get-ClusterNode -Name $Role.OwnerNode.Name
if ($PrimaryNode.State -ne "Up"){
    Start-ClusterNode -Name $PrimaryNode.Name
    
    $gbolChange = $True
    UpdateLogMessage -strMsg ("Restarted: Primary Node - " + $PrimaryNode.Name)
    #if ($LogMessage.Length -eq 0){$LogMessage =(Get-Date -Format o).ToString() +  " Restarted: Primary Node - " + $PrimaryNode.Name}
    #Else {$LogMessage =$LogMessage.trim() + $CRLF + (Get-Date -Format o).ToString() +  " Restarted: Primary Node - " + $PrimaryNode.Name}
}

$Nodes = Get-ClusterNode
For ($i = 0;$i -le ($Nodes.Count -1);$i++){
    if ($Nodes[$i].State -ne "Up"){
        Start-ClusterNode -Name $Nodes[$i].Name
        
        $gbolChange = $True
        UpdateLogMessage -strMsg ("Restarted: Node - " + $Nodes[$i].Name)
        #if ($LogMessage.Length -eq 0){$LogMessage =(Get-Date -Format o).ToString() +  " Restarted: Node - " + $Nodes[$i].Name}
        #Else {$LogMessage =$LogMessage.trim() + $CRLF + (Get-Date -Format o).ToString() +  " Restarted: Node - " + $Nodes[$i].Name}

    }
}

#If Change(s) made, Write to Application log
If ($gbolChange -eq $True){
    #Primary Node
    $LogMessage = "Primary Node: " + $PrimaryNode.Name + $CRLF + $LogMessage

    #Run the Job 'AG-AlwaysOnMonitor.Subplan_1' so that each server's backup jobs are in sync
    For ($i = 0;$i -le ($Nodes.Count -1);$i++){
        if ($Nodes[$i].State -eq "Up"){
            Write-Host "Run SQL Job AG-AlwaysOnMonitor on " $Nodes[$i].Name
            SQL-AG-AlwaysOnMonitor -ServerName $Nodes[$i].Name -JobName "AG-AlwaysOnMonitor.Subplan_1" -StepName "Subplan_1"
        }
    }


    #Determine if any Error's occured
    If ($gbolError -eq $True){
        #Write To Application Log
        Write-EventLog -LogName $strLogName -Source $strSource -EventId 3000 -EntryType Error -Message $LogMessage

        #Clear Error Count
        $Error.Clear()

    }
    Else{
        #No Errors Generated 
        #Write To Application Log
        Write-EventLog -LogName $strLogName -Source $strSource -EventId 1000 -EntryType Information -Message $LogMessage
    }

    #Send Email to Slack Server-Alerts Channel
    Write-Host "Sending Email"
    $EmailBody = $LogMessage
    
    #SMTP server name
    $SMTPServer = "smtp.nyu.edu"
    #Create a Mail object
    $SMTPMessage = New-Object Net.Mail.MailMessage($EmailFrom, $EmailTo, $EmailSubject, $EmailBody)
    #Create an Attachment Objct
    #$Attachment = New-Object Net.Mail.Attachment($EmailAttachment)
    #Add Attachment Object
    #$SMTPMessage.Attachments.Add($Attachment)

    #Create SMTP server object
    If($Global:gbolProd -eq $True){$
        $SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer) 
        #Send
        $SMTPClient.Send($SMTPMessage)
    
        Write-Host "Email Sent!"
    }
}