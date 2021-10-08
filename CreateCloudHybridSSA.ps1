<#
.SYNOPSIS
    When you run this script creat a cloud SSA ready to onboard for Cloud Hybrid Search.
    The script validates that the Server roles support Search
    The script waits for the critical tasks to start before proceeding
.PARAMETER SearchServiceAppName
    SharePoint Online portal URL, for example 'Cloud SSA'.
.PARAMETER SearchServerName
    Name of the first Search Server to host the cloud SSA
.PARAMETER SearchServerName2
    [Optional] Name of the second Search Server to host the cloud SSA
.PARAMETER DatabaseServerName
    Name of the SQL Server instance to host the cloud SSA Databases
.PARAMETER Credential
    Logon credential for Service account. Will prompt for credential if not specified.
.LAST UPDATED
    2021-09-24 by RainerA
    Added compatibility for SP 2016
    2021-10-08 by RainerA
#>
Param(
    [Parameter(Mandatory=$true, HelpMessage="Name of the Cloud Search Search Service Application ie. 'Cloud SSA'")]
    [ValidateNotNullOrEmpty()]
    [string] $SearchServiceAppName,

    [Parameter(Mandatory=$true, HelpMessage="Name of the first search Server that hosts the Search Service Application")]
    [ValidateNotNullOrEmpty()]
    [string] $SearchServerName,

    [Parameter(Mandatory=$false, HelpMessage="[Optional] Name of the second search servicer to host the Search Service Application")]
    [string] $SearchServerName2,

    [Parameter(Mandatory=$true, HelpMessage="Name of the SQL Instance to host the search databases ie. 'SQL\SearchDBs'")]
    [ValidateNotNullOrEmpty()]
    [string] $DatabaseServerName,

    [Parameter(Mandatory=$true, HelpMessage="Credential of the Service Account to run the services")]
    [PSCredential] $SearchServiceAccount

)

$IsDataMove = $false
$isSubscriptionEdition = $false;

#Handle SP Subscription Edition vs. older SP Versions
if ((get-module SharePointServer) -eq $null)
{
    Add-PSSnapin Microsoft.SharePoint.PowerShell
}

function Log($logstring)
{
    $d = get-date -Format "yyyy-MM-dd HH:mm:ss"
    write-host "$d : $logstring"
}

cls

log "Confirming SharePoint version"
$buildVersion = (Get-SPFarm).BuildVersion
log "SharePoint BuildVersion : $BuildVersion"
if (($buildVersion.Major -eq 16) -and ($buildVersion.Build -gt 10701))
{
    $isSubscriptionEdition = $true
    if ($buildVersion.Build -lt 14326)
    {
        throw ("Cannot create a Hybrid Search Service Application with a pre-release build of SharePoint Subscription Edition.")
    }
} 

log "Checking for existing Cloud SSA"

$cloudSsa = Get-SPEnterpriseSearchServiceApplication | Where { $_.CloudIndex -eq $true }
if ($cloudSsa -ne $null) 
{        
    Switch ($cloudSsa.Count)
    {
        0 
        {
            $IsDataMove = $false;
        }
        1 
        {
            if ($cloudSsa.Name -eq $SearchServiceAppName)
            {
                log "A cloud SSA with the name '$SearchServiceAppName' already exists.  Doing nothing."
                return
            }
            elseif(!$isSubscriptionEdition)
            {
                log "This script is designed to run on SharePoint Subscription Edition"
                return
            }
            else
            {
                $IsDataMove = $true 
            }
        }
        default 
        {
            throw "No cloud SSA will be created because this farm already has two or more Cloud SSAs.";
        }
    }
}

log "Validating service account"
if ($SearchServiceAccount -eq $null)
{
    $SearchServiceAccount = Get-Credential -Message "Enter the credentials for the Service Account"
}


$appPoolName = $SearchServiceAppName + "_AppPool"

## Validate if the supplied account exists in Active Directory and whether it’s supplied as domain\username 

if ($SearchServiceAccount.UserName.Contains("\")) # if True then domain\username was used 
{ 
    $Account = $SearchServiceAccount.UserName.Split("\") 
    $Account = $Account[1] 
} 
else # no domain was specified at account entry 
{ 
    $Account = $SearchServiceAccount 
} 

$domainRoot = [ADSI]'' 
$dirSearcher = New-Object System.DirectoryServices.DirectorySearcher($domainRoot) 
$dirSearcher.filter = "(&(objectClass=user)(sAMAccountName=$Account))" 
$results = $dirSearcher.findall() 

if ($results.Count -eq 0) # Test for user not found 
{  
     throw "The account $($SearchServiceAccount.UserName) does not exist in Active Directory. Please create this account in Active Directory and try again."
}

log "Confirming that the account $($SearchServiceAccount.Username) is registered as managed account"
if (!(Get-SPManagedAccount | ? { $_.username -eq $SearchServiceAccount.UserName}))
{
    log "Create Managed Account for $($SearchServiceAccount.UserName)"
    New-SPManagedAccount -Credential $SearchServiceAccount 
    
    #validating account again
    if (!(Get-SPManagedAccount | ? { $_.username -eq $SearchServiceAccount.UserName}))
    {
        throw "The account $($SearchServiceAccount.UserName) could not be registered as managed account in SharePoint. Please register this account and try again."
    }
}

log "Validating different Servernames"
if ($SearchServerName -eq $SearchServerName2)
{
    throw "The Names for SearchServer and SearchServer2 must be different."    
}

log "Confirm that servers have Search in their Role, or are custom servers"
$SearchServer1Role = ((Get-SPServer $SearchServerName) | select role).Role
if (!(($SearchServer1Role -match "Search") -or ($SearchServer1Role -eq "Custom") -or ($SearchServer1Role -eq "SingleServerFarm") ))
{
    throw "The server '$SearchServerName' has the role '$SearchServer1Role'. This role cannot host a Search Service Instance."
}

if($SearchServerName2 -ne [String]::Empty)
{
    $SearchServer2Role = ((Get-SPServer $SearchServerName2) | select role).Role
    if (!(($SearchServer2Role -match "Search") -or ($SearchServer2Role -match "Custom") )) #Single Server Farm cannot have 2 servers
    {
        throw "The server '$SearchServerName2' has the role '$SearchServer2Role'. This role cannot host a Search Service Instance."
    }
}

log "Check if the app pool already exists"
$appPool = (Get-SPServiceApplicationPool  | ? {$_.name -eq $appPoolName})
if ($appPool)
{
    if ($appPool.ProcessAccountName -ne $SearchServiceAccount.UserName)
    {
    
        log "Run  Get-SPServiceApplicationPool $appPoolName | Remove-SPServiceApplicationPool"
        throw "An application pool with name: $appPoolName using a different Service Account already exists. Please delete this service application pool an try again."
    }
    else 
    {
        log "Application pool already exists"
    }
}
else
{
    log "Create Service application pool $appPoolName"
    $appPool = New-SPServiceApplicationPool -name $appPoolName -account $SearchServiceAccount.UserName

    $appPool = (Get-SPServiceApplicationPool  | ? {$_.name -eq $appPoolName})
    if ($appPool -eq $null) 
    {
        throw "Could not create a new app pool with name: $appPoolName and account: $SearchServiceAccount. Please check if the parameters are valid."
    }
}

log "Starting Service Instance on $SearchServerName"
$searchInstance = Get-SPEnterpriseSearchServiceInstance  | ? {$_.server.name -eq  $SearchServerName}
if ($SearchInstance.Status -ne "Online")
{ 
    Start-SPEnterpriseSearchServiceInstance $SearchServerName 
    log "Starting Search Service Instance on Server '$SearchServerName'" 
}


if($SearchServerName2 -ne [String]::Empty)
{
    log  "Starting Service Instance on $SearchServerName2"
    $searchInstance2 = Get-SPEnterpriseSearchServiceInstance  | ? {$_.server.name -eq $SearchServerName2}
    if ($SearchInstance2.Status -ne "Online")
    { 
        Start-SPEnterpriseSearchServiceInstance $SearchServerName2 
        log "Starting Search Service Instance on Server '$SearchServerName2'" 
    }
}

log "Wait for the Search Service Instance comes up on Server $SearchServerName"
$searchInstance = Get-SPEnterpriseSearchServiceInstance $SearchServerName
$timeoutTime = (Get-Date).AddMinutes(10)
if ($SearchInstance.Status -ne "Online")
{
    while (($SearchInstance.Status -ne "Online") -and ($timeoutTime -ge (Get-Date)))
    { 
        write-host "." -NoNewline
        Start-Sleep 10; 
        $searchInstance = Get-SPEnterpriseSearchServiceInstance $SearchServerName
    }
    write-host
}

$searchInstance = Get-SPEnterpriseSearchServiceInstance $SearchServerName
if ($SearchInstance.Status -ne "Online")
{ 
    throw "Search Service Instance could not be initialized on Server '$SearchServerName'" 
}

if($SearchServerName2 -ne [String]::Empty)
{
    log "Wait for the Search Service Instance comes up on server $SearchServerName2"
    $searchInstance2 = Get-SPEnterpriseSearchServiceInstance $SearchServerName2
    $timeoutTime = (Get-Date).AddMinutes(1)
    if ($SearchInstance2.Status -ne "Online") 
    {        
        while (($SearchInstance2.Status -ne "Online") -and ($timeoutTime -ge (Get-Date)))
        { 
            write-host "." -NoNewline
            Start-Sleep 10; 
            $searchInstance2 = Get-SPEnterpriseSearchServiceInstance $SearchServerName2
        } 
        $searchInstance2 = Get-SPEnterpriseSearchServiceInstance $SearchServerName2
        if ($SearchInstance2.Status -ne "Online")
        { 
            throw "Search Service Instance could not be initialized on Server '$SearchServerName2'" 
        }
        write-host
    }
}

log "Generating new Search Service Application this will take a few minutes"
$searchApp = New-SPEnterpriseSearchServiceApplication -Name $SearchServiceAppName -ApplicationPool $appPool -DatabaseServer $DatabaseServerName -CloudIndex $true
if ($searchApp -eq $null) 
{
    throw "Could not create a new search service application with name: $SearchServiceAppName, app pool: $appPool, database server: $DatabaseServerName and cloud index flag: true. Please check if the parameters are valid."
}

log "Setting Search Admin component"
$searchInstance = Get-SPEnterpriseSearchServiceInstance $SearchServerName
$searchApp | get-SPEnterpriseSearchAdministrationComponent | set-SPEnterpriseSearchAdministrationComponent -SearchServiceInstance $searchInstance
$admin = ($searchApp | get-SPEnterpriseSearchAdministrationComponent)
$timeoutTime = (Get-Date).AddMinutes(20)
 while ((-not $admin.Initialized) -and ($timeoutTime -ge (Get-Date))) { 
    write-host "." -NoNewline
    Start-Sleep 10; 
}

if (-not $admin.Initialized) 
{ 
    throw 'Admin Component could not be initialized' 
}

log "Adding Topology components"
$searchApp = Get-SPEnterpriseSearchServiceApplication $SearchServiceAppName
$topology = $searchApp.ActiveTopology.Clone()
$topology.Store()
$oldComponents = @($topology.GetComponents())

if (@($oldComponents | ? { $_.GetType().Name -eq "AdminComponent" }).Length -eq 0)
{
    $topology.AddComponent((New-Object Microsoft.Office.Server.Search.Administration.Topology.AdminComponent $SearchServerName))
}

$topology.AddComponent((New-Object Microsoft.Office.Server.Search.Administration.Topology.CrawlComponent $SearchServerName))
$topology.AddComponent((New-Object Microsoft.Office.Server.Search.Administration.Topology.ContentProcessingComponent $SearchServerName))
$topology.AddComponent((New-Object Microsoft.Office.Server.Search.Administration.Topology.AnalyticsProcessingComponent $SearchServerName))
$topology.AddComponent((New-Object Microsoft.Office.Server.Search.Administration.Topology.QueryProcessingComponent $SearchServerName))
$topology.AddComponent((New-Object Microsoft.Office.Server.Search.Administration.Topology.IndexComponent $SearchServerName,0))

if($SearchServerName2 -ne [String]::Empty)
{
    $topology.AddComponent((New-Object Microsoft.Office.Server.Search.Administration.Topology.AdminComponent $SearchServerName2))
    $topology.AddComponent((New-Object Microsoft.Office.Server.Search.Administration.Topology.CrawlComponent $SearchServerName2))    
    $topology.AddComponent((New-Object Microsoft.Office.Server.Search.Administration.Topology.ContentProcessingComponent $SearchServerName2))    
    $topology.AddComponent((New-Object Microsoft.Office.Server.Search.Administration.Topology.AnalyticsProcessingComponent $SearchServerName2))   
    $topology.AddComponent((New-Object Microsoft.Office.Server.Search.Administration.Topology.QueryProcessingComponent $SearchServerName2))    
    $topology.AddComponent((New-Object Microsoft.Office.Server.Search.Administration.Topology.IndexComponent $SearchServerName2,0))
}
$oldComponents | ? { $_.GetType().Name -ne "AdminComponent" } | foreach { $topology.RemoveComponent($_) }

log "Activating Topology"
$topology.Activate()

if ($searchApp.GetTopology($topology.TopologyId).State -ne "Active")
{ 
    throw 'Could not activate the search topology'
}

log "Creating Search Service Application Proxy"
$searchAppProxy = New-SPEnterpriseSearchServiceApplicationProxy -name ($SearchServiceAppName + "_proxy") -SearchApplication $searchApp

log "Clean up inactive topology"
$searchApp | Get-SPEnterpriseSearchTopology | ? {$_.state -ne "Active"} | Remove-SPEnterpriseSearchTopology -Confirm:$false

log "Done"