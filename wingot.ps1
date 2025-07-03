param(
	[Parameter()]
	[bool]$DownloadOnly = $false,
    [Parameter()]
	[bool]$MCMImportOnly = $false,
	[Parameter()]
	[string]$DownloadLocation = $PWD,
    [Parameter()]
	[string]$ImportContentPath = "$PWD\Wingot_Downloads",
    [Parameter()]
    [bool]$makeTaskSequence = $false
)

#High Level Variables
$MCMSiteCode = "CHQ"
$MCMPrimarySiteServer = "CM1.corp.contoso.com"
$MCMApplicationLibraryLocation = "\\localhost\c$\Packages\Apps"
$DPGroups = @("Corp DPs")
$MCMFolderName = "Wingot"

$appsToDownload = @(
    "DominikReichl.KeePass",
    "Microsoft.VCRedist.2015+.x64"
    "Microsoft.VCRedist.2015+.x86",
    "Oracle.JavaRuntimeEnvironment",
    "Citrix.Workspace.LTSR"
)

Try{
    Import-Module (Join-Path $PSScriptRoot "powershell-yaml")
    $here = $PWD
}
Catch{
    Write-Host "Unable to find yaml module, exiting" -ForegroundColor Red
    Break
}

If(!($DownloadOnly)){
    try{
        $SiteCode = $MCMSiteCode 
        $ProviderMachineName = $MCMPrimarySiteServer
        if((Get-Module ConfigurationManager) -eq $null) {
            Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" -ErrorAction Stop
        }
        if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName -ErrorAction Stop
        }
    }
    Catch{
        Write-Host "Unable to connect with MCM. Is the console installed?" -ForegroundColor Red
        break
    }
}




Function Download($App){
    If (!(Test-Path "$DownloadLocation\Wingot_Downloads")){
        New-Item $DownloadLocation -Name "Wingot_Downloads" -ItemType Directory | Out-Null
        Write-Host "Wingot_Downloads folder not found, created it." -ForegroundColor Green
    }
    Else{
        Write-Host "Wingot_Downloads folder found under $DownloadLocation, continuing." -ForegroundColor Yellow
    }
    If (Test-Path "$DownloadLocation\Wingot_Downloads\$App"){
        Remove-Item "$DownloadLocation\Wingot_Downloads\$App" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "Previously download of $App found, probably from the last time this was run. Cleaned up." -ForegroundColor Magenta
    }
    New-Item (Join-Path $DownloadLocation "Wingot_Downloads") -Name $App -ItemType Directory | Out-Null
    Write-Host "Created $App folder under $DownloadLocation\Wingot_Downloads as a temp location" -ForegroundColor Green
    $DownloadPath = "$DownloadLocation\Wingot_Downloads\$App"
    Try{
        Write-Host "Attempting to download $App with winget" -ForegroundColor Blue
        Start-Process winget -ArgumentList "download $App -d $DownloadPath" -Wait -NoNewWindow
        Write-Host "Download Successful" -ForegroundColor Green
    }
    Catch{
        Write-Host "Failed to download from winget. Please investigate." -ForegroundColor Red
        Write-Host "Error: $Error[0]"
    }
}

Function GetDetails($App, $contentPath){

    $global:appName = $null
    $global:author = $null
    $global:version = $null
    $global:installCode = $null
    $global:installContent = $null
    $global:installArgument = $null  
    Try{
        Write-Host "Attemping to access downloaded .yaml file at $contentpath\$app" -ForegroundColor Blue
        $yaml = Get-Content (Get-ChildItem "$contentPath\$App" -Recurse -Filter "*.yaml").FullName | ConvertFrom-Yaml
        Write-Host "Successfully found and accessed .yaml file" -ForegroundColor Green
    }
    Catch{
        Write-Host "Failed to access file." -ForegroundColor Red
    }
    $global:appName = $yaml.PackageName
    $global:author = $yaml.Author
    $global:author = $global:author.Trimend('.')
    $global:version = $yaml.PackageVersion
    $global:installCode = $yaml.Installers.ProductCode
    $global:installContent = (Get-ChildItem "$contentPath\$App" -Recurse -Exclude "*.yaml").Name
    $global:installArgument = $yaml.Installers.InstallerSwitches.Silent
}

Function CopytoLibrary($App, $contentPath){
    $global:appLibraryPath = $null
    try{
        If (!(Test-Path "$MCMApplicationLibraryLocation\$global:author\")){
            New-Item "$MCMApplicationLibraryLocation" -Name "$global:author" -ItemType Directory | Out-Null
            Write-Host "Created $global:author folder under $MCMApplicationLibraryLocation" -ForegroundColor Green
        }
        Else{
            Write-Host "Found $global:author folder under $MCMApplicationLibraryLocation" -ForegroundColor Yellow
        }
        If (!(Test-Path "$MCMApplicationLibraryLocation\$global:author\$global:appName")){
            New-Item "$MCMApplicationLibraryLocation\$global:author" -Name "$global:appName" -ItemType Directory | Out-Null
            Write-Host "Created $global:appName folder under $MCMApplicationLibraryLocation\$global:author" -ForegroundColor Green
        }
        Else{
            Write-Host "Found $global:appName folder under $MCMApplicationLibraryLocation\$global:author" -ForegroundColor Yellow
            }
        If (!(Test-Path "$MCMApplicationLibraryLocation\$global:author\$global:appName\$global:version")){
            New-Item "$MCMApplicationLibraryLocation\$global:author\$global:appName" -Name "$global:version" -ItemType Directory | Out-Null
            Write-Host "Created $global:version folder under $MCMApplicationLibraryLocation\$global:author\$global:appName" -ForegroundColor Green
        }
        else{
            Write-Host "Folder for this version of $global:appName already found. Check if this is needed. Exiting." -ForegroundColor Red
            Break
        }
    }
    Catch{
        Write-Host "Failed to make folders" -ForegroundColor Red
        Break
    }

    $global:appLibraryPath = "$MCMApplicationLibraryLocation\$global:author\$global:appName\$global:version"
    Try{
        Write-Host "Attemping to copy from temp directory to $global:appLibraryPath" -ForegroundColor Blue
        Copy-Item "$contentPath\$App\*" -Destination $global:appLibraryPath -Recurse -Exclude "*.yaml"
        Write-Host "Successfully copied content" -ForegroundColor Green
    }
    Catch{
        Write-Host "Failed to copy content." -ForegroundColor Red
    }
}

Function MCMWork(){
    Set-Location "$($SiteCode):\"
    Try{
        Write-Host "Attemping to make MCM Application for $global:appName $global:version." -ForegroundColor Blue
        If ($global:installCode -ne $null){
            $Detection = New-CMDetectionClauseRegistryKeyValue -Hive LocalMachine -KeyName "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$global:installCode" -PropertyType Version -ValueName "DisplayVersion" -Value -ExpectedValue $global:version -ExpressionOperator GreaterEquals
            New-CMApplication -Name "$global:appName $global:version" -Publisher $global:author -SoftwareVersion $global:version -AutoInstall $true | Out-Null
            Add-CMScriptDeploymentType -ApplicationName "$global:appName $global:version" -DeploymentTypeName "Install $global:appName $global:version" -InstallCommand "`"$global:installContent`" $global:installArgument" -AddDetectionClause $Detection -ContentLocation $global:appLibraryPath -InstallationBehaviorType InstallForSystem -LogonRequirementType WhetherOrNotUserLoggedOn | Out-Null
            Write-Host "Created $global:appName $global:version application." -ForegroundColor Green
            [bool]$AppDeploy = $true
        }
        else{
            Write-Host "Detection code for $global:appName doesn't exist in the YAML. This will need to be a package"
            New-CMPackage -Name "$global:appName $global:version" -Manufacturer $global:author -Version $global:version -Path $global:appLibraryPath | Out-Null
            New-CMProgram -StandardProgramName "Install $global:appName" -CommandLine "`"$global:installContent`" $global:installArgument" -PackageName "$global:appName $global:version" -ProgramRunType WhetherOrNotUserIsLoggedOn -RunMode RunWithAdministrativeRights -RunType Hidden | Out-Null
        }
        
    }
    Catch{
        Write-Host "Failed to make application" -ForegroundColor Red
        Break
    }

    Try{
        Write-Host "Starting distribution of $global:appName to $DPGroups" -ForegroundColor Blue
        If ($AppDeploy){
            Start-CMContentDistribution -ApplicationName "$global:appName $global:version" -DistributionPointGroupName $DPGroups | Out-Null
        }
        Else{
            Start-CMContentDistribution -PackageName "$global:appName $global:version" -DistributionPointGroupName $DPGroups | Out-Null
        }
        Write-Host "Successfully started distribution of $global:appName $global:version." -ForegroundColor Green
    }
    Catch{
        Write-Host "Failed to start distribution" -ForegroundColor Red
    }

    $date = Get-Date -Format 'MM-yyyy'
    If ($AppDeploy){
        if (!(Get-CMFolder -name "$MCMFolderName" -ParentFolderPath "Application")){
            Write-Host "Attempting to make $MCMFolderName folder." -ForegroundColor Blue
            New-CMFolder -ParentFolderPath "Application" -Name "$MCMFolderName" | Out-Null
            Write-Host "$MCMFolderName folder created" -ForegroundColor Green
        }
        else{
            Write-Host "$MCMFolderName folder already exists. Continuing with that." -ForegroundColor Yellow
        }
        If (!(Get-CMFolder -Name $date -ParentFolderPath "Application\$MCMFolderName")){
            Write-Host "Attempting to make folder for $date." -ForegroundColor Blue
            New-CMFolder -ParentFolderPath "Application\$MCMFolderName" -Name $date | Out-Null
            Write-Host "$date folder created" -ForegroundColor Green
            }
        else{
            Write-Host "$date folder already exists. Continuing with that." -ForegroundColor Yellow
        }
        $prog = Get-CMApplication -Name "$global:appName $global:version"
        Move-CMObject -FolderPath "Application\$MCMFolderName\$date" -InputObject $prog
        Write-Host "Moved to Applications\$MCMFolderName\$date folder." -ForegroundColor Green

        Write-Host ""
        Write-Host "Application creation complete" -ForegroundColor Green
        Write-Host ""
    }
    else{
        if (!(Get-CMFolder -name "$MCMFolderName" -ParentFolderPath "Package")){
            Write-Host "Attempting to make $MCMFolderName folder." -ForegroundColor Blue
            New-CMFolder -ParentFolderPath "Package" -Name "$MCMFolderName" | Out-Null
            Write-Host "$MCMFolderName folder created" -ForegroundColor Green
        }
        else{
            Write-Host "$MCMFolderName folder already exists. Continuing with that." -ForegroundColor Yellow
        }
        If (!(Get-CMFolder -Name $date -ParentFolderPath "Package\$MCMFolderName")){
            Write-Host "Attempting to make folder for $date." -ForegroundColor Blue
            New-CMFolder -ParentFolderPath "Package\$MCMFolderName" -Name $date | Out-Null
            Write-Host "$date folder created" -ForegroundColor Green
        }
        else{
            Write-Host "$date folder already exists. Continuing with that." -ForegroundColor Yellow
        }
        $prog = Get-CMPackage -Name "$global:appName $global:version" -Fast
        Move-CMObject -FolderPath "Package\$MCMFolderName\$date" -InputObject $prog
        Write-Host "Moved to Packages\$MCMFolderName\$date folder." -ForegroundColor Green

        Write-Host ""
        Write-Host "Package creation complete" -ForegroundColor Green
        Write-Host ""
    }
    


    If ($makeTaskSequence){
        if (!(Get-CMTaskSequence -name $date -Fast)){
            Write-Host "$date Task Sequence not found, creating." -ForegroundColor Blue
            New-CMTaskSequence -CustomTaskSequence -Name $date | Out-Null
            Write-Host "Created TS." -ForegroundColor Green
        }
        else{
            Write-Host "$date TS already exits, using that." -ForegroundColor Yellow
            }
        Try{
            Write-Host "Attempting to add $global:appName to $date TS" -ForegroundColor Blue
            $shortName = $global:appName.subString(0, [System.Math]::Min(40, $global:appName.Length)) 
            if ($AppDeploy){
                $prog = Get-CMApplication -Name "$global:appName $global:version"
                $step = New-CMTSStepInstallApplication -Name "Install $shortName" -Application $prog
                $ts = Get-CMTaskSequence -Name $date -Fast
                $ts | Add-CMTaskSequenceStep -Step $step
            }
            Else{
                $prog = Get-CMProgram -PackageName "$global:appName $global:version" -ProgramName "Install $global:appName"
                $step = New-CMTSStepInstallSoftware -Name "Install $shortName" -Program $prog
                $ts = Get-CMTaskSequence -Name $date -Fast
                $ts | Add-CMTaskSequenceStep -Step $step
            }

            Write-Host "Successfully added $global:appName $global:version to TS." -ForegroundColor Green
        }
        Catch{
            Write-Host "Failed to add to TS." -ForegroundColor Red
        }

    }
    Write-Host ""
}
    

foreach ($application in $appsToDownload){
    Set-Location $here
    if ($DownloadOnly){
        Download $application
    }
    elseif ($MCMImportOnly){
        GetDetails $application $ImportContentPath
        CopytoLibrary $application $ImportContentPath
        MCMWork
    }
    Else{
        Download $application
        GetDetails $application $ImportContentPath
        CopytoLibrary $application $ImportContentPath
        MCMWork
        [bool]$CleanNeeded = $true
    }
}
if ($CleanNeeded){
    Set-Location $here
    Write-Host "Cleaning up temp directory" -ForegroundColor Blue
    Remove-Item "$DownloadLocation\Wingot_Downloads" -Recurse -Force
    Write-Host "$DownloadLocation\Wingot_Downloads Removed." -ForegroundColor Green
}
