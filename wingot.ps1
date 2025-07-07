[CmdletBinding()]
param(
    [Parameter(HelpMessage = "Only download applications without importing to MCM")]
    [switch]$DownloadOnly = $false,
    
    [Parameter(HelpMessage = "Only import to MCM without downloading")]
    [switch]$MCMImportOnly = $false,
    
    [Parameter(HelpMessage = "Location to download applications")]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$DownloadLocation = $PWD,
    
    [Parameter(HelpMessage = "Path to content for import")]
    [string]$ImportContentPath = "$PWD\Wingot_Downloads",
    
    [Parameter(HelpMessage = "Create task sequence for applications")]
    [switch]$MakeTaskSequence = $false
)

#region Configuration
$Config = @{
    MCMSiteCode = "CHQ"
    MCMPrimarySiteServer = "CM1.corp.contoso.com"
    MCMApplicationLibraryLocation = "\\localhost\c$\Packages\Apps"
    DPGroups = @("Corp DPs")
    MCMFolderName = "Wingot"
    TempFolderName = "Wingot_Downloads"
}

$AppsToDownload = @(
    "DominikReichl.KeePass",
    "Microsoft.VCRedist.2015+.x64",
    "Microsoft.VCRedist.2015+.x86",
    "Oracle.JavaRuntimeEnvironment",
    "Citrix.Workspace.LTSR"
)
#endregion

#region Helper Functions
function Write-ColorOutput {
    param(
        [string]$Message,
        [ValidateSet("Info", "Success", "Warning", "Error")]
        [string]$Type = "Info"
    )
    
    $colors = @{
        Info = "Blue"
        Success = "Green"
        Warning = "Yellow"
        Error = "Red"
    }
    
    Write-Host $Message -ForegroundColor $colors[$Type]
}

function Test-Prerequisites {
    if (-not $DownloadOnly) {
        # Test YAML module
        try {
            Import-Module (Join-Path $PSScriptRoot "powershell-yaml") -ErrorAction Stop
            Write-ColorOutput "YAML module loaded successfully" "Success"
        }
        catch {
            Write-ColorOutput "Unable to find YAML module. Please install it first." "Error"
            return $false
        }
        
        # Test Configuration Manager module
        try {
            if (-not (Get-Module ConfigurationManager)) {
                $CMPath = "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
                if (-not (Test-Path $CMPath)) {
                    throw "Configuration Manager console not found"
                }
                Import-Module $CMPath -ErrorAction Stop
            }
            
            if (-not (Get-PSDrive -Name $Config.MCMSiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
                New-PSDrive -Name $Config.MCMSiteCode -PSProvider CMSite -Root $Config.MCMPrimarySiteServer -ErrorAction Stop | Out-Null
            }
            
            Write-ColorOutput "Configuration Manager connection established" "Success"
        }
        catch {
            Write-ColorOutput "Unable to connect to Configuration Manager: $($_.Exception.Message)" "Error"
            return $false
        }
    }
    
    return $true
}

function Initialize-DownloadFolder {
    param([string]$Path)
    
    $downloadPath = Join-Path $Path $Config.TempFolderName
    
    if (-not (Test-Path $downloadPath)) {
        New-Item -Path $downloadPath -ItemType Directory | Out-Null
        Write-ColorOutput "Created download folder: $downloadPath" "Success"
    }
    else {
        Write-ColorOutput "Download folder already exists: $downloadPath" "Warning"
    }
    
    return $downloadPath
}
#endregion

#region Core Functions
function Invoke-ApplicationDownload {
    param(
        [string]$AppName,
        [string]$DownloadPath
    )
    
    Write-ColorOutput "Starting download for: $AppName" "Info"
    
    $appDownloadPath = Join-Path $DownloadPath $AppName
    
    # Clean up existing download if present
    if (Test-Path $appDownloadPath) {
        Remove-Item $appDownloadPath -Recurse -Force -ErrorAction SilentlyContinue
        Write-ColorOutput "Cleaned up previous download for: $AppName" "Warning"
    }
    
    # Create app-specific folder
    New-Item -Path $appDownloadPath -ItemType Directory | Out-Null
    
    try {
        $processArgs = @{
            FilePath = "winget"
            ArgumentList = @("download", $AppName, "-d", $appDownloadPath)
            Wait = $true
            NoNewWindow = $true
            ErrorAction = "Stop"
        }
        
        Start-Process @processArgs
        Write-ColorOutput "Successfully downloaded: $AppName" "Success"
        return $true
    }
    catch {
        Write-ColorOutput "Failed to download $AppName`: $($_.Exception.Message)" "Error"
        return $false
    }
}

function Get-ApplicationDetails {
    param(
        [string]$AppName,
        [string]$ContentPath
    )
    
    Write-ColorOutput "Extracting details for: $AppName" "Info"
    
    $appPath = Join-Path $ContentPath $AppName
    
    try {
        $yamlFile = Get-ChildItem -Path $appPath -Recurse -Filter "*.yaml" | Select-Object -First 1
        if (-not $yamlFile) {
            throw "No YAML file found for $AppName"
        }
        
        $yaml = Get-Content $yamlFile.FullName | ConvertFrom-Yaml
        
        $appDetails = @{
            Name = $yaml.PackageName
            Author = $yaml.Author.TrimEnd('.')
            Version = $yaml.PackageVersion
            InstallCode = $yaml.Installers.ProductCode
            InstallContent = (Get-ChildItem -Path $appPath -Recurse -Exclude "*.yaml" | Select-Object -First 1).Name
            InstallArgument = $yaml.Installers.InstallerSwitches.Silent
        }
        
        Write-ColorOutput "Successfully extracted details for: $AppName" "Success"
        return $appDetails
    }
    catch {
        Write-ColorOutput "Failed to extract details for $AppName`: $($_.Exception.Message)" "Error"
        return $null
    }
}

function Copy-ToLibrary {
    param(
        [string]$AppName,
        [string]$ContentPath,
        [hashtable]$AppDetails
    )
    
    Write-ColorOutput "Copying $AppName to library" "Info"
    
    $libraryPath = Join-Path $Config.MCMApplicationLibraryLocation $AppDetails.Author
    $appLibraryPath = Join-Path $libraryPath $AppDetails.Name
    $versionPath = Join-Path $appLibraryPath $AppDetails.Version
    
    try {
        # Create directory structure
        @($libraryPath, $appLibraryPath) | ForEach-Object {
            if (-not (Test-Path $_)) {
                New-Item -Path $_ -ItemType Directory | Out-Null
                Write-ColorOutput "Created directory: $_" "Success"
            }
        }
        
        # Check if version already exists
        if (Test-Path $versionPath) {
            Write-ColorOutput "Version $($AppDetails.Version) already exists for $($AppDetails.Name)" "Error"
            return $null
        }
        
        # Create version directory and copy content
        New-Item -Path $versionPath -ItemType Directory | Out-Null
        $sourcePath = Join-Path $ContentPath $AppName
        Copy-Item -Path "$sourcePath\*" -Destination $versionPath -Recurse -Exclude "*.yaml"
        
        Write-ColorOutput "Successfully copied content to: $versionPath" "Success"
        return $versionPath
    }
    catch {
        Write-ColorOutput "Failed to copy to library: $($_.Exception.Message)" "Error"
        return $null
    }
}

function Invoke-MCMWork {
    param(
        [hashtable]$AppDetails,
        [string]$LibraryPath
    )
    
    $originalLocation = Get-Location
    
    try {
        Set-Location "$($Config.MCMSiteCode):\"
        
        $isApplication = $null -ne $AppDetails.InstallCode
        $objectName = "$($AppDetails.Name) $($AppDetails.Version)"
        
        if ($isApplication) {
            Write-ColorOutput "Creating MCM Application: $objectName" "Info"
            
            $detection = New-CMDetectionClauseRegistryKeyValue -Hive LocalMachine -KeyName "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$($AppDetails.InstallCode)" -PropertyType Version -ValueName "DisplayVersion" -Value -ExpectedValue $AppDetails.Version -ExpressionOperator GreaterEquals
            
            New-CMApplication -Name $objectName -Publisher $AppDetails.Author -SoftwareVersion $AppDetails.Version -AutoInstall $true | Out-Null
            
            Add-CMScriptDeploymentType -ApplicationName $objectName -DeploymentTypeName "Install $objectName" -InstallCommand "`"$($AppDetails.InstallContent)`" $($AppDetails.InstallArgument)" -AddDetectionClause $detection -ContentLocation $LibraryPath -InstallationBehaviorType InstallForSystem -LogonRequirementType WhetherOrNotUserLoggedOn | Out-Null
            
            $folderType = "Application"
        }
        else {
            Write-ColorOutput "Creating MCM Package: $objectName" "Info"
            
            New-CMPackage -Name $objectName -Manufacturer $AppDetails.Author -Version $AppDetails.Version -Path $LibraryPath | Out-Null
            
            New-CMProgram -StandardProgramName "Install $($AppDetails.Name)" -CommandLine "`"$($AppDetails.InstallContent)`" $($AppDetails.InstallArgument)" -PackageName $objectName -ProgramRunType WhetherOrNotUserIsLoggedOn -RunMode RunWithAdministrativeRights -RunType Hidden | Out-Null
            
            $folderType = "Package"
        }
        
        # Distribute content
        Write-ColorOutput "Starting distribution to: $($Config.DPGroups -join ', ')" "Info"
        if ($isApplication) {
            Start-CMContentDistribution -ApplicationName $objectName -DistributionPointGroupName $Config.DPGroups | Out-Null
        }
        else {
            Start-CMContentDistribution -PackageName $objectName -DistributionPointGroupName $Config.DPGroups | Out-Null
        }
        
        # Organize in folders
        $dateFolder = Get-Date -Format 'MM-yyyy'
        $folderPath = "$folderType\$($Config.MCMFolderName)\$dateFolder"
        
        New-MCMFolderStructure -FolderType $folderType -DateFolder $dateFolder
        
        if ($isApplication) {
            $object = Get-CMApplication -Name $objectName
        }
        else {
            $object = Get-CMPackage -Name $objectName -Fast
        }
        
        Move-CMObject -FolderPath $folderPath -InputObject $object
        Write-ColorOutput "Moved to: $folderPath" "Success"
        
        # Add to task sequence if requested
        if ($MakeTaskSequence) {
            Add-ToTaskSequence -AppDetails $AppDetails -IsApplication $isApplication -DateFolder $dateFolder
        }
        
        Write-ColorOutput "MCM work completed for: $($AppDetails.Name)" "Success"
    }
    catch {
        Write-ColorOutput "MCM work failed: $($_.Exception.Message)" "Error"
    }
    finally {
        Set-Location $originalLocation
    }
}

function New-MCMFolderStructure {
    param(
        [string]$FolderType,
        [string]$DateFolder
    )
    
    $mainFolderPath = "$FolderType\$($Config.MCMFolderName)"
    $dateFolderPath = "$mainFolderPath\$DateFolder"
    
    @(
        @{Path = $mainFolderPath; Name = $Config.MCMFolderName; Parent = $FolderType}
        @{Path = $dateFolderPath; Name = $DateFolder; Parent = $mainFolderPath}
    ) | ForEach-Object {
        if (-not (Get-CMFolder -Name $_.Name -ParentFolderPath $_.Parent -ErrorAction SilentlyContinue)) {
            New-CMFolder -ParentFolderPath $_.Parent -Name $_.Name | Out-Null
            Write-ColorOutput "Created folder: $($_.Path)" "Success"
        }
    }
}

function Add-ToTaskSequence {
    param(
        [hashtable]$AppDetails,
        [bool]$IsApplication,
        [string]$DateFolder
    )
    
    try {
        if (-not (Get-CMTaskSequence -Name $DateFolder -Fast)) {
            New-CMTaskSequence -CustomTaskSequence -Name $DateFolder | Out-Null
            Write-ColorOutput "Created task sequence: $DateFolder" "Success"
        }
        
        $objectName = "$($AppDetails.Name) $($AppDetails.Version)"
        $shortName = $AppDetails.Name.Substring(0, [Math]::Min(40, $AppDetails.Name.Length))
        
        if ($IsApplication) {
            $object = Get-CMApplication -Name $objectName
            $step = New-CMTSStepInstallApplication -Name "Install $shortName" -Application $object
        }
        else {
            $object = Get-CMProgram -PackageName $objectName -ProgramName "Install $($AppDetails.Name)"
            $step = New-CMTSStepInstallSoftware -Name "Install $shortName" -Program $object
        }
        
        $taskSequence = Get-CMTaskSequence -Name $DateFolder -Fast
        $taskSequence | Add-CMTaskSequenceStep -Step $step
        
        Write-ColorOutput "Added $objectName to task sequence" "Success"
    }
    catch {
        Write-ColorOutput "Failed to add to task sequence: $($_.Exception.Message)" "Error"
    }
}

function Remove-TempDirectory {
    param([string]$Path)
    
    if (Test-Path $Path) {
        Remove-Item $Path -Recurse -Force
        Write-ColorOutput "Cleaned up temporary directory: $Path" "Success"
    }
}
#endregion

#region Main Execution
function Main {
    Write-ColorOutput "Starting Wingot Application Deployment Script" "Info"
    
    if (-not (Test-Prerequisites)) {
        Write-ColorOutput "Prerequisites check failed. Exiting." "Error"
        return
    }
    
    $originalLocation = Get-Location
    $cleanupNeeded = $false
    
    try {
        foreach ($app in $AppsToDownload) {
            Set-Location $originalLocation
            Write-ColorOutput "Processing: $app" "Info"
            
            if ($DownloadOnly) {
                $downloadPath = Initialize-DownloadFolder -Path $DownloadLocation
                Invoke-ApplicationDownload -AppName $app -DownloadPath $downloadPath
            }
            elseif ($MCMImportOnly) {
                $appDetails = Get-ApplicationDetails -AppName $app -ContentPath $ImportContentPath
                if ($appDetails) {
                    $libraryPath = Copy-ToLibrary -AppName $app -ContentPath $ImportContentPath -AppDetails $appDetails
                    if ($libraryPath) {
                        Invoke-MCMWork -AppDetails $appDetails -LibraryPath $libraryPath
                    }
                }
            }
            else {
                $downloadPath = Initialize-DownloadFolder -Path $DownloadLocation
                
                if (Invoke-ApplicationDownload -AppName $app -DownloadPath $downloadPath) {
                    $appDetails = Get-ApplicationDetails -AppName $app -ContentPath $downloadPath
                    if ($appDetails) {
                        $libraryPath = Copy-ToLibrary -AppName $app -ContentPath $downloadPath -AppDetails $appDetails
                        if ($libraryPath) {
                            Invoke-MCMWork -AppDetails $appDetails -LibraryPath $libraryPath
                            $cleanupNeeded = $true
                        }
                    }
                }
            }
            
            Write-ColorOutput "Completed processing: $app" "Success"
        }
        
        if ($cleanupNeeded) {
            Remove-TempDirectory -Path (Join-Path $DownloadLocation $Config.TempFolderName)
        }
        
        Write-ColorOutput "Script execution completed successfully" "Success"
    }
    catch {
        Write-ColorOutput "Script execution failed: $($_.Exception.Message)" "Error"
    }
    finally {
        Set-Location $originalLocation
    }
}

# Execute main function
Main
#endregion