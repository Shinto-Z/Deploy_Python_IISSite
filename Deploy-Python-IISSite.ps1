# Deploy IIS Python Site
# Supports deploying Flask/Python apps via FastCGI with proper config management
# Usage: .\Deploy-Python-IISSite.ps1 -Mode Deploy|Undeploy [-SiteName <name>] [-SiteLocation <path>] [-AppPoolName <name>] [-Port <port>] [-KeepAppPool <$true/$false] [-KeepLocation <$true/$false>] [-NoConfirm <$true/$false>] 
param(
    [ValidateSet('Deploy', 'Undeploy', 'Reset')]
    [string]$Mode,
    [string]$SiteName,
    [string]$SiteLocation,
    [string]$AppPoolName,
    [int]$Port = 8080,
    [bool]$KeepAppPool = $true,
    [bool]$KeepLocation = $true,
    [switch]$NoConfirm = $true
)

if (-not $PSBoundParameters.ContainsKey('KeepLocation')) { $KeepLocation = $true }
if (-not $PSBoundParameters.ContainsKey('KeepAppPool')) { $KeepAppPool = $true }

function Write-Success {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Green
}

function Write-Error-Custom {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Red
}

function Prompt-YesNo {
    param([string]$Question)
    $response = Read-Host "$Question (Y/N)"
    return $response -eq 'Y' -or $response -eq 'y'
}

function Test-PythonPackage {
    param(
        [string]$PackageName,
        [string]$PythonExe
    )

    $cmd = "`"$PythonExe`" -c `"import $PackageName`""
    cmd /c $cmd 2>$null | Out-Null
    return ($LASTEXITCODE -eq 0)
}

function Test-NetworkAvailable {
    try {
        Test-Connection -ComputerName "pypi.org" -Count 1 -Quiet -ErrorAction Stop
    }
    catch {
        return $false
    }
}

function Test-IISFeatures {

    $RequiredFeatures = @("Web-Default-Doc","Web-Dir-Browsing","Web-Http-Errors","Web-Static-Content","Web-Http-Logging","Web-Stat-Compression","Web-Filtering","Web-CGI","Web-Mgmt-Console","Web-Scripting-Tools")

    Write-Host ""
    Write-Host "Checking IIS feature installation status..."

    foreach ($feature in $RequiredFeatures) {

        $featureState = Get-WindowsFeature -Name $feature

        if ($null -eq $featureState) {
            Write-Host "Feature '$feature' does not exist on this OS."
            continue
        }

        if (!$featureState.Installed) {
            Write-Host "$feature is NOT installed."

            $answer = Read-Host "Install $feature now? (Y/N)"

            if ($answer -match '^[Yy]$') {
                Write-Host "Installing $feature..."
                Install-WindowsFeature -Name $feature -IncludeManagementTools | Out-Null

                $featureState = Get-WindowsFeature -Name $feature
                if ($featureState.Installed) { Write-Host "Successfully installed $feature." }
                else { Write-Host "Failed to install $feature." }
            }
            else { Write-Host "Skipping installation of $feature." }
        }
    }

    Write-Host ""
    Write-Host "Feature check complete."
}

function Add-FastCgiApplication {
    param(
        [string]$PythonExe,
        [string]$ArgumentsPath,
        [string]$SiteLocation
    )

    Import-Module WebAdministration

    $fcgiPath = "system.webServer/fastCgi"

    # Check if it already exists
    $existing = Get-WebConfigurationProperty -Filter "$fcgiPath/application" -Name "." |
        Where-Object { $_.fullPath -eq $PythonExe -and $_.arguments -eq $ArgumentsPath }

    if ($existing) {
        Write-Host "FastCGI application already exists for $PythonExe"
        return
    }
    else {
        Write-Host "Adding FastCGI application for $PythonExe"
        
        #Add the FastCGI application WITHOUT env vars
        Add-WebConfigurationProperty `
            -pspath 'MACHINE/WEBROOT/APPHOST' `
            -filter $fcgiPath `
            -name "." `
            -value @{
                fullPath  = $PythonExe
                arguments = $ArgumentsPath
            }
    }

    $fastCgiOk = $true 
    try {
        # 2. Add environment variables as child elements
        $envPath = "$fcgiPath/application[@fullPath='$PythonExe' and @arguments='$ArgumentsPath']/environmentVariables"

        Add-WebConfigurationProperty `
            -pspath 'MACHINE/WEBROOT/APPHOST' `
            -filter $envPath `
            -name "." `
            -value @{ name="WSGI_HANDLER"; value="app.app" } -ErrorAction Stop

        Add-WebConfigurationProperty `
            -pspath 'MACHINE/WEBROOT/APPHOST' `
            -filter $envPath `
            -name "." `
            -value @{ name="PYTHONPATH"; value=$SiteLocation } -ErrorAction Stop
    }
    catch {
        $fastCgiOk = $false
        Write-Error-Custom "Error adding FastCGI environment variables: $($_.Exception.Message)"
    }
    if($fastCgiOk){ Write-Host "FastCGI application added with environment variables." }
}

function Remove-FastCgiApplication {
    param(
        [string]$PythonExe,
        [string]$ArgumentsPath
    )

    $configPath = 'C:\Windows\System32\inetsrv\config\applicationHost.config'
    [xml]$xml = Get-Content $configPath

    $xpath = "//fastCgi/application[@fullPath='$PythonExe' and @arguments='$ArgumentsPath']"
    $nodes = $xml.SelectNodes($xpath)

    if ($nodes.Count -gt 0) {
        foreach ($node in $nodes) {
            Write-Host "Removing FastCGI entry: $($node.fullPath) | $($node.arguments)"
            $node.ParentNode.RemoveChild($node) | Out-Null
        }

        $xml.Save($configPath)
        Write-Host "FastCGI entry removed."
    }
    else { Write-Host "No FastCGI entry found for: $PythonExe | $ArgumentsPath" }
}

function Get-PotentiallyInvalidFastCgiApplications {
    [CmdletBinding()]
    param()

    Import-Module WebAdministration

    $fcgiPath = "system.webServer/fastCgi"

    # Get all FastCGI applications
    $apps = Get-WebConfigurationProperty -Filter $fcgiPath -Name "application"

    $invalid = @()

    foreach ($app in $apps) {
        $path = $app.fullPath

        # Skip relative paths (IIS may resolve these)
        if (-not [System.IO.Path]::IsPathRooted($path)) { continue }

        # Skip UNC paths (Test-Path may fail due to identity permissions)
        if ($path.StartsWith("\\")) { continue }

        # If the executable truly does not exist, flag it
        if (-not (Test-Path $path)) { $invalid += $app }
    }

    return $invalid
}

function Ensure-CleanApplicationHostConfig {
    $cleanBackup = Join-Path $PSScriptRoot 'applicationHost.config.clean'
    
    if (-not (Test-Path $cleanBackup)) {
        Write-Host "`nClean backup not found at: $cleanBackup" -ForegroundColor Yellow
        Write-Host "This file is required as a baseline for rollback operations."
        
        if (-not (Prompt-YesNo "Create clean backup from current applicationHost.config?")) {
            Write-Error-Custom "Cannot proceed without clean backup."
            exit 1
        }
        
        # Create the clean backup
        try {
            Copy-Item -Path 'C:\Windows\System32\inetsrv\config\applicationHost.config' -Destination $cleanBackup -Force
            Write-Success "Clean backup created: $cleanBackup"
        } catch {
            Write-Error-Custom "Failed to create clean backup: $_"
            exit 1
        }
        
        # Verify it was created
        if (-not (Test-Path $cleanBackup)) {
            Write-Error-Custom "Clean backup verification failed - file does not exist."
            exit 1
        }
    }
    
    return $cleanBackup
}

function Backup-ApplicationHostConfig {
    $configPath = 'C:\Windows\System32\inetsrv\config\applicationHost.config'
    $backupPath = "$configPath.backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    
    Copy-Item -Path $configPath -Destination $backupPath -Force
    Write-Host "Backup created: $backupPath" -ForegroundColor Cyan
    return $backupPath
}

function Sync-IISNativeModules {
    [CmdletBinding()]
    param()

    Import-Module WebAdministration

    # Mapping: Windows Feature → Required IIS Modules
    $FeatureToModules = @{
        "Web-CGI"              = @("CgiModule", "FastCgiModule")
        "Web-ASP"              = @("AspModule")
        "Web-ISAPI-Ext"        = @("IsapiModule")
        "Web-ISAPI-Filter"     = @("IsapiFilterModule")
        "Web-WebSockets"       = @("WebSocketModule")
        "Web-Http-Logging"     = @("HttpLoggingModule")
        "Web-Stat-Compression" = @("StaticCompressionModule")
        "Web-Dyn-Compression"  = @("DynamicCompressionModule")
        "Web-Filtering"        = @("RequestFilteringModule")
    }

    # Determine installed features
    $installed = Get-WindowsFeature | Where-Object { $_.Installed } | Select-Object -ExpandProperty Name

    # Build list of required modules
    $requiredModules = @()
    foreach ($feature in $installed) {
        if ($FeatureToModules.ContainsKey($feature)) { $requiredModules += $FeatureToModules[$feature] }
    }

    # Remove duplicates
    $requiredModules = $requiredModules | Sort-Object -Unique

    # Get current modules in applicationHost.config
    $currentModules = (Get-WebConfiguration "//globalModules/add").name

    foreach ($module in $requiredModules) {
        if ($currentModules -notcontains $module) {
            Write-Host "Adding missing IIS module: $module"

            # Determine DLL path
            switch ($module) {
                "CgiModule"              { $dll = "$env:windir\System32\inetsrv\cgi.dll" }
                "FastCgiModule"          { $dll = "$env:windir\System32\inetsrv\iisfcgi.dll" }
                "AspModule"              { $dll = "$env:windir\System32\inetsrv\asp.dll" }
                "IsapiModule"            { $dll = "$env:windir\System32\inetsrv\isapi.dll" }
                "IsapiFilterModule"      { $dll = "$env:windir\System32\inetsrv\filter.dll" }
                "WebSocketModule"        { $dll = "$env:windir\System32\inetsrv\websocket.dll" }
                "StaticCompressionModule" { $dll = "$env:windir\System32\inetsrv\compstat.dll" }
                "DynamicCompressionModule" { $dll = "$env:windir\System32\inetsrv\compdyn.dll" }
                "RequestFilteringModule" { $dll = "$env:windir\System32\inetsrv\modrqflt.dll" }
                default                  { continue }
            }

            Add-WebConfiguration "//globalModules" -Value @{
                name  = $module
                image = $dll
            }
        }
    }

    Write-Host "IIS module synchronization complete."
}

function Reset-ApplicationHostConfig {
    $cleanBackup = Ensure-CleanApplicationHostConfig
    $configPath = 'C:\Windows\System32\inetsrv\config\applicationHost.config'
    
    Stop-Service W3SVC -Force
    Stop-Service WAS -Force
    Copy-Item -Path $cleanBackup -Destination $configPath -Force
    Start-Service WAS
    Start-Service W3SVC
    
    Sync-IISNativeModules
    $potentialUnusedFastCGI = Get-PotentiallyInvalidFastCgiApplications

    foreach ($app in $potentialUnusedFastCGI) {
        Write-Host "FastCGI $($app.fullPath) : $($app.arguments) entry appears invalid or unused."
        Write-Host ""

        $response = Read-Host "Would you like to delete this FastCGI entry? (y/n)"

        if ($response -eq 'y') {
            Clear-WebConfiguration -Filter "system.webServer/fastCgi/application[@fullPath='$($app.fullPath)' and @arguments='$($app.arguments)']"

            # Verify deletion
            $fastCgiStillExists = Get-WebConfigurationProperty `
                -Filter "system.webServer/fastCgi/application" `
                -Name "." | Where-Object { 
                    $_.fullPath -eq $app.fullPath -and
                    $_.arguments -eq $app.arguments
                } 
                
            if ($fastCgiStillExists) { Write-Error-Custom "Failed to delete FastCGI entry." }
            else { Write-Host "Entry deleted." }

        } else { Write-Host "Entry skipped." }

        Write-Host ""
    }

    Write-Host "ApplicationHost.config cleaning has concluded." -ForegroundColor Cyan
}

function Copy-PythonPackage {
    param(
        [string]$PackageName,
        [string]$SourceSitePackages,
        [string]$TargetSitePackages
    )

    Write-Host "Copying $PackageName from $SourceSitePackages to $TargetSitePackages"

    $copied = $false

    # Detect actual folder name (case-insensitive)
    $pkgFolder = Get-ChildItem $SourceSitePackages -Directory |
                 Where-Object { $_.Name -ieq $PackageName } |
                 Select-Object -First 1

    if ($pkgFolder) {
        Write-Host "Found folder package: $($pkgFolder.FullName)"
        Copy-Item -Path $pkgFolder.FullName -Destination (Join-Path $TargetSitePackages $pkgFolder.Name) -Recurse -Force
        $copied = $true
    } else { Write-Host "No folder package found for $PackageName" }

    # Single-file module
    $pkgFile = Get-ChildItem $SourceSitePackages -File |
               Where-Object { $_.Name -ieq "$PackageName.py" } |
               Select-Object -First 1

    if ($pkgFile) {
        Write-Host "Found file module: $($pkgFile.FullName)"
        Copy-Item -Path $pkgFile.FullName -Destination (Join-Path $TargetSitePackages $pkgFile.Name) -Force
        $copied = $true
    } else { Write-Host "No file module found for $PackageName.py" }

    # dist-info metadata
    $distInfo = Get-ChildItem $SourceSitePackages -Directory | Where-Object { $_.Name -match "^$PackageName-[\d\.]+\.dist-info$" }

    if ($distInfo) {
        foreach ($info in $distInfo) {
            Write-Host "Found dist-info: $($info.FullName)"
            Copy-Item -Path $info.FullName -Destination (Join-Path $TargetSitePackages $info.Name) -Recurse -Force
        }
        $copied = $true
    } else { Write-Host "No dist-info found for $PackageName" }

    return $copied
}

function Get-SystemPython {
    # Try python.exe on PATH
    $cmd = Get-Command python.exe -ErrorAction SilentlyContinue
    $candidate = if ($cmd) { $cmd.Source } else { $null }

    # Reject Windows Store shim
    if ($candidate -and $candidate -notmatch "AppData\\Local\\Microsoft\\WindowsApps") { return $candidate }

    # Try py launcher
    try {
        $pyPath = py -3 -c "import sys; print(sys.executable)" 2>$null
        if ($pyPath -and (Test-Path $pyPath)) { return $pyPath }
    } catch {}

    # Try common install locations
    $common = @(
        "C:\Python311\python.exe",
        "C:\Python312\python.exe",
        "$env:ProgramFiles\Python311\python.exe",
        "$env:ProgramFiles\Python312\python.exe"
    )

    foreach ($path in $common) { if (Test-Path $path) { return $path } }

    return $null
}

function Deploy-IISSite {
    Write-Host "`n=== IIS Python Site Deployment ===" -ForegroundColor Cyan
    
    # Verify IIS is installed
    if (-not (Get-Service W3SVC -ErrorAction SilentlyContinue)) {
        Write-Error-Custom "IIS is not installed. Please install IIS first."
        exit 1
    }
    
    # Get deployment parameters (from command-line args or prompt user)
    if (-not $SiteLocation) {
        $SiteLocation = Read-Host "Enter site location (default: C:\iis-python-site)"
        if ([string]::IsNullOrWhiteSpace($SiteLocation)) { $SiteLocation = "C:\iis-python-site" }
    }
    
    if (-not $SiteName) {
        $SiteName = Read-Host "Enter site name (default: PythonSite)"
        if ([string]::IsNullOrWhiteSpace($SiteName)) { $SiteName = "PythonSite" }
    }
    
    if ($Port -eq 0) {
        $portInput = Read-Host "Enter port number (default: 8080)"
        if (-not [string]::IsNullOrWhiteSpace($portInput)) { $Port = [int]$portInput } else { $Port = 8080 }
    }

    if(-not $AppPoolName) {
        # Get or create application pool
        $appPoolInput = Read-Host "Enter application pool name (default: DefaultAppPool) or press Enter to create new"
        if ([string]::IsNullOrWhiteSpace($appPoolInput)) {
            $AppPoolName = "DefaultAppPool"
        } elseif ($appPoolInput -eq "new") {
            $AppPoolName = "$siteName-AppPool"
        } else {
            $AppPoolName = $appPoolInput
        }
    }
    
    # Check if app pool exists, create if needed
    $existingPool = Get-Item "IIS:\AppPools\$AppPoolName" -ErrorAction SilentlyContinue
    if (-not $existingPool) {
        Write-Host "Creating application pool: $AppPoolName"

        # Try to create the pool
        try { New-WebAppPool -Name $AppPoolName -ErrorAction Stop | Out-Null }
        catch {
            Write-Error-Custom "Failed to create application pool: $($_.Exception.Message)"
            return
        }

        Start-Sleep -Milliseconds 200

        # Verify creation
        $poolExists = Get-Item "IIS:\AppPools\$AppPoolName" -ErrorAction SilentlyContinue

        if ($poolExists) { Write-Success "Application pool created." } 
        else {
            Write-Error-Custom "Application pool creation failed."
            return
        }
    } 
    else { Write-Host "Using existing application pool: $AppPoolName" }
    
    # Create site directory
    if (-not (Test-Path $SiteLocation)) {
        try { New-Item -ItemType Directory -Path $SiteLocation -Force -ErrorAction Stop | Out-Null }
        catch {
            Write-Error-Custom "Failed to create directory '$SiteLocation': $($_.Exception.Message)"
            return
        }

        # Verify creation
        if (Test-Path $SiteLocation) { Write-Success "Created directory: $SiteLocation"
        } else {
            Write-Error-Custom "Directory '$SiteLocation' could not be created."
            return
        }
    }

    #Test Network
    $testNet = Test-NetworkAvailable

    #Get System Python Exe
    $SystemPython = Get-SystemPython

    if (-not $SystemPython) { Write-Error-Custom "No valid Python interpreter found. Install Python from python.org." exit 1 }

    # Packages required
    $packages = @("flask", "wfastcgi","click", "itsdangerous", "jinja2", "markupsafe", "werkzeug", "blinker")

    #ONLINE
    if ($testNet) {
        Write-Host "Network available. Installing packages directly into virtual environment."

        # Create venv 
        $venvPath = Join-Path $SiteLocation '.venv'
        if (-not (Test-Path $venvPath)) {
            Write-Host "Creating Python virtual environment..."
            & $SystemPython -m venv $venvPath
            Write-Host "Virtual environment created"
        }
        # venv paths
        $pythonExe = Join-Path $venvPath 'Scripts\python.exe'
        $pipExe = Join-Path $venvPath 'Scripts\pip.exe'

        # Install packages into venv
        Write-Host "Installing required packages into venv..."
        & $pipExe --disable-pip-version-check install @($packages)
        $exitCode = $LASTEXITCODE
        
        if ($exitCode -ne 0) {
            Write-Error-Custom "Package installation failed."
            return
        }
        
        # Verify imports using venv Python
        $failed = @()
        foreach ($pkg in $packages) {
            $cmd = "`"$pythonExe`" -c `"import $pkg`""
            cmd /c $cmd 2>$null | Out-Null
            if ($LASTEXITCODE -ne 0) { $failed += $pkg }
        }
        
        if ($failed.Count -eq 0) {
            Write-Success "All packages installed successfully in venv."
        }
        else {
            Write-Error-Custom "Failed to import packages in venv: $($failed -join ', ')"
        }
    }
    #OFFLINE
    else{
        Write-Host "Machine is offline. Checking system Python for required packages..."    

        # Check system Python for required packages
        $missing = @()
        foreach ($pkg in $packages) {
            if (-not (Test-PythonPackage -PackageName $pkg -PythonExe $SystemPython)) {
                $missing += $pkg
            }
        }
        
        if ($missing.Count -gt 0) {
            Write-Error-Custom "Offline mode: missing required packages in system Python: $($missing -join ', ')"
            return
        }
        
        Write-Host "All required packages found in system Python. Proceeding with offline installation."
        
        # Create venv
        $venvPath = Join-Path $SiteLocation '.venv'
        if (-not (Test-Path $venvPath)) {
            Write-Host "Creating Python virtual environment..."
            & $SystemPython -m venv $venvPath
            Write-Host "Virtual environment created"
        } 
        
        # venv paths
        $pythonExe = Join-Path $venvPath 'Scripts\python.exe'
        $venvSitePackages = Join-Path $venvPath "Lib\site-packages"

        # Detect system site-packages
        $systemSitePackages = & $SystemPython -c "import sysconfig; print(sysconfig.get_paths()['purelib'])"
        if (-not (Test-Path $systemSitePackages)) {
            Write-Error-Custom "Could not determine system site-packages directory."
            return
        }
        
        Write-Host "Copying packages from system Python into venv..."
        
        foreach ($pkg in $packages) {
            if (-not (Copy-PythonPackage -PackageName $pkg `
                -SourceSitePackages $systemSitePackages `
                -TargetSitePackages $venvSitePackages)) { Write-Error-Custom "Failed to copy package: $pkg" return
            }
        } 

        # Verify imports using venv Python
        $failed = @()
        foreach ($pkg in $packages) {
            $cmd = "`"$pythonExe`" -c `"import $pkg`""
            cmd /c $cmd 2>$null | Out-Null
            if ($LASTEXITCODE -ne 0) { $failed += $pkg }
        }
        
        if ($failed.Count -eq 0) {
            Write-Success "Offline installation complete. All packages imported successfully in venv."
        }
        else {
            Write-Error-Custom "Offline install: failed to import packages in venv: $($failed -join ', ')" 
        }
    }

    # Create Flask app
    $appPy = Join-Path $SiteLocation 'app.py'
    $appContent = @'
from flask import Flask
app = Flask(__name__)

@app.route('/')
def hello():
    return 'Hello from IIS + Flask via wfastcgi!'

if __name__ == '__main__':
    app.run()
'@
    Set-Content -Path $appPy -Value $appContent
    Write-Host "Created app.py"
    
    # Create web.config
    $pythonExe = Join-Path $venvPath 'Scripts\python.exe'
    $wfastcgiPath = Join-Path $venvPath 'Lib\site-packages\wfastcgi.py'
    
    $webConfig = Join-Path $SiteLocation 'web.config'
    $webConfigContent = @"
<?xml version="1.0" encoding="utf-8"?>
<configuration>
    <system.webServer>
        <handlers>
            <add name="PythonHandler" path="*" verb="*" modules="FastCgiModule" scriptProcessor="$pythonExe|$wfastcgiPath" resourceType="Unspecified" requireAccess="Script" />
        </handlers>
    </system.webServer>
</configuration>
"@
    Set-Content -Path $webConfig -Value $webConfigContent
    Write-Host "Created web.config"
    
    # Backup applicationHost.config before ANY modifications
    Write-Host "`nBacking up applicationHost.config before modifications..."
    $backup = Backup-ApplicationHostConfig
    
    # Verify all prerequisites BEFORE making any changes
    Write-Host "`nVerifying prerequisites..."
    $configPath = 'C:\Windows\System32\inetsrv\config\applicationHost.config'
    
    # Check if site name already exists
    $existingSite = Get-Website -Name $siteName -ErrorAction SilentlyContinue
    if ($existingSite) {
        Write-Error-Custom "Site '$siteName' already exists"
        exit 1
    }
    
    # Verify FastCGI section exists in config
    [xml]$xmlCheck = Get-Content $configPath
    $fastCgiNode = $xmlCheck.SelectSingleNode("//fastCgi")
    if (-not $fastCgiNode) {
        Write-Error-Custom "fastCgi section not found in applicationHost.config"
        exit 1
    }
    
    Write-Success "All prerequisites verified"

    try {
        Write-Host "Creating IIS site: $siteName..."
        New-WebSite -Name $siteName -PhysicalPath $SiteLocation -Port $port -ApplicationPool $AppPoolName -Force
        Write-Success "Site created: $siteName"
        Start-Sleep -Milliseconds 200

        Import-Module WebAdministration

        # --- Register FastCGI application ---
        Write-Host "Registering FastCGI application..."

        # Ensure no existing FastCGI entry for this Python/wfastcgi pair
        Remove-FastCgiApplication -PythonExe $pythonExe -ArgumentsPath $wfastcgiPath

        # Add FastCGI entry (single, canonical definition)
        Add-FastCgiApplication -PythonExe $pythonExe -ArgumentsPath $wfastcgiPath -SiteLocation $SiteLocation

        Write-Success "FastCGI registration complete"

        # --- Add handler mapping ---
        Write-Host "Adding handler mapping..."

        # Remove existing handler (if present)
        Remove-WebConfigurationProperty `
            -PSPath "IIS:\Sites\$siteName" `
            -Filter "system.webServer/handlers" `
            -Name "." `
            -AtElement @{ name = "PythonHandler" } `
            -ErrorAction SilentlyContinue

        Start-Sleep -Milliseconds 200

        # Add handler mapping
        Add-WebConfigurationProperty `
            -PSPath "IIS:\Sites\$siteName" `
            -Filter "system.webServer/handlers" `
            -Name "." `
            -Value @{
                name="PythonHandler";
                path="*";
                verb="*";
                modules="FastCgiModule";
                scriptProcessor="$pythonExe|$wfastcgiPath";
                resourceType="Unspecified";
                requireAccess="Script"
            }

        Start-Sleep -Milliseconds 200

        Write-Success "Handler mapping added"

        # --- Set permissions ---
        Write-Host "Setting permissions..."
        $identity = "IIS AppPool\$AppPoolName"

        $acl = Get-Acl $SiteLocation
        $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            $identity, "ReadAndExecute", "ContainerInherit,ObjectInherit", "None", "Allow"
        )
        $acl.AddAccessRule($rule)
        Set-Acl -Path $SiteLocation -AclObject $acl

        Write-Host "Permissions set for $identity"

        # --- Test the site ---
        Start-Sleep 2
        Write-Host "`nTesting site..."

        try {
            $uri = "http://localhost:$port/"
            $response = Invoke-WebRequest -Uri $uri -UseBasicParsing -TimeoutSec 5
            if ($response.StatusCode -eq 200) {
                Write-Success "Site is accessible at $uri"
                Write-Success "Response: $($response.Content)"
            }
        }
        catch {
            Write-Error-Custom "Could not access site: $_"
        }

        Write-Success "`nDeployment complete!`nSite: $siteName`nLocation: $SiteLocation`nPort: $port"
    }
    catch {
        Write-Error-Custom "Deployment failed: $_"
        Write-Host "`nRolling back to clean state..."

        # Remove the site if it was created
        $createdSite = Get-Website -Name $siteName -ErrorAction SilentlyContinue
        if ($createdSite) {
            Remove-Website -Name $siteName
            Write-Host "Removed partially created site"
        }

        # Remove the site folder if it was created
        Remove-Item -Path $SiteLocation -Recurse -Force -ErrorAction SilentlyContinue

        if (-not (Test-Path $SiteLocation)) {  Write-Success "Site folder removed." } 
        else { Write-Error-Custom "Failed to remove site folder." }

        exit 1
    }
}

function Undeploy-IISSite {

    Write-Host "`n=== IIS Python Site Undeployment ===" -ForegroundColor Cyan

    # Listing Sites
    $sites = @(Get-Website | Where-Object {$_.Name -ne 'Default Web Site'} | Select-Object -ExpandProperty Name)
    
    if ($sites.Count -eq 0) {
        Write-Host "No custom sites to undeploy."
        return
    }

    Write-Host "`nAvailable sites to undeploy:"
    for ($i = 0; $i -lt $sites.Count; $i++) {
        Write-Host "$($i + 1). $($sites[$i])"
    }

    if(!$SiteName){
    # Selecting Sites to Undeploy
        $siteIndex = Read-Host "Select site number to undeploy"
        $parsedIndex = 0
        if (-not ([int]::TryParse($siteIndex, [ref]$parsedIndex)) -or 
            $parsedIndex -lt 1 -or 
            $parsedIndex -gt $sites.Count) {

            Write-Error-Custom "Invalid selection."
            return
        }
    
        $SiteName = $sites[$parsedIndex - 1]

    }

    $site = Get-Website -Name $SiteName
    $sitePhysicalPath = $site.PhysicalPath
    $siteAppPool = $site.ApplicationPool
    $venvPath = Join-Path $sitePhysicalPath '.venv'

    # Paths
    $pythonExe = Join-Path $venvPath 'Scripts\python.exe'
    $venvWFastCGIPath = Join-Path $venvPath 'Lib\site-packages\wfastcgi.py'

    if ($NoConfirm -eq $false) {
        if (-not (Prompt-YesNo "Undeploy site '$SiteName' at '$sitePhysicalPath'?")) {
            return
        }
    }
    
    Write-Host "Undeploying '$SiteName'..."
    
    # Backup applicationHost.config
    Write-Host "Backing up applicationHost.config..."
    Backup-ApplicationHostConfig
    
    Write-Host "Removing fastcgi applications for '$SiteName'..."
    $fastCgiRemoved = $true

    try { Remove-FastCgiApplication -PythonExe $pythonExe -ArgumentsPath $venvWFastCGIPath -ErrorAction Stop }
    catch {
        $fastCgiRemoved = $false
        Write-Error-Custom "FastCGI removal threw an error: $($_.Exception.Message)"
    }

    # Verify removal
    $fastCgiExists = Get-WebConfigurationProperty -Filter "system.webServer/fastCgi/application" -Name "." |
    Where-Object {
        $_.fullPath  -eq $pythonExe -and
        $_.arguments -eq $venvWFastCGIPath
    }

    if ($fastCgiRemoved -and -not $fastCgiExists) { Write-Success "FastCGI entries removed." }
    elseif ($fastCgiExists) { Write-Error-Custom "FastCGI entry still exists after removal attempt." }
    else { Write-Error-Custom "FastCGI removal may have failed." }

    Start-Sleep -Milliseconds 300

    # Remove site
    $siteRemoved = $true

    try { Remove-Website -Name $SiteName -ErrorAction Stop }
    catch {
        $siteRemoved = $false
        Write-Error-Custom "Failed to remove site '$SiteName': $($_.Exception.Message)"
    }

    # Verify removal
    $siteExists = Get-Website -Name $SiteName -ErrorAction SilentlyContinue

    if ($siteRemoved -and -not $siteExists) { Write-Success "Site removed." }
    elseif ($siteExists) { Write-Error-Custom "Site '$SiteName' still exists after removal attempt." }
    else { Write-Error-Custom "Site removal may have failed (provider threw an error)." }

    # Check if app pool is used by other sites
    if ($siteAppPool -and $siteAppPool -ne "DefaultAppPool") {
        if ($KeepAppPool -eq $false) {
            Write-Host "`nChecking application pool usage..."
            $remainingSitesUsingPool = @(Get-Website | Where-Object {$_.ApplicationPool -eq $siteAppPool} | Select-Object -ExpandProperty Name)
            
            if ($remainingSitesUsingPool.Count -eq 0) {
                Write-Host "Application pool '$siteAppPool' is not used by any other sites."

                # Only prompt if NoConfirm is false
                $shouldRemove = $true
                if ($NoConfirm -eq $false) { $shouldRemove = Prompt-YesNo "Remove application pool '$siteAppPool'?" }

                if ($shouldRemove){
                    Write-Host "Stopping app pool '$siteAppPool'..."
                    $poolStopped = $true

                    try { Stop-WebAppPool -Name $siteAppPool -ErrorAction Stop }
                    catch {
                        $poolStopped = $false
                        Write-Error-Custom "Failed to stop application pool '$siteAppPool': $($_.Exception.Message)"
                    }

                    # Poll until fully stopped
                    if ($poolStopped) {
                        for ($i = 1; $i -le 20; $i++) {
                            $state = (Get-WebAppPoolState -Name $siteAppPool -ErrorAction SilentlyContinue).Value
                            if ($state -eq "Stopped") { break } 
                            Start-Sleep -Milliseconds 250
                        }

                        # Re-check final state
                        $finalState = (Get-WebAppPoolState -Name $siteAppPool -ErrorAction SilentlyContinue).Value
                        if ($finalState -ne "Stopped") {
                            $poolStopped = $false 
                            Write-Error-Custom "Application pool '$siteAppPool' did not fully stop (state: $finalState)."
                        }
                    }

                    if ($poolStopped) {
                        Write-Host "Removing app pool '$siteAppPool'..."
                        $poolRemoved = $true

                        try { Remove-WebAppPool -Name $siteAppPool -ErrorAction Stop }
                        catch {
                            $poolRemoved = $false
                            Write-Error-Custom "Failed to remove application pool '$siteAppPool': $($_.Exception.Message)"
                        }

                        # Verify removal
                        $poolStillExists = Get-Item "IIS:\AppPools\$siteAppPool" -ErrorAction SilentlyContinue

                        if ($poolRemoved -and -not $poolStillExists) { Write-Success "Application pool '$siteAppPool' removed." }
                        elseif ($poolStillExists) { Write-Error-Custom "Application pool '$siteAppPool' still exists after removal attempt." }
                        else { Write-Error-Custom "Application pool removal may have failed." }
                    }
                }
            } else {
                Write-Host "Application pool '$siteAppPool' is used by: $($remainingSitesUsingPool -join ', ')"
            }
        }
    }

    # Remove physical path
    if (Test-Path $sitePhysicalPath) {
        if ($KeepLocation -eq $false) {

            # Kill any w3wp.exe processes still using the site path
            Get-WmiObject Win32_Process |
                Where-Object { $_.Name -eq "w3wp.exe" -and $_.CommandLine -match [regex]::Escape($sitePhysicalPath) } |
                ForEach-Object { $_.Terminate() }

            Remove-Item -Path $sitePhysicalPath -Recurse -Force -ErrorAction SilentlyContinue
            if (-not (Test-Path $sitePhysicalPath)) { Write-Success "Site folder removed." } 
            else { Write-Error-Custom "Failed to remove site folder." }
        }
    }

    Write-Success "Undeployment complete!"
}

Test-IISFeatures

# Ensure module loads
if (-not (Get-Module -Name WebAdministration)) {
    Import-Module WebAdministration -ErrorAction Stop
}

# Main menu
Write-Host "IIS Python Site Deployment Tool" -ForegroundColor Cyan

# Verify clean backup exists before any operations
$cleanBackup = Ensure-CleanApplicationHostConfig

if (-not $Mode) {
    Write-Host "1. Deploy new site"
    Write-Host "2. Undeploy site"
    Write-Host "3. Reset to clean state (hard reset)"
    $choice = Read-Host "Choose option"
    
    switch ($choice) {
        '1' { Deploy-IISSite }
        '2' { Undeploy-IISSite }
        '3' { Reset-ApplicationHostConfig }
        default { Write-Error-Custom "Invalid choice"; exit 1 }
    }
} else {
    switch ($Mode) {
        'Deploy' { Deploy-IISSite }
        'Undeploy' { Undeploy-IISSite }
        'Reset' { Reset-ApplicationHostConfig }
    }
}