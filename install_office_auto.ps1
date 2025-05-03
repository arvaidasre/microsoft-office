# Self-elevation mechanism
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
    Start-Process -FilePath PowerShell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Ensure proper encoding for English characters
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# Ensure proper encoding for all operations
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo("en-US")
[System.Threading.Thread]::CurrentThread.CurrentUICulture = [System.Globalization.CultureInfo]::GetCultureInfo("en-US")

# XAML UI in English
[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Office LTSC 2021 Installation Tool" Height="500" Width="650" WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="FontFamily" Value="Segoe UI"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="FontFamily" Value="Segoe UI"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="FontFamily" Value="Consolas"/>
        </Style>
    </Window.Resources>
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <StackPanel Grid.Row="0">
            <TextBlock Text="Office LTSC 2021 Installation Tool" FontSize="24" FontWeight="Bold" Margin="0,0,0,20"/>
            <TextBlock x:Name="StatusText" Text="Checking Office status..." Margin="0,0,0,10"/>
            <ProgressBar x:Name="ProgressBar" Height="10" IsIndeterminate="True" Visibility="Visible"/>
        </StackPanel>
        
        <ScrollViewer Grid.Row="1" Margin="0,10,0,10">
            <TextBox x:Name="InfoTextBox" IsReadOnly="True" TextWrapping="Wrap" Background="#F0F0F0" Padding="10"/>
        </ScrollViewer>
        
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="ActivateButton" Content="Activate Office" Width="150" Height="35" Margin="0,0,10,0"/>
            <Button x:Name="UninstallButton" Content="Uninstall Office" Width="150" Height="35" Margin="0,0,10,0" IsEnabled="False"/>
            <Button x:Name="InstallButton" Content="Install Office" Width="150" Height="35" Margin="0,0,10,0" IsEnabled="False"/>
            <Button x:Name="CloseButton" Content="Close" Width="100" Height="35"/>
        </StackPanel>
    </Grid>
</Window>
"@

# Create window
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get controls
$statusText = $window.FindName("StatusText")
$progressBar = $window.FindName("ProgressBar")
$infoTextBox = $window.FindName("InfoTextBox")
$installButton = $window.FindName("InstallButton")
$uninstallButton = $window.FindName("UninstallButton")
$closeButton = $window.FindName("CloseButton")
$activateButton = $window.FindName("ActivateButton")

# Define common English terms for consistency
$enStrings = @{
    Checking = "Checking"
    Installed = "Installed"
    NotInstalled = "Not installed"
    Installing = "Installing"
    InstallComplete = "Installation complete"
    InstallFailed = "Installation failed"
    Error = "Error"
    VersionInfo = "Version information"
    ActivationStatus = "Activation status"
    ActivationInfo = "Activation information"
    Found = "Found"
    NotFound = "Not found"
    FilesMissing = "Files missing"
    Activating = "Activating"
    ActivationComplete = "Activation complete"
    Uninstalling = "Uninstalling"
    UninstallComplete = "Uninstallation complete"
    UninstallFailed = "Uninstallation failed"
}

# Function to download setup.exe from GitHub if not present locally
function Ensure-SetupExists {
    $setupPath = "$PSScriptRoot\setup.exe"
    $setupExists = Test-Path $setupPath
    
    if (-not $setupExists) {
        try {
            $statusText.Text = "Downloading setup.exe from GitHub..."
            $progressBar.Visibility = "Visible"
            $progressBar.IsIndeterminate = $true
            $infoTextBox.Text = "setup.exe not found locally. Downloading from GitHub...`r`n"
            
            # Setup GitHub direct download URL (raw file)
            $setupUrl = "https://github.com/arvaidasre/microsoft-office/raw/master/setup.exe"
            
            # Create WebClient and add download progress event
            $webClient = New-Object System.Net.WebClient
            
            # Download the file
            $infoTextBox.AppendText("Starting download from $setupUrl`r`n")
            $webClient.DownloadFile($setupUrl, $setupPath)
            
            $infoTextBox.AppendText("Download completed successfully.`r`n")
            return $true
        }
        catch {
            $infoTextBox.AppendText("Error downloading setup.exe: $_`r`n")
            $infoTextBox.AppendText("Please download setup.exe manually and place it in the same folder as this script.`r`n")
            return $false
        }
        finally {
            $progressBar.IsIndeterminate = $false
            $progressBar.Visibility = "Collapsed"
        }
    }
    
    return $true
}

# Function to check Office installation
function Check-OfficeInstallation {
    $statusText.Text = "$($enStrings.Checking) if Office is already installed on this system..."
    $infoTextBox.Text = "Checking in progress..."
    
    try {
        $officeInstalled = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "ProductReleaseIds" -ErrorAction SilentlyContinue
        
        if ($officeInstalled -and $officeInstalled.ProductReleaseIds -like "*2021*") {
            $statusText.Text = "Office LTSC 2021 is already installed on this computer."
            $progressBar.Visibility = "Collapsed"
            
            # Get version info
            $versionInfo = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "VersionToReport" -ErrorAction SilentlyContinue
            $version = $versionInfo.VersionToReport
            
            # Get activation status
            $osppPath = if (Test-Path "${env:ProgramFiles}\Microsoft Office\Office16\OSPP.VBS") {
                "${env:ProgramFiles}\Microsoft Office\Office16\OSPP.VBS"
            } elseif (Test-Path "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OSPP.VBS") {
                "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OSPP.VBS"
            } else {
                $null
            }
            
            $activationStatus = if ($osppPath) {
                cscript //nologo $osppPath /dstatus
            } else {
                "Could not find OSPP.VBS file."
            }
            
            # Get activation date
            $activationInfo = Get-WinEvent -FilterHashtable @{ProviderName='Microsoft-Windows-Security-SPP'; Id=1003} -MaxEvents 1 -ErrorAction SilentlyContinue | Format-List TimeCreated, Message | Out-String
            
            # Display all information
            $infoTextBox.Text = "--- $($enStrings.VersionInfo) ---`r`nVersion: $version`r`n`r`n"
            $infoTextBox.Text += "--- $($enStrings.ActivationStatus) ---`r`n$($activationStatus -join "`r`n")`r`n`r`n"
            $infoTextBox.Text += "--- $($enStrings.ActivationInfo) ---`r`n$activationInfo"
            
            $installButton.IsEnabled = $false
            $uninstallButton.IsEnabled = $true
        } else {
            $statusText.Text = "Office LTSC 2021 is not installed. You can install it."
            $progressBar.Visibility = "Collapsed"
            
            # Check for setup files and download if missing
            $setupExists = Test-Path "$PSScriptRoot\setup.exe"
            $configExists = Test-Path "$PSScriptRoot\configuration.xml"
            
            if ($setupExists) {
                $infoTextBox.Text = "Installation files found. Click the 'Install Office' button to start installation."
                $installButton.IsEnabled = $true
            } else {
                $infoTextBox.Text = "setup.exe not found locally. Click 'Install Office' to download it automatically from GitHub and start installation."
                $installButton.IsEnabled = $true
            }
            
            $uninstallButton.IsEnabled = $false
        }
    } catch {
        $statusText.Text = "An error occurred while checking Office status."
        $progressBar.Visibility = "Collapsed"
        $infoTextBox.Text = "$($enStrings.Error): $_"
        $installButton.IsEnabled = $false
        $uninstallButton.IsEnabled = $false
    }
}

# Function to install Office
function Install-Office {
    $statusText.Text = "Running Office installation..."
    $progressBar.Visibility = "Visible"
    $progressBar.IsIndeterminate = $true
    $installButton.IsEnabled = $false
    $uninstallButton.IsEnabled = $false
    $infoTextBox.Text = "[1/3] Starting Office LTSC 2021 installation...`r`n"
    
    try {
        # Ensure setup.exe exists, download from GitHub if needed
        $setupReady = Ensure-SetupExists
        if (-not $setupReady) {
            $infoTextBox.AppendText("Cannot proceed with installation because setup.exe is missing.`r`n")
            $statusText.Text = $enStrings.InstallFailed
            $progressBar.Visibility = "Collapsed"
            $installButton.IsEnabled = $true
            return
        }
        
        $infoTextBox.AppendText("Using setup.exe to install Office...`r`n")
        
        # Create configuration XML directly in the script
        $configPath = [System.IO.Path]::GetTempFileName()
        $configXml = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="PerpetualVL2021">
    <Product ID="ProPlus2021Volume">
      <Language ID="lt-lt" />
      <ExcludeApp ID="Access" />
      <ExcludeApp ID="Publisher" />
      <ExcludeApp ID="OneNote" />
      <ExcludeApp ID="Outlook" />
      <ExcludeApp ID="SkypeforBusiness" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="Teams" />
      <ExcludeApp ID="OneDrive" />
    </Product>
    <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
    <Property Name="AUTOACTIVATE" Value="0" />
    <Property Name="SharedComputerLicensing" Value="0" />
    <Property Name="PinIconsToTaskbar" Value="FALSE" />
    <Property Name="EXCLUDEAPP" Value="Bing,Search" />
    <Property Name="UseOfficeForDarkSystemTheme" Value="TRUE" />
    <Property Name="MigrateArch" Value="TRUE" />
  </Add>
  <RemoveMSI />
  <Display Level="None" AcceptEULA="TRUE" />
  <Updates Enabled="FALSE" />
  <Property Name="DONOTCREATEDESKTOPSHORTCUT" Value="TRUE" />
  <DisableLogging>TRUE</DisableLogging>
</Configuration>
"@
        Set-Content -Path $configPath -Value $configXml -Encoding UTF8
        $infoTextBox.AppendText("Configuration created for installation...`r`n")
        
        $process = Start-Process -FilePath "$PSScriptRoot\setup.exe" -ArgumentList "/configure", "`"$configPath`"" -NoNewWindow -PassThru -Wait
        
        # Clean up temp file
        Remove-Item -Path $configPath -Force
        
        if ($process.ExitCode -eq 0) {
            $infoTextBox.Text += "[2/3] Installation running successfully...`r`n"
            $infoTextBox.Text += "[3/3] Installation complete.`r`n`r`n"
            $infoTextBox.Text += "Office LTSC 2021 installed successfully. Restart the tool to see updated information."
            $statusText.Text = $enStrings.InstallComplete
        } else {
            $infoTextBox.Text += "Installation ended with an error. Exit code: $($process.ExitCode)"
            $statusText.Text = $enStrings.InstallFailed
        }
    } catch {
        $infoTextBox.Text += "$($enStrings.Error) during installation: $_"
        $statusText.Text = "$($enStrings.InstallFailed) due to error."
    }
    
    $progressBar.Visibility = "Collapsed"
    Check-OfficeInstallation
}

# Function to uninstall Office
function Uninstall-Office {
    $statusText.Text = "$($enStrings.Uninstalling) Office..."
    $progressBar.Visibility = "Visible"
    $progressBar.IsIndeterminate = $true
    $installButton.IsEnabled = $false
    $uninstallButton.IsEnabled = $false
    $activateButton.IsEnabled = $false
    $infoTextBox.Text = "Starting Office LTSC 2021 uninstallation...`r`n"
    
    try {
        # Ensure setup.exe exists, download from GitHub if needed
        $setupReady = Ensure-SetupExists
        if (-not $setupReady) {
            $infoTextBox.AppendText("Cannot proceed with uninstallation because setup.exe is missing.`r`n")
            $statusText.Text = $enStrings.UninstallFailed
            $progressBar.Visibility = "Collapsed"
            $uninstallButton.IsEnabled = $true
            return
        }
        
        # Create removal configuration XML
        $configPath = [System.IO.Path]::GetTempFileName()
        $configXml = @"
<Configuration>
  <Remove All="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
</Configuration>
"@
        Set-Content -Path $configPath -Value $configXml -Encoding UTF8
        $infoTextBox.AppendText("Uninstallation configuration created...`r`n")
        
        $infoTextBox.AppendText("Closing Office applications...`r`n")
        Get-Process -Name "*excel*", "*word*", "*powerpoint*", "*outlook*", "*onenote*", "*publisher*", "*access*" -ErrorAction SilentlyContinue | ForEach-Object { 
            try { $_.CloseMainWindow() | Out-Null } catch { }
            Start-Sleep -Seconds 1
            try { $_ | Stop-Process -Force -ErrorAction SilentlyContinue } catch { }
        }
        
        $infoTextBox.AppendText("Running Office uninstaller...`r`n")
        $process = Start-Process -FilePath "$PSScriptRoot\setup.exe" -ArgumentList "/configure", "`"$configPath`"" -NoNewWindow -PassThru -Wait
        
        # Clean up temp file
        Remove-Item -Path $configPath -Force
        
        if ($process.ExitCode -eq 0) {
            $infoTextBox.AppendText("Uninstallation completed successfully.`r`n")
            $infoTextBox.AppendText("Office has been removed from your system.`r`n")
            $statusText.Text = $enStrings.UninstallComplete
        } else {
            $infoTextBox.AppendText("Uninstallation ended with an error. Exit code: $($process.ExitCode)`r`n")
            $statusText.Text = $enStrings.UninstallFailed
        }
    } catch {
        $infoTextBox.AppendText("$($enStrings.Error) during uninstallation: $_`r`n")
        $statusText.Text = "$($enStrings.UninstallFailed) due to error."
    }
    
    $progressBar.Visibility = "Collapsed"
    Check-OfficeInstallation
}

# Function to activate Office
function Activate-Office {
    $statusText.Text = "$($enStrings.Activating) Office..."
    $progressBar.Visibility = "Visible"
    $progressBar.IsIndeterminate = $true
    $activateButton.IsEnabled = $false
    $installButton.IsEnabled = $false
    $uninstallButton.IsEnabled = $false
    $infoTextBox.Text = "Running activation script...`r`n"
    
    try {
        # Run the activation script
        $infoTextBox.AppendText("Starting activation process...`r`n")
        $result = Invoke-Expression -Command "(Invoke-RestMethod -Uri 'https://get.activated.win') | Invoke-Expression"
        $infoTextBox.AppendText("Activation completed.`r`n")
        $statusText.Text = $enStrings.ActivationComplete
    } catch {
        $infoTextBox.AppendText("Error during activation: $_`r`n")
        $statusText.Text = "Activation error"
    } finally {
        $progressBar.Visibility = "Collapsed"
        $activateButton.IsEnabled = $true
        Check-OfficeInstallation
    }
}

# Button event handlers
$installButton.Add_Click({
    Install-Office
})

$uninstallButton.Add_Click({
    # Add confirmation dialog
    $result = [System.Windows.MessageBox]::Show("Are you sure you want to uninstall Office? This will remove all Office applications and data.", "Confirm Uninstallation", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        Uninstall-Office
    }
})

$activateButton.Add_Click({
    Activate-Office
})

$closeButton.Add_Click({
    $window.Close()
})

# Run initial check
Check-OfficeInstallation

# Show window
$window.ShowDialog() | Out-Null
