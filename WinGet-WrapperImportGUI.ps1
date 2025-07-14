# Soren Lundt - 12-02-2024
# URL: https://github.com/SorenLundt/WinGet-Wrapper
# License: https://raw.githubusercontent.com/SorenLundt/WinGet-Wrapper/main/LICENSE.txt
# Graphical interface for WinGet-Wrapper Import - WPF Version
# Package content is stored under Packages\Package.ID-Context-UpdateOnly-UserName-yyyy-mm-dd-hhssmm

# Requirements:
# Requires Script files and IntuneWinAppUtil.exe to be present in script directory
#
# Version History
# Version 1.0 - 12-02-2024 SorenLundt - Initial Version
# Version 1.1 - 21-02-2024 SorenLundt - Fixed issue where only 1 package was imported to InTune (Script assumed there was just one row)
# Version 1.2 - 30-10-2024 SorenLundt - Various improvements (Ability to get details and available versions for packages, loading GUI progressbar, link to repo, unblock script files, bug fixes.)
# Version 1.3 - 31-10-2024 SorenLundt - Added Column Definitions help button, added SKIPMODULECHECK param to skip check if required modules are up-to-date. (testing use)
# Version 2.0 - 14-07-2025 - Converted to WPF for modern UI

#Parameters
Param (
    #Skip module checking (for testing purposes)
    [Switch]$SKIPMODULECHECK = $false
)

# Greeting
Write-Host ""
Write-Host "****************************************************"
Write-Host "                  WinGet-Wrapper)"
Write-Host "  https://github.com/SorenLundt/WinGet-Wrapper"
Write-Host ""
Write-Host "          GNU General Public License v3"
Write-Host "****************************************************"
Write-Host "   WinGet-WrapperImportGUI Starting up.."
Write-Host ""

# Function to show loading progress bar in console.
function Show-ConsoleProgress {
    param (
        [string]$Activity = "Loading Winget-Wrapper Import GUI",
        [string]$Status = "",
        [int]$PercentComplete = 0
    )
    Write-Host "$Status - [$PercentComplete%]"
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete
}

# Update ConsoleProgress
Show-ConsoleProgress -PercentComplete 0 -Status "Initializing..."

# Add required assemblies for WPF
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Set the timestamp for log file
$timestamp = Get-Date -Format "yyyyMMddHHmmss"

#Find Script root path  
if (-not $PSScriptRoot) {
    $scriptRoot = (Get-Location -PSProvider FileSystem).ProviderPath
}
else {
    $scriptRoot = $PSScriptRoot
}

# Create logs folder if it doesn't exist
$LogFolder = Join-Path -Path $scriptRoot -ChildPath "Logs"
if (-not (Test-Path -Path $LogFolder)) {
    New-Item -Path $LogFolder -ItemType Directory | Out-Null
}

# Install and load required modules
$intuneWin32AppModule = "IntuneWin32App"
$microsoftGraphIntuneModule = "Microsoft.Graph.Intune"
$microsoftGraphAuthenticationModule = "Microsoft.Graph.Authentication"

if (-not $SKIPMODULECHECK) {
    # Check IntuneWin32App module
    Show-ConsoleProgress -PercentComplete 10 -Status "Checking and updating $intuneWin32AppModule.."  
    $moduleInstalled = Get-InstalledModule -Name $intuneWin32AppModule -ErrorAction SilentlyContinue
    if (-not $moduleInstalled) {
        Install-Module -Name $intuneWin32AppModule -Force
    }
    else {
        $latestVersion = (Find-Module -Name $intuneWin32AppModule).Version
        if ($moduleInstalled.Version -lt $latestVersion) {
            Update-Module -Name $intuneWin32AppModule -Force
        }
        else {
            Write-Host "Module $intuneWin32AppModule is already up-to-date." -ForegroundColor Green
        }
    }

    # Check Microsoft.Graph.Intune module
    Show-ConsoleProgress -PercentComplete 40 -Status "Checking and updating $microsoftGraphIntuneModule.."  
    $moduleInstalled = Get-InstalledModule -Name $microsoftGraphIntuneModule -ErrorAction SilentlyContinue
    if (-not $moduleInstalled) {
        Install-Module -Name $microsoftGraphIntuneModule -Force
    }
    else {
        $latestVersion = (Find-Module -Name $microsoftGraphIntuneModule).Version
        if ($moduleInstalled.Version -lt $latestVersion) {
            Update-Module -Name $microsoftGraphIntuneModule -Force
        }
        else {
            Write-Host "Module $microsoftGraphIntuneModule is already up-to-date." -ForegroundColor Green
        }
    }

    # Check Microsoft.Graph.Authentication module
    Show-ConsoleProgress -PercentComplete 40 -Status "Checking and updating $microsoftGraphAuthenticationModule.."  
    $moduleInstalled = Get-InstalledModule -Name $microsoftGraphAuthenticationModule -ErrorAction SilentlyContinue
    if (-not $moduleInstalled) {
        Install-Module -Name $microsoftGraphAuthenticationModule -Force
    }
    else {
        $latestVersion = (Find-Module -Name $microsoftGraphAuthenticationModule).Version
        if ($moduleInstalled.Version -lt $latestVersion) {
            Update-Module -Name $microsoftGraphAuthenticationModule -Force
        }
        else {
            Write-Host "Module $microsoftGraphAuthenticationModule is already up-to-date." -ForegroundColor Green
        }
    }
}

#Import modules
Show-ConsoleProgress -PercentComplete 60 -Status "Importing module $intuneWin32AppModule.."
Import-Module -Name "IntuneWin32App"

Show-ConsoleProgress -PercentComplete 80 -Status "Importing module $microsoftGraphIntuneModule.."
Import-Module -Name "Microsoft.Graph.Intune"

Show-ConsoleProgress -PercentComplete 90 -Status "Unblocking script files (Unblock-File)"
# Unblock all files in the current directory
$files = Get-ChildItem -Path . -File
foreach ($file in $files) {
    try {
        Unblock-File -Path $file.FullName
    }
    catch {
        Write-Host "Failed to unblock: $($file.FullName) - $_"
    }
}

# Update ConsoleProgress
Show-ConsoleProgress -PercentComplete 80 -Status "Loading functions.."   

# Define WPF XAML
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="WinGet-Wrapper Import GUI - https://github.com/SorenLundt/WinGet-Wrapper" 
        Height="1000" Width="1500"
        Background="#F8F9FA" 
        FontFamily="Segoe UI" 
        FontSize="12"
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanResize">
    
    <Window.Resources>
        <!-- Modern Button Style -->
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="#007ACC"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="Margin" Value="4"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                CornerRadius="4"
                                BorderThickness="0">
                            <ContentPresenter HorizontalAlignment="Center" 
                                            VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#005A9E"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#004578"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Secondary Button Style -->
        <Style x:Key="SecondaryButton" TargetType="Button" BasedOn="{StaticResource ModernButton}">
            <Setter Property="Background" Value="#6C757D"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#5A6268"/>
                </Trigger>
                <Trigger Property="IsPressed" Value="True">
                    <Setter Property="Background" Value="#495057"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- Danger Button Style -->
        <Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource ModernButton}">
            <Setter Property="Background" Value="#DC3545"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#C82333"/>
                </Trigger>
                <Trigger Property="IsPressed" Value="True">
                    <Setter Property="Background" Value="#BD2130"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- Modern TextBox Style -->
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="8,8"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="BorderBrush" Value="#CED4DA"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="Height" Value="52"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border Background="{TemplateBinding Background}" 
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4"
                                Height="{TemplateBinding Height}">
                            <ScrollViewer x:Name="PART_ContentHost" 
                                        Margin="8,0,8,0"
                                        VerticalAlignment="Center"
                                        VerticalScrollBarVisibility="Hidden"
                                        HorizontalScrollBarVisibility="Hidden"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="BorderBrush" Value="#80BDFF"/>
                            </Trigger>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter Property="BorderBrush" Value="#007ACC"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Modern DataGrid Style -->
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#DEE2E6"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
            <Setter Property="HorizontalGridLinesBrush" Value="#F8F9FA"/>
            <Setter Property="VerticalGridLinesBrush" Value="Transparent"/>
            <Setter Property="RowBackground" Value="White"/>
            <Setter Property="AlternatingRowBackground" Value="#F8F9FA"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="SelectionMode" Value="Extended"/>
            <Setter Property="SelectionUnit" Value="FullRow"/>
            <Setter Property="CanUserAddRows" Value="False"/>
            <Setter Property="CanUserDeleteRows" Value="False"/>
            <Setter Property="AutoGenerateColumns" Value="False"/>
        </Style>

        <!-- DataGrid Header Style -->
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#495057"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="BorderBrush" Value="#495057"/>
            <Setter Property="BorderThickness" Value="0,0,1,0"/>
        </Style>

        <!-- Card Style -->
        <Style x:Key="Card" TargetType="Border">
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#DEE2E6"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="8"/>
            <Setter Property="Padding" Value="16"/>
            <Setter Property="Margin" Value="8"/>
            <Setter Property="Effect">
                <Setter.Value>
                    <DropShadowEffect Color="#000000" Opacity="0.1" BlurRadius="10" ShadowDepth="2"/>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="300"/>
        </Grid.RowDefinitions>

        <!-- Search Section -->
        <Border Grid.Row="0" Style="{StaticResource Card}" Margin="0,0,0,8">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                
                <StackPanel Grid.Column="0" Orientation="Horizontal">
                    <TextBox Name="SearchBox" Width="400" MinWidth="300" VerticalContentAlignment="Center"
                           Text="Search for software (e.g., VLC, Chrome, Firefox)"
                           Foreground="Gray"
                           ToolTip="Enter software name, ex. VLC, 7-zip, etc."/>
                    <Button Name="SearchButton" Content="Search" Style="{StaticResource ModernButton}" 
                          Width="130" MinWidth="100" Height="52" Margin="8,0,0,0"/>
                </StackPanel>
                
                <StackPanel Grid.Column="1" Orientation="Vertical" HorizontalAlignment="Right" VerticalAlignment="Center">
                    <TextBlock Name="GitHubLink" Text="Visit GitHub Repository" 
                             Foreground="#007ACC" TextDecorations="Underline" 
                             Cursor="Hand" FontSize="11" HorizontalAlignment="Right"/>
                    <TextBlock Name="ColumnHelp" Text="Show Column Definitions" 
                             Foreground="#6C757D" TextDecorations="Underline" 
                             Cursor="Hand" FontSize="10" HorizontalAlignment="Right" Margin="0,4,0,0"/>
                </StackPanel>
                
                <TextBlock Name="SearchStatus" Grid.Column="1" VerticalAlignment="Bottom" 
                         Margin="16,0,0,0" FontSize="12" Foreground="#6C757D" TextWrapping="Wrap"/>
            </Grid>
        </Border>

        <!-- Main Content Area -->
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- Search Results Panel -->
            <Border Grid.Column="0" Style="{StaticResource Card}" Margin="0,0,4,0">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock Grid.Row="0" Text="WinGet Packages" FontWeight="SemiBold" 
                             FontSize="14" Margin="0,0,0,8"/>
                    
                    <DataGrid Name="SearchResults" Grid.Row="1">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="0.7*" MinWidth="140"/>
                            <DataGridTextColumn Header="ID" Binding="{Binding ID}" Width="250"/>
                            <DataGridTextColumn Header="Version" Binding="{Binding Version}" Width="120"/>
                        </DataGrid.Columns>
                    </DataGrid>
                    
                    <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,8,0,0">
                        <Button Name="GetPackageDetails" Content="Get Details" 
                              Style="{StaticResource SecondaryButton}" Width="120"/>
                        <Button Name="GetPackageVersions" Content="Get Versions" 
                              Style="{StaticResource SecondaryButton}" Width="120"/>
                    </StackPanel>
                </Grid>
            </Border>

            <!-- Move Button -->
            <Grid Grid.Column="1" Width="80" VerticalAlignment="Center">
                <Button Name="MoveButton" Content="Move" FontSize="12" 
                      Style="{StaticResource ModernButton}" Width="70" Height="50"
                      ToolTip="Add Selected Package(s) to Import List"/>
            </Grid>

            <!-- Import List Panel -->
            <Border Grid.Column="2" Style="{StaticResource Card}" Margin="4,0,0,0">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock Grid.Row="0" Text="InTune Import List" FontWeight="SemiBold" 
                             FontSize="14" Margin="0,0,0,8"/>
                    
                    <DataGrid Name="ImportList" Grid.Row="1">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="PackageID" Binding="{Binding PackageID}" Width="150"/>
                            <DataGridTextColumn Header="Context" Binding="{Binding Context}" Width="80"/>
                            <DataGridTextColumn Header="AcceptNewer" Binding="{Binding AcceptNewerVersion}" Width="90"/>
                            <DataGridTextColumn Header="UpdateOnly" Binding="{Binding UpdateOnly}" Width="80"/>
                            <DataGridTextColumn Header="TargetVersion" Binding="{Binding TargetVersion}" Width="100"/>
                            <DataGridTextColumn Header="StopProcessInstall" Binding="{Binding StopProcessInstall}" Width="130"/>
                            <DataGridTextColumn Header="StopProcessUninstall" Binding="{Binding StopProcessUninstall}" Width="140"/>
                            <DataGridTextColumn Header="PreScriptInstall" Binding="{Binding PreScriptInstall}" Width="120"/>
                            <DataGridTextColumn Header="PostScriptInstall" Binding="{Binding PostScriptInstall}" Width="130"/>
                            <DataGridTextColumn Header="PreScriptUninstall" Binding="{Binding PreScriptUninstall}" Width="130"/>
                            <DataGridTextColumn Header="PostScriptUninstall" Binding="{Binding PostScriptUninstall}" Width="140"/>
                            <DataGridTextColumn Header="CustomArgInstall" Binding="{Binding CustomArgumentListInstall}" Width="130"/>
                            <DataGridTextColumn Header="CustomArgUninstall" Binding="{Binding CustomArgumentListUninstall}" Width="140"/>
                            <DataGridTextColumn Header="InstallIntent" Binding="{Binding InstallIntent}" Width="100"/>
                            <DataGridTextColumn Header="Notification" Binding="{Binding Notification}" Width="90"/>
                            <DataGridTextColumn Header="GroupID" Binding="{Binding GroupID}" Width="80"/>
                        </DataGrid.Columns>
                    </DataGrid>
                    
                    <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,8,0,0">
                        <Button Name="ImportCSV" Content="Import CSV" 
                              Style="{StaticResource SecondaryButton}" Width="100"/>
                        <Button Name="ExportCSV" Content="Export CSV" 
                              Style="{StaticResource SecondaryButton}" Width="100"/>
                        <Button Name="DeleteSelected" Content="Delete" 
                              Style="{StaticResource DangerButton}" Width="100"/>
                        <Button Name="ImportToInTune" Content="Import to InTune" 
                              Style="{StaticResource ModernButton}" Width="120"/>
                    </StackPanel>
                </Grid>
            </Border>
        </Grid>

        <!-- InTune Configuration Section -->
        <Border Grid.Row="2" Style="{StaticResource Card}" Margin="0,8,0,8">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>                    // ...existing code...
                    
                            <!-- Console Output -->
                            <Border Grid.Row="3" Style="{StaticResource Card}" Margin="0,0,0,0">
                                <Grid>
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="*"/>
                                    </Grid.RowDefinitions>
                                    
                                    <TextBlock Grid.Row="0" Text="Console Output" FontWeight="SemiBold" 
                                             FontSize="14" Margin="0,0,0,8"/>
                                    
                                    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" 
                                                HorizontalScrollBarVisibility="Auto"
                                                VerticalAlignment="Stretch"
                                                HorizontalAlignment="Stretch">
                                        <TextBox Name="ConsoleOutput" IsReadOnly="True" 
                                               Background="Transparent" BorderThickness="0"
                                               FontFamily="Consolas" FontSize="11"
                                               TextWrapping="Wrap" AcceptsReturn="True"
                                               Height="Auto" MinHeight="250"
                                               VerticalAlignment="Stretch"
                                               HorizontalAlignment="Stretch"/>
                                    </ScrollViewer>
                                </Grid>
                            </Border>
                    
                    // ...existing code...
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <TextBlock Grid.Row="0" Text="Microsoft InTune Import Configuration" 
                         FontWeight="SemiBold" FontSize="12" Margin="0,0,0,6"/>
                
                <Grid Grid.Row="1">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*" MinWidth="200"/>
                        <ColumnDefinition Width="*" MinWidth="200"/>
                        <ColumnDefinition Width="*" MinWidth="250"/>
                    </Grid.ColumnDefinitions>
                    
                    <StackPanel Grid.Column="0" Margin="0,0,8,0">
                        <TextBlock Text="Tenant ID" FontWeight="SemiBold" FontSize="11" Margin="0,0,0,2"/>
                        <TextBox Name="TenantID" Height="36"
                               Text="company.onmicrosoft.com"
                               Foreground="Black"/>
                    </StackPanel>
                    
                    <StackPanel Grid.Column="1" Margin="8,0">
                        <TextBlock Text="Client ID" FontWeight="SemiBold" FontSize="11" Margin="0,0,0,2"/>
                        <TextBox Name="ClientID" Height="36"
                               Text="14d82eec-204b-4c2f-b7e8-296a70dab67e"/>
                    </StackPanel>
                    
                    <StackPanel Grid.Column="2" Margin="8,0,0,0">
                        <TextBlock Text="Redirect URI" FontWeight="SemiBold" FontSize="11" Margin="0,0,0,2"/>
                        <TextBox Name="RedirectURI" Height="36"
                               Text="https://login.microsoftonline.com/common/oauth2/nativeclient"/>
                    </StackPanel>
                </Grid>
            </Grid>
        </Border>

        <!-- Console Output -->
        <Border Grid.Row="3" Style="{StaticResource Card}" Margin="0,0,0,0">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                
                <TextBlock Grid.Row="0" Text="Console Output" FontWeight="SemiBold" 
                         FontSize="14" Margin="0,0,0,8"/>
                
                <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" 
                            HorizontalScrollBarVisibility="Auto"
                            VerticalAlignment="Stretch"
                            HorizontalAlignment="Stretch">
                    <TextBox Name="ConsoleOutput" IsReadOnly="True" 
                           Background="Transparent" BorderThickness="0"
                           FontFamily="Consolas" FontSize="11"
                           TextWrapping="Wrap" AcceptsReturn="True"
                           Height="Auto" MinHeight="250"
                           VerticalAlignment="Stretch"
                           HorizontalAlignment="Stretch"/>
                </ScrollViewer>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Load XAML
try {
    $reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
}
catch {
    Write-Error "Failed to load XAML: $_"
    exit 1
}

# Get WPF controls
$searchBox = $window.FindName("SearchBox")
$searchButton = $window.FindName("SearchButton")
$searchStatus = $window.FindName("SearchStatus")
$searchResults = $window.FindName("SearchResults")
$moveButton = $window.FindName("MoveButton")
$importList = $window.FindName("ImportList")
$getPackageDetails = $window.FindName("GetPackageDetails")
$getPackageVersions = $window.FindName("GetPackageVersions")
$importCSV = $window.FindName("ImportCSV")
$exportCSV = $window.FindName("ExportCSV")
$deleteSelected = $window.FindName("DeleteSelected")
$tenantID = $window.FindName("TenantID")
$clientID = $window.FindName("ClientID")
$redirectURI = $window.FindName("RedirectURI")
$importToInTune = $window.FindName("ImportToInTune")
$consoleOutput = $window.FindName("ConsoleOutput")
$gitHubLink = $window.FindName("GitHubLink")
$columnHelp = $window.FindName("ColumnHelp")

#Functions
function Write-ConsoleTextBox {
    param (
        [string]$Message,
        [switch]$NoTimeStamp
    )

    if (-not $NoTimeStamp) {
        $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $Message = "[$TimeStamp] $Message"
    }

    # Output to the console
    Write-Host $Message
    
    # Append to the console textbox in WPF
    $window.Dispatcher.Invoke([System.Action]{
        $consoleOutput.AppendText("$Message`r`n")
        # Auto-scroll to bottom
        $consoleOutput.ScrollToEnd()
    })
}

# Function to read log file and update the GUI
function Update-GUIFromLogFile {
    param (
        [string]$logFilePath
    )
    # Read the log file content
    $logContent = Get-Content -Path $logFilePath

    # Assuming Write-ConsoleTextBox adds each line to the GUI's textbox
    foreach ($line in $logContent) {
        Write-ConsoleTextBox -Message $line -NoTimeStamp
    }
}

# Function to get WinGet package details for a single/selected package
function WinGetPackageDetails {
    param (
        [string]$PackageID
    )

    # Get package details and available versions
    $WingetPackageDetails = winget show --id $PackageID --source WinGet --accept-source-agreements --disable-interactivity

    # Output package details line by line to maintain formatting
    Write-ConsoleTextBox "$PackageID - Details:"
    $WingetPackageDetails -split "`r?`n" | ForEach-Object { Write-ConsoleTextBox $_ }
    Write-ConsoleTextBox "_"  # Separator for readability

    # Optionally return the available versions array for further processing
    return $WinGetPackageDetails
}

# Function to get WinGet package versions for a single/selected package
function WinGetPackageVersions {
    param (
        [string]$PackageID
    )
    # Get package details and available versions
    $WingetPackageVersionsOutput = winget show --id $PackageID --source WinGet --versions

    # Output available versions header
    Write-ConsoleTextBox "$PackageID - Available Versions:"

    # Initialize an array for the available versions
    $WinGetPackageVersions = @()  

    # Split version details into lines and filter out empty lines
    $versionLines = $WingetPackageVersionsOutput -split "`r?`n" | Where-Object { 
        -not [string]::IsNullOrWhiteSpace($_) 
    }

    # Skip the first three lines and process the remaining lines
    foreach ($line in $versionLines[3..($versionLines.Length - 1)]) {
        $trimmedLine = $line.Trim()  # Trim whitespace
        if (-not [string]::IsNullOrWhiteSpace($trimmedLine) -and $trimmedLine -notmatch "^-+$") {
            # Check it's not empty or dashes
            $WinGetPackageVersions += $trimmedLine  # Add to available versions array
            Write-ConsoleTextBox $trimmedLine  # Display each version line
        }
    }

    # Return the available versions array for further processing
    return $WinGetPackageVersions
}

# Define a function to parse the search results
function ParseSearchResults($searchResult) {
    Write-ConsoleTextBox "Parsing data..."
    $parsedData = @()
    $pattern = "^(.+?)\s+((?:[\w.-]+(?:\.[\w.-]+)+))\s+(\S.*?)\s*$"
    $searchResult -split "`n" | Where-Object { $_ -match $pattern } | ForEach-Object {
        $parsedName = $Matches[1].Trim()
        $parsedID = $Matches[2].Trim()
        $parsedID = $parsedID -replace 'ÔÇª', ''  # Remove ellipsis character from ID
        $parsedVersion = $Matches[3].Trim()

        # Add parsed and cleaned data to the result
        $parsedData += [PSCustomObject]@{
            'Name' = $parsedName
            'ID' = $parsedID
            'Version' = $parsedVersion
        }
    }
    Write-ConsoleTextBox "Finished"
    return $parsedData
}

# Define the PerformSearch function that uses the parsing function
function PerformSearch {
    $searchString = $searchBox.Text
    $searchStatus.Text = "Searching for '$searchString'..."

    #Update winget sources (to prevent source updating as part of results, weird output)
    @(winget source update)

    # Search Logic
    if (![string]::IsNullOrWhiteSpace($searchString)) {
        Write-ConsoleTextBox "winget search --query $searchString --source WinGet --accept-source-agreements --disable-interactivity"
        $searchResult = @(winget search --query $searchString --source WinGet --accept-source-agreements --disable-interactivity)
        
        # Splitting the search result into lines for logging purposes
        $lines = $searchResult -split "`r`n"

        # Writing each line to the consoleTextBox
        foreach ($line in $lines) {
            Write-ConsoleTextBox $line
        }
        
        if ($searchResult -contains "No package found matching input criteria.") {
            $searchResults.ItemsSource = $null
            $searchStatus.Text = "No WinGet package found for search query '$searchString'"
        }
        else {
            # Parse the search result using the ParseSearchResults function
            $parsedSearchResult = ParseSearchResults -searchResult $searchResult |
            Where-Object { $null -ne $_.Name -and $_.Name -ne "" -and $_.Name.Trim() -ne "" }

            # Set ItemsSource for WPF DataGrid
            $searchResults.ItemsSource = $parsedSearchResult
            $searchStatus.Text = "Found $($parsedSearchResult.Count) packages for '$searchString'"
        }
    }
    else {
        $searchResults.ItemsSource = $null
        $searchStatus.Text = "Please enter a search query."
    }
}

# Function to write Column Definitions to ConsoleWriteBox
Function GetColumnDefinitions {
    # Fetch the content of the README.md file
    $url = "https://raw.githubusercontent.com/SorenLundt/WinGet-Wrapper/main/README.md"
    Write-ConsoleTextBox "$url"
    
    try {
        $response = Invoke-WebRequest -Uri $url

        # Check if the request was successful
        if ($response.StatusCode -eq 200) {
            $content = $response.Content

            # Get column names from the ImportList DataGrid
            $columnNames = @("PackageID", "Context", "AcceptNewerVersion", "UpdateOnly", "TargetVersion", 
                           "StopProcessInstall", "StopProcessUninstall", "PreScriptInstall", "PostScriptInstall", 
                           "PreScriptUninstall", "PostScriptUninstall", "CustomArgumentListInstall", 
                           "CustomArgumentListUninstall", "InstallIntent", "Notification", "GroupID")
            
            Write-ConsoleTextBox "Column Name --> Description"
            Write-ConsoleTextBox "****************************"        
            
            # Loop through each column name
            foreach ($columnName in $columnNames) {
                # Use regex to find the line corresponding to the column name
                if ($content -match "\* $columnName\s*=\s*(.*?)<br>") {
                    $description = $matches[1] -replace "<br>", "`n" -replace "\* ", "" # Clean up the description
                    Write-ConsoleTextBox "$columnName --> $description"
                }
                else {
                    Write-ConsoleTextBox "$columnName --> No description found."
                }
            }
            Write-ConsoleTextBox "****************************"
        }
        else {
            Write-ConsoleTextBox "Failed to retrieve the README file. Status code: $($response.StatusCode)"
        }
    }
    catch {
        Write-ConsoleTextBox "Error fetching README: $($_.Exception.Message)"
    }
}

# Create a global observable collection for the import list
$script:importListItems = New-Object System.Collections.ObjectModel.ObservableCollection[object]

# Set up event handlers

# Search functionality
$searchButton.Add_Click({ PerformSearch })

# Allow Enter key to trigger search
$window.Add_KeyDown({
    param($sender, $e)
    if ($e.Key -eq "Enter") {
        PerformSearch
    }
})

# Move packages to import list
$moveButton.Add_Click({
    $selectedItems = $searchResults.SelectedItems
    
    if ($selectedItems.Count -eq 0) {
        Write-ConsoleTextBox "No packages selected. Please select one or more packages to move."
        return
    }
    
    # Initialize the import list if it's empty
    if (-not $importList.ItemsSource) {
        $importList.ItemsSource = $script:importListItems
    }
    
    foreach ($item in $selectedItems) {
        # Check if package already exists in the list
        $existingItem = $script:importListItems | Where-Object { $_.PackageID -eq $item.ID }
        if ($existingItem) {
            Write-ConsoleTextBox "Package '$($item.ID)' already exists in import list"
            continue
        }
        
        $newItem = [PSCustomObject]@{
            PackageID = $item.ID
            Context = "Machine"
            AcceptNewerVersion = "1"
            UpdateOnly = "0"
            TargetVersion = ""
            StopProcessInstall = ""
            StopProcessUninstall = ""
            PreScriptInstall = ""
            PostScriptInstall = ""
            PreScriptUninstall = ""
            PostScriptUninstall = ""
            CustomArgumentListInstall = ""
            CustomArgumentListUninstall = ""
            InstallIntent = ""
            Notification = ""
            GroupID = ""
        }
        
        $script:importListItems.Add($newItem)
        Write-ConsoleTextBox "Added '$($item.ID)' to import list"
    }
})

# Delete selected items from import list
$deleteSelected.Add_Click({
    $selectedItems = @($importList.SelectedItems)
    
    if ($selectedItems.Count -eq 0) {
        Write-ConsoleTextBox "No items selected for deletion."
        return
    }
    
    foreach ($item in $selectedItems) {
        $script:importListItems.Remove($item)
        Write-ConsoleTextBox "Removed '$($item.PackageID)' from import list"
    }
})

# Package details
$getPackageDetails.Add_Click({
    $selectedItems = $searchResults.SelectedItems
    foreach ($item in $selectedItems) {
        WinGetPackageDetails -PackageID $item.ID
    }
})

# Package versions
$getPackageVersions.Add_Click({
    $selectedItems = $searchResults.SelectedItems
    foreach ($item in $selectedItems) {
        WinGetPackageVersions -PackageID $item.ID
    }
})

# Import CSV
$importCSV.Add_Click({
    $openFileDialog = New-Object Microsoft.Win32.OpenFileDialog
    $openFileDialog.InitialDirectory = (Get-Location).Path
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv"
    
    if ($openFileDialog.ShowDialog()) {
        $csvFilePath = $openFileDialog.FileName
        
        try {
            $importedData = Import-Csv -Path $csvFilePath
            
            # Clear existing items and add imported data
            $script:importListItems.Clear()
            foreach ($item in $importedData) {
                $script:importListItems.Add($item)
            }
            
            # Set the ItemsSource if not already set
            if (-not $importList.ItemsSource) {
                $importList.ItemsSource = $script:importListItems
            }
            
            Write-ConsoleTextBox "Imported CSV: $csvFilePath"
        }
        catch {
            Write-ConsoleTextBox "Error importing CSV: $($_.Exception.Message)"
        }
    }
})

# Export CSV
$exportCSV.Add_Click({
    if ($script:importListItems.Count -gt 0) {
        $saveFileDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveFileDialog.InitialDirectory = (Get-Location).Path
        $saveFileDialog.FileName = "WinGet-WrapperImportGUI-$timestamp.csv"
        $saveFileDialog.Filter = "CSV files (*.csv)|*.csv"
        
        if ($saveFileDialog.ShowDialog()) {
            $csvFilePath = $saveFileDialog.FileName
            
            try {
                $script:importListItems | Export-Csv -Path $csvFilePath -NoTypeInformation
                Write-ConsoleTextBox "Exported: $csvFilePath"
            }
            catch {
                Write-ConsoleTextBox "Error exporting CSV: $($_.Exception.Message)"
            }
        }
    }
    else {
        Write-ConsoleTextBox "No data to export."
    }
})

# GitHub link
$gitHubLink.Add_MouseDown({
    Start-Process "https://github.com/SorenLundt/WinGet-Wrapper"
})

# Column help
$columnHelp.Add_MouseDown({
    GetColumnDefinitions
})

# Search box placeholder handling
$defaultSearchText = "Search for software (e.g., VLC, Chrome, Firefox)"
$searchBox.Add_GotFocus({
    if ($searchBox.Text -eq $defaultSearchText) {
        $searchBox.Text = ""
        $searchBox.Foreground = "Black"
    }
})

$searchBox.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($searchBox.Text)) {
        $searchBox.Text = $defaultSearchText
        $searchBox.Foreground = "Gray"
    }
})

# Tenant ID placeholder handling
$defaultTenantText = "company.onmicrosoft.com"
$tenantID.Add_GotFocus({
    if ($tenantID.Text -eq $defaultTenantText) {
        $tenantID.Text = ""
        $tenantID.Foreground = "Black"
    }
})

$tenantID.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($tenantID.Text)) {
        $tenantID.Text = $defaultTenantText
        $tenantID.Foreground = "Gray"
    }
})

# Import to InTune
$importToInTune.Add_Click({
    Write-ConsoleTextBox "Started import to InTune.."

    # Check if TenantID is valid
    if ([string]::IsNullOrWhiteSpace($tenantID.Text) -or $tenantID.Text -eq $defaultTenantText -or -not ($tenantID.Text -like "*.*")) {
        Write-ConsoleTextBox "Please enter a valid Tenant ID before importing to InTune."
        return
    }

    # List of files to check
    $filesToCheck = @(
        "WinGet-Wrapper.ps1",
        "WinGet-WrapperDetection.ps1", 
        "WinGet-WrapperRequirements.ps1",
        "WinGet-WrapperImportFromCSV.ps1",
        "IntuneWinAppUtil.exe"
    )

    $foundAllFiles = $true
    foreach ($file in $filesToCheck) {
        $fileFullPath = Join-Path -Path $scriptRoot -ChildPath $file

        if (-not (Test-Path -Path $fileFullPath -PathType Leaf)) {
            Write-ConsoleTextBox "File '$file' was not found."
            $foundAllFiles = $false
        }
        else {
            Write-ConsoleTextBox "File '$file' was found."
        }
    }

    if ($foundAllFiles) {
        Write-ConsoleTextBox "All required files found. Continue import to InTune..."

        if ($script:importListItems.Count -gt 0) {
            $fileName = "TempExport-$timestamp.csv"
            $csvFilePath = Join-Path -Path $scriptRoot -ChildPath $fileName
            
            try {
                $script:importListItems | Export-Csv -Path $csvFilePath -NoTypeInformation
                Write-ConsoleTextBox "Exported: $csvFilePath"

                # Prepare the Import script
                $logFile = "$scriptRoot\Logs\WinGet_WrapperImportFromCSV_$($TimeStamp).log"
                $importScriptPath = Join-Path -Path $scriptRoot -ChildPath "Winget-WrapperImportFromCSV.ps1"
                Write-ConsoleTextBox "ImportScriptPath: $importScriptPath"

                Write-ConsoleTextBox "****************************************************"
                Write-ConsoleTextBox "See log file for progress: $logFile"
                Write-ConsoleTextBox "****************************************************"

                # Run The Import Script
                $arguments = "-csvFile `"$csvFilePath`" -TenantID $($tenantID.Text) -ClientID $($clientID.Text) -RedirectURL `"$($redirectURI.Text)`" -ScriptRoot `"$scriptRoot`" -SkipConfirmation -SkipModuleCheck"
                Write-ConsoleTextBox "Arguments to be passed: $arguments"
                
                Start-Process powershell -ArgumentList "-NoProfile", "-ExecutionPolicy Bypass", "-File `"$importScriptPath`"", $arguments -Wait -NoNewWindow

                # Update GUI from log file
                Start-Sleep -Seconds 5
                Update-GUIFromLogFile -logFilePath "$logFile"

                # Remove temporary CSV
                if (Test-Path $csvFilePath) {
                    Remove-Item $csvFilePath -Force
                    Write-ConsoleTextBox "File $csvFilePath deleted successfully."
                }

                Write-ConsoleTextBox "****************************************************"
                Write-ConsoleTextBox "Import Log File: $logFile"
                Write-ConsoleTextBox "****************************************************"
            }
            catch {
                Write-ConsoleTextBox "Error during import: $($_.Exception.Message)"
            }
        }
        else {
            Write-ConsoleTextBox "No data to import."
        }
    }
    else {
        Write-ConsoleTextBox "Not all required files were found. Code will not run."
    }
})

# Update ConsoleProgress
Show-ConsoleProgress -PercentComplete 100 -Status "Successfully loaded Winget-Wrapper Import GUI"

# Initial greeting
Write-ConsoleTextBox "****************************************************"
Write-ConsoleTextBox "                           WinGet-Wrapper"
Write-ConsoleTextBox "  https://github.com/SorenLundt/WinGet-Wrapper"
Write-ConsoleTextBox ""
Write-ConsoleTextBox "              GNU General Public License v3"
Write-ConsoleTextBox "****************************************************"

# Show the WPF window
$window.ShowDialog() | Out-Null