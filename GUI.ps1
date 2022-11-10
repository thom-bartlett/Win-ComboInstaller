$Global:syncHash = [hashtable]::Synchronized(@{})
$newRunspace =[runspacefactory]::CreateRunspace()
$newRunspace.ApartmentState = "STA"
$newRunspace.ThreadOptions = "ReuseThread"
$newRunspace.Open()
$newRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)

# Load WPF assembly if necessary
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')

$psCmd = [PowerShell]::Create().AddScript({
param ($CurrentDir)
[xml]$form = @"
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        Title="WPromote Combo Installer" Height="525" Width="800" UseLayoutRounding="False" ResizeMode="NoResize">
    <Window.Resources>
        <Style x:Key="ImageStyle1" TargetType="{x:Type Image}">
            <Setter Property="Effect">
                <Setter.Value>
                    <DropShadowEffect/>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid Background="#FFEBEBEB" Margin="0,0,0,-1">
        <Image x:Name="Large_Icon" Style="{DynamicResource ImageStyle1}" HorizontalAlignment="Left" Height="177" Margin="31,131,0,0" VerticalAlignment="Top" Width="177" Source="$CurrentDir/Local.png">
        </Image>
        <Border BorderThickness="1" HorizontalAlignment="Left" Height="285" Margin="235,104,0,0" VerticalAlignment="Top" Width="530" Background="White" CornerRadius="5,5,5,5">
            <ScrollViewer VerticalScrollBarVisibility="Hidden" HorizontalScrollBarVisibility="Disabled" Margin="1,-1,-1,1">
                <Grid Margin="0,9,0,9">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="50"/>
                        <RowDefinition Height="50"/>
                        <RowDefinition Height="50"/>
                        <RowDefinition Height="50"/>
                        <RowDefinition Height="50"/>
                        <RowDefinition Height="50"/>
                        <RowDefinition Height="50"/>
                    </Grid.RowDefinitions>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="74"/>
                        <ColumnDefinition Width="Auto" MinWidth="75"/>
                        <ColumnDefinition Width="75"/>
                    </Grid.ColumnDefinitions>
                    <Image HorizontalAlignment="Center" Height="44" VerticalAlignment="Center" Width="34" Source="$CurrentDir/Chrome.png" Grid.Row="0"/>
                    <Label Grid.Column="1" Content="Google Chrome" HorizontalAlignment="Center" VerticalAlignment="Center" FontSize="22" Grid.Row="0"/>
                    <Line Grid.Row="0" Stroke="LightGray" StrokeThickness="1" VerticalAlignment="Bottom" Grid.ColumnSpan="3" Stretch="Uniform" X1="0" X2="120" Margin="10,0,-152,0"></Line>
                    <Image HorizontalAlignment="Center" Height="40" Grid.Row="1" VerticalAlignment="Center" Width="42" Source="$CurrentDir/Drive.png"/>
                    <ProgressBar x:Name="Progress_Chrome" Grid.Column="2" Height="20" Margin="165,0,-205,0" VerticalAlignment="Center"/>
                    <Label Grid.Column="1" Content="Google Drive" HorizontalAlignment="Left" Grid.Row="1" VerticalAlignment="Center" VerticalContentAlignment="Center" FontSize="22"/>
                    <ProgressBar x:Name="Progress_Drive" Grid.Column="2" Height="20" Margin="165,0,-205,0" VerticalAlignment="Center" Grid.Row="1"/>
                    <Line Grid.Row="1" Stroke="LightGray" StrokeThickness="1" VerticalAlignment="Bottom" Grid.ColumnSpan="3" Stretch="Uniform" X1="0" X2="120" Margin="10,0,-152,0"></Line>
                    <Image HorizontalAlignment="Center" Height="40" Grid.Row="2" VerticalAlignment="Center" Width="42" Source="$CurrentDir/Firefox.png"/>
                    <Label Grid.Column="1" Content="Firefox" HorizontalAlignment="Left" Grid.Row="2" VerticalAlignment="Center" VerticalContentAlignment="Center" FontSize="22"/>
                    <ProgressBar x:Name="Progress_Firefox" Grid.Column="2" Height="20" Margin="165,0,-205,0" VerticalAlignment="Center" Grid.Row="2"/>
                    <Line Grid.Row="2" Stroke="LightGray" StrokeThickness="1" VerticalAlignment="Bottom" Grid.ColumnSpan="3" Stretch="Uniform" X1="0" X2="120" Margin="10,0,-152,0"></Line>
                    <Image HorizontalAlignment="Center" Height="40" Grid.Row="3" VerticalAlignment="Center" Width="42" Source="$CurrentDir/Slack.png"/>
                    <Label Grid.Column="1" Content="Slack" HorizontalAlignment="Left" Grid.Row="3" VerticalAlignment="Center" VerticalContentAlignment="Center" FontSize="22"/>
                    <ProgressBar x:Name="Progress_Slack" Grid.Column="2" Height="20" Margin="165,0,-205,0" VerticalAlignment="Center" Grid.Row="3"/>
                    <Line Grid.Row="3" Stroke="LightGray" StrokeThickness="1" VerticalAlignment="Bottom" Grid.ColumnSpan="3" Stretch="Uniform" X1="0" X2="120" Margin="10,0,-152,0"></Line>
                    <Image HorizontalAlignment="Center" Height="40" Grid.Row="4" VerticalAlignment="Center" Width="42" Source="$CurrentDir/Splashtop.png"/>
                    <Label Grid.Column="1" Content="Splashtop" HorizontalAlignment="Left" Grid.Row="4" VerticalAlignment="Center" VerticalContentAlignment="Center" FontSize="22"/>
                    <ProgressBar x:Name="Progress_Splashtop" Grid.Column="2" Height="20" Margin="165,0,-205,0" VerticalAlignment="Center" Grid.Row="4"/>
                    <Line Grid.Row="4" Stroke="LightGray" StrokeThickness="1" VerticalAlignment="Bottom" Grid.ColumnSpan="3" Stretch="Uniform" X1="0" X2="120" Margin="10,0,-152,0"></Line>
                    <Image HorizontalAlignment="Center" Height="40" Grid.Row="5" VerticalAlignment="Center" Width="42" Source="$CurrentDir/Evernote.png"/>
                    <Label Grid.Column="1" Content="Evernote" HorizontalAlignment="Left" Grid.Row="5" VerticalAlignment="Center" VerticalContentAlignment="Center" FontSize="22"/>
                    <ProgressBar x:Name="Progress_Evernote" Grid.Column="2" Height="20" Margin="165,0,-205,0" VerticalAlignment="Center" Grid.Row="5"/>
                    <Line Grid.Row="5" Stroke="LightGray" StrokeThickness="1" VerticalAlignment="Bottom" Grid.ColumnSpan="3" Stretch="Uniform" X1="0" X2="120" Margin="10,0,-152,0"></Line>
                    <Image HorizontalAlignment="Center" Height="40" Grid.Row="6" VerticalAlignment="Center" Width="42" Source="$CurrentDir/Falcon.png"/>
                    <Label Grid.Column="1" Content="Falcon" HorizontalAlignment="Left" Grid.Row="6" VerticalAlignment="Center" VerticalContentAlignment="Center" FontSize="22"/>
                    <ProgressBar x:Name="Progress_Falcon" Grid.Column="2" Height="20" Margin="165,0,-205,0" VerticalAlignment="Center" Grid.Row="6"/>
                    <Line Grid.Row="6" Stroke="LightGray" StrokeThickness="1" VerticalAlignment="Bottom" Grid.ColumnSpan="3" Stretch="Uniform" X1="0" X2="120" Margin="10,0,-152,0"></Line>
                    <Label x:Name="Progress_Chrome_Label" Grid.Column="2" Content="Not started" Height="28" VerticalAlignment="Center" HorizontalContentAlignment="Center" VerticalContentAlignment="Center" FontSize="14" Margin="165,0,-205,0" ScrollViewer.VerticalScrollBarVisibility="Disabled"/>
                    <Label x:Name="Progress_Drive_Label" Grid.Column="2" Content="Not started" Height="28" VerticalAlignment="Center" HorizontalContentAlignment="Center" VerticalContentAlignment="Center" FontSize="14" Margin="165,0,-205,0" ScrollViewer.VerticalScrollBarVisibility="Disabled" Grid.Row="1"/>
                    <Label x:Name="Progress_Firefox_Label" Grid.Column="2" Content="Not started" Height="28" VerticalAlignment="Center" HorizontalContentAlignment="Center" VerticalContentAlignment="Center" FontSize="14" Margin="165,0,-205,0" ScrollViewer.VerticalScrollBarVisibility="Disabled" Grid.Row="2"/>
                    <Label x:Name="Progress_Slack_Label" Grid.Column="2" Content="Not started" Height="28" VerticalAlignment="Center" HorizontalContentAlignment="Center" VerticalContentAlignment="Center" FontSize="14" Margin="165,0,-205,0" ScrollViewer.VerticalScrollBarVisibility="Disabled" Grid.Row="3"/>
                    <Label x:Name="Progress_Splashtop_Label" Grid.Column="2" Content="Not started" Height="28" VerticalAlignment="Center" HorizontalContentAlignment="Center" VerticalContentAlignment="Center" FontSize="14" Margin="165,0,-205,0" ScrollViewer.VerticalScrollBarVisibility="Disabled" Grid.Row="4"/>
                    <Label x:Name="Progress_Evernote_Label" Grid.Column="2" Content="Not started" Height="28" VerticalAlignment="Center" HorizontalContentAlignment="Center" VerticalContentAlignment="Center" FontSize="14" Margin="165,0,-205,0" ScrollViewer.VerticalScrollBarVisibility="Disabled" Grid.Row="5"/>
                    <Label x:Name="Progress_Falcon_Label" Grid.Column="2" Content="Not started" Height="28" VerticalAlignment="Center" HorizontalContentAlignment="Center" VerticalContentAlignment="Center" FontSize="14" Margin="165,0,-205,0" ScrollViewer.VerticalScrollBarVisibility="Disabled" Grid.Row="6"/>
                </Grid>
            </ScrollViewer>
        </Border>
        <Label Content="Installing Applications" HorizontalAlignment="Center" Height="60" Margin="0,-1,0,0" VerticalAlignment="Top" Width="480" HorizontalContentAlignment="Center" FontSize="30" FontWeight="Bold"/>
        <Line Stroke="LightGray" StrokeThickness="1" VerticalAlignment="Top" Stretch="Uniform" X1="0" X2="120" Margin="20,51,20,0"></Line>
        <Label Content="Please wait while the following apps are downloaded and installed." HorizontalAlignment="Left" Height="25" Margin="240,69,0,0" VerticalAlignment="Top" Width="525" HorizontalContentAlignment="Center"/>
        <ProgressBar x:Name="Overall_Progress" HorizontalAlignment="Center" Height="5" Margin="50,399,50,0" VerticalAlignment="Top" Width="790" Foreground="#FF0585D4"/>
        <Label Margin="75,409,75,0" VerticalAlignment="Top" HorizontalContentAlignment="Center" Height="45">
            <TextBlock x:Name="App_Description" TextAlignment="Center" TextWrapping="WrapWithOverflow" Text="" VerticalAlignment="Center" />
        </Label>
        <Button x:Name="Done" Content="Done" HorizontalAlignment="Left" Height="25" Margin="690,456,0,0" VerticalAlignment="Top" Width="90" IsEnabled="False"/>
    </Grid>
</Window>
"@

$Reader = (New-Object System.Xml.XmlNodeReader $form)
$syncHash.Window = [Windows.Markup.XamlReader]::Load($Reader)
$syncHash.Chrome = $syncHash.Window.FindName('Progress_Chrome')
$syncHash.ChromeLabel = $syncHash.Window.FindName('Progress_Chrome_Label')
$syncHash.Drive = $syncHash.Window.FindName('Progress_Drive')
$syncHash.DriveLabel = $syncHash.Window.FindName('Progress_Drive_Label')
$syncHash.Firefox = $syncHash.Window.FindName('Progress_Firefox')
$syncHash.FirefoxLabel = $syncHash.Window.FindName('Progress_Firefox_Label')
$syncHash.Slack = $syncHash.Window.FindName('Progress_Slack')
$syncHash.SlackLabel = $syncHash.Window.FindName('Progress_Slack_Label')
$syncHash.Splashtop = $syncHash.Window.FindName('Progress_Splashtop')
$syncHash.SplashtopLabel = $syncHash.Window.FindName('Progress_Splashtop_Label')
$syncHash.Evernote = $syncHash.Window.FindName('Progress_Evernote')
$syncHash.EvernoteLabel = $syncHash.Window.FindName('Progress_Evernote_Label')
$syncHash.Falcon = $syncHash.Window.FindName('Progress_Falcon')
$syncHash.FalconLabel = $syncHash.Window.FindName('Progress_Falcon_Label')
$syncHash.Description = $syncHash.Window.FindName('App_Description')
$syncHash.OverallProgress = $syncHash.Window.FindName('Overall_Progress')
$syncHash.LargeIcon = $syncHash.Window.FindName('Large_Icon')
$syncHash.Done = $syncHash.Window.FindName('Done')
$syncHash.Done.Add_Click({
    $syncHash.Window.Close()
})
$syncHash.Window.ShowDialog() | Out-Null
})


function Get-LocalVersion-Splashtop{
    [version] $Current = (Get-Package -name 'Splashtop Streamer' -ErrorAction SilentlyContinue).Version
    return $Current
}

function Get-OnlineVersion-Splashtop{
    # no reliable way found to pull online version
    $Download = Get-RedirectedURL https://redirect.splashtop.com/src/win
    [version] $version = [regex]::Match($Download, '_v([\d\.]+)\.exe$').Groups[1].Value
    return $version
}

function Get-Software-Splashtop{
    mkdir -Path $env:temp\softwareinstall -erroraction SilentlyContinue | Out-Null
    $Download = join-path $env:temp\softwareinstall Splashtop_installer.exe
    (new-object System.Net.WebClient).DownloadFile('https://my.splashtop.com/csrs/win',$Download)
    $sig = Get-AuthenticodeSignature $Download
    if ($sig.SignerCertificate.Thumbprint -eq "B980FE0A5338B8A39A9A1B732CEAEA90B69D7222") {
        #Invoke-Expression "$Download /silent /install"
        Start-Process $Download -argumentlist "prevercheck /s /i dcode=ZH2P522JLST3,confirm_d=0,hidewindow=1" -wait   
    }
    else {
        Write-Host "Signatures do not match... exiting..."
        exit 1
    }
}

function Get-Splashtop{
    Update-AppProgress -App Splashtop -Phase CheckVersion
    Write-Host 'Getting versions...'
    $local = Get-LocalVersion-Splashtop
    Write-Host "Local version: $local"
    $online = Get-OnlineVersion-Splashtop
    Write-Host "Online version: $online"
    
    if ($local -lt $online) {
        Write-Host 'Splashtop needs to be installed/updated. Downloading...'
        Update-AppProgress -App Splashtop -Phase Install
        Get-Software-Splashtop
    }

    Update-AppProgress -App Splashtop -Phase Done
}

function Get-LocalVersion-Chrome{
    $software_location = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe'
    if (Test-Path $software_Location){
        $version = (Get-Item (Get-ItemProperty $software_location).'(Default)').VersionInfo
        Write-Host $version.ProductVersion
        return $version.ProductVersion
    }
    else {
        Write-Host "Chrome is not installed"
        return $false}
}

function Get-OnlineVersion-Chrome{
    $raw = Invoke-RestMethod 'http://omahaproxy.appspot.com/all?os=win&amp;channel=stable'
    $win64 = $raw.split("`n") | Select-String -Pattern "win,stable" | select-object -First 1
    $version = $win64.Line.split(",")[2]
    Write-Host "Online Version" $version
    return $version
}

function Get-Software-Chrome{
    mkdir -Path $env:temp\softwareinstall -erroraction SilentlyContinue | Out-Null
    $Download = join-path $env:temp\softwareinstall googlechromestandaloneenterprise64.msi
    (new-object System.Net.WebClient).DownloadFile('https://dl.google.com/tag/s/dl/chrome/install/googlechromestandaloneenterprise64.msi',$Download)
    $sig = Get-AuthenticodeSignature $Download
    if ($sig.SignerCertificate.Thumbprint -eq "2673EA6CC23BEFFDA49AC715B121544098A1284C") {
        #Invoke-Expression "$Download /silent /install"
        Start-Process $Download -argumentlist "/qn" -wait   
    }
    else {
        Write-Host "Signatures do not match... exiting..."
        return
    }
}
function Get-Chrome{
    Update-AppProgress -App Chrome -Phase CheckVersion
    $online = Get-OnlineVersion-Chrome
    $local = Get-LocalVersion-Chrome
    
    if ($local -ne $online) {
        Update-AppProgress -App Chrome -Phase Install
        Get-Software-Chrome
    }

    Update-AppProgress -App Chrome -Phase Done
}

function Get-Drive{
    Update-AppProgress -App Drive -Phase CheckVersion
    Write-Host 'Getting latest Drive versions'
    $DriveVersions = Invoke-WebRequest -UseBasicParsing -Uri https://support.google.com/a/answer/7577057?hl=en
    Write-Host 'Getting latest 2 versions'
    $Latests = [regex]::Matches($DriveVersions.Content, 'Version (\d+\.\d+)')[0..1]
    [version] $Latest = $Latests[0].Groups[1].Value
    Write-Output "Latest version: $Latest"
    [version] $SecondLatest = $Latests[1].Groups[1].Value
    Write-Output "Second latest: $SecondLatest"
    [version] $Current = (Get-Package -Name '*Google Drive*' -ErrorAction SilentlyContinue).Version
    Write-Output "Current version: $Current"
    if (-not ($Current -gt $Latest -or $Current -gt $SecondLatest)) {
        Write-Host "Drive not installed or latest version. Downloading..."
        Update-AppProgress -App Drive -Phase Install
        $DownloadDir = "$env:temp\driveinstall"
        $Installer = "$DownloadDir\googledrive.exe"
        New-Item -Path $DownloadDir -ItemType Directory -Force > $null
        (new-object System.Net.WebClient).DownloadFile('https://dl.google.com/drive-file-stream/GoogleDriveSetup.exe',$Installer)
        $sig = Get-AuthenticodeSignature $Installer
        if ($sig.SignerCertificate.Thumbprint -eq "2673EA6CC23BEFFDA49AC715B121544098A1284C") {
            Write-Host "Signatures Match... Installing Drive..."
            Start-Process $Installer -argumentlist "--silent --desktop_shortcut" -wait   
        }
        else {
            Write-Host "Signatures do not match... exiting..."
            exit 1
        }
    }
    
    Update-AppProgress -App Drive -Phase Done
}

function Get-LocalVersion-Firefox{
    $firefox_Location = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\firefox.exe'
    if (Test-Path $firefox_Location){
        $firefox = (Get-Item (Get-ItemProperty  $firefox_Location).'(Default)').VersionInfo
        Write-Host "Local Version" $firefox.ProductVersion
        return [version]($firefox.ProductVersion)
    }
    else {
        Write-Host "Firefox is not installed"
        return $false}
}

function Get-OnlineVersion-Firefox{
    $data = Invoke-WebRequest -uri https://product-details.mozilla.org/1.0/firefox_versions.json | ConvertFrom-Json
    [version] $version = $data.LATEST_FIREFOX_VERSION
    Write-Host "Online Version" $version
    return $version
} 

function Get-Software-Firefox{
    mkdir -Path $env:temp\firefoxinstall -erroraction SilentlyContinue | Out-Null
    $Download = join-path $env:temp\firefoxinstall firefox_installer.exe
    Invoke-WebRequest 'https://download.mozilla.org/?product=firefox-latest&os=win64&lang=en-US' -outfile $Download
    $sig = Get-AuthenticodeSignature $Download
    if ($sig.SignerCertificate.Thumbprint -eq "1326B39C3D5D2CA012F66FB439026F7B59CB1974") {
        Write-Host "Signatures Match... will install Firefox..."
        #Invoke-Expression "$Download /silent /install"
        Start-Process $Download -argumentlist "/S" -wait   
    }
    else {
        Write-Host "Signatures do not match... exiting..."
    }
}

function Get-Firefox{
    Update-AppProgress -App Firefox -Phase CheckVersion
    $online = Get-OnlineVersion-Firefox
    $local = Get-LocalVersion-Firefox

    if ($local -lt $online){
        Update-AppProgress -App Firefox -Phase Install
        Get-Software-Firefox
    }

    Update-AppProgress -App Firefox -Phase Done
}

function Get-LocalVersion-Slack{
    [version] $Current = (Get-Package -Name 'Slack (Machine - MSI)' -ErrorAction SilentlyContinue).Version
    return $Current
}

Function Get-RedirectedUrl{
    Param (
        [Parameter(Mandatory=$true)]
        [String]$URL
    )

    $request = [System.Net.WebRequest]::Create($url)
    $request.AllowAutoRedirect=$false
    $response=$request.GetResponse()

    If ($response.StatusCode -eq "Found")
    {
        $response.GetResponseHeader("Location")
    }
}

function Get-OnlineVersion-Slack{
    $releases = 'https://slack.com/intl/en-nl/downloads/windows'
    $download_page = Invoke-WebRequest -Uri $releases -UseBasicParsing
    $url64 = Get-RedirectedUrl 'https://slack.com/ssb/download-win64-msi'
    $re = "Version (.+\d)</span>"
    [version] $version = [regex]::Match($download_page.RawContent, $re).Groups[1].Value
    write-host "Slack online version: $version"
    return $version
}

function Get-Software-Slack{
    mkdir -Path $env:temp\softwareinstall -erroraction SilentlyContinue | Out-Null
    $Download = join-path $env:temp\softwareinstall slack.msi
    (new-object System.Net.WebClient).DownloadFile('https://slack.com/ssb/download-win64-msi',$Download)
    $sig = Get-AuthenticodeSignature $Download
    if ($sig.SignerCertificate.Thumbprint -eq "9E28163A554C92D8B7DEFE126DF2D1ABE767563F") {
        #Invoke-Expression "$Download /silent /install"
        Start-Process $Download -argumentlist "/qn /norestart" -wait   
    }
    else {
        Write-Host "Signatures do not match... exiting..."
        exit 1
    }
}

function Get-Slack{
    Update-AppProgress -App Slack -Phase CheckVersion
    $online = Get-OnlineVersion-Slack
    $local = Get-LocalVersion-Slack

    if ($local -lt $online){
        Update-AppProgress -App Slack -Phase Install
        Get-Software-Slack
    }

    Update-AppProgress -App Slack -Phase Done
}

function Get-LocalVersion-Evernote{
    $software_location = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\e4251011-875e-51f3-a464-121adaff5aaa'
    if (Test-Path $software_Location){
        $software = Get-ItemProperty $software_Location
        [version] $version = $software.DisplayVersion
        Write-Host "Evernote Local version: $version"
        return $version
    }
    else {
        Write-Host "Evernote is not installed"
        return $false}
}

function Get-OnlineVersion-Evernote{
    $releases = 'https://evernote.com/download/'
    $download_page = Invoke-WebRequest -Uri $releases -UseBasicParsing
    $re = "Evernote[_-](.+)(-setup?).exe"
    $url32 = $download_page.Links | Where-Object href -match $re | Select-Object -First 1 -expand href
    $longversion = ([regex]::Match($download_page.RawContent, $re)).Captures.Groups[1].value
    [version] $version = $longversion.split("-")[0]
    Write-Host "Evernote online version: $version"
    return $version
}

function Get-Software-Evernote{
    $releases = 'https://evernote.com/download/'
    $re = "Evernote[_-](.+)(-setup?).exe"
    $download_page = Invoke-WebRequest -Uri $releases -UseBasicParsing
    $url32 = $download_page.Links | Where-Object href -match $re | Select-Object -First 1 -expand href
    mkdir -Path $env:temp\softwareinstall -erroraction SilentlyContinue | Out-Null
    $Download = join-path $env:temp\softwareinstall evernote.exe
    (new-object System.Net.WebClient).DownloadFile($url32, $Download)
    $sig = Get-AuthenticodeSignature $Download
    if ($sig.SignerCertificate.Thumbprint -eq "A4059A822F41530EA76E8F037B92CADD37716F97") {
        #Invoke-Expression "$Download /silent /install"
        Start-Process $Download -argumentlist "/AllUsers /S" -wait   
    }
    else {
        Write-Host "Signatures do not match... exiting..."
        exit 1
    }
}

function Get-Evernote{
    Update-AppProgress -App Evernote -Phase CheckVersion
    $online = Get-OnlineVersion-Evernote
    $local = Get-LocalVersion-Evernote

    if ($local -lt $online){
        Update-AppProgress -App Evernote -Phase Install
        Get-Software-Evernote
    }

    Update-AppProgress -App Evernote -Phase Done
}

function Install-Falcon {
    Update-AppProgress -App 'Falcon' -Phase CheckVersion
    $license = Get-Content "$PSSCRIPTROOT\license.txt"
    if (Test-Path 'C:\Program Files\CrowdStrike\CSFalconService.exe'){
        write-host "Falcon Installed already... Skipping... "
        "Falcon Installed"
    }
    else {
        Update-AppProgress -App 'Falcon' -Phase Install
        Write-Host "Installing Falcon"
        try {
            $Splat = @{   
                'FilePath' = "$PSScriptRoot\WindowsSensor.exe"
                'ArgumentList' = "/install /quiet /norestart CID=$license"
                'Wait' = $true
                'PassThru' = $false
                'ErrorAction' = 'Stop'
            }
            try {
                $Process = Start-Process @Splat
                $Result = 'Done'
            }
            catch {
                $Result = 'Fail'
            }
        }
        catch {
            $Result = 'Fail'
        }
        if ($Process.ExitCode -ne 0) {
            $Result = 'Fail'
        }
    }
    Update-AppProgress -App 'Falcon' -Phase $Result
}

function Install-Nextiva {
    Update-AppProgress -App 'Nextiva' -Phase CheckVersion
    if (Test-Path 'C:\Program Files (x86)\Nextiva, Inc\Nextiva App\Communicator.exe'){
        write-host "Nextiva Installed already... Skipping... "
    }
    else {
        Update-AppProgress -App 'Nextiva' -Phase Install
        Write-Host "Installing Nextiva"
        $Splat = @{
            'FilePath' = "$PSScriptRoot\Nextiva_App.bc-uc.win-22.9.33.39.msi"
            'ArgumentList' = "ALLUSERS=1 /qn"
            'Wait' = $true
            'PassThru' = $true
            'ErrorAction' = 'Stop'
        }
        try {
            $Process = Start-Process @Splat
            $Result = 'Done'
        }
        catch {
            $Result = 'Fail'
        }
        if ($Process.ExitCode -ne 0) {
            $Result = 'Fail'
        }
    }
    Update-AppProgress -App 'Nextiva' -Phase $Result
}

function Update-WindowElement {
    param(
        [ValidateSet(
            'Chrome',
            'ChromeLabel',
            'Drive',
            'DriveLabel',
            'Firefox',
            'FirefoxLabel',
            'Slack',
            'SlackLabel',
            'Splashtop',
            'SplashtopLabel',
            'Evernote',
            'EvernoteLabel',
            'Falcon',
            'FalconLabel',
            'Nextiva',
            'NextivaLabel',
            'Description',
            'OverallProgress',
            'LargeIcon',
            'Done'
        )]
        [string]$Element,
        [string]$Attribute,
        $Value
    )

    $Global:syncHash.Window.Dispatcher.Invoke(
        [action]{$syncHash.$Element.$Attribute = $Value},"Normal"
    )
}

function Update-AppProgress {
    param (
        [string] $App,
        [ValidateSet(
            'CheckVersion',
            'Install',
            'Done',
            'Fail'
        )]
        [string] $Phase
    )

    switch ($Phase) {
        'CheckVersion' {
            $Progress = 33
            $Description = 'Checking for updates'
        }
        'Install' {
            $Progress = 66
            $Description = 'Installing'
        }
        'Done' {
            $Progress = 100
            $Description = 'Installed'
        }
        'Fail' {
            $Progress = 100
            $Description = 'Failed'
            Update-WindowElement -Element $App -Attribute Foreground -Value '#D2042D'
        }
    }

    Update-WindowElement -Element $App -Attribute Value -Value $Progress
    Update-WindowElement -Element "$App`Label" -Attribute Content -Value $Description
}

function Install-Application {
    param(
        [string]$Name
    )

    switch ($Name) {
        'Chrome' {Get-Chrome}
        'Drive' {Get-Drive}
        'Firefox' {Get-Firefox}
        'Slack' {Get-Slack}
        'Splashtop' {Get-Splashtop}
        'Evernote' {Get-Evernote}
        'Falcon' {Install-Falcon}
        'Nextiva' {Install-Nextiva}
    }
}

function main{
    $Apps = [ordered]@{
        'Chrome' = 'A fast, secure, and free web browser built for the modern web. Chrome syncs bookmarks across all your devices, fills out forms automatically, and so much more.'
        'Drive' = 'Google Drive allows users to store files in the cloud, synchronize files across devices, and share files.'
        'Firefox' = 'Firefox is a free web browser backed by Mozilla, a non-profit dedicated to internet health and privacy.'
        'Slack' = "Slack is a new way to communicate with your team. It's faster, better organized, and more secure than email"
        'Splashtop' = 'Splashtop enables users to remotely access or remotely support computers from desktop and mobile devices.'
        'Evernote' = 'Evernote is a powerful tool that can help executives, entrepreneurs and creative people capture and arrange their ideas.'
        'Falcon' = 'Falcon stops breaches via a unified set of cloud-delivered technologies that prevent all types of attacks - including malware and much more.'
    }
    $TotalApps = $Apps.Keys.Count
    $Processed = 0
    foreach ($app in $Apps.Keys) {
        try {
            Write-Host -Object "Starting $App"
            Update-WindowElement -Element LargeIcon -Attribute Source -Value "$PSScriptRoot\$app.png"
            Update-WindowElement -Element Description -Attribute Text -Value $Apps.$app
            Install-Application -Name $app
        }
        catch {
            Update-AppProgress -App $app -Phase Fail
        }
        $Processed++
        Write-Host -Object "Setting overall progress to $((($Processed / $TotalApps) * 100))"
        Update-WindowElement -Element OverallProgress -Attribute Value -Value (($Processed / $TotalApps) * 100)
        Write-Host -Object "Finished $App"
    }

    Write-Host "Cleaning up"
    Remove-Item $env:temp\softwareinstall -Force -Recurse

    Update-WindowElement -Element Done -Attribute IsEnabled -Value $true
    Update-WindowElement -Element LargeIcon -Attribute Source -Value "$PSScriptRoot\Done.png"
    Update-WindowElement -Element Description -Attribute Text -Value 'Finished! Click the Done button to exit.'

    # Wait for window to close before exiting script
    while ($data.IsCompleted -eq $false) {
        Start-Sleep -Milliseconds 100
    }
}


$psCmd.AddArgument($PWD)
$psCmd.Runspace = $newRunspace
$data = $psCmd.BeginInvoke()

while (-not $syncHash.Done) {
    Start-Sleep -Milliseconds 500
}

main