<# 
    Network Adapter Manager (WPF)
    Fixes:
      - ScheduledTask LogonType=Interactive
      - Correct MessageBox overload (YesNo + icon)
      - Fixed join precedence in confirm list
      - Selections persist after actions
      - Auto-load profile on startup (if present)
      - Window widened to 1300
      - "Create No-UAC..." checkbox checked by default
#>

#region --- Self-Elevation + STA relaunch ---
function Start-SelfElevatedSta {
    try {
        $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
                   ).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
        $isSta   = [Threading.Thread]::CurrentThread.ApartmentState -eq 'STA'
        if ($isAdmin -and $isSta) { return }

        $exe  = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
        $scriptPath = if ($PSCommandPath) { $PSCommandPath } elseif ($MyInvocation.MyCommand.Path) { $MyInvocation.MyCommand.Path } else { $null }
        if (-not $scriptPath) { throw "Cannot determine script path for relaunch." }

        $args = @('-NoProfile','-ExecutionPolicy','Bypass','-STA','-File',"`"$scriptPath`"")
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName  = $exe
        $psi.Arguments = ($args -join ' ')
        $psi.Verb      = 'runas'
        [Diagnostics.Process]::Start($psi) | Out-Null
        exit
    } catch {
        try { Add-Type -AssemblyName PresentationFramework } catch {}
        [System.Windows.MessageBox]::Show(
            "Failed to self-elevate or relaunch in STA:`n$($_.Exception.Message)",
            "Startup Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
        exit 1
    }
}
Start-SelfElevatedSta
#endregion

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

#region --- Paths, helpers ---
$AppTitle = "Network Adapter Manager"
$Script:AppTitle = $AppTitle

$ScriptPath = if ($PSCommandPath) { $PSCommandPath } elseif ($MyInvocation.MyCommand.Path) { $MyInvocation.MyCommand.Path } else { $null }
$ScriptDir  = if ($ScriptPath) { Split-Path -Path $ScriptPath -Parent } else { (Get-Location).Path }
$ProfilePath = Join-Path $ScriptDir 'adapter-profile.json'

function Show-ErrorDialog([string]$msg, [string]$caption="$AppTitle - Error") {
    [System.Windows.MessageBox]::Show($msg, $caption, [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
}
function Update-Status([string]$text, [bool]$isError=$false) {
    if ($StatusTextBlock) {
        $StatusTextBlock.Text = $text
        $StatusTextBlock.Foreground = if ($isError) { 'Tomato' } else { 'LimeGreen' }
    }
}
function New-Brush([string]$color) {
    try { ([System.Windows.Media.BrushConverter]::new()).ConvertFromString($color) } catch { [System.Windows.Media.Brushes]::Transparent }
}
#endregion

#region --- Adapter discovery & resolution ---
function Get-AdapterInventory {
    $adapters = @()
    try { $net = Get-NetAdapter -IncludeHidden -ErrorAction SilentlyContinue } catch { $net = @() }
    $cimAll = @(); try { $cimAll = Get-CimInstance -ClassName Win32_NetworkAdapter -ErrorAction SilentlyContinue } catch {}

    foreach ($n in $net) {
        $pnpId=$null;$desc=$n.InterfaceDescription;$guid=$n.InterfaceGuid;$ifIndex=$n.ifIndex
        if ($ifIndex -ne $null) {
            $cim = $cimAll | Where-Object { $_.InterfaceIndex -eq $ifIndex } | Select-Object -First 1
            if ($cim) { $pnpId=$cim.PNPDeviceID }
        }
        elseif ($guid) {
            try {
                $guidStr=$guid.ToString()
                $cim=$cimAll | Where-Object { $_.GUID -and $_.GUID -eq $guidStr } | Select-Object -First 1
                if($cim){$pnpId=$cim.PNPDeviceID}
            } catch {}
        }
        $type = switch -Regex ("$($n.InterfaceDescription) $($n.Name)") {
            'wi[-\s]?fi|wireless' {'Wi-Fi'}
            'ethernet|gigabit|realtek pci|intel\(r\).*ethernet' {'Ethernet'}
            'vpn|virtual|hyper-v|vmware|loopback|tap|wan miniport|wireguard' {'Virtual/VPN'}
            default { if ($n.MediaType -eq 'Native802_11') {'Wi-Fi'} else {$n.MediaType} }
        }
        $adapters += [pscustomobject]@{
            Selected=$false
            Alias=$n.Name
            Status="$($n.Status)"
            InterfaceType="$type"
            InterfaceDescription=$desc
            MacAddress=($n.MacAddress -replace '-',':')
            InterfaceGuid= if($guid){$guid.ToString()} else {$null}
            IfIndex=$n.ifIndex
            PnPInstanceId=$pnpId
        }
    }
    foreach ($c in $cimAll) {
        $guidStr=$null; try { if ($c.GUID) { $guidStr=[string]$c.GUID } } catch {}
        if ($guidStr -and ($adapters.InterfaceGuid -contains $guidStr)) { continue }
        if ($c.PNPClass -ne 'NET' -and $c.NetConnectionID -eq $null -and $c.AdapterType -eq $null) { continue }
        $alias = if ($c.NetConnectionID) { $c.NetConnectionID } else { $c.Name }
        $status = switch ($c.NetEnabled) { $true{'Up'} $false{'Disabled'} default{'Unknown'} }
        $type = switch -Regex ("$($c.Name) $($c.Description)") {
            'wi[-\s]?fi|wireless'{'Wi-Fi'}
            'ethernet|gigabit|realtek pci|intel\(r\).*ethernet'{'Ethernet'}
            'vpn|virtual|hyper-v|vmware|loopback|tap|wan miniport|wireguard'{'Virtual/VPN'}
            default{$c.AdapterType}
        }
        $adapters += [pscustomobject]@{
            Selected=$false
            Alias=$alias
            Status=$status
            InterfaceType="$type"
            InterfaceDescription=$c.Description
            MacAddress=($c.MACAddress -replace '-',':')
            InterfaceGuid=$guidStr
            IfIndex=$c.InterfaceIndex
            PnPInstanceId=$c.PNPDeviceID
        }
    }
    $adapters |
        Sort-Object Alias |
        Group-Object { "{0}|{1}|{2}" -f $_.InterfaceGuid,$_.PnPInstanceId,$_.InterfaceDescription } |
        ForEach-Object { $_.Group | Select-Object -First 1 }
}
function Resolve-Adapter {
    param([hashtable]$Identity)
    try {
        $candidates = Get-NetAdapter -IncludeHidden -ErrorAction SilentlyContinue
        if (-not $candidates) { return $null }
        $scored = foreach($c in $candidates){
            $s=0
            if ($Identity.InterfaceGuid -and $c.InterfaceGuid -and ($c.InterfaceGuid.ToString() -ieq $Identity.InterfaceGuid)) {$s+=100}
            if ($Identity.Alias -and $c.Name -and ($c.Name -ieq $Identity.Alias)) {$s+=40}
            if ($Identity.MacAddress -and $c.MacAddress -and (($c.MacAddress -replace '-|:','') -ieq ($Identity.MacAddress -replace '-|:',''))) {$s+=30}
            if ($Identity.InterfaceDescription -and $c.InterfaceDescription -and ($c.InterfaceDescription -ieq $Identity.InterfaceDescription)) {$s+=20}
            if ($Identity.PnPInstanceId) {
                try {
                    $cim=Get-CimInstance Win32_NetworkAdapter -Filter "InterfaceIndex=$($c.ifIndex)" -ErrorAction SilentlyContinue
                    if($cim -and $cim.PNPDeviceID -and ($cim.PNPDeviceID -ieq $Identity.PnPInstanceId)){$s+=10}
                } catch {}
            }
            [pscustomobject]@{Score=$s;Adapter=$c}
        }
        $best=$scored|Sort-Object Score -Descending|Select-Object -First 1
        if ($best.Score -gt 0) {$best.Adapter} else {$null}
    } catch { $null }
}
#endregion

#region --- Exported script template (escaped braces) ---
function Get-ActionScriptContent {
param(
    [ValidateSet('Enable','Disable')] [string] $Action,
    [Parameter(Mandatory=$true)] [array] $Identities
)
    $identityLines = @()
    foreach ($id in $Identities) {
        $identityLines += ('    @{{ InterfaceGuid="{0}"; Alias="{1}"; MacAddress="{2}"; InterfaceDescription="{3}"; PnPInstanceId="{4}" }}' -f `
            ($id.InterfaceGuid -replace '"','`"'),
            ($id.Alias -replace '"','`"'),
            ($id.MacAddress -replace '"','`"'),
            ($id.InterfaceDescription -replace '"','`"'),
            ($id.PnPInstanceId -replace '"','`"')
        )
    }
    $identitiesBlock = $identityLines -join [Environment]::NewLine

    # IMPORTANT: correct here-string delimiters for a literal block
    $template = @'
<# Auto-generated: __ACTION__ adapters
    - Self-elevates & ensures STA
    - Robust resolution (GUID>Alias>MAC>Description>PnP)
#>
function Start-SelfElevatedSta {
    try {
        $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
        $isSta   = [Threading.Thread]::CurrentThread.ApartmentState -eq 'STA'
        if ($isAdmin -and $isSta) { return }
        $exe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
        $args = @('-NoProfile','-ExecutionPolicy','Bypass','-STA','-File',"`"$($MyInvocation.MyCommand.Path)`"")
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $exe; $psi.Arguments = ($args -join ' '); $psi.Verb='runas'
        [Diagnostics.Process]::Start($psi) | Out-Null; exit
    } catch {
        try { Add-Type -AssemblyName PresentationFramework } catch {}
        [System.Windows.MessageBox]::Show("Failed to self-elevate:`n$($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null; exit 1
    }
}
Add-Type -AssemblyName PresentationFramework | Out-Null
Start-SelfElevatedSta

function Resolve-Adapter { param([hashtable]$Identity)
    try {
        $candidates = Get-NetAdapter -IncludeHidden -ErrorAction SilentlyContinue
        if (-not $candidates) { return $null }
        $scored = foreach ($c in $candidates) {
            $s=0
            if ($Identity.InterfaceGuid -and $c.InterfaceGuid -and ($c.InterfaceGuid.ToString() -ieq $Identity.InterfaceGuid)) { $s+=100 }
            if ($Identity.Alias         -and $c.Name -and ($c.Name -ieq $Identity.Alias)) { $s+=40 }
            if ($Identity.MacAddress    -and $c.MacAddress -and (($c.MacAddress -replace '-|:','') -ieq ($Identity.MacAddress -replace '-|:',''))) { $s+=30 }
            if ($Identity.InterfaceDescription -and $c.InterfaceDescription -and ($c.InterfaceDescription -ieq $Identity.InterfaceDescription)) { $s+=20 }
            if ($Identity.PnPInstanceId) {
                try {
                    $cim = Get-CimInstance -ClassName Win32_NetworkAdapter -Filter "InterfaceIndex=$($c.ifIndex)" -ErrorAction SilentlyContinue
                    if ($cim -and $cim.PNPDeviceID -and ($cim.PNPDeviceID -ieq $Identity.PnPInstanceId)) { $s+=10 }
                } catch {}
            }
            [PSCustomObject]@{Score=$s; Adapter=$c}
        }
        $best = $scored | Sort-Object Score -Descending | Select-Object -First 1
        if ($best.Score -gt 0) { return $best.Adapter } else { return $null }
    } catch { return $null }
}

$identities = @(
__IDENTITIES__
)

$action = '__ACTION__'
$errors = @()
$ok = 0
$names = @()

foreach ($id in $identities) {
    $a = Resolve-Adapter $id
    if (-not $a) { 
        $errors += ("Not found: Alias=""{0}"" GUID=""{1}""" -f $id.Alias, $id.InterfaceGuid)
        continue
    }
    try {
        if ($action -eq 'Enable') { Enable-NetAdapter -Name $a.Name -Confirm:$false -ErrorAction Stop }
        else { Disable-NetAdapter -Name $a.Name -Confirm:$false -ErrorAction Stop }
        $ok++; $names += $a.Name
    } catch {
        $errors += "Failed $action for '$($a.Name)': $($_.Exception.Message)"
    }
}
if ($errors.Count -gt 0) {
    [System.Windows.MessageBox]::Show(($errors -join "`n"), "$action completed with errors", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
} else {
    [System.Windows.MessageBox]::Show(("Successfully completed: {0} on {1} adapter(s).`n{2}" -f $action, $names.Count, (( $names | ForEach-Object { " - $_" }) -join "`r`n")), "$action complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
}
'@

    $template.Replace('__ACTION__', $Action).Replace('__IDENTITIES__', $identitiesBlock)
}
#endregion

#region --- WPF UI ---
# IMPORTANT: correct here-string delimiters for interpolated XAML (needs $AppTitle)
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$AppTitle" Width="1700" Height="800" WindowStartupLocation="CenterScreen"
        FontSize="14" Background="#1E1E1E" Foreground="#FFFFFF">
    <DockPanel x:Name="RootDock" LastChildFill="True">

        <!-- Top Controls -->
        <StackPanel DockPanel.Dock="Top" Orientation="Vertical" Margin="10,10,10,6">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="Auto" />
                </Grid.ColumnDefinitions>

                <!-- Left: Primary Actions -->
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Left" >
                    <Button x:Name="BtnEnable" Content="Enable Selected" Padding="16,10" Margin="0,0,8,0" />
                    <Button x:Name="BtnDisable" Content="Disable Selected" Padding="16,10" Margin="0,0,8,0" />
                    <Button x:Name="BtnRefresh" Content="Refresh" Padding="16,10" />
                </StackPanel>

                <!-- Right: Helpers + Theme + Profile + Exports -->
                <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right" >
                    <TextBlock VerticalAlignment="Center" Margin="0,0,6,0" Text="Theme:"/>
                    <ComboBox x:Name="CmbTheme" Width="180" SelectedIndex="0">
                        <ComboBoxItem Content="Dark (High Contrast)" />
                        <ComboBoxItem Content="Light" />
                    </ComboBox>
                    <Separator Width="10"/>
                    <Button x:Name="BtnCheckAll" Content="Check All (Visible)" Padding="12,8" Margin="0,0,8,0" />
                    <Button x:Name="BtnUncheckAll" Content="Uncheck All (Visible)" Padding="12,8" Margin="0,0,8,0" />
                    <Button x:Name="BtnInvert" Content="Invert (Visible)" Padding="12,8" Margin="0,0,8,0" />
                    <Separator Width="12"/>
                    <Button x:Name="BtnSaveProfile" Content="Save Profile" Padding="12,8" Margin="0,0,8,0" />
                    <Button x:Name="BtnLoadProfile" Content="Load Profile" Padding="12,8" Margin="0,0,8,0" />
                    <Separator Width="12"/>
                    <Button x:Name="BtnExportEnable" Content="Export Enable Script" Padding="12,8" Margin="0,0,8,0" />
                    <Button x:Name="BtnExportDisable" Content="Export Disable Script" Padding="12,8" />
                </StackPanel>
            </Grid>

            <!-- Filter + No-UAC -->
            <Grid Margin="0,10,0,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="Auto" />
                </Grid.ColumnDefinitions>
                <DockPanel Grid.Column="0">
                    <TextBlock Text="Filter:" VerticalAlignment="Center" Margin="0,0,8,0" />
                    <TextBox x:Name="TxtFilter" Width="600" />
                </DockPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right">
                    <CheckBox x:Name="ChkCreateShortcut" Content="Create No-UAC Desktop Shortcut after export" VerticalAlignment="Center" />
                </StackPanel>
            </Grid>
        </StackPanel>

        <!-- Data Grid -->
        <DataGrid x:Name="GridAdapters" Margin="10,0,10,6" AutoGenerateColumns="False" CanUserAddRows="False"
                  HeadersVisibility="Column" RowHeaderWidth="0" GridLinesVisibility="Horizontal"
                  Background="#11151C" Foreground="#FFFFFF" AlternatingRowBackground="#171C24"
                  RowBackground="#11151C" SelectionMode="Extended" SelectionUnit="FullRow" IsReadOnly="False"
                  RowHeight="28" >
            <DataGrid.ColumnHeaderStyle>
                <Style TargetType="DataGridColumnHeader">
                    <Setter Property="Background" Value="#0F131A"/>
                    <Setter Property="Foreground" Value="#FFFFFF"/>
                </Style>
            </DataGrid.ColumnHeaderStyle>
            <DataGrid.Columns>
                <DataGridTemplateColumn Header="✔" Width="48">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <CheckBox IsChecked="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" HorizontalAlignment="Center"/>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTextColumn Header="Adapter Name (Alias)" Binding="{Binding Alias}" Width="260" IsReadOnly="True"/>
                <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="120" IsReadOnly="True"/>
                <DataGridTextColumn Header="Interface Type" Binding="{Binding InterfaceType}" Width="160" IsReadOnly="True"/>
                <DataGridTextColumn Header="Interface Description" Binding="{Binding InterfaceDescription}" Width="*" IsReadOnly="True"/>
            </DataGrid.Columns>
        </DataGrid>

        <!-- Status Bar -->
        <StatusBar x:Name="MainStatusBar" DockPanel.Dock="Bottom" Background="#0B0D12" Foreground="#FFFFFF">
            <StatusBarItem>
                <TextBlock x:Name="TxtTotals" Text="Adapters: 0 | Selected: 0 | Visible Selected: 0" />
            </StatusBarItem>
            <Separator />
            <StatusBarItem>
                <TextBlock x:Name="TxtStatus" Text="Ready." />
            </StatusBarItem>
        </StatusBar>

    </DockPanel>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)
#endregion

# Find controls
$BtnEnable        = $window.FindName('BtnEnable')
$BtnDisable       = $window.FindName('BtnDisable')
$BtnRefresh       = $window.FindName('BtnRefresh')
$BtnCheckAll      = $window.FindName('BtnCheckAll')
$BtnUncheckAll    = $window.FindName('BtnUncheckAll')
$BtnInvert        = $window.FindName('BtnInvert')
$TxtFilter        = $window.FindName('TxtFilter')
$CmbTheme         = $window.FindName('CmbTheme')
$ChkCreateShortcut= $window.FindName('ChkCreateShortcut')
$GridAdapters     = $window.FindName('GridAdapters')
$BtnSaveProfile   = $window.FindName('BtnSaveProfile')
$BtnLoadProfile   = $window.FindName('BtnLoadProfile')
$BtnExportEnable  = $window.FindName('BtnExportEnable')
$BtnExportDisable = $window.FindName('BtnExportDisable')
$TotalsTextBlock  = $window.FindName('TxtTotals')
$StatusTextBlock  = $window.FindName('TxtStatus')
$MainStatusBar    = $window.FindName('MainStatusBar')

#region --- Data, filtering ---
$Script:AdapterCollection = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
$GridAdapters.ItemsSource = $Script:AdapterCollection
$view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($GridAdapters.ItemsSource)
$view.Filter = {
    if ([string]::IsNullOrWhiteSpace($TxtFilter.Text)) { return $true }
    $t = $TxtFilter.Text.Trim()
    $_.Alias -like "*$t*" -or $_.InterfaceDescription -like "*$t*"
}
function Get-VisibleItems { $l=@(); foreach($i in $view){$l+=$i}; $l }
function Update-Totals {
    $TotalsTextBlock.Text = "Adapters: {0} | Selected: {1} | Visible Selected: {2}" -f `
        $AdapterCollection.Count, `
        ($AdapterCollection | Where-Object Selected).Count, `
        (Get-VisibleItems | Where-Object Selected).Count
}
function Load-Adapters {
    $AdapterCollection.Clear()
    (Get-AdapterInventory) | ForEach-Object { [void]$AdapterCollection.Add($_) }
    $view.Refresh(); Update-Totals
    Update-Status "Loaded $($AdapterCollection.Count) adapter(s)."
}
Load-Adapters
#endregion

#region --- Theme (incl. checkbox foreground) ---
$ThemePalette = @{
    'Dark (High Contrast)' = @{
        WindowBg = New-Brush '#1E1E1E'; Fore = New-Brush '#FFFFFF'
        GridBg   = New-Brush '#11151C'; AltBg  = New-Brush '#171C24'
        HeaderBg = New-Brush '#0F131A'; StatusBg = New-Brush '#0B0D12'
        TextBoxBg= New-Brush '#11151C'; TextBoxFg= New-Brush '#FFFFFF'
        CheckFg  = New-Brush '#FFFFFF'
    }
    'Light' = @{
        WindowBg = New-Brush '#FFFFFF'; Fore = New-Brush '#111111'
        GridBg   = New-Brush '#FFFFFF'; AltBg = New-Brush '#F5F7FA'
        HeaderBg = New-Brush '#E9ECF1'; StatusBg = New-Brush '#F0F2F5'
        TextBoxBg= New-Brush '#FFFFFF'; TextBoxFg= New-Brush '#111111'
        CheckFg  = New-Brush '#111111'
    }
}
function Apply-Theme([string]$name) {
    $p = $ThemePalette[$name]; if (-not $p) { return }
    $window.Background = $p.WindowBg; $window.Foreground = $p.Fore
    if ($MainStatusBar) { $MainStatusBar.Background = $p.StatusBg; $MainStatusBar.Foreground = $p.Fore }
    if ($TxtFilter) { $TxtFilter.Background = $p.TextBoxBg; $TxtFilter.Foreground = $p.TextBoxFg }
    if ($ChkCreateShortcut) { $ChkCreateShortcut.Foreground = $p.CheckFg }
    if ($GridAdapters) {
        $GridAdapters.Background=$p.GridBg; $GridAdapters.Foreground=$p.Fore
        $GridAdapters.RowBackground=$p.GridBg; $GridAdapters.AlternatingRowBackground=$p.AltBg
        $style = [System.Windows.Style]::new([System.Windows.Controls.Primitives.DataGridColumnHeader])
        $style.Setters.Add([System.Windows.Setter]::new([System.Windows.Controls.Control]::BackgroundProperty,$p.HeaderBg))
        $style.Setters.Add([System.Windows.Setter]::new([System.Windows.Controls.Control]::ForegroundProperty,$p.Fore))
        $GridAdapters.ColumnHeaderStyle = $style
    }
}
if ($CmbTheme) {
    $CmbTheme.Add_SelectionChanged({
        $name = ($CmbTheme.SelectedItem.Content).ToString()
        Apply-Theme $name
    })
}
Apply-Theme 'Dark (High Contrast)'
#endregion

#region --- Selection helpers ---
$BtnCheckAll.Add_Click({ Get-VisibleItems | ForEach-Object { $_.Selected = $true }; $view.Refresh(); Update-Totals })
$BtnUncheckAll.Add_Click({ Get-VisibleItems | ForEach-Object { $_.Selected = $false }; $view.Refresh(); Update-Totals })
$BtnInvert.Add_Click({ Get-VisibleItems | ForEach-Object { $_.Selected = -not $_.Selected }; $view.Refresh(); Update-Totals })
$GridAdapters.Add_CurrentCellChanged({ Update-Totals })
$TxtFilter.Add_TextChanged({ $view.Refresh(); Update-Totals })
$BtnRefresh.Add_Click({ Load-Adapters })
#endregion

#region --- Actions (preserve selections) ---
function Get-SelectedIdentities {
    $AdapterCollection | Where-Object Selected | ForEach-Object {
        @{ InterfaceGuid=$_.InterfaceGuid; Alias=$_.Alias; MacAddress=$_.MacAddress; InterfaceDescription=$_.InterfaceDescription; PnPInstanceId=$_.PnPInstanceId }
    }
}
function Reapply-Selection($ids) {
    if (-not $ids) { return }
    foreach ($id in $ids) {
        $item = $AdapterCollection | Where-Object { $_.InterfaceGuid -and ($_.InterfaceGuid -ieq $id.InterfaceGuid) } | Select-Object -First 1
        if (-not $item) { $item = $AdapterCollection | Where-Object { $_.Alias -and ($_.Alias -ieq $id.Alias) } | Select-Object -First 1 }
        if (-not $item -and $id.MacAddress) { $clean = ($id.MacAddress -replace '-|:',''); $item = $AdapterCollection | Where-Object { ($_.MacAddress -replace '-|:','') -ieq $clean } | Select-Object -First 1 }
        if (-not $item -and $id.InterfaceDescription) { $item = $AdapterCollection | Where-Object { $_.InterfaceDescription -ieq $id.InterfaceDescription } | Select-Object -First 1 }
        if (-not $item -and $id.PnPInstanceId) { $item = $AdapterCollection | Where-Object { $_.PnPInstanceId -and ($_.PnPInstanceId -ieq $id.PnPInstanceId) } | Select-Object -First 1 }
        if ($item) { $item.Selected = $true }
    }
    $view.Refresh(); Update-Totals
}
function Do-ActionOnSelected([ValidateSet('Enable','Disable')]$Action) {
    $ids = Get-SelectedIdentities
    if (-not $ids -or $ids.Count -eq 0) { Update-Status "No adapters selected." $true; return }
    $names = foreach($id in $ids){ if($id.Alias){$id.Alias}elseif($id.InterfaceDescription){$id.InterfaceDescription}else{$id.InterfaceGuid} }
    $list = ( $names | ForEach-Object { " - $_" } ) -join "`r`n"
    $msg = "{0} the following {1} adapter(s)?`r`n{2}" -f $Action, $names.Count, $list
    $icon = if ($Action -eq 'Disable') { [System.Windows.MessageBoxImage]::Warning } else { [System.Windows.MessageBoxImage]::Question }
    $res  = [System.Windows.MessageBox]::Show($msg, "$AppTitle - Confirm", [System.Windows.MessageBoxButton]::YesNo, $icon)
    if ($res -ne [System.Windows.MessageBoxResult]::Yes) { return }

    $ok=0;$err=0;$messages=@()
    foreach ($id in $ids) {
        $adapter = Resolve-Adapter $id
        if (-not $adapter) { $err++; $messages += "Not found: Alias=""$($id.Alias)"", GUID=""$($id.InterfaceGuid)"""; continue }
        try {
            if ($Action -eq 'Enable') { Enable-NetAdapter -Name $adapter.Name -Confirm:$false -ErrorAction Stop }
            else { Disable-NetAdapter -Name $adapter.Name -Confirm:$false -ErrorAction Stop }
            $ok++
        } catch { $err++; $messages += "Failed $Action '$($adapter.Name)': $($_.Exception.Message)" }
    }
    Load-Adapters
    Reapply-Selection $ids   # keep them checked after action
    if ($err -gt 0) { Update-Status "$Action completed: $ok ok, $err error(s)." $true; Show-ErrorDialog(($messages -join "`r`n"), "$Action completed with errors") }
    else { Update-Status "Successfully $($Action.ToLower())d $ok adapter(s)." }
}
$BtnEnable.Add_Click({ Do-ActionOnSelected 'Enable' })
$BtnDisable.Add_Click({ Do-ActionOnSelected 'Disable' })
#endregion

#region --- Profile Save/Load (+ auto-load on startup) ---
$BtnSaveProfile.Add_Click({
    try {
        $ids = Get-SelectedIdentities
        if (-not $ids -or $ids.Count -eq 0) { Update-Status "No adapters selected to save." $true; return }
        ($ids | ConvertTo-Json -Depth 4) | Out-File -LiteralPath $ProfilePath -Encoding UTF8
        Update-Status "Profile saved to $ProfilePath"
    } catch { Update-Status "Error: Could not save profile. $($_.Exception.Message)" $true; Show-ErrorDialog("Could not save profile:`n$($_.Exception.Message)") }
})
$BtnLoadProfile.Add_Click({
    try {
        if (-not (Test-Path -LiteralPath $ProfilePath)) { Update-Status "Profile not found at $ProfilePath" $true; return }
        $data = Get-Content -LiteralPath $ProfilePath -Raw | ConvertFrom-Json -ErrorAction Stop
        Reapply-Selection $data
        Update-Status "Profile loaded. Checked $($data.Count) adapter(s)."
    } catch { Update-Status "Error: Could not load profile. $($_.Exception.Message)" $true; Show-ErrorDialog("Could not load profile:`n$($_.Exception.Message)") }
})

# Auto-load profile at startup if present
if (Test-Path -LiteralPath $ProfilePath) {
    try {
        $data = Get-Content -LiteralPath $ProfilePath -Raw | ConvertFrom-Json -ErrorAction Stop
        Reapply-Selection $data
        Update-Status "Profile auto-loaded."
    } catch {}
}
#endregion

#region --- Export scripts + No-UAC Scheduled Task + Desktop Shortcut ---
function New-UniqueTaskName([string]$prefix) { "$prefix-$(Get-Date -Format 'yyyyMMddHHmmss')-$([guid]::NewGuid().ToString('N').Substring(0,6))" }
function Create-ScheduledTaskAndShortcut {
param([string]$ScriptPath,[string]$TaskName,[string]$ShortcutName)
    $errors=@()
    try {
        $user = if ($env:USERDOMAIN) { "$env:USERDOMAIN\$env:USERNAME" } else { $env:USERNAME }
        $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument ("-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`"")
        # FIX: LogonType Interactive (not InteractiveToken)
        $principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Highest
        $task = New-ScheduledTask -Action $action -Principal $principal
        try {
            if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
                Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
            }
        } catch {}
        Register-ScheduledTask -TaskName $TaskName -InputObject $task -Force | Out-Null
    } catch { $errors += "Error creating Scheduled Task: $($_.Exception.Message)" }

    $desktop = [Environment]::GetFolderPath('Desktop'); $lnkPath = Join-Path $desktop ($ShortcutName + '.lnk')
    try {
        $wsh = New-Object -ComObject WScript.Shell
        $lnk = $wsh.CreateShortcut($lnkPath)
        $lnk.TargetPath = "$env:SystemRoot\System32\schtasks.exe"
        $lnk.Arguments  = "/Run /TN `"$TaskName`""
        $lnk.IconLocation = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe,0"
        $lnk.WorkingDirectory = Split-Path -Parent $ScriptPath
        $lnk.Save()
    } catch { $errors += "Error creating desktop shortcut: $($_.Exception.Message)" }

    if ($errors.Count -gt 0) {
        $msg = $errors -join "`r`n"
        Update-Status $msg $true
        Show-ErrorDialog($msg, "$AppTitle - Scheduled Task/Shortcut Error")
    } else {
        Update-Status "Created Scheduled Task '$TaskName' and desktop shortcut '$ShortcutName'."
    }
}
function Export-ActionScript {
param([ValidateSet('Enable','Disable')][string]$Action)
    $ids = Get-SelectedIdentities
    if (-not $ids -or $ids.Count -eq 0) { Update-Status "No adapters selected to export." $true; return }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.FileName = "Adapters-$Action-$(Get-Date -Format 'yyyyMMdd-HHmmss').ps1"
    $dlg.InitialDirectory = $ScriptDir; $dlg.Filter = "PowerShell Script (*.ps1)|*.ps1"
    if (-not $dlg.ShowDialog()) { return }
    $path = $dlg.FileName
    try {
        (Get-ActionScriptContent -Action $Action -Identities $ids) | Out-File -LiteralPath $path -Encoding UTF8
        Update-Status "Exported $Action script to $path"
    } catch {
        Update-Status "Failed to export script: $($_.Exception.Message)" $true
        Show-ErrorDialog("Failed to export script:`n$($_.Exception.Message)")
        return
    }
    if ($ChkCreateShortcut.IsChecked) {
        try {
            Create-ScheduledTaskAndShortcut -ScriptPath $path -TaskName (New-UniqueTaskName "NetManager-$Action") -ShortcutName "$Action Selected Adapters (No UAC)"
        } catch {
            Update-Status "Error: Could not create Scheduled Task/Shortcut. $($_.Exception.Message)" $true
            Show-ErrorDialog("Error while creating Scheduled Task/Shortcut:`n$($_.Exception.Message)")
        }
    }
}
$BtnExportEnable.Add_Click({ Export-ActionScript -Action 'Enable' })
$BtnExportDisable.Add_Click({ Export-ActionScript -Action 'Disable' })
#endregion

# Default: check the No-UAC shortcut option ON
if ($ChkCreateShortcut) { $ChkCreateShortcut.IsChecked = $true }

# Show window
$window.Add_ContentRendered({ Update-Totals; Update-Status "Ready." })
[void]$window.ShowDialog()
