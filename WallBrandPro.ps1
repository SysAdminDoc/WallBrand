<#
.SYNOPSIS
    WallBrand Pro - Enterprise Wallpaper Branding & System Info Tool
    
.DESCRIPTION
    Professional-grade wallpaper branding tool with:
    - Modern WPF GUI with dark theme
    - Multi-monitor support with per-monitor wallpapers
    - Dynamic system information overlay (BGInfo replacement)
    - Logo and text branding with advanced effects
    - Template system with configurable regions
    - Enterprise deployment (GPO, Intune, SCCM)
    - Lock screen and desktop wallpaper support
    - Batch processing and silent deployment
    
.PARAMETER ConfigPath
    Path to JSON configuration file
    
.PARAMETER Silent
    Run without GUI, apply configuration directly
    
.PARAMETER Apply
    Apply wallpaper immediately after generation
    
.PARAMETER ExportGPO
    Export Group Policy deployment script
    
.PARAMETER ExportIntune
    Export Intune deployment package
    
.NOTES
    Author: WallBrand Pro
    Version: 2.0.0
    Requires: PowerShell 5.1+, Windows 10/11
#>

[CmdletBinding()]
param(
    [string]$ConfigPath,
    [switch]$Silent,
    [switch]$Apply,
    [string]$ExportGPO,
    [string]$ExportIntune,
    [string]$OutputPath
)

# ============================================================================
# INITIALIZATION
# ============================================================================

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

# Windows API for wallpaper and monitor detection
Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Drawing;

public class WallpaperAPI {
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
    
    public const int SPI_SETDESKWALLPAPER = 0x0014;
    public const int SPIF_UPDATEINIFILE = 0x01;
    public const int SPIF_SENDCHANGE = 0x02;
    
    [DllImport("user32.dll")]
    public static extern bool EnumDisplayMonitors(IntPtr hdc, IntPtr lprcClip, MonitorEnumProc lpfnEnum, IntPtr dwData);
    
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFOEX lpmi);
    
    public delegate bool MonitorEnumProc(IntPtr hMonitor, IntPtr hdcMonitor, ref RECT lprcMonitor, IntPtr dwData);
    
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left, Top, Right, Bottom;
    }
    
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    public struct MONITORINFOEX {
        public int Size;
        public RECT Monitor;
        public RECT WorkArea;
        public uint Flags;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string DeviceName;
    }
}
"@

# ============================================================================
# CONFIGURATION CLASS
# ============================================================================

class WallBrandConfig {
    # Source files
    [string]$WallpaperPath = ""
    [string]$LogoPath = ""
    
    # Branding text
    [string]$PrimaryText = ""
    [string]$SecondaryText = ""
    
    # Position & Layout
    [string]$Position = "BottomRight"
    [int]$Margin = 40
    [int]$CustomX = -1
    [int]$CustomY = -1
    [int]$CustomTextX = -1
    [int]$CustomTextY = -1
    
    # Logo settings
    [double]$LogoScale = 1.0
    [int]$LogoOpacity = 255
    [int]$CustomLogoW = -1
    [int]$CustomLogoH = -1
    
    # Text styling
    [string]$FontName = "Segoe UI"
    [int]$FontSize = 14
    [string]$FontColor = "FFFFFF"
    [bool]$FontBold = $false
    [bool]$EnableShadow = $true
    [int]$ShadowOpacity = 180
    [int]$ShadowOffset = 2
    
    # Backdrop/Glow
    [bool]$EnableBackdrop = $true
    [string]$BackdropMode = "Glow"  # Box, Glow, Outline
    [string]$BackdropColor = "000000"
    [int]$BackdropOpacity = 180
    [int]$BackdropRadius = 12
    [int]$BackdropPadding = 20
    
    # System Info (BGInfo-style)
    [bool]$EnableSystemInfo = $false
    [string]$SystemInfoPosition = "TopLeft"
    [int]$SystemInfoCustomX = -1
    [int]$SystemInfoCustomY = -1
    [string[]]$SystemInfoFields = @("HostName", "IPAddress", "UserName", "OSVersion")
    [string]$SystemInfoFontName = "Consolas"
    [int]$SystemInfoFontSize = 11
    [string]$SystemInfoFontColor = "FFFFFF"
    [string]$SystemInfoLabelColor = "88CCFF"
    [bool]$SystemInfoShadow = $true
    
    # Multi-monitor
    [string]$MultiMonitorMode = "Clone"  # Clone, Span, PerMonitor
    [hashtable]$PerMonitorConfig = @{}
    
    # Output
    [bool]$MatchResolution = $true
    [string]$OutputFormat = "PNG"
    [int]$JpegQuality = 95
    
    # Lock screen
    [bool]$ApplyToLockScreen = $false
    
    # Converts to JSON
    [string] ToJson() {
        return $this | ConvertTo-Json -Depth 10
    }
    
    # Load from JSON
    static [WallBrandConfig] FromJson([string]$json) {
        $obj = $json | ConvertFrom-Json
        $config = [WallBrandConfig]::new()
        foreach ($prop in $obj.PSObject.Properties) {
            if ($null -ne $config.PSObject.Properties[$prop.Name]) {
                $config.($prop.Name) = $prop.Value
            }
        }
        return $config
    }
}

# ============================================================================
# SYSTEM INFORMATION GATHERING
# ============================================================================

function Get-SystemInfoData {
    [CmdletBinding()]
    param()
    
    $info = [ordered]@{}
    
    # Computer info
    $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $bios = Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue
    
    $info["HostName"] = $env:COMPUTERNAME
    $info["UserName"] = "$env:USERDOMAIN\$env:USERNAME"
    $info["MachineDomain"] = if ($cs.PartOfDomain) { $cs.Domain } else { "WORKGROUP" }
    
    # OS Info
    $info["OSVersion"] = $os.Caption
    $info["OSBuild"] = $os.BuildNumber
    $info["OSArchitecture"] = $os.OSArchitecture
    $info["InstallDate"] = $os.InstallDate.ToString("yyyy-MM-dd")
    $info["LastBoot"] = $os.LastBootUpTime.ToString("yyyy-MM-dd HH:mm")
    
    # Hardware
    $info["Manufacturer"] = $cs.Manufacturer
    $info["Model"] = $cs.Model
    $info["SerialNumber"] = $bios.SerialNumber
    $info["BIOSVersion"] = $bios.SMBIOSBIOSVersion
    
    # CPU
    $cpu = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
    $info["CPU"] = $cpu.Name -replace '\s+', ' '
    $info["CPUCores"] = "$($cpu.NumberOfCores) Cores / $($cpu.NumberOfLogicalProcessors) Threads"
    
    # Memory
    $totalRAM = [math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
    $freeRAM = [math]::Round($os.FreePhysicalMemory / 1MB, 1)
    $info["TotalRAM"] = "$totalRAM GB"
    $info["FreeRAM"] = "$freeRAM GB"
    
    # Disk
    $disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction SilentlyContinue
    if ($disk) {
        $totalDisk = [math]::Round($disk.Size / 1GB, 1)
        $freeDisk = [math]::Round($disk.FreeSpace / 1GB, 1)
        $info["DiskTotal"] = "$totalDisk GB"
        $info["DiskFree"] = "$freeDisk GB"
    }
    
    # Network
    $adapters = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True" -ErrorAction SilentlyContinue
    $ips = @()
    $macs = @()
    foreach ($adapter in $adapters) {
        if ($adapter.IPAddress) {
            $ipv4 = $adapter.IPAddress | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' } | Select-Object -First 1
            if ($ipv4) { $ips += $ipv4 }
        }
        if ($adapter.MACAddress) { $macs += $adapter.MACAddress }
    }
    $info["IPAddress"] = ($ips -join ", ")
    $info["MACAddress"] = ($macs | Select-Object -First 1)
    
    # Default Gateway
    $gateway = ($adapters | Where-Object { $_.DefaultIPGateway } | Select-Object -First 1).DefaultIPGateway
    $info["Gateway"] = if ($gateway) { $gateway[0] } else { "N/A" }
    
    # DNS
    $dns = ($adapters | Where-Object { $_.DNSServerSearchOrder } | Select-Object -First 1).DNSServerSearchOrder
    $info["DNSServers"] = if ($dns) { $dns -join ", " } else { "N/A" }
    
    # Current time
    $info["CurrentTime"] = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $info["TimeZone"] = [System.TimeZoneInfo]::Local.DisplayName
    
    # Uptime
    $uptime = (Get-Date) - $os.LastBootUpTime
    $info["Uptime"] = "{0}d {1}h {2}m" -f $uptime.Days, $uptime.Hours, $uptime.Minutes
    
    # Windows Update (if available)
    try {
        $hotfix = Get-HotFix -ErrorAction SilentlyContinue | Sort-Object InstalledOn -Descending | Select-Object -First 1
        if ($hotfix) {
            $info["LastUpdate"] = $hotfix.InstalledOn.ToString("yyyy-MM-dd")
        }
    } catch { }
    
    return $info
}

# ============================================================================
# MONITOR DETECTION
# ============================================================================

function Get-MonitorInfo {
    [CmdletBinding()]
    param()
    
    $monitors = @()
    
    # Use WMI for basic info
    $wmiMonitors = Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue
    $screenInfo = [System.Windows.Forms.Screen]::AllScreens
    
    $index = 0
    foreach ($screen in $screenInfo) {
        $monitor = [PSCustomObject]@{
            Index = $index
            Name = $screen.DeviceName
            Primary = $screen.Primary
            Width = $screen.Bounds.Width
            Height = $screen.Bounds.Height
            X = $screen.Bounds.X
            Y = $screen.Bounds.Y
            WorkingArea = $screen.WorkingArea
            BitsPerPixel = $screen.BitsPerPixel
            Resolution = "$($screen.Bounds.Width)x$($screen.Bounds.Height)"
        }
        $monitors += $monitor
        $index++
    }
    
    return $monitors
}

function Get-VirtualScreenBounds {
    # Get the combined bounds of all monitors
    $screens = [System.Windows.Forms.Screen]::AllScreens
    
    $minX = ($screens | Measure-Object -Property { $_.Bounds.X } -Minimum).Minimum
    $minY = ($screens | Measure-Object -Property { $_.Bounds.Y } -Minimum).Minimum
    $maxX = ($screens | ForEach-Object { $_.Bounds.X + $_.Bounds.Width } | Measure-Object -Maximum).Maximum
    $maxY = ($screens | ForEach-Object { $_.Bounds.Y + $_.Bounds.Height } | Measure-Object -Maximum).Maximum
    
    return [PSCustomObject]@{
        X = $minX
        Y = $minY
        Width = $maxX - $minX
        Height = $maxY - $minY
    }
}

# ============================================================================
# IMAGE RENDERING ENGINE
# ============================================================================

function Render-BrandedWallpaper {
    [CmdletBinding()]
    param(
        [WallBrandConfig]$Config,
        [int]$TargetWidth = 0,
        [int]$TargetHeight = 0,
        [int]$MonitorIndex = -1
    )
    
    # Determine output size
    if ($TargetWidth -eq 0 -or $TargetHeight -eq 0) {
        if ($Config.MatchResolution) {
            $primary = [System.Windows.Forms.Screen]::PrimaryScreen
            $TargetWidth = $primary.Bounds.Width
            $TargetHeight = $primary.Bounds.Height
        } else {
            $TargetWidth = 1920
            $TargetHeight = 1080
        }
    }
    
    # Create output bitmap
    $output = New-Object System.Drawing.Bitmap($TargetWidth, $TargetHeight)
    $graphics = [System.Drawing.Graphics]::FromImage($output)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    
    # Draw wallpaper background
    if ($Config.WallpaperPath -and (Test-Path $Config.WallpaperPath)) {
        $wallpaper = [System.Drawing.Image]::FromFile($Config.WallpaperPath)
        
        # Scale to fill (cover mode)
        $scaleX = $TargetWidth / $wallpaper.Width
        $scaleY = $TargetHeight / $wallpaper.Height
        $scale = [Math]::Max($scaleX, $scaleY)
        
        $newWidth = [int]($wallpaper.Width * $scale)
        $newHeight = [int]($wallpaper.Height * $scale)
        $x = [int](($TargetWidth - $newWidth) / 2)
        $y = [int](($TargetHeight - $newHeight) / 2)
        
        $graphics.DrawImage($wallpaper, $x, $y, $newWidth, $newHeight)
        $wallpaper.Dispose()
    } else {
        # Default gradient background
        $rect = New-Object System.Drawing.Rectangle(0, 0, $TargetWidth, $TargetHeight)
        $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
            $rect,
            [System.Drawing.Color]::FromArgb(20, 30, 48),
            [System.Drawing.Color]::FromArgb(36, 59, 85),
            [System.Drawing.Drawing2D.LinearGradientMode]::Vertical
        )
        $graphics.FillRectangle($brush, $rect)
        $brush.Dispose()
    }
    
    # Draw system info if enabled
    if ($Config.EnableSystemInfo) {
        Draw-SystemInfo $graphics $Config $TargetWidth $TargetHeight
    }
    
    # Draw branding (logo + text)
    if ($Config.LogoPath -or $Config.PrimaryText) {
        Draw-Branding $graphics $Config $TargetWidth $TargetHeight
    }
    
    $graphics.Dispose()
    return $output
}

function Draw-SystemInfo {
    param(
        [System.Drawing.Graphics]$g,
        [WallBrandConfig]$Config,
        [int]$Width,
        [int]$Height
    )
    
    $sysInfo = Get-SystemInfoData
    
    # Reference resolution (what the preview uses)
    $refWidth = 1920
    $refHeight = 1080
    $scaleX = $Width / $refWidth
    $scaleY = $Height / $refHeight
    
    # Scale font size to target resolution
    $scaledFontSize = [Math]::Max(8, [int]($Config.SystemInfoFontSize * $scaleY))
    
    # Fonts
    $labelFont = New-Object System.Drawing.Font($Config.SystemInfoFontName, $scaledFontSize, [System.Drawing.FontStyle]::Bold)
    $valueFont = New-Object System.Drawing.Font($Config.SystemInfoFontName, $scaledFontSize)
    
    # Colors
    $labelColor = [System.Drawing.ColorTranslator]::FromHtml("#$($Config.SystemInfoLabelColor)")
    $valueColor = [System.Drawing.ColorTranslator]::FromHtml("#$($Config.SystemInfoFontColor)")
    $shadowColor = [System.Drawing.Color]::FromArgb(180, 0, 0, 0)
    
    $labelBrush = New-Object System.Drawing.SolidBrush($labelColor)
    $valueBrush = New-Object System.Drawing.SolidBrush($valueColor)
    $shadowBrush = New-Object System.Drawing.SolidBrush($shadowColor)
    
    # Calculate position with scaling
    $lineHeight = [int](($Config.SystemInfoFontSize + 6) * $scaleY)
    $scaledMargin = [int](30 * $scaleX)
    $labelWidth = [int](140 * $scaleX)
    
    # Use custom position if specified - scale to target resolution
    if ($Config.SystemInfoCustomX -ge 0 -and $Config.SystemInfoCustomY -ge 0) {
        $x = [int]($Config.SystemInfoCustomX * $scaleX)
        $y = [int]($Config.SystemInfoCustomY * $scaleY)
    } else {
        switch ($Config.SystemInfoPosition) {
            "TopLeft" { $x = $scaledMargin; $y = $scaledMargin }
            "TopRight" { $x = $Width - [int](400 * $scaleX); $y = $scaledMargin }
            "BottomLeft" { $x = $scaledMargin; $y = $Height - ($Config.SystemInfoFields.Count * $lineHeight) - $scaledMargin }
            "BottomRight" { $x = $Width - [int](400 * $scaleX); $y = $Height - ($Config.SystemInfoFields.Count * $lineHeight) - $scaledMargin }
            default { $x = $scaledMargin; $y = $scaledMargin }
        }
    }
    
    # Field name mappings for display
    $fieldNames = @{
        "HostName" = "Host Name"
        "UserName" = "User Name"
        "MachineDomain" = "Domain"
        "OSVersion" = "OS"
        "OSBuild" = "Build"
        "OSArchitecture" = "Architecture"
        "IPAddress" = "IP Address"
        "MACAddress" = "MAC Address"
        "Gateway" = "Gateway"
        "DNSServers" = "DNS"
        "CPU" = "CPU"
        "CPUCores" = "CPU Cores"
        "TotalRAM" = "Total RAM"
        "FreeRAM" = "Free RAM"
        "DiskTotal" = "Disk Size"
        "DiskFree" = "Disk Free"
        "Manufacturer" = "Manufacturer"
        "Model" = "Model"
        "SerialNumber" = "Serial"
        "CurrentTime" = "Time"
        "LastBoot" = "Last Boot"
        "Uptime" = "Uptime"
    }
    
    foreach ($field in $Config.SystemInfoFields) {
        if ($sysInfo.Contains($field)) {
            $displayName = if ($fieldNames.Contains($field)) { $fieldNames[$field] } else { $field }
            $value = $sysInfo[$field]
            
            $label = "${displayName}:"
            $text = "$value"
            
            # Draw shadow
            if ($Config.SystemInfoShadow) {
                $shadowOffset = [Math]::Max(1, [int](2 * $scaleX))
                $g.DrawString($label, $labelFont, $shadowBrush, ($x + $shadowOffset), ($y + $shadowOffset))
                $g.DrawString($text, $valueFont, $shadowBrush, ($x + $labelWidth + $shadowOffset), ($y + $shadowOffset))
            }
            
            # Draw text
            $g.DrawString($label, $labelFont, $labelBrush, $x, $y)
            $g.DrawString($text, $valueFont, $valueBrush, ($x + $labelWidth), $y)
            
            $y += $lineHeight
        }
    }
    
    # Cleanup
    $labelFont.Dispose()
    $valueFont.Dispose()
    $labelBrush.Dispose()
    $valueBrush.Dispose()
    $shadowBrush.Dispose()
}

function Draw-Branding {
    param(
        [System.Drawing.Graphics]$g,
        [WallBrandConfig]$Config,
        [int]$Width,
        [int]$Height
    )
    
    $margin = $Config.Margin
    $padding = $Config.BackdropPadding
    
    # Reference resolution (what the preview uses)
    $refWidth = 1920
    $refHeight = 1080
    $scaleX = $Width / $refWidth
    $scaleY = $Height / $refHeight
    
    # Load logo if specified
    $logo = $null
    $logoWidth = 0
    $logoHeight = 0
    
    if ($Config.LogoPath -and (Test-Path $Config.LogoPath)) {
        $logo = [System.Drawing.Image]::FromFile($Config.LogoPath)
        
        # Use custom size if specified, otherwise use scale
        if ($Config.CustomLogoW -gt 0 -and $Config.CustomLogoH -gt 0) {
            # Scale custom size to target resolution
            $logoWidth = [int]($Config.CustomLogoW * $scaleX)
            $logoHeight = [int]($Config.CustomLogoH * $scaleY)
        } else {
            $logoWidth = [int]($logo.Width * $Config.LogoScale)
            $logoHeight = [int]($logo.Height * $Config.LogoScale)
        }
    }
    
    # Measure text
    # Scale font size to target resolution
    $scaledFontSize = [Math]::Max(8, [int]($Config.FontSize * $scaleY))
    $primaryFont = New-Object System.Drawing.Font(
        $Config.FontName,
        $scaledFontSize,
        $(if ($Config.FontBold) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular })
    )
    $secondaryFont = New-Object System.Drawing.Font($Config.FontName, [int]($scaledFontSize * 0.85))
    
    $primarySize = $g.MeasureString($Config.PrimaryText, $primaryFont)
    $secondarySize = if ($Config.SecondaryText) { $g.MeasureString($Config.SecondaryText, $secondaryFont) } else { [System.Drawing.SizeF]::Empty }
    
    # Scale margin and padding
    $scaledMargin = [int]($margin * $scaleX)
    $scaledPadding = [int]($padding * $scaleX)
    
    # Calculate content block size for logo
    $contentWidth = [Math]::Max($logoWidth, [Math]::Max($primarySize.Width, $secondarySize.Width))
    $contentHeight = $logoHeight
    if ($Config.PrimaryText) { $contentHeight += $primarySize.Height + [int](8 * $scaleY) }
    if ($Config.SecondaryText) { $contentHeight += $secondarySize.Height + [int](4 * $scaleY) }
    
    # Calculate logo position - scale custom positions to target resolution
    if ($Config.CustomX -ge 0 -and $Config.CustomY -ge 0) {
        $blockX = [int]($Config.CustomX * $scaleX)
        $blockY = [int]($Config.CustomY * $scaleY)
    } else {
        switch ($Config.Position) {
            "TopLeft" { $blockX = $scaledMargin; $blockY = $scaledMargin }
            "TopCenter" { $blockX = ($Width - $contentWidth) / 2; $blockY = $scaledMargin }
            "TopRight" { $blockX = $Width - $contentWidth - $scaledMargin; $blockY = $scaledMargin }
            "BottomLeft" { $blockX = $scaledMargin; $blockY = $Height - $contentHeight - $scaledMargin }
            "BottomCenter" { $blockX = ($Width - $contentWidth) / 2; $blockY = $Height - $contentHeight - $scaledMargin }
            "BottomRight" { $blockX = $Width - $contentWidth - $scaledMargin; $blockY = $Height - $contentHeight - $scaledMargin }
            "CenterLeft" { $blockX = $scaledMargin; $blockY = ($Height - $contentHeight) / 2 }
            "Center" { $blockX = ($Width - $contentWidth) / 2; $blockY = ($Height - $contentHeight) / 2 }
            "CenterRight" { $blockX = $Width - $contentWidth - $scaledMargin; $blockY = ($Height - $contentHeight) / 2 }
            default { $blockX = $Width - $contentWidth - $scaledMargin; $blockY = $Height - $contentHeight - $scaledMargin }
        }
    }
    
    # Calculate text position (separate from logo if custom position set)
    if ($Config.CustomTextX -ge 0 -and $Config.CustomTextY -ge 0) {
        $textBlockX = [int]($Config.CustomTextX * $scaleX)
        $textBlockY = [int]($Config.CustomTextY * $scaleY)
        $useCustomTextPos = $true
    } else {
        $textBlockX = $blockX
        $textBlockY = $blockY + $logoHeight + [int](8 * $scaleY)
        $useCustomTextPos = $false
    }
    
    # Draw backdrop
    if ($Config.EnableBackdrop) {
        $backdropColor = [System.Drawing.ColorTranslator]::FromHtml("#$($Config.BackdropColor)")
        $scaledRadius = [int]($Config.BackdropRadius * $scaleX)
        
        switch ($Config.BackdropMode) {
            "Box" {
                $boxColor = [System.Drawing.Color]::FromArgb($Config.BackdropOpacity, $backdropColor)
                $boxBrush = New-Object System.Drawing.SolidBrush($boxColor)
                
                $boxRect = New-Object System.Drawing.Rectangle(
                    ([int]$blockX - $scaledPadding),
                    ([int]$blockY - $scaledPadding),
                    ([int]$contentWidth + $scaledPadding * 2),
                    ([int]$contentHeight + $scaledPadding * 2)
                )
                
                # Rounded rectangle
                $path = New-Object System.Drawing.Drawing2D.GraphicsPath
                $path.AddArc($boxRect.X, $boxRect.Y, $scaledRadius * 2, $scaledRadius * 2, 180, 90)
                $path.AddArc($boxRect.Right - $scaledRadius * 2, $boxRect.Y, $scaledRadius * 2, $scaledRadius * 2, 270, 90)
                $path.AddArc($boxRect.Right - $scaledRadius * 2, $boxRect.Bottom - $scaledRadius * 2, $scaledRadius * 2, $scaledRadius * 2, 0, 90)
                $path.AddArc($boxRect.X, $boxRect.Bottom - $scaledRadius * 2, $scaledRadius * 2, $scaledRadius * 2, 90, 90)
                $path.CloseFigure()
                
                $g.FillPath($boxBrush, $path)
                $boxBrush.Dispose()
                $path.Dispose()
            }
            "Glow" {
                if ($logo) {
                    Draw-LogoGlow $g $logo $blockX $blockY $logoWidth $logoHeight $backdropColor $Config.BackdropOpacity $scaledRadius
                }
            }
            "Outline" {
                if ($logo) {
                    Draw-LogoOutline $g $logo $blockX $blockY $logoWidth $logoHeight $backdropColor $Config.BackdropOpacity ([int]($scaledRadius / 3))
                }
            }
        }
    }
    
    # Draw logo
    $currentY = $blockY
    if ($logo) {
        # Apply opacity
        if ($Config.LogoOpacity -lt 255) {
            $colorMatrix = New-Object System.Drawing.Imaging.ColorMatrix
            $colorMatrix.Matrix33 = $Config.LogoOpacity / 255.0
            $imageAttr = New-Object System.Drawing.Imaging.ImageAttributes
            $imageAttr.SetColorMatrix($colorMatrix)
            
            $destRect = New-Object System.Drawing.Rectangle([int]$blockX, [int]$currentY, $logoWidth, $logoHeight)
            $g.DrawImage($logo, $destRect, 0, 0, $logo.Width, $logo.Height, [System.Drawing.GraphicsUnit]::Pixel, $imageAttr)
            $imageAttr.Dispose()
        } else {
            $g.DrawImage($logo, [int]$blockX, [int]$currentY, $logoWidth, $logoHeight)
        }
        $currentY += $logoHeight + [int](8 * $scaleY)
        $logo.Dispose()
    }
    
    # Draw text - use custom position if set
    $fontColor = [System.Drawing.ColorTranslator]::FromHtml("#$($Config.FontColor)")
    $textBrush = New-Object System.Drawing.SolidBrush($fontColor)
    $shadowBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb($Config.ShadowOpacity, 0, 0, 0))
    $scaledShadowOffset = [Math]::Max(1, [int]($Config.ShadowOffset * $scaleX))
    
    # Use custom text position or follow logo
    $textX = $textBlockX
    $textY = if ($useCustomTextPos) { $textBlockY } else { $currentY }
    
    if ($Config.PrimaryText) {
        if ($Config.EnableShadow) {
            $g.DrawString($Config.PrimaryText, $primaryFont, $shadowBrush, ($textX + $scaledShadowOffset), ($textY + $scaledShadowOffset))
        }
        $g.DrawString($Config.PrimaryText, $primaryFont, $textBrush, $textX, $textY)
        $textY += $primarySize.Height + [int](4 * $scaleY)
    }
    
    if ($Config.SecondaryText) {
        if ($Config.EnableShadow) {
            $g.DrawString($Config.SecondaryText, $secondaryFont, $shadowBrush, ($textX + $scaledShadowOffset), ($textY + $scaledShadowOffset))
        }
        $g.DrawString($Config.SecondaryText, $secondaryFont, $textBrush, $textX, $textY)
    }
    
    # Cleanup
    $primaryFont.Dispose()
    $secondaryFont.Dispose()
    $textBrush.Dispose()
    $shadowBrush.Dispose()
}

function Draw-LogoGlow {
    param(
        [System.Drawing.Graphics]$g,
        [System.Drawing.Image]$logo,
        [int]$x,
        [int]$y,
        [int]$width,
        [int]$height,
        [System.Drawing.Color]$color,
        [int]$opacity,
        [int]$radius
    )
    
    $steps = [Math]::Min($radius, 15)
    
    for ($i = $steps; $i -ge 1; $i--) {
        $expand = $i * 2
        $alpha = [int](($opacity / $steps) * ($steps - $i + 1) * 0.6)
        $alpha = [Math]::Min(255, [Math]::Max(0, $alpha))
        
        # Create color matrix for silhouette
        $matrix = New-Object System.Drawing.Imaging.ColorMatrix
        $matrix.Matrix00 = 0; $matrix.Matrix11 = 0; $matrix.Matrix22 = 0
        $matrix.Matrix40 = $color.R / 255.0
        $matrix.Matrix41 = $color.G / 255.0
        $matrix.Matrix42 = $color.B / 255.0
        $matrix.Matrix33 = $alpha / 255.0
        
        $attr = New-Object System.Drawing.Imaging.ImageAttributes
        $attr.SetColorMatrix($matrix)
        
        $destRect = New-Object System.Drawing.Rectangle(
            ($x - $expand),
            ($y - $expand),
            ($width + $expand * 2),
            ($height + $expand * 2)
        )
        
        $g.DrawImage($logo, $destRect, 0, 0, $logo.Width, $logo.Height, [System.Drawing.GraphicsUnit]::Pixel, $attr)
        $attr.Dispose()
    }
}

function Draw-LogoOutline {
    param(
        [System.Drawing.Graphics]$g,
        [System.Drawing.Image]$logo,
        [int]$x,
        [int]$y,
        [int]$width,
        [int]$height,
        [System.Drawing.Color]$color,
        [int]$opacity,
        [int]$thickness
    )
    
    $thickness = [Math]::Max(1, $thickness)
    
    # Create silhouette matrix
    $matrix = New-Object System.Drawing.Imaging.ColorMatrix
    $matrix.Matrix00 = 0; $matrix.Matrix11 = 0; $matrix.Matrix22 = 0
    $matrix.Matrix40 = $color.R / 255.0
    $matrix.Matrix41 = $color.G / 255.0
    $matrix.Matrix42 = $color.B / 255.0
    $matrix.Matrix33 = $opacity / 255.0
    
    $attr = New-Object System.Drawing.Imaging.ImageAttributes
    $attr.SetColorMatrix($matrix)
    
    # Draw in 8 directions
    $offsets = @(
        @(-$thickness, -$thickness), @(0, -$thickness), @($thickness, -$thickness),
        @(-$thickness, 0), @($thickness, 0),
        @(-$thickness, $thickness), @(0, $thickness), @($thickness, $thickness)
    )
    
    foreach ($offset in $offsets) {
        $destRect = New-Object System.Drawing.Rectangle(
            ($x + $offset[0]),
            ($y + $offset[1]),
            $width,
            $height
        )
        $g.DrawImage($logo, $destRect, 0, 0, $logo.Width, $logo.Height, [System.Drawing.GraphicsUnit]::Pixel, $attr)
    }
    
    $attr.Dispose()
}

# ============================================================================
# WALLPAPER APPLICATION
# ============================================================================

function Set-DesktopWallpaper {
    param(
        [string]$Path,
        [string]$Style = "Fill"  # Fill, Fit, Stretch, Tile, Center, Span
    )
    
    # Set wallpaper style in registry
    $styleMap = @{
        "Fill" = @{ WallpaperStyle = "10"; TileWallpaper = "0" }
        "Fit" = @{ WallpaperStyle = "6"; TileWallpaper = "0" }
        "Stretch" = @{ WallpaperStyle = "2"; TileWallpaper = "0" }
        "Tile" = @{ WallpaperStyle = "0"; TileWallpaper = "1" }
        "Center" = @{ WallpaperStyle = "0"; TileWallpaper = "0" }
        "Span" = @{ WallpaperStyle = "22"; TileWallpaper = "0" }
    }
    
    $regPath = "HKCU:\Control Panel\Desktop"
    $settings = $styleMap[$Style]
    
    Set-ItemProperty -Path $regPath -Name WallpaperStyle -Value $settings.WallpaperStyle
    Set-ItemProperty -Path $regPath -Name TileWallpaper -Value $settings.TileWallpaper
    
    # Apply wallpaper
    [WallpaperAPI]::SystemParametersInfo(
        [WallpaperAPI]::SPI_SETDESKWALLPAPER,
        0,
        $Path,
        [WallpaperAPI]::SPIF_UPDATEINIFILE -bor [WallpaperAPI]::SPIF_SENDCHANGE
    ) | Out-Null
}

function Set-LockScreenWallpaper {
    param([string]$Path)
    
    # Copy to system location
    $lockScreenPath = "$env:WINDIR\Web\Screen"
    $destPath = Join-Path $lockScreenPath "lockscreen.jpg"
    
    try {
        # Take ownership if needed
        $acl = Get-Acl $lockScreenPath
        Copy-Item $Path $destPath -Force
        
        # Set registry for lock screen
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }
        
        Set-ItemProperty -Path $regPath -Name "LockScreenImagePath" -Value $destPath
        Set-ItemProperty -Path $regPath -Name "LockScreenImageStatus" -Value 1
        
        return $true
    } catch {
        Write-Warning "Failed to set lock screen: $($_.Exception.Message)"
        return $false
    }
}

# ============================================================================
# ENTERPRISE DEPLOYMENT HELPERS
# ============================================================================

function Export-GPODeploymentScript {
    param(
        [string]$OutputPath,
        [WallBrandConfig]$Config,
        [string]$WallpaperNetworkPath
    )
    
    # Build script content using string array to avoid here-string nesting issues
    $lines = @(
        '<#'
        '.SYNOPSIS'
        '    WallBrand Pro - GPO Deployment Script'
        '    Auto-generated deployment script for Group Policy'
        ''
        '.DESCRIPTION'
        '    Deploy this script via GPO User Logon Script'
        "    Ensure wallpaper image is accessible at: $WallpaperNetworkPath"
        '#>'
        ''
        '$ErrorActionPreference = ''SilentlyContinue'''
        ''
        '# Configuration'
        "`$WallpaperSource = `"$WallpaperNetworkPath`""
        '$LocalCache = "$env:LOCALAPPDATA\WallBrand"'
        '$LocalWallpaper = "$LocalCache\wallpaper.png"'
        ''
        '# Ensure cache directory exists'
        'if (-not (Test-Path $LocalCache)) {'
        '    New-Item -Path $LocalCache -ItemType Directory -Force | Out-Null'
        '}'
        ''
        '# Copy wallpaper to local cache (for offline access)'
        'if (Test-Path $WallpaperSource) {'
        '    Copy-Item $WallpaperSource $LocalWallpaper -Force'
        '}'
        ''
        '# Apply wallpaper if local copy exists'
        'if (Test-Path $LocalWallpaper) {'
        '    # Set wallpaper style'
        '    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -Value "10"'
        '    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper -Value "0"'
        ''
        '    # Apply wallpaper via API'
        '    Add-Type -TypeDefinition @'''
        '    using System;'
        '    using System.Runtime.InteropServices;'
        '    public class WallpaperHelper {'
        '        [DllImport("user32.dll", CharSet = CharSet.Auto)]'
        '        public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);'
        '    }'
        '''@'
        '    [WallpaperHelper]::SystemParametersInfo(0x0014, 0, $LocalWallpaper, 0x03) | Out-Null'
        '}'
    )
    
    $lines -join "`r`n" | Out-File -FilePath $OutputPath -Encoding UTF8
    return $OutputPath
}

function Export-IntunePackage {
    param(
        [string]$OutputFolder,
        [WallBrandConfig]$Config,
        [string]$WallpaperPath
    )
    
    # Create package folder
    if (-not (Test-Path $OutputFolder)) {
        New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
    }
    
    # Copy wallpaper
    $wallpaperDest = Join-Path $OutputFolder "wallpaper.png"
    Copy-Item $WallpaperPath $wallpaperDest -Force
    
    # Create install script using string array
    $installLines = @(
        '<#'
        '.SYNOPSIS'
        '    WallBrand Pro - Intune Installation Script'
        '#>'
        ''
        '$ErrorActionPreference = ''Stop'''
        ''
        'try {'
        '    # Determine paths'
        '    $ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path'
        '    $WallpaperSource = Join-Path $ScriptPath "wallpaper.png"'
        '    $DestFolder = "$env:ProgramData\WallBrand"'
        '    $DestWallpaper = "$DestFolder\wallpaper.png"'
        ''
        '    # Create destination folder'
        '    if (-not (Test-Path $DestFolder)) {'
        '        New-Item -Path $DestFolder -ItemType Directory -Force | Out-Null'
        '    }'
        ''
        '    # Copy wallpaper'
        '    Copy-Item $WallpaperSource $DestWallpaper -Force'
        ''
        '    # Apply wallpaper using registry method (simpler, works for current user)'
        '    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -Value "10"'
        '    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper -Value "0"'
        '    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name Wallpaper -Value $DestWallpaper'
        ''
        '    # Refresh desktop'
        '    rundll32.exe user32.dll, UpdatePerUserSystemParameters 1, True'
        ''
        '    exit 0'
        '} catch {'
        '    Write-Error $_.Exception.Message'
        '    exit 1'
        '}'
    )
    
    $installLines -join "`r`n" | Out-File -FilePath (Join-Path $OutputFolder "Install.ps1") -Encoding UTF8
    
    # Create detection script
    $detectLines = @(
        '# Detection script for Intune'
        '$WallpaperPath = "$env:ProgramData\WallBrand\wallpaper.png"'
        'if (Test-Path $WallpaperPath) {'
        '    Write-Output "Installed"'
        '    exit 0'
        '} else {'
        '    exit 1'
        '}'
    )
    
    $detectLines -join "`r`n" | Out-File -FilePath (Join-Path $OutputFolder "Detect.ps1") -Encoding UTF8
    
    # Create uninstall script
    $uninstallLines = @(
        '# Uninstall script'
        'Remove-Item "$env:ProgramData\WallBrand" -Recurse -Force -ErrorAction SilentlyContinue'
    )
    
    $uninstallLines -join "`r`n" | Out-File -FilePath (Join-Path $OutputFolder "Uninstall.ps1") -Encoding UTF8
    
    return $OutputFolder
}

# ============================================================================
# WPF GUI
# ============================================================================

function Show-WallBrandProGUI {
    param([WallBrandConfig]$InitialConfig)
    
    $script:Config = if ($InitialConfig) { $InitialConfig } else { [WallBrandConfig]::new() }
    $script:PreviewBitmap = $null
    $script:Monitors = Get-MonitorInfo
    
    # XAML Definition
    [xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="WallBrand Pro"
    Width="1400"
    Height="900"
    MinWidth="1200"
    MinHeight="700"
    WindowStartupLocation="CenterScreen"
    Background="#0d1117">
    
    <Window.Resources>
        <!-- Dark Theme Colors -->
        <SolidColorBrush x:Key="BgDark" Color="#0d1117"/>
        <SolidColorBrush x:Key="BgPanel" Color="#161b22"/>
        <SolidColorBrush x:Key="BgInput" Color="#21262d"/>
        <SolidColorBrush x:Key="BgHover" Color="#30363d"/>
        <SolidColorBrush x:Key="Border" Color="#30363d"/>
        <SolidColorBrush x:Key="BorderLight" Color="#484f58"/>
        <SolidColorBrush x:Key="TextPrimary" Color="#e6edf3"/>
        <SolidColorBrush x:Key="TextSecondary" Color="#8b949e"/>
        <SolidColorBrush x:Key="TextMuted" Color="#6e7681"/>
        <SolidColorBrush x:Key="Accent" Color="#2f81f7"/>
        <SolidColorBrush x:Key="AccentHover" Color="#58a6ff"/>
        <SolidColorBrush x:Key="Success" Color="#3fb950"/>
        <SolidColorBrush x:Key="Warning" Color="#d29922"/>
        
        <!-- ComboBox Toggle Button Template -->
        <ControlTemplate x:Key="ComboBoxToggleButton" TargetType="ToggleButton">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition/>
                    <ColumnDefinition Width="30"/>
                </Grid.ColumnDefinitions>
                <Border x:Name="Border" Grid.ColumnSpan="2" Background="#21262d" BorderBrush="#30363d" BorderThickness="1" CornerRadius="4"/>
                <Border Grid.Column="0" Background="Transparent" Margin="1"/>
                <Path x:Name="Arrow" Grid.Column="1" Fill="#8b949e" HorizontalAlignment="Center" VerticalAlignment="Center" Data="M 0 0 L 6 6 L 12 0 Z"/>
            </Grid>
            <ControlTemplate.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter TargetName="Border" Property="BorderBrush" Value="#484f58"/>
                    <Setter TargetName="Arrow" Property="Fill" Value="#e6edf3"/>
                </Trigger>
                <Trigger Property="IsChecked" Value="True">
                    <Setter TargetName="Border" Property="BorderBrush" Value="#2f81f7"/>
                </Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate>
        
        <!-- ComboBox TextBox Template -->
        <ControlTemplate x:Key="ComboBoxTextBox" TargetType="TextBox">
            <Border x:Name="PART_ContentHost" Focusable="False" Background="Transparent"/>
        </ControlTemplate>
        
        <!-- Full ComboBox Dark Style -->
        <Style TargetType="ComboBox">
            <Setter Property="Foreground" Value="#e6edf3"/>
            <Setter Property="Background" Value="#21262d"/>
            <Setter Property="BorderBrush" Value="#30363d"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="Height" Value="32"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton x:Name="ToggleButton" Template="{StaticResource ComboBoxToggleButton}" 
                                          Focusable="False" IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" 
                                          ClickMode="Press"/>
                            <ContentPresenter x:Name="ContentSite" IsHitTestVisible="False" 
                                              Content="{TemplateBinding SelectionBoxItem}" 
                                              ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}" 
                                              ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}" 
                                              Margin="10,3,30,3" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                            <TextBox x:Name="PART_EditableTextBox" Style="{x:Null}" Template="{StaticResource ComboBoxTextBox}" 
                                     HorizontalAlignment="Left" VerticalAlignment="Center" Margin="10,3,30,3" Focusable="True" 
                                     Background="Transparent" Foreground="#e6edf3" Visibility="Hidden" IsReadOnly="{TemplateBinding IsReadOnly}"/>
                            <Popup x:Name="Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}" 
                                   AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                                <Grid x:Name="DropDown" SnapsToDevicePixels="True" MinWidth="{TemplateBinding ActualWidth}" MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                    <Border x:Name="DropDownBorder" Background="#21262d" BorderBrush="#484f58" BorderThickness="1" CornerRadius="4" Margin="0,2,0,0">
                                        <ScrollViewer Margin="4" SnapsToDevicePixels="True">
                                            <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained"/>
                                        </ScrollViewer>
                                    </Border>
                                </Grid>
                            </Popup>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- ComboBoxItem Dark Style -->
        <Style TargetType="ComboBoxItem">
            <Setter Property="Foreground" Value="#e6edf3"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border x:Name="Border" Padding="{TemplateBinding Padding}" Background="{TemplateBinding Background}" CornerRadius="4" Margin="2">
                            <ContentPresenter/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#30363d"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#2f81f7"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- Button Style -->
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="#21262d"/>
            <Setter Property="Foreground" Value="#e6edf3"/>
            <Setter Property="BorderBrush" Value="#30363d"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#30363d"/>
                    <Setter Property="BorderBrush" Value="#484f58"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <!-- Primary Button Style -->
        <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource ModernButton}">
            <Setter Property="Background" Value="#238636"/>
            <Setter Property="BorderBrush" Value="#238636"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#2ea043"/>
                    <Setter Property="BorderBrush" Value="#2ea043"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <!-- TextBox Style -->
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#21262d"/>
            <Setter Property="Foreground" Value="#e6edf3"/>
            <Setter Property="BorderBrush" Value="#30363d"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="Height" Value="32"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="CaretBrush" Value="#e6edf3"/>
        </Style>
        
        <!-- CheckBox Style -->
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="#e6edf3"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        
        <!-- Section Header Style -->
        <Style x:Key="SectionHeader" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#e6edf3"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Margin" Value="0,16,0,12"/>
        </Style>
        
        <!-- Label Style -->
        <Style x:Key="FieldLabel" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#8b949e"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Margin" Value="0,0,0,6"/>
        </Style>
    </Window.Resources>
    
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="380"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        
        <!-- Left Sidebar -->
        <Border Grid.Column="0" Background="#161b22" BorderBrush="#30363d" BorderThickness="0,0,1,0">
            <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
                <StackPanel Margin="20">
                    
                    <!-- App Header -->
                    <StackPanel Margin="0,0,0,20">
                        <TextBlock Text="WallBrand Pro" FontSize="24" FontWeight="Bold" Foreground="#e6edf3"/>
                        <TextBlock Text="Enterprise Wallpaper Branding" FontSize="12" Foreground="#8b949e" Margin="0,4,0,0"/>
                    </StackPanel>
                    
                    <!-- Source Files -->
                    <TextBlock Text="SOURCE FILES" Style="{StaticResource SectionHeader}"/>
                    
                    <TextBlock Text="Wallpaper Image" Style="{StaticResource FieldLabel}"/>
                    <Grid Margin="0,0,0,12">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="txtWallpaper" Grid.Column="0"/>
                        <Button x:Name="btnBrowseWallpaper" Grid.Column="1" Content="Browse" Style="{StaticResource ModernButton}" Margin="8,0,0,0" Padding="12,6"/>
                    </Grid>
                    
                    <TextBlock Text="Logo Image (PNG)" Style="{StaticResource FieldLabel}"/>
                    <Grid Margin="0,0,0,12">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="txtLogo" Grid.Column="0"/>
                        <Button x:Name="btnBrowseLogo" Grid.Column="1" Content="Browse" Style="{StaticResource ModernButton}" Margin="8,0,0,0" Padding="12,6"/>
                    </Grid>
                    
                    <!-- Branding Text -->
                    <TextBlock Text="BRANDING" Style="{StaticResource SectionHeader}"/>
                    
                    <TextBlock Text="Primary Text" Style="{StaticResource FieldLabel}"/>
                    <TextBox x:Name="txtPrimaryText" Margin="0,0,0,12"/>
                    
                    <TextBlock Text="Secondary Text" Style="{StaticResource FieldLabel}"/>
                    <TextBox x:Name="txtSecondaryText" Margin="0,0,0,12"/>
                    
                    <!-- Position -->
                    <TextBlock Text="POSITION" Style="{StaticResource SectionHeader}"/>
                    
                    <TextBlock Text="Drag elements in preview or use preset:" Style="{StaticResource FieldLabel}" Foreground="#3fb950"/>
                    
                    <Grid Margin="0,0,0,12">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="80"/>
                        </Grid.ColumnDefinitions>
                        <Grid.RowDefinitions>
                            <RowDefinition/>
                            <RowDefinition/>
                        </Grid.RowDefinitions>
                        
                        <TextBlock Text="Preset" Style="{StaticResource FieldLabel}" Grid.Column="0"/>
                        <TextBlock Text="Margin" Style="{StaticResource FieldLabel}" Grid.Column="1"/>
                        
                        <ComboBox x:Name="cmbPosition" Grid.Row="1" Grid.Column="0" Margin="0,0,8,0">
                            <ComboBoxItem Content="TopLeft"/>
                            <ComboBoxItem Content="TopCenter"/>
                            <ComboBoxItem Content="TopRight"/>
                            <ComboBoxItem Content="CenterLeft"/>
                            <ComboBoxItem Content="Center"/>
                            <ComboBoxItem Content="CenterRight"/>
                            <ComboBoxItem Content="BottomLeft"/>
                            <ComboBoxItem Content="BottomCenter"/>
                            <ComboBoxItem Content="BottomRight" IsSelected="True"/>
                            <ComboBoxItem Content="Custom"/>
                        </ComboBox>
                        <TextBox x:Name="txtMargin" Grid.Row="1" Grid.Column="1" Text="40"/>
                    </Grid>
                    
                    <!-- Custom position display -->
                    <Border Background="#21262d" CornerRadius="6" Padding="12" Margin="0,0,0,12">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Grid.Column="0">
                                <TextBlock Text="Logo Position" Foreground="#8b949e" FontSize="11"/>
                                <TextBlock x:Name="lblLogoPos" Text="Preset" Foreground="#58a6ff" FontWeight="SemiBold"/>
                            </StackPanel>
                            <StackPanel Grid.Column="1">
                                <TextBlock Text="Text Position" Foreground="#8b949e" FontSize="11"/>
                                <TextBlock x:Name="lblTextPos" Text="Preset" Foreground="#58a6ff" FontWeight="SemiBold"/>
                            </StackPanel>
                        </Grid>
                    </Border>
                    
                    <Button x:Name="btnResetPositions" Content="Reset to Preset" Style="{StaticResource ModernButton}" HorizontalAlignment="Left" Margin="0,0,0,12"/>
                    
                    <!-- Logo Settings -->
                    <TextBlock Text="LOGO" Style="{StaticResource SectionHeader}"/>
                    
                    <TextBlock Text="Scale (or drag corners in preview)" Style="{StaticResource FieldLabel}"/>
                    <Grid Margin="0,0,0,12">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="50"/>
                        </Grid.ColumnDefinitions>
                        <Slider x:Name="sliderScale" Minimum="0.1" Maximum="3" Value="1" TickFrequency="0.1" VerticalAlignment="Center"/>
                        <TextBlock x:Name="lblScaleValue" Grid.Column="1" Text="1.0x" Foreground="#58a6ff" FontWeight="SemiBold" VerticalAlignment="Center" HorizontalAlignment="Right"/>
                    </Grid>
                    
                    <!-- Logo Crop -->
                    <TextBlock Text="Crop Logo" Style="{StaticResource FieldLabel}"/>
                    <Border Background="#21262d" CornerRadius="6" Padding="12" Margin="0,0,0,12">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition/>
                                <RowDefinition/>
                                <RowDefinition/>
                            </Grid.RowDefinitions>
                            
                            <StackPanel Grid.Column="0" Grid.Row="0" Margin="0,0,4,8">
                                <TextBlock Text="Left %" Foreground="#8b949e" FontSize="10"/>
                                <TextBox x:Name="txtCropLeft" Text="0" Height="28"/>
                            </StackPanel>
                            <StackPanel Grid.Column="1" Grid.Row="0" Margin="4,0,0,8">
                                <TextBlock Text="Right %" Foreground="#8b949e" FontSize="10"/>
                                <TextBox x:Name="txtCropRight" Text="0" Height="28"/>
                            </StackPanel>
                            <StackPanel Grid.Column="0" Grid.Row="1" Margin="0,0,4,8">
                                <TextBlock Text="Top %" Foreground="#8b949e" FontSize="10"/>
                                <TextBox x:Name="txtCropTop" Text="0" Height="28"/>
                            </StackPanel>
                            <StackPanel Grid.Column="1" Grid.Row="1" Margin="4,0,0,8">
                                <TextBlock Text="Bottom %" Foreground="#8b949e" FontSize="10"/>
                                <TextBox x:Name="txtCropBottom" Text="0" Height="28"/>
                            </StackPanel>
                            <Button x:Name="btnResetCrop" Grid.ColumnSpan="2" Grid.Row="2" Content="Reset Crop" Style="{StaticResource ModernButton}" Padding="8,4"/>
                        </Grid>
                    </Border>
                    
                    <!-- Text Style -->
                    <TextBlock Text="TEXT STYLE" Style="{StaticResource SectionHeader}"/>
                    
                    <Grid Margin="0,0,0,12">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="70"/>
                            <ColumnDefinition Width="50"/>
                        </Grid.ColumnDefinitions>
                        <Grid.RowDefinitions>
                            <RowDefinition/>
                            <RowDefinition/>
                        </Grid.RowDefinitions>
                        
                        <TextBlock Text="Font" Style="{StaticResource FieldLabel}" Grid.Column="0"/>
                        <TextBlock Text="Size" Style="{StaticResource FieldLabel}" Grid.Column="1"/>
                        <TextBlock Text="Color" Style="{StaticResource FieldLabel}" Grid.Column="2"/>
                        
                        <ComboBox x:Name="cmbFont" Grid.Row="1" Grid.Column="0" Margin="0,0,8,0">
                            <ComboBoxItem Content="Segoe UI" IsSelected="True"/>
                            <ComboBoxItem Content="Arial"/>
                            <ComboBoxItem Content="Calibri"/>
                            <ComboBoxItem Content="Consolas"/>
                            <ComboBoxItem Content="Verdana"/>
                            <ComboBoxItem Content="Tahoma"/>
                            <ComboBoxItem Content="Georgia"/>
                        </ComboBox>
                        <TextBox x:Name="txtFontSize" Grid.Row="1" Grid.Column="1" Text="14" Margin="0,0,8,0"/>
                        <Border x:Name="btnFontColor" Grid.Row="1" Grid.Column="2" Background="White" CornerRadius="4" Cursor="Hand" BorderBrush="#30363d" BorderThickness="1" Height="32"/>
                    </Grid>
                    
                    <CheckBox x:Name="chkShadow" Content="Text Shadow" IsChecked="True" Margin="0,0,0,8"/>
                    <CheckBox x:Name="chkBold" Content="Bold Text" Margin="0,0,0,12"/>
                    
                    <!-- Backdrop -->
                    <TextBlock Text="BACKDROP / GLOW" Style="{StaticResource SectionHeader}"/>
                    
                    <CheckBox x:Name="chkBackdrop" Content="Enable Backdrop" IsChecked="True" Margin="0,0,0,12"/>
                    
                    <Grid Margin="0,0,0,12">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="70"/>
                            <ColumnDefinition Width="50"/>
                        </Grid.ColumnDefinitions>
                        <Grid.RowDefinitions>
                            <RowDefinition/>
                            <RowDefinition/>
                        </Grid.RowDefinitions>
                        
                        <TextBlock Text="Mode" Style="{StaticResource FieldLabel}" Grid.Column="0"/>
                        <TextBlock Text="Opacity" Style="{StaticResource FieldLabel}" Grid.Column="1"/>
                        <TextBlock Text="Color" Style="{StaticResource FieldLabel}" Grid.Column="2"/>
                        
                        <ComboBox x:Name="cmbBackdropMode" Grid.Row="1" Grid.Column="0" Margin="0,0,8,0">
                            <ComboBoxItem Content="Box"/>
                            <ComboBoxItem Content="Glow" IsSelected="True"/>
                            <ComboBoxItem Content="Outline"/>
                        </ComboBox>
                        <TextBox x:Name="txtBackdropOpacity" Grid.Row="1" Grid.Column="1" Text="180" Margin="0,0,8,0"/>
                        <Border x:Name="btnBackdropColor" Grid.Row="1" Grid.Column="2" Background="Black" CornerRadius="4" Cursor="Hand" BorderBrush="#30363d" BorderThickness="1" Height="32"/>
                    </Grid>
                    
                    <!-- System Info -->
                    <TextBlock Text="SYSTEM INFO (BGInfo)" Style="{StaticResource SectionHeader}"/>
                    
                    <CheckBox x:Name="chkSystemInfo" Content="Enable System Information Overlay" Margin="0,0,0,12"/>
                    
                    <TextBlock Text="Position" Style="{StaticResource FieldLabel}"/>
                    <ComboBox x:Name="cmbSystemInfoPosition" Margin="0,0,0,12">
                        <ComboBoxItem Content="TopLeft" IsSelected="True"/>
                        <ComboBoxItem Content="TopRight"/>
                        <ComboBoxItem Content="BottomLeft"/>
                        <ComboBoxItem Content="BottomRight"/>
                    </ComboBox>
                    
                    <!-- System Info Styling -->
                    <TextBlock Text="System Info Font" Style="{StaticResource FieldLabel}"/>
                    <Grid Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="60"/>
                        </Grid.ColumnDefinitions>
                        <ComboBox x:Name="cmbSysInfoFont" Grid.Column="0" Margin="0,0,8,0">
                            <ComboBoxItem Content="Consolas" IsSelected="True"/>
                            <ComboBoxItem Content="Segoe UI"/>
                            <ComboBoxItem Content="Courier New"/>
                            <ComboBoxItem Content="Lucida Console"/>
                            <ComboBoxItem Content="Arial"/>
                        </ComboBox>
                        <TextBox x:Name="txtSysInfoFontSize" Grid.Column="1" Text="11"/>
                    </Grid>
                    
                    <Grid Margin="0,0,0,12">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Grid.Column="0" Margin="0,0,4,0">
                            <TextBlock Text="Label Color" Foreground="#8b949e" FontSize="10"/>
                            <Border x:Name="btnSysInfoLabelColor" Background="#88CCFF" CornerRadius="4" Cursor="Hand" BorderBrush="#30363d" BorderThickness="1" Height="28"/>
                        </StackPanel>
                        <StackPanel Grid.Column="1" Margin="4,0,0,0">
                            <TextBlock Text="Value Color" Foreground="#8b949e" FontSize="10"/>
                            <Border x:Name="btnSysInfoValueColor" Background="#FFFFFF" CornerRadius="4" Cursor="Hand" BorderBrush="#30363d" BorderThickness="1" Height="28"/>
                        </StackPanel>
                    </Grid>
                    
                    <TextBlock Text="Fields to Display" Style="{StaticResource FieldLabel}"/>
                    <Border Background="#21262d" CornerRadius="6" Padding="12" Margin="0,0,0,12">
                        <StackPanel>
                            <CheckBox x:Name="chkHostName" Content="Host Name" IsChecked="True" Margin="0,0,0,4"/>
                            <CheckBox x:Name="chkIPAddress" Content="IP Address" IsChecked="True" Margin="0,0,0,4"/>
                            <CheckBox x:Name="chkUserName" Content="User Name" IsChecked="True" Margin="0,0,0,4"/>
                            <CheckBox x:Name="chkOSVersion" Content="OS Version" IsChecked="True" Margin="0,0,0,4"/>
                            <CheckBox x:Name="chkCPU" Content="CPU" Margin="0,0,0,4"/>
                            <CheckBox x:Name="chkRAM" Content="RAM" Margin="0,0,0,4"/>
                            <CheckBox x:Name="chkDisk" Content="Disk Space" Margin="0,0,0,4"/>
                            <CheckBox x:Name="chkUptime" Content="Uptime" Margin="0,0,0,4"/>
                            <CheckBox x:Name="chkSerial" Content="Serial Number" Margin="0"/>
                        </StackPanel>
                    </Border>
                    
                    <!-- Multi-Monitor -->
                    <TextBlock Text="MULTI-MONITOR" Style="{StaticResource SectionHeader}"/>
                    
                    <TextBlock Text="Mode" Style="{StaticResource FieldLabel}"/>
                    <ComboBox x:Name="cmbMultiMonitor" Margin="0,0,0,12">
                        <ComboBoxItem Content="Clone (same on all)" IsSelected="True"/>
                        <ComboBoxItem Content="Span (stretch across)"/>
                        <ComboBoxItem Content="Per-Monitor (individual)"/>
                    </ComboBox>
                    
                    <TextBlock x:Name="lblMonitorInfo" Text="Detected: 1 monitor(s)" Foreground="#6e7681" FontSize="11" Margin="0,0,0,20"/>
                    
                    <!-- Auto-Update Scheduled Task -->
                    <TextBlock Text="AUTO-UPDATE (Scheduled Task)" Style="{StaticResource SectionHeader}"/>
                    
                    <TextBlock Text="Keep system info current by automatically refreshing the wallpaper on a schedule." Foreground="#8b949e" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,12"/>
                    
                    <TextBlock Text="Update Frequency" Style="{StaticResource FieldLabel}"/>
                    <ComboBox x:Name="cmbUpdateFrequency" Margin="0,0,0,12">
                        <ComboBoxItem Content="Every 5 minutes"/>
                        <ComboBoxItem Content="Every 15 minutes" IsSelected="True"/>
                        <ComboBoxItem Content="Every 30 minutes"/>
                        <ComboBoxItem Content="Every hour"/>
                        <ComboBoxItem Content="Every 4 hours"/>
                        <ComboBoxItem Content="Daily"/>
                    </ComboBox>
                    
                    <TextBlock Text="Additional Triggers" Style="{StaticResource FieldLabel}"/>
                    <Border Background="#21262d" CornerRadius="6" Padding="12" Margin="0,0,0,12">
                        <StackPanel>
                            <CheckBox x:Name="chkTriggerLogon" Content="On user logon" IsChecked="True" Margin="0,0,0,4"/>
                            <CheckBox x:Name="chkTriggerUnlock" Content="On workstation unlock" IsChecked="True" Margin="0,0,0,4"/>
                            <CheckBox x:Name="chkTriggerResume" Content="On resume from sleep" Margin="0"/>
                        </StackPanel>
                    </Border>
                    
                    <TextBlock Text="Options" Style="{StaticResource FieldLabel}"/>
                    <Border Background="#21262d" CornerRadius="6" Padding="12" Margin="0,0,0,12">
                        <StackPanel>
                            <CheckBox x:Name="chkRunHidden" Content="Run silently (no window)" IsChecked="True" Margin="0,0,0,4"/>
                            <CheckBox x:Name="chkRunAsAdmin" Content="Run with highest privileges" Margin="0"/>
                        </StackPanel>
                    </Border>
                    
                    <TextBlock x:Name="lblTaskStatus" Text="Status: No scheduled task" Foreground="#6e7681" FontSize="11" Margin="0,0,0,8"/>
                    
                    <Grid Margin="0,0,0,12">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Button x:Name="btnCreateTask" Content="Create Task" Grid.Column="0" Margin="0,0,4,0" Padding="8,6"/>
                        <Button x:Name="btnRemoveTask" Content="Remove Task" Grid.Column="1" Margin="4,0,0,0" Padding="8,6"/>
                    </Grid>
                    
                    <Button x:Name="btnRefreshNow" Content="Refresh Wallpaper Now" Padding="8,6" Margin="0,0,0,20"/>
                    
                </StackPanel>
            </ScrollViewer>
        </Border>
        
        <!-- Main Content Area -->
        <Grid Grid.Column="1">
            <Grid.RowDefinitions>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            
            <!-- Preview Area -->
            <Border Grid.Row="0" Background="#0d1117" Margin="20">
                <Grid x:Name="previewContainer">
                    <Border x:Name="previewBorder" Background="#161b22" CornerRadius="8" BorderBrush="#30363d" BorderThickness="1">
                        <Grid>
                            <Image x:Name="imgPreview" Stretch="Uniform" RenderOptions.BitmapScalingMode="HighQuality"/>
                            <!-- Interactive overlay canvas -->
                            <Canvas x:Name="dragCanvas" Background="Transparent" ClipToBounds="True">
                                
                                <!-- Logo container with actual image -->
                                <Canvas x:Name="logoContainer" Visibility="Collapsed">
                                    <!-- Actual logo image -->
                                    <Image x:Name="logoImage" Stretch="Fill" Cursor="SizeAll" RenderOptions.BitmapScalingMode="HighQuality"/>
                                    <!-- Selection border -->
                                    <Border x:Name="logoSelectionBorder" BorderBrush="#2f81f7" BorderThickness="2" Background="Transparent" IsHitTestVisible="False"/>
                                    <!-- Resize handles -->
                                    <Border x:Name="logoResizeNW" Width="12" Height="12" Background="#2f81f7" BorderBrush="White" BorderThickness="1" Cursor="SizeNWSE" CornerRadius="2"/>
                                    <Border x:Name="logoResizeNE" Width="12" Height="12" Background="#2f81f7" BorderBrush="White" BorderThickness="1" Cursor="SizeNESW" CornerRadius="2"/>
                                    <Border x:Name="logoResizeSW" Width="12" Height="12" Background="#2f81f7" BorderBrush="White" BorderThickness="1" Cursor="SizeNESW" CornerRadius="2"/>
                                    <Border x:Name="logoResizeSE" Width="12" Height="12" Background="#2f81f7" BorderBrush="White" BorderThickness="1" Cursor="SizeNWSE" CornerRadius="2"/>
                                </Canvas>
                                
                                <!-- Text container with live editable text -->
                                <Canvas x:Name="textContainer" Visibility="Collapsed">
                                    <!-- Background for text area - used for dragging -->
                                    <Border x:Name="textBackground" Background="#60000000" CornerRadius="4" Cursor="SizeAll"/>
                                    <!-- Primary text - click to edit -->
                                    <TextBlock x:Name="primaryTextBlock" Foreground="White" FontWeight="SemiBold" Cursor="IBeam" ToolTip="Click to edit"/>
                                    <TextBox x:Name="primaryTextEdit" Background="#333333" Foreground="White" BorderBrush="#2f81f7" BorderThickness="2" Visibility="Collapsed" AcceptsReturn="False"/>
                                    <!-- Secondary text - click to edit -->
                                    <TextBlock x:Name="secondaryTextBlock" Foreground="#CCCCCC" Cursor="IBeam" ToolTip="Click to edit"/>
                                    <TextBox x:Name="secondaryTextEdit" Background="#333333" Foreground="White" BorderBrush="#2f81f7" BorderThickness="2" Visibility="Collapsed" AcceptsReturn="False"/>
                                    <!-- Selection border -->
                                    <Border x:Name="textSelectionBorder" BorderBrush="#3fb950" BorderThickness="2" Background="Transparent" IsHitTestVisible="False"/>
                                    <!-- Resize handles for text -->
                                    <Border x:Name="textResizeNW" Width="12" Height="12" Background="#3fb950" BorderBrush="White" BorderThickness="1" Cursor="SizeNWSE" CornerRadius="2"/>
                                    <Border x:Name="textResizeNE" Width="12" Height="12" Background="#3fb950" BorderBrush="White" BorderThickness="1" Cursor="SizeNESW" CornerRadius="2"/>
                                    <Border x:Name="textResizeSW" Width="12" Height="12" Background="#3fb950" BorderBrush="White" BorderThickness="1" Cursor="SizeNESW" CornerRadius="2"/>
                                    <Border x:Name="textResizeSE" Width="12" Height="12" Background="#3fb950" BorderBrush="White" BorderThickness="1" Cursor="SizeNWSE" CornerRadius="2"/>
                                </Canvas>
                                
                                <!-- System Info container -->
                                <Canvas x:Name="sysInfoContainer" Visibility="Collapsed">
                                    <!-- Background - use this for dragging -->
                                    <Border x:Name="sysInfoBackground" Background="#80000000" CornerRadius="4" Cursor="SizeAll"/>
                                    <!-- System info text panel - rebuilt dynamically -->
                                    <StackPanel x:Name="sysInfoPanel" IsHitTestVisible="False">
                                        <!-- Lines added dynamically -->
                                    </StackPanel>
                                    <!-- Selection border -->
                                    <Border x:Name="sysInfoSelectionBorder" BorderBrush="#d29922" BorderThickness="2" Background="Transparent" IsHitTestVisible="False"/>
                                    <!-- Resize handles -->
                                    <Border x:Name="sysInfoResizeNW" Width="12" Height="12" Background="#d29922" BorderBrush="White" BorderThickness="1" Cursor="SizeNWSE" CornerRadius="2"/>
                                    <Border x:Name="sysInfoResizeNE" Width="12" Height="12" Background="#d29922" BorderBrush="White" BorderThickness="1" Cursor="SizeNESW" CornerRadius="2"/>
                                    <Border x:Name="sysInfoResizeSW" Width="12" Height="12" Background="#d29922" BorderBrush="White" BorderThickness="1" Cursor="SizeNESW" CornerRadius="2"/>
                                    <Border x:Name="sysInfoResizeSE" Width="12" Height="12" Background="#d29922" BorderBrush="White" BorderThickness="1" Cursor="SizeNWSE" CornerRadius="2"/>
                                </Canvas>
                                
                            </Canvas>
                        </Grid>
                    </Border>
                    
                    <!-- Preview placeholder -->
                    <StackPanel x:Name="previewPlaceholder" VerticalAlignment="Center" HorizontalAlignment="Center">
                        <TextBlock Text="PREVIEW" FontSize="32" FontWeight="Light" Foreground="#30363d" HorizontalAlignment="Center"/>
                        <TextBlock Text="Select a wallpaper to begin" FontSize="14" Foreground="#6e7681" HorizontalAlignment="Center" Margin="0,8,0,0"/>
                        <TextBlock Text="Drag to move, corners to resize" FontSize="12" Foreground="#3fb950" HorizontalAlignment="Center" Margin="0,4,0,0"/>
                    </StackPanel>
                </Grid>
            </Border>
            
            <!-- Bottom Action Bar -->
            <Border Grid.Row="1" Background="#161b22" BorderBrush="#30363d" BorderThickness="0,1,0,0" Padding="20,16">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    
                    <!-- Left buttons -->
                    <StackPanel Grid.Column="0" Orientation="Horizontal">
                        <Button x:Name="btnRefresh" Content="Refresh Preview" Style="{StaticResource ModernButton}"/>
                        <Button x:Name="btnSaveImage" Content="Save Image" Style="{StaticResource ModernButton}" Margin="8,0,0,0"/>
                    </StackPanel>
                    
                    <!-- Status -->
                    <TextBlock x:Name="lblStatus" Grid.Column="1" Text="Ready - Drag elements in preview to position" Foreground="#6e7681" VerticalAlignment="Center" HorizontalAlignment="Center"/>
                    
                    <!-- Right buttons -->
                    <StackPanel Grid.Column="2" Orientation="Horizontal">
                        <Button x:Name="btnExportGPO" Content="Export GPO" Style="{StaticResource ModernButton}"/>
                        <Button x:Name="btnExportIntune" Content="Export Intune" Style="{StaticResource ModernButton}" Margin="8,0,0,0"/>
                        <Button x:Name="btnLoadConfig" Content="Load Config" Style="{StaticResource ModernButton}" Margin="8,0,0,0"/>
                        <Button x:Name="btnSaveConfig" Content="Save Config" Style="{StaticResource ModernButton}" Margin="8,0,0,0"/>
                        <Button x:Name="btnApply" Content="Apply Wallpaper" Style="{StaticResource PrimaryButton}" Margin="8,0,0,0"/>
                    </StackPanel>
                </Grid>
            </Border>
        </Grid>
    </Grid>
</Window>
"@
    
    # Parse XAML
    $reader = New-Object System.Xml.XmlNodeReader($xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
    
# codex-branding:start
                try {
                    $brandingIconPath = Join-Path $PSScriptRoot 'icon.ico'
                    if (Test-Path $brandingIconPath) {
                        $window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create((New-Object System.Uri($brandingIconPath)))
                    }
                } catch {
                }
                # codex-branding:end
    # Get controls
    $txtWallpaper = $window.FindName("txtWallpaper")
    $btnBrowseWallpaper = $window.FindName("btnBrowseWallpaper")
    $txtLogo = $window.FindName("txtLogo")
    $btnBrowseLogo = $window.FindName("btnBrowseLogo")
    $txtPrimaryText = $window.FindName("txtPrimaryText")
    $txtSecondaryText = $window.FindName("txtSecondaryText")
    $cmbPosition = $window.FindName("cmbPosition")
    $txtMargin = $window.FindName("txtMargin")
    $sliderScale = $window.FindName("sliderScale")
    $lblScaleValue = $window.FindName("lblScaleValue")
    $cmbFont = $window.FindName("cmbFont")
    $txtFontSize = $window.FindName("txtFontSize")
    $btnFontColor = $window.FindName("btnFontColor")
    $chkShadow = $window.FindName("chkShadow")
    $chkBold = $window.FindName("chkBold")
    $chkBackdrop = $window.FindName("chkBackdrop")
    $cmbBackdropMode = $window.FindName("cmbBackdropMode")
    $txtBackdropOpacity = $window.FindName("txtBackdropOpacity")
    $btnBackdropColor = $window.FindName("btnBackdropColor")
    $chkSystemInfo = $window.FindName("chkSystemInfo")
    $cmbSystemInfoPosition = $window.FindName("cmbSystemInfoPosition")
    $cmbMultiMonitor = $window.FindName("cmbMultiMonitor")
    $lblMonitorInfo = $window.FindName("lblMonitorInfo")
    $imgPreview = $window.FindName("imgPreview")
    $previewPlaceholder = $window.FindName("previewPlaceholder")
    $btnRefresh = $window.FindName("btnRefresh")
    $btnSaveImage = $window.FindName("btnSaveImage")
    $btnExportGPO = $window.FindName("btnExportGPO")
    $btnExportIntune = $window.FindName("btnExportIntune")
    $btnLoadConfig = $window.FindName("btnLoadConfig")
    $btnSaveConfig = $window.FindName("btnSaveConfig")
    $btnApply = $window.FindName("btnApply")
    $lblStatus = $window.FindName("lblStatus")
    
    # System info checkboxes
    $chkHostName = $window.FindName("chkHostName")
    $chkIPAddress = $window.FindName("chkIPAddress")
    $chkUserName = $window.FindName("chkUserName")
    $chkOSVersion = $window.FindName("chkOSVersion")
    $chkCPU = $window.FindName("chkCPU")
    $chkRAM = $window.FindName("chkRAM")
    $chkDisk = $window.FindName("chkDisk")
    $chkUptime = $window.FindName("chkUptime")
    $chkSerial = $window.FindName("chkSerial")
    
    # New controls for drag functionality
    $dragCanvas = $window.FindName("dragCanvas")
    $lblLogoPos = $window.FindName("lblLogoPos")
    $lblTextPos = $window.FindName("lblTextPos")
    $btnResetPositions = $window.FindName("btnResetPositions")
    $previewContainer = $window.FindName("previewContainer")
    
    # Logo container and elements
    $logoContainer = $window.FindName("logoContainer")
    $logoImage = $window.FindName("logoImage")
    $logoSelectionBorder = $window.FindName("logoSelectionBorder")
    $logoResizeNW = $window.FindName("logoResizeNW")
    $logoResizeNE = $window.FindName("logoResizeNE")
    $logoResizeSW = $window.FindName("logoResizeSW")
    $logoResizeSE = $window.FindName("logoResizeSE")
    
    # Text container and elements
    $textContainer = $window.FindName("textContainer")
    $textBackground = $window.FindName("textBackground")
    $primaryTextBlock = $window.FindName("primaryTextBlock")
    $primaryTextEdit = $window.FindName("primaryTextEdit")
    $secondaryTextBlock = $window.FindName("secondaryTextBlock")
    $secondaryTextEdit = $window.FindName("secondaryTextEdit")
    $textSelectionBorder = $window.FindName("textSelectionBorder")
    $textResizeNW = $window.FindName("textResizeNW")
    $textResizeNE = $window.FindName("textResizeNE")
    $textResizeSW = $window.FindName("textResizeSW")
    $textResizeSE = $window.FindName("textResizeSE")
    
    # System Info container and elements
    $sysInfoContainer = $window.FindName("sysInfoContainer")
    $sysInfoBackground = $window.FindName("sysInfoBackground")
    $sysInfoPanel = $window.FindName("sysInfoPanel")
    $sysInfoSelectionBorder = $window.FindName("sysInfoSelectionBorder")
    $sysInfoResizeNW = $window.FindName("sysInfoResizeNW")
    $sysInfoResizeNE = $window.FindName("sysInfoResizeNE")
    $sysInfoResizeSW = $window.FindName("sysInfoResizeSW")
    $sysInfoResizeSE = $window.FindName("sysInfoResizeSE")
    
    # System Info styling controls
    $cmbSysInfoFont = $window.FindName("cmbSysInfoFont")
    $txtSysInfoFontSize = $window.FindName("txtSysInfoFontSize")
    $btnSysInfoLabelColor = $window.FindName("btnSysInfoLabelColor")
    $btnSysInfoValueColor = $window.FindName("btnSysInfoValueColor")
    
    # Scheduled task controls
    $cmbUpdateFrequency = $window.FindName("cmbUpdateFrequency")
    $chkTriggerLogon = $window.FindName("chkTriggerLogon")
    $chkTriggerUnlock = $window.FindName("chkTriggerUnlock")
    $chkTriggerResume = $window.FindName("chkTriggerResume")
    $chkRunHidden = $window.FindName("chkRunHidden")
    $chkRunAsAdmin = $window.FindName("chkRunAsAdmin")
    $lblTaskStatus = $window.FindName("lblTaskStatus")
    $btnCreateTask = $window.FindName("btnCreateTask")
    $btnRemoveTask = $window.FindName("btnRemoveTask")
    $btnRefreshNow = $window.FindName("btnRefreshNow")
    
    # Crop controls
    $txtCropLeft = $window.FindName("txtCropLeft")
    $txtCropRight = $window.FindName("txtCropRight")
    $txtCropTop = $window.FindName("txtCropTop")
    $txtCropBottom = $window.FindName("txtCropBottom")
    $btnResetCrop = $window.FindName("btnResetCrop")
    
    # Color variables
    $script:FontColorHex = "FFFFFF"
    $script:BackdropColorHex = "000000"
    $script:SysInfoLabelColorHex = "88CCFF"
    $script:SysInfoValueColorHex = "FFFFFF"
    
    # Drag state variables
    $script:IsDraggingLogo = $false
    $script:IsDraggingText = $false
    $script:IsDraggingSysInfo = $false
    $script:IsResizingLogo = $false
    $script:IsResizingText = $false
    $script:IsResizingSysInfo = $false
    $script:ResizeCorner = ""
    $script:DragOffsetX = 0
    $script:DragOffsetY = 0
    $script:ResizeStartX = 0
    $script:ResizeStartY = 0
    $script:ResizeStartW = 0
    $script:ResizeStartH = 0
    $script:CustomLogoX = -1
    $script:CustomLogoY = -1
    $script:CustomLogoW = -1
    $script:CustomLogoH = -1
    $script:CustomTextX = -1
    $script:CustomTextY = -1
    $script:CustomFontSize = -1
    $script:UseCustomLogoPosition = $false
    $script:UseCustomTextPosition = $false
    $script:PreviewScale = 1.0
    $script:PreviewOffsetX = 0
    $script:PreviewOffsetY = 0
    $script:WallpaperWidth = 1920
    $script:WallpaperHeight = 1080
    $script:OriginalLogoWidth = 0
    $script:OriginalLogoHeight = 0
    $script:LogoBitmapSource = $null
    
    # Crop values
    $script:CropLeft = 0
    $script:CropRight = 0
    $script:CropTop = 0
    $script:CropBottom = 0
    
    # Current positions for text alignment
    $script:CurrentLogoX = 0
    $script:CurrentLogoY = 0
    $script:CurrentLogoW = 100
    $script:CurrentLogoH = 100
    
    # System info position
    $script:CustomSysInfoX = -1
    $script:CustomSysInfoY = -1
    $script:UseCustomSysInfoPosition = $false
    $script:CustomSysInfoFontSize = -1
    
    # Cache for system info data (expensive CIM calls)
    $script:SysInfoCache = $null
    
    # Initialize from config
    $txtWallpaper.Text = $script:Config.WallpaperPath
    $txtLogo.Text = $script:Config.LogoPath
    $txtPrimaryText.Text = $script:Config.PrimaryText
    $txtSecondaryText.Text = $script:Config.SecondaryText
    $sliderScale.Value = $script:Config.LogoScale
    
    # Update monitor info
    $lblMonitorInfo.Text = "Detected: $($script:Monitors.Count) monitor(s)"
    
    # Helper to get config from UI
    $getConfigFromUI = {
        $cfg = [WallBrandConfig]::new()
        $cfg.WallpaperPath = $txtWallpaper.Text
        $cfg.LogoPath = $txtLogo.Text
        $cfg.PrimaryText = $txtPrimaryText.Text
        $cfg.SecondaryText = $txtSecondaryText.Text
        
        # Handle custom positions
        if ($script:UseCustomLogoPosition) {
            $cfg.Position = "Custom"
            $cfg.CustomX = [int]$script:CustomLogoX
            $cfg.CustomY = [int]$script:CustomLogoY
        } else {
            $cfg.Position = $cmbPosition.SelectedItem.Content
        }
        
        # Custom text position
        if ($script:UseCustomTextPosition -and $script:CustomTextX -ge 0) {
            $cfg.CustomTextX = [int]$script:CustomTextX
            $cfg.CustomTextY = [int]$script:CustomTextY
        }
        
        # Custom logo size
        if ($script:CustomLogoW -gt 0) {
            $cfg.CustomLogoW = [int]$script:CustomLogoW
            $cfg.CustomLogoH = [int]$script:CustomLogoH
        }
        
        $cfg.Margin = [int]$txtMargin.Text
        $cfg.LogoScale = $sliderScale.Value
        $cfg.FontName = $cmbFont.SelectedItem.Content
        $cfg.FontSize = [int]$txtFontSize.Text
        $cfg.FontColor = $script:FontColorHex
        $cfg.FontBold = $chkBold.IsChecked
        $cfg.EnableShadow = $chkShadow.IsChecked
        $cfg.EnableBackdrop = $chkBackdrop.IsChecked
        $cfg.BackdropMode = $cmbBackdropMode.SelectedItem.Content
        $cfg.BackdropColor = $script:BackdropColorHex
        $cfg.BackdropOpacity = [int]$txtBackdropOpacity.Text
        $cfg.EnableSystemInfo = $chkSystemInfo.IsChecked
        $cfg.SystemInfoPosition = $cmbSystemInfoPosition.SelectedItem.Content
        
        # Custom system info position
        if ($script:UseCustomSysInfoPosition -and $script:CustomSysInfoX -ge 0) {
            $cfg.SystemInfoCustomX = [int]$script:CustomSysInfoX
            $cfg.SystemInfoCustomY = [int]$script:CustomSysInfoY
        }
        
        # System info styling
        $cfg.SystemInfoFontName = if ($cmbSysInfoFont.SelectedItem) { $cmbSysInfoFont.SelectedItem.Content } else { "Consolas" }
        $cfg.SystemInfoFontSize = [int]$txtSysInfoFontSize.Text
        $cfg.SystemInfoLabelColor = $script:SysInfoLabelColorHex
        $cfg.SystemInfoFontColor = $script:SysInfoValueColorHex
        
        # Build system info fields list
        $fields = @()
        if ($chkHostName.IsChecked) { $fields += "HostName" }
        if ($chkIPAddress.IsChecked) { $fields += "IPAddress" }
        if ($chkUserName.IsChecked) { $fields += "UserName" }
        if ($chkOSVersion.IsChecked) { $fields += "OSVersion" }
        if ($chkCPU.IsChecked) { $fields += "CPU" }
        if ($chkRAM.IsChecked) { $fields += "TotalRAM"; $fields += "FreeRAM" }
        if ($chkDisk.IsChecked) { $fields += "DiskTotal"; $fields += "DiskFree" }
        if ($chkUptime.IsChecked) { $fields += "Uptime" }
        if ($chkSerial.IsChecked) { $fields += "SerialNumber" }
        $cfg.SystemInfoFields = $fields
        
        return $cfg
    }
    
    # Calculate preview scaling
    $calculatePreviewScale = {
        if (-not $imgPreview.Source) { return }
        
        $containerWidth = $imgPreview.ActualWidth
        $containerHeight = $imgPreview.ActualHeight
        
        if ($containerWidth -le 0 -or $containerHeight -le 0) { return }
        
        $imgWidth = $script:WallpaperWidth
        $imgHeight = $script:WallpaperHeight
        
        $scaleX = $containerWidth / $imgWidth
        $scaleY = $containerHeight / $imgHeight
        $script:PreviewScale = [Math]::Min($scaleX, $scaleY)
        
        # Calculate offset for centering
        $scaledWidth = $imgWidth * $script:PreviewScale
        $scaledHeight = $imgHeight * $script:PreviewScale
        $script:PreviewOffsetX = ($containerWidth - $scaledWidth) / 2
        $script:PreviewOffsetY = ($containerHeight - $scaledHeight) / 2
    }
    
    # Convert preview coordinates to wallpaper coordinates
    $previewToWallpaper = {
        param($previewX, $previewY)
        $wpX = [int](($previewX - $script:PreviewOffsetX) / $script:PreviewScale)
        $wpY = [int](($previewY - $script:PreviewOffsetY) / $script:PreviewScale)
        return @{ X = $wpX; Y = $wpY }
    }
    
    # Convert wallpaper coordinates to preview coordinates
    $wallpaperToPreview = {
        param($wpX, $wpY)
        $prevX = ($wpX * $script:PreviewScale) + $script:PreviewOffsetX
        $prevY = ($wpY * $script:PreviewScale) + $script:PreviewOffsetY
        return @{ X = $prevX; Y = $prevY }
    }
    
    # Update drag handles positions
    $updateDragHandles = {
        & $calculatePreviewScale
        
        $hasLogo = -not [string]::IsNullOrEmpty($txtLogo.Text) -and (Test-Path $txtLogo.Text)
        $hasText = -not [string]::IsNullOrEmpty($txtPrimaryText.Text)
        
        # Position logo with actual image
        if ($hasLogo -and $script:PreviewScale -gt 0) {
            try {
                # Load logo to get dimensions and display
                $logoBitmap = New-Object System.Windows.Media.Imaging.BitmapImage
                $logoBitmap.BeginInit()
                $logoBitmap.UriSource = New-Object System.Uri($txtLogo.Text)
                $logoBitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                $logoBitmap.EndInit()
                $logoBitmap.Freeze()
                
                $script:OriginalLogoWidth = $logoBitmap.PixelWidth
                $script:OriginalLogoHeight = $logoBitmap.PixelHeight
                
                # Apply cropping to dimensions and image
                $cropL = [Math]::Max(0, [Math]::Min(49, [int]$txtCropLeft.Text))
                $cropR = [Math]::Max(0, [Math]::Min(49, [int]$txtCropRight.Text))
                $cropT = [Math]::Max(0, [Math]::Min(49, [int]$txtCropTop.Text))
                $cropB = [Math]::Max(0, [Math]::Min(49, [int]$txtCropBottom.Text))
                
                if ($cropL -gt 0 -or $cropR -gt 0 -or $cropT -gt 0 -or $cropB -gt 0) {
                    $srcW = $logoBitmap.PixelWidth
                    $srcH = $logoBitmap.PixelHeight
                    $cropX = [int]($srcW * $cropL / 100)
                    $cropY = [int]($srcH * $cropT / 100)
                    $cropW = [int]($srcW * (100 - $cropL - $cropR) / 100)
                    $cropH = [int]($srcH * (100 - $cropT - $cropB) / 100)
                    
                    $rect = New-Object System.Windows.Int32Rect($cropX, $cropY, [Math]::Max(1, $cropW), [Math]::Max(1, $cropH))
                    $croppedBitmap = New-Object System.Windows.Media.Imaging.CroppedBitmap($logoBitmap, $rect)
                    $croppedBitmap.Freeze()
                    $logoImage.Source = $croppedBitmap
                    
                    $script:OriginalLogoWidth = $cropW
                    $script:OriginalLogoHeight = $cropH
                } else {
                    $logoImage.Source = $logoBitmap
                }
                
                # Calculate logo size
                $scale = $sliderScale.Value
                if ($script:CustomLogoW -gt 0) {
                    $logoW = $script:CustomLogoW
                    $logoH = $script:CustomLogoH
                } else {
                    $logoW = $script:OriginalLogoWidth * $scale
                    $logoH = $script:OriginalLogoHeight * $scale
                }
                
                # Get positions
                $margin = [int]$txtMargin.Text
                $pos = $cmbPosition.SelectedItem.Content
                
                # Calculate logo position on wallpaper
                if ($script:UseCustomLogoPosition -and $script:CustomLogoX -ge 0) {
                    $logoX = $script:CustomLogoX
                    $logoY = $script:CustomLogoY
                } else {
                    $textHeight = 60
                    switch ($pos) {
                        "TopLeft" { $logoX = $margin; $logoY = $margin }
                        "TopCenter" { $logoX = ($script:WallpaperWidth - $logoW) / 2; $logoY = $margin }
                        "TopRight" { $logoX = $script:WallpaperWidth - $logoW - $margin; $logoY = $margin }
                        "CenterLeft" { $logoX = $margin; $logoY = ($script:WallpaperHeight - $logoH) / 2 }
                        "Center" { $logoX = ($script:WallpaperWidth - $logoW) / 2; $logoY = ($script:WallpaperHeight - $logoH) / 2 }
                        "CenterRight" { $logoX = $script:WallpaperWidth - $logoW - $margin; $logoY = ($script:WallpaperHeight - $logoH) / 2 }
                        "BottomLeft" { $logoX = $margin; $logoY = $script:WallpaperHeight - $logoH - $textHeight - $margin }
                        "BottomCenter" { $logoX = ($script:WallpaperWidth - $logoW) / 2; $logoY = $script:WallpaperHeight - $logoH - $textHeight - $margin }
                        "BottomRight" { $logoX = $script:WallpaperWidth - $logoW - $margin; $logoY = $script:WallpaperHeight - $logoH - $textHeight - $margin }
                        default { $logoX = $script:WallpaperWidth - $logoW - $margin; $logoY = $script:WallpaperHeight - $logoH - $textHeight - $margin }
                    }
                }
                
                # Position logo container in preview
                $previewPos = & $wallpaperToPreview $logoX $logoY
                $previewW = $logoW * $script:PreviewScale
                $previewH = $logoH * $script:PreviewScale
                
                [System.Windows.Controls.Canvas]::SetLeft($logoContainer, $previewPos.X)
                [System.Windows.Controls.Canvas]::SetTop($logoContainer, $previewPos.Y)
                
                # Size and position logo image
                $logoImage.Width = $previewW
                $logoImage.Height = $previewH
                [System.Windows.Controls.Canvas]::SetLeft($logoImage, 0)
                [System.Windows.Controls.Canvas]::SetTop($logoImage, 0)
                
                # Size and position selection border
                $logoSelectionBorder.Width = $previewW
                $logoSelectionBorder.Height = $previewH
                [System.Windows.Controls.Canvas]::SetLeft($logoSelectionBorder, 0)
                [System.Windows.Controls.Canvas]::SetTop($logoSelectionBorder, 0)
                
                # Position resize handles (slightly outside corners)
                [System.Windows.Controls.Canvas]::SetLeft($logoResizeNW, -6)
                [System.Windows.Controls.Canvas]::SetTop($logoResizeNW, -6)
                [System.Windows.Controls.Canvas]::SetLeft($logoResizeNE, $previewW - 6)
                [System.Windows.Controls.Canvas]::SetTop($logoResizeNE, -6)
                [System.Windows.Controls.Canvas]::SetLeft($logoResizeSW, -6)
                [System.Windows.Controls.Canvas]::SetTop($logoResizeSW, $previewH - 6)
                [System.Windows.Controls.Canvas]::SetLeft($logoResizeSE, $previewW - 6)
                [System.Windows.Controls.Canvas]::SetTop($logoResizeSE, $previewH - 6)
                
                $logoContainer.Visibility = "Visible"
                
                if ($script:UseCustomLogoPosition) {
                    $lblLogoPos.Text = "$([int]$logoX), $([int]$logoY)"
                } else {
                    $lblLogoPos.Text = "Preset"
                }
                
                # Store for text positioning
                $script:CurrentLogoX = $logoX
                $script:CurrentLogoY = $logoY
                $script:CurrentLogoW = $logoW
                $script:CurrentLogoH = $logoH
                $script:CurrentPreviewW = $previewW
                $script:CurrentPreviewH = $previewH
                
            } catch {
                $logoContainer.Visibility = "Collapsed"
                $lblLogoPos.Text = "Error"
            }
        } else {
            $logoContainer.Visibility = "Collapsed"
            $lblLogoPos.Text = "No Logo"
        }
        
        # Position text with actual content
        if ($hasText -and $script:PreviewScale -gt 0) {
            $fontSize = if ($script:CustomFontSize -gt 0) { $script:CustomFontSize } else { [int]$txtFontSize.Text }
            $previewFontSize = [Math]::Max(8, $fontSize * $script:PreviewScale * 0.7)
            
            # Configure text blocks
            $fontFamily = New-Object System.Windows.Media.FontFamily($cmbFont.SelectedItem.Content)
            $fontWeight = if ($chkBold.IsChecked) { "Bold" } else { "Normal" }
            $fontColor = [System.Windows.Media.Color]::FromRgb(
                [Convert]::ToByte($script:FontColorHex.Substring(0,2), 16),
                [Convert]::ToByte($script:FontColorHex.Substring(2,2), 16),
                [Convert]::ToByte($script:FontColorHex.Substring(4,2), 16)
            )
            
            $primaryTextBlock.Text = $txtPrimaryText.Text
            $primaryTextBlock.FontFamily = $fontFamily
            $primaryTextBlock.FontSize = $previewFontSize
            $primaryTextBlock.FontWeight = $fontWeight
            $primaryTextBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush($fontColor)
            
            $secondaryTextBlock.Text = $txtSecondaryText.Text
            $secondaryTextBlock.FontFamily = $fontFamily
            $secondaryTextBlock.FontSize = $previewFontSize * 0.85
            $secondaryTextBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush($fontColor)
            
            # Measure text for sizing
            $primaryTextBlock.Measure([System.Windows.Size]::new([Double]::PositiveInfinity, [Double]::PositiveInfinity))
            $secondaryTextBlock.Measure([System.Windows.Size]::new([Double]::PositiveInfinity, [Double]::PositiveInfinity))
            
            $textW = [Math]::Max($primaryTextBlock.DesiredSize.Width, $secondaryTextBlock.DesiredSize.Width) + 16
            $textH = $primaryTextBlock.DesiredSize.Height + $secondaryTextBlock.DesiredSize.Height + 12
            
            # Calculate text position
            if ($script:UseCustomTextPosition -and $script:CustomTextX -ge 0) {
                $textX = $script:CustomTextX
                $textY = $script:CustomTextY
            } elseif ($hasLogo) {
                $textX = $script:CurrentLogoX
                $textY = $script:CurrentLogoY + $script:CurrentLogoH + 10
            } else {
                $margin = [int]$txtMargin.Text
                $textX = $script:WallpaperWidth - ($textW / $script:PreviewScale) - $margin
                $textY = $script:WallpaperHeight - ($textH / $script:PreviewScale) - $margin
            }
            
            $previewPos = & $wallpaperToPreview $textX $textY
            
            [System.Windows.Controls.Canvas]::SetLeft($textContainer, $previewPos.X)
            [System.Windows.Controls.Canvas]::SetTop($textContainer, $previewPos.Y)
            
            # Position text elements
            $textBackground.Width = $textW
            $textBackground.Height = $textH
            [System.Windows.Controls.Canvas]::SetLeft($textBackground, 0)
            [System.Windows.Controls.Canvas]::SetTop($textBackground, 0)
            
            [System.Windows.Controls.Canvas]::SetLeft($primaryTextBlock, 8)
            [System.Windows.Controls.Canvas]::SetTop($primaryTextBlock, 4)
            [System.Windows.Controls.Canvas]::SetLeft($primaryTextEdit, 8)
            [System.Windows.Controls.Canvas]::SetTop($primaryTextEdit, 4)
            $primaryTextEdit.Width = $textW - 16
            
            [System.Windows.Controls.Canvas]::SetLeft($secondaryTextBlock, 8)
            [System.Windows.Controls.Canvas]::SetTop($secondaryTextBlock, $primaryTextBlock.DesiredSize.Height + 6)
            [System.Windows.Controls.Canvas]::SetLeft($secondaryTextEdit, 8)
            [System.Windows.Controls.Canvas]::SetTop($secondaryTextEdit, $primaryTextBlock.DesiredSize.Height + 6)
            $secondaryTextEdit.Width = $textW - 16
            
            # Selection border
            $textSelectionBorder.Width = $textW
            $textSelectionBorder.Height = $textH
            [System.Windows.Controls.Canvas]::SetLeft($textSelectionBorder, 0)
            [System.Windows.Controls.Canvas]::SetTop($textSelectionBorder, 0)
            
            # Resize handles
            [System.Windows.Controls.Canvas]::SetLeft($textResizeNW, -6)
            [System.Windows.Controls.Canvas]::SetTop($textResizeNW, -6)
            [System.Windows.Controls.Canvas]::SetLeft($textResizeNE, $textW - 6)
            [System.Windows.Controls.Canvas]::SetTop($textResizeNE, -6)
            [System.Windows.Controls.Canvas]::SetLeft($textResizeSW, -6)
            [System.Windows.Controls.Canvas]::SetTop($textResizeSW, $textH - 6)
            [System.Windows.Controls.Canvas]::SetLeft($textResizeSE, $textW - 6)
            [System.Windows.Controls.Canvas]::SetTop($textResizeSE, $textH - 6)
            
            $textContainer.Visibility = "Visible"
            
            if ($script:UseCustomTextPosition) {
                $lblTextPos.Text = "$([int]$textX), $([int]$textY)"
            } else {
                $lblTextPos.Text = "With Logo"
            }
        } else {
            $textContainer.Visibility = "Collapsed"
            $lblTextPos.Text = "No Text"
        }
        
        # Position system info overlay
        if ($chkSystemInfo.IsChecked -and $script:PreviewScale -gt 0) {
            $sysInfoFontSize = if ($script:CustomSysInfoFontSize -gt 0) { $script:CustomSysInfoFontSize } else { [int]$txtSysInfoFontSize.Text }
            $previewSysFontSize = [Math]::Max(6, $sysInfoFontSize * $script:PreviewScale * 0.7)
            
            # Build system info lines
            $sysInfoPanel.Children.Clear()
            $sysInfoFontName = if ($cmbSysInfoFont.SelectedItem) { $cmbSysInfoFont.SelectedItem.Content } else { "Consolas" }
            $sysInfoFont = New-Object System.Windows.Media.FontFamily($sysInfoFontName)
            
            $labelColor = [System.Windows.Media.Color]::FromRgb(
                [Convert]::ToByte($script:SysInfoLabelColorHex.Substring(0,2), 16),
                [Convert]::ToByte($script:SysInfoLabelColorHex.Substring(2,2), 16),
                [Convert]::ToByte($script:SysInfoLabelColorHex.Substring(4,2), 16)
            )
            $valueColor = [System.Windows.Media.Color]::FromRgb(
                [Convert]::ToByte($script:SysInfoValueColorHex.Substring(0,2), 16),
                [Convert]::ToByte($script:SysInfoValueColorHex.Substring(2,2), 16),
                [Convert]::ToByte($script:SysInfoValueColorHex.Substring(4,2), 16)
            )
            
            # Use cached system info data (fetch once, reuse)
            if (-not $script:SysInfoCache) {
                $script:SysInfoCache = @{
                    HostName = $env:COMPUTERNAME
                    UserName = $env:USERNAME
                    IPAddress = "192.168.1.100"
                    OS = "Windows 11 Pro"
                    CPU = "Intel Core i7"
                    RAM = "16 GB"
                    Disk = "256 GB Free"
                    Uptime = "2d 5h 30m"
                    Serial = "ABC123XYZ"
                }
                # Fetch actual values with error handling
                try { 
                    $ip = (Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.InterfaceAlias -notlike "*Loopback*" } | Select-Object -First 1).IPAddress
                    if ($ip) { $script:SysInfoCache.IPAddress = $ip }
                } catch { }
                try { 
                    $os = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption
                    if ($os) { $script:SysInfoCache.OS = $os }
                } catch { }
                try { 
                    $cpu = (Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1).Name
                    if ($cpu) { $script:SysInfoCache.CPU = $cpu }
                } catch { }
                try { 
                    $ramBytes = (Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue).TotalPhysicalMemory
                    if ($ramBytes) { $script:SysInfoCache.RAM = "{0} GB" -f [Math]::Round($ramBytes / 1GB) }
                } catch { }
                try { 
                    $diskFree = (Get-PSDrive C -ErrorAction SilentlyContinue).Free
                    if ($diskFree) { $script:SysInfoCache.Disk = "{0} GB Free" -f [Math]::Round($diskFree / 1GB) }
                } catch { }
                try { 
                    $bootTime = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).LastBootUpTime
                    if ($bootTime) { 
                        $u = (Get-Date) - $bootTime
                        $script:SysInfoCache.Uptime = "{0}d {1}h {2}m" -f $u.Days, $u.Hours, $u.Minutes 
                    }
                } catch { }
                try { 
                    $serial = (Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue).SerialNumber
                    if ($serial) { $script:SysInfoCache.Serial = $serial }
                } catch { }
            }
            
            # Build display lines from cache
            $sysInfoLines = @()
            if ($chkHostName.IsChecked) { $sysInfoLines += @{ Label = "Host Name:"; Value = $script:SysInfoCache.HostName } }
            if ($chkIPAddress.IsChecked) { $sysInfoLines += @{ Label = "IP Address:"; Value = $script:SysInfoCache.IPAddress } }
            if ($chkUserName.IsChecked) { $sysInfoLines += @{ Label = "User Name:"; Value = $script:SysInfoCache.UserName } }
            if ($chkOSVersion.IsChecked) { $sysInfoLines += @{ Label = "OS:"; Value = $script:SysInfoCache.OS } }
            if ($chkCPU.IsChecked) { $sysInfoLines += @{ Label = "CPU:"; Value = $script:SysInfoCache.CPU } }
            if ($chkRAM.IsChecked) { $sysInfoLines += @{ Label = "RAM:"; Value = $script:SysInfoCache.RAM } }
            if ($chkDisk.IsChecked) { $sysInfoLines += @{ Label = "Disk:"; Value = $script:SysInfoCache.Disk } }
            if ($chkUptime.IsChecked) { $sysInfoLines += @{ Label = "Uptime:"; Value = $script:SysInfoCache.Uptime } }
            if ($chkSerial.IsChecked) { $sysInfoLines += @{ Label = "Serial:"; Value = $script:SysInfoCache.Serial } }
            
            # Only show if we have lines to display
            if ($sysInfoLines.Count -gt 0) {
                foreach ($line in $sysInfoLines) {
                    $sp = New-Object System.Windows.Controls.StackPanel
                    $sp.Orientation = "Horizontal"
                    $sp.Margin = New-Object System.Windows.Thickness(0, 1, 0, 1)
                    
                    $lbl = New-Object System.Windows.Controls.TextBlock
                    $lbl.Text = $line.Label
                    $lbl.FontFamily = $sysInfoFont
                    $lbl.FontSize = $previewSysFontSize
                    $lbl.Foreground = New-Object System.Windows.Media.SolidColorBrush($labelColor)
                    $lbl.Width = 80 * $script:PreviewScale
                    
                    $val = New-Object System.Windows.Controls.TextBlock
                    $val.Text = $line.Value
                    $val.FontFamily = $sysInfoFont
                    $val.FontSize = $previewSysFontSize
                    $val.Foreground = New-Object System.Windows.Media.SolidColorBrush($valueColor)
                    $val.Margin = New-Object System.Windows.Thickness(4, 0, 0, 0)
                    
                    $sp.Children.Add($lbl)
                    $sp.Children.Add($val)
                    $sysInfoPanel.Children.Add($sp)
                }
                
                # Measure panel
                $sysInfoPanel.Measure([System.Windows.Size]::new([Double]::PositiveInfinity, [Double]::PositiveInfinity))
                $sysW = [Math]::Max(100, $sysInfoPanel.DesiredSize.Width + 20)
                $sysH = [Math]::Max(30, $sysInfoPanel.DesiredSize.Height + 16)
                
                # Calculate system info position
                $margin = 20
                $sysInfoPos = if ($cmbSystemInfoPosition.SelectedItem) { $cmbSystemInfoPosition.SelectedItem.Content } else { "TopLeft" }
                
                if ($script:UseCustomSysInfoPosition -and $script:CustomSysInfoX -ge 0) {
                    $sysX = $script:CustomSysInfoX
                    $sysY = $script:CustomSysInfoY
                } else {
                    $wpSysW = $sysW / $script:PreviewScale
                    $wpSysH = $sysH / $script:PreviewScale
                    switch ($sysInfoPos) {
                        "TopLeft" { $sysX = $margin; $sysY = $margin }
                        "TopRight" { $sysX = $script:WallpaperWidth - $wpSysW - $margin; $sysY = $margin }
                        "BottomLeft" { $sysX = $margin; $sysY = $script:WallpaperHeight - $wpSysH - $margin }
                        "BottomRight" { $sysX = $script:WallpaperWidth - $wpSysW - $margin; $sysY = $script:WallpaperHeight - $wpSysH - $margin }
                        default { $sysX = $margin; $sysY = $margin }
                    }
                }
                
                $previewPos = & $wallpaperToPreview $sysX $sysY
                
                [System.Windows.Controls.Canvas]::SetLeft($sysInfoContainer, $previewPos.X)
                [System.Windows.Controls.Canvas]::SetTop($sysInfoContainer, $previewPos.Y)
                
                # Position elements
                $sysInfoBackground.Width = $sysW
                $sysInfoBackground.Height = $sysH
                [System.Windows.Controls.Canvas]::SetLeft($sysInfoBackground, 0)
                [System.Windows.Controls.Canvas]::SetTop($sysInfoBackground, 0)
                
                [System.Windows.Controls.Canvas]::SetLeft($sysInfoPanel, 10)
                [System.Windows.Controls.Canvas]::SetTop($sysInfoPanel, 8)
                
                $sysInfoSelectionBorder.Width = $sysW
                $sysInfoSelectionBorder.Height = $sysH
                [System.Windows.Controls.Canvas]::SetLeft($sysInfoSelectionBorder, 0)
                [System.Windows.Controls.Canvas]::SetTop($sysInfoSelectionBorder, 0)
                
                # Resize handles
                [System.Windows.Controls.Canvas]::SetLeft($sysInfoResizeNW, -6)
                [System.Windows.Controls.Canvas]::SetTop($sysInfoResizeNW, -6)
                [System.Windows.Controls.Canvas]::SetLeft($sysInfoResizeNE, $sysW - 6)
                [System.Windows.Controls.Canvas]::SetTop($sysInfoResizeNE, -6)
                [System.Windows.Controls.Canvas]::SetLeft($sysInfoResizeSW, -6)
                [System.Windows.Controls.Canvas]::SetTop($sysInfoResizeSW, $sysH - 6)
                [System.Windows.Controls.Canvas]::SetLeft($sysInfoResizeSE, $sysW - 6)
                [System.Windows.Controls.Canvas]::SetTop($sysInfoResizeSE, $sysH - 6)
                
                $sysInfoContainer.Visibility = "Visible"
            } else {
                $sysInfoContainer.Visibility = "Collapsed"
            }
        } else {
            $sysInfoContainer.Visibility = "Collapsed"
        }
    }
    
    # Update preview function - shows wallpaper only, overlay handles branding elements
    $updatePreview = {
        try {
            $script:WallpaperWidth = 1920
            $script:WallpaperHeight = 1080
            
            # Just load and scale the wallpaper background (no branding)
            $wallpaperPath = $txtWallpaper.Text
            if ([string]::IsNullOrEmpty($wallpaperPath) -or -not (Test-Path $wallpaperPath)) {
                $lblStatus.Text = "No wallpaper selected"
                return
            }
            
            # Load wallpaper
            $srcImage = [System.Drawing.Image]::FromFile($wallpaperPath)
            $bitmap = New-Object System.Drawing.Bitmap($script:WallpaperWidth, $script:WallpaperHeight)
            $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
            $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
            
            # Scale to fill (cover mode)
            $srcRatio = $srcImage.Width / $srcImage.Height
            $dstRatio = $script:WallpaperWidth / $script:WallpaperHeight
            
            if ($srcRatio -gt $dstRatio) {
                $scaledHeight = $script:WallpaperHeight
                $scaledWidth = [int]($srcImage.Width * ($script:WallpaperHeight / $srcImage.Height))
                $x = -[int](($scaledWidth - $script:WallpaperWidth) / 2)
                $y = 0
            } else {
                $scaledWidth = $script:WallpaperWidth
                $scaledHeight = [int]($srcImage.Height * ($script:WallpaperWidth / $srcImage.Width))
                $x = 0
                $y = -[int](($scaledHeight - $script:WallpaperHeight) / 2)
            }
            
            $graphics.DrawImage($srcImage, $x, $y, $scaledWidth, $scaledHeight)
            $graphics.Dispose()
            $srcImage.Dispose()
            
            # Convert to WPF image
            $ms = New-Object System.IO.MemoryStream
            $bitmap.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
            $ms.Position = 0
            
            $bitmapImage = New-Object System.Windows.Media.Imaging.BitmapImage
            $bitmapImage.BeginInit()
            $bitmapImage.StreamSource = $ms
            $bitmapImage.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
            $bitmapImage.EndInit()
            $bitmapImage.Freeze()
            
            $imgPreview.Source = $bitmapImage
            $previewPlaceholder.Visibility = "Collapsed"
            
            if ($script:PreviewBitmap) { $script:PreviewBitmap.Dispose() }
            $script:PreviewBitmap = $bitmap
            
            # Update drag handles after preview is loaded
            $imgPreview.Dispatcher.BeginInvoke([Action]{
                & $updateDragHandles
            })
            
            $lblStatus.Text = "Drag frames to position - corners to resize"
        } catch {
            $lblStatus.Text = "Error: $($_.Exception.Message)"
        }
    }
    
    # Logo drag handler (on the image itself)
    $logoImage.Add_MouseLeftButtonDown({
        param($sender, $e)
        $script:IsDraggingLogo = $true
        $pos = $e.GetPosition($dragCanvas)
        $containerLeft = [System.Windows.Controls.Canvas]::GetLeft($logoContainer)
        $containerTop = [System.Windows.Controls.Canvas]::GetTop($logoContainer)
        $script:DragOffsetX = $pos.X - $containerLeft
        $script:DragOffsetY = $pos.Y - $containerTop
        $logoImage.CaptureMouse()
        $e.Handled = $true
    })
    
    $logoImage.Add_MouseMove({
        param($sender, $e)
        if ($script:IsDraggingLogo) {
            $pos = $e.GetPosition($dragCanvas)
            $newX = $pos.X - $script:DragOffsetX
            $newY = $pos.Y - $script:DragOffsetY
            
            # Clamp to valid area
            $maxX = $dragCanvas.ActualWidth - $logoImage.ActualWidth
            $maxY = $dragCanvas.ActualHeight - $logoImage.ActualHeight
            $newX = [Math]::Max(0, [Math]::Min($maxX, $newX))
            $newY = [Math]::Max(0, [Math]::Min($maxY, $newY))
            
            [System.Windows.Controls.Canvas]::SetLeft($logoContainer, $newX)
            [System.Windows.Controls.Canvas]::SetTop($logoContainer, $newY)
            
            # Convert to wallpaper coordinates
            $wpCoords = & $previewToWallpaper $newX $newY
            $script:CustomLogoX = [Math]::Max(0, $wpCoords.X)
            $script:CustomLogoY = [Math]::Max(0, $wpCoords.Y)
            $script:UseCustomLogoPosition = $true
            
            $lblLogoPos.Text = "$($script:CustomLogoX), $($script:CustomLogoY)"
            $lblStatus.Text = "Logo: $($script:CustomLogoX), $($script:CustomLogoY)"
        }
    })
    
    $logoImage.Add_MouseLeftButtonUp({
        param($sender, $e)
        if ($script:IsDraggingLogo) {
            $script:IsDraggingLogo = $false
            $logoImage.ReleaseMouseCapture()
            foreach ($item in $cmbPosition.Items) {
                if ($item.Content -eq "Custom") { $cmbPosition.SelectedItem = $item; break }
            }
            & $updatePreview
        }
    })
    
    # Logo resize handlers
    $setupLogoResize = {
        param($handle, $corner)
        
        $handle.Add_MouseLeftButtonDown({
            param($s, $e)
            $script:IsResizingLogo = $true
            $script:ResizeCorner = $corner
            $pos = $e.GetPosition($dragCanvas)
            $script:ResizeStartX = $pos.X
            $script:ResizeStartY = $pos.Y
            $script:ResizeStartW = $logoImage.Width
            $script:ResizeStartH = $logoImage.Height
            $s.CaptureMouse()
            $e.Handled = $true
        }.GetNewClosure())
        
        $handle.Add_MouseMove({
            param($s, $e)
            if ($script:IsResizingLogo) {
                $pos = $e.GetPosition($dragCanvas)
                $deltaX = $pos.X - $script:ResizeStartX
                $deltaY = $pos.Y - $script:ResizeStartY
                
                # Maintain aspect ratio
                $aspect = $script:OriginalLogoHeight / [Math]::Max(1, $script:OriginalLogoWidth)
                $delta = [Math]::Max($deltaX, $deltaY)
                if ($corner -eq "NW" -or $corner -eq "SW") { $delta = -$deltaX }
                if ($corner -eq "NE" -or $corner -eq "SE") { $delta = $deltaX }
                
                $newW = [Math]::Max(20, $script:ResizeStartW + $delta)
                $newH = $newW * $aspect
                
                # Update logo image and selection border size
                $logoImage.Width = $newW
                $logoImage.Height = $newH
                $logoSelectionBorder.Width = $newW
                $logoSelectionBorder.Height = $newH
                
                # Update resize handles positions
                [System.Windows.Controls.Canvas]::SetLeft($logoResizeNE, $newW - 6)
                [System.Windows.Controls.Canvas]::SetLeft($logoResizeSE, $newW - 6)
                [System.Windows.Controls.Canvas]::SetTop($logoResizeSW, $newH - 6)
                [System.Windows.Controls.Canvas]::SetTop($logoResizeSE, $newH - 6)
                
                # Calculate wallpaper size
                $wpW = $newW / $script:PreviewScale
                $wpH = $newH / $script:PreviewScale
                $lblStatus.Text = "Logo size: $([int]$wpW) x $([int]$wpH)"
            }
        }.GetNewClosure())
        
        $handle.Add_MouseLeftButtonUp({
            param($s, $e)
            if ($script:IsResizingLogo) {
                $script:IsResizingLogo = $false
                $s.ReleaseMouseCapture()
                
                # Store custom size
                $script:CustomLogoW = $logoImage.Width / $script:PreviewScale
                $script:CustomLogoH = $logoImage.Height / $script:PreviewScale
                
                # Update scale slider to match
                $newScale = $script:CustomLogoW / [Math]::Max(1, $script:OriginalLogoWidth)
                $sliderScale.Value = [Math]::Max(0.1, [Math]::Min(3, $newScale))
                
                & $updatePreview
            }
        }.GetNewClosure())
    }
    
    & $setupLogoResize $logoResizeNW "NW"
    & $setupLogoResize $logoResizeNE "NE"
    & $setupLogoResize $logoResizeSW "SW"
    & $setupLogoResize $logoResizeSE "SE"
    
    # Text drag handlers - use background for dragging
    $textBackground.Add_MouseLeftButtonDown({
        param($sender, $e)
        $script:IsDraggingText = $true
        $pos = $e.GetPosition($dragCanvas)
        $containerLeft = [System.Windows.Controls.Canvas]::GetLeft($textContainer)
        $containerTop = [System.Windows.Controls.Canvas]::GetTop($textContainer)
        $script:DragOffsetX = $pos.X - $containerLeft
        $script:DragOffsetY = $pos.Y - $containerTop
        $textBackground.CaptureMouse()
        $e.Handled = $true
    })
    
    $textBackground.Add_MouseMove({
        param($sender, $e)
        if ($script:IsDraggingText) {
            $pos = $e.GetPosition($dragCanvas)
            $newX = $pos.X - $script:DragOffsetX
            $newY = $pos.Y - $script:DragOffsetY
            
            $maxX = $dragCanvas.ActualWidth - $textBackground.Width
            $maxY = $dragCanvas.ActualHeight - $textBackground.Height
            $newX = [Math]::Max(0, [Math]::Min($maxX, $newX))
            $newY = [Math]::Max(0, [Math]::Min($maxY, $newY))
            
            [System.Windows.Controls.Canvas]::SetLeft($textContainer, $newX)
            [System.Windows.Controls.Canvas]::SetTop($textContainer, $newY)
            
            $wpCoords = & $previewToWallpaper $newX $newY
            $script:CustomTextX = [Math]::Max(0, $wpCoords.X)
            $script:CustomTextY = [Math]::Max(0, $wpCoords.Y)
            $script:UseCustomTextPosition = $true
            
            $lblTextPos.Text = "$($script:CustomTextX), $($script:CustomTextY)"
            $lblStatus.Text = "Text: $($script:CustomTextX), $($script:CustomTextY)"
        }
    })
    
    $textBackground.Add_MouseLeftButtonUp({
        param($sender, $e)
        if ($script:IsDraggingText) {
            $script:IsDraggingText = $false
            $textBackground.ReleaseMouseCapture()
            & $updatePreview
        }
    })
    
    # Click-to-edit primary text
    $primaryTextBlock.Add_MouseLeftButtonDown({
        param($sender, $e)
        $primaryTextBlock.Visibility = "Collapsed"
        $primaryTextEdit.Text = $primaryTextBlock.Text
        $primaryTextEdit.Visibility = "Visible"
        $primaryTextEdit.Focus()
        $primaryTextEdit.SelectAll()
        $e.Handled = $true
    })
    
    $primaryTextEdit.Add_LostFocus({
        $txtPrimaryText.Text = $primaryTextEdit.Text
        $primaryTextBlock.Text = $primaryTextEdit.Text
        $primaryTextEdit.Visibility = "Collapsed"
        $primaryTextBlock.Visibility = "Visible"
        $primaryTextEdit.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles })
    })
    
    $primaryTextEdit.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq "Return" -or $e.Key -eq "Escape") {
            if ($e.Key -eq "Return") {
                $txtPrimaryText.Text = $primaryTextEdit.Text
                $primaryTextBlock.Text = $primaryTextEdit.Text
            }
            $primaryTextEdit.Visibility = "Collapsed"
            $primaryTextBlock.Visibility = "Visible"
            $primaryTextEdit.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles })
            $e.Handled = $true
        }
    })
    
    # Click-to-edit secondary text
    $secondaryTextBlock.Add_MouseLeftButtonDown({
        param($sender, $e)
        $secondaryTextBlock.Visibility = "Collapsed"
        $secondaryTextEdit.Text = $secondaryTextBlock.Text
        $secondaryTextEdit.Visibility = "Visible"
        $secondaryTextEdit.Focus()
        $secondaryTextEdit.SelectAll()
        $e.Handled = $true
    })
    
    $secondaryTextEdit.Add_LostFocus({
        $txtSecondaryText.Text = $secondaryTextEdit.Text
        $secondaryTextBlock.Text = $secondaryTextEdit.Text
        $secondaryTextEdit.Visibility = "Collapsed"
        $secondaryTextBlock.Visibility = "Visible"
        $secondaryTextEdit.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles })
    })
    
    $secondaryTextEdit.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq "Return" -or $e.Key -eq "Escape") {
            if ($e.Key -eq "Return") {
                $txtSecondaryText.Text = $secondaryTextEdit.Text
                $secondaryTextBlock.Text = $secondaryTextEdit.Text
            }
            $secondaryTextEdit.Visibility = "Collapsed"
            $secondaryTextBlock.Visibility = "Visible"
            $secondaryTextEdit.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles })
            $e.Handled = $true
        }
    })
    
    # Text resize handlers (changes font size)
    $setupTextResize = {
        param($handle, $corner)
        
        $handle.Add_MouseLeftButtonDown({
            param($s, $e)
            $script:IsResizingText = $true
            $script:ResizeCorner = $corner
            $pos = $e.GetPosition($dragCanvas)
            $script:ResizeStartX = $pos.X
            $script:ResizeStartY = $pos.Y
            $script:ResizeStartW = [int]$txtFontSize.Text
            $s.CaptureMouse()
            $e.Handled = $true
        }.GetNewClosure())
        
        $handle.Add_MouseMove({
            param($s, $e)
            if ($script:IsResizingText) {
                $pos = $e.GetPosition($dragCanvas)
                $deltaX = $pos.X - $script:ResizeStartX
                $deltaY = $pos.Y - $script:ResizeStartY
                
                $delta = if ($corner -eq "SE" -or $corner -eq "NE") { $deltaX } else { -$deltaX }
                $sizeChange = [int]($delta / 5)
                
                $newSize = [Math]::Max(8, [Math]::Min(72, $script:ResizeStartW + $sizeChange))
                $script:CustomFontSize = $newSize
                $txtFontSize.Text = $newSize.ToString()
                
                $lblStatus.Text = "Font size: $newSize pt"
            }
        }.GetNewClosure())
        
        $handle.Add_MouseLeftButtonUp({
            param($s, $e)
            if ($script:IsResizingText) {
                $script:IsResizingText = $false
                $s.ReleaseMouseCapture()
                & $updatePreview
            }
        }.GetNewClosure())
    }
    
    & $setupTextResize $textResizeNW "NW"
    & $setupTextResize $textResizeNE "NE"
    & $setupTextResize $textResizeSW "SW"
    & $setupTextResize $textResizeSE "SE"
    
    # System Info drag handlers - use background to avoid conflicts with panel rebuild
    $sysInfoBackground.Add_MouseLeftButtonDown({
        param($sender, $e)
        $script:IsDraggingSysInfo = $true
        $pos = $e.GetPosition($dragCanvas)
        $containerLeft = [System.Windows.Controls.Canvas]::GetLeft($sysInfoContainer)
        $containerTop = [System.Windows.Controls.Canvas]::GetTop($sysInfoContainer)
        $script:DragOffsetX = $pos.X - $containerLeft
        $script:DragOffsetY = $pos.Y - $containerTop
        $sysInfoBackground.CaptureMouse()
        $e.Handled = $true
    })
    
    $sysInfoBackground.Add_MouseMove({
        param($sender, $e)
        if ($script:IsDraggingSysInfo) {
            $pos = $e.GetPosition($dragCanvas)
            $newX = $pos.X - $script:DragOffsetX
            $newY = $pos.Y - $script:DragOffsetY
            
            $maxX = $dragCanvas.ActualWidth - $sysInfoBackground.Width
            $maxY = $dragCanvas.ActualHeight - $sysInfoBackground.Height
            $newX = [Math]::Max(0, [Math]::Min($maxX, $newX))
            $newY = [Math]::Max(0, [Math]::Min($maxY, $newY))
            
            [System.Windows.Controls.Canvas]::SetLeft($sysInfoContainer, $newX)
            [System.Windows.Controls.Canvas]::SetTop($sysInfoContainer, $newY)
            
            $wpCoords = & $previewToWallpaper $newX $newY
            $script:CustomSysInfoX = [Math]::Max(0, $wpCoords.X)
            $script:CustomSysInfoY = [Math]::Max(0, $wpCoords.Y)
            $script:UseCustomSysInfoPosition = $true
            
            $lblStatus.Text = "System Info: $($script:CustomSysInfoX), $($script:CustomSysInfoY)"
        }
    })
    
    $sysInfoBackground.Add_MouseLeftButtonUp({
        param($sender, $e)
        if ($script:IsDraggingSysInfo) {
            $script:IsDraggingSysInfo = $false
            $sysInfoBackground.ReleaseMouseCapture()
            # Don't call updateDragHandles here - just keep current position
            $lblStatus.Text = "System Info positioned at $($script:CustomSysInfoX), $($script:CustomSysInfoY)"
        }
    })
    
    # System Info resize handlers (changes font size)
    $setupSysInfoResize = {
        param($handle, $corner)
        
        $handle.Add_MouseLeftButtonDown({
            param($s, $e)
            $script:IsResizingSysInfo = $true
            $script:ResizeCorner = $corner
            $pos = $e.GetPosition($dragCanvas)
            $script:ResizeStartX = $pos.X
            $script:ResizeStartY = $pos.Y
            $script:ResizeStartW = [int]$txtSysInfoFontSize.Text
            $s.CaptureMouse()
            $e.Handled = $true
        }.GetNewClosure())
        
        $handle.Add_MouseMove({
            param($s, $e)
            if ($script:IsResizingSysInfo) {
                $pos = $e.GetPosition($dragCanvas)
                $deltaX = $pos.X - $script:ResizeStartX
                $deltaY = $pos.Y - $script:ResizeStartY
                
                $delta = if ($corner -eq "SE" -or $corner -eq "NE") { $deltaX } else { -$deltaX }
                $sizeChange = [int]($delta / 5)
                
                $newSize = [Math]::Max(8, [Math]::Min(24, $script:ResizeStartW + $sizeChange))
                $script:CustomSysInfoFontSize = $newSize
                $txtSysInfoFontSize.Text = $newSize.ToString()
                
                $lblStatus.Text = "System Info font: $newSize pt"
            }
        }.GetNewClosure())
        
        $handle.Add_MouseLeftButtonUp({
            param($s, $e)
            if ($script:IsResizingSysInfo) {
                $script:IsResizingSysInfo = $false
                $s.ReleaseMouseCapture()
                # Schedule update after event completes to avoid reentrancy
                $s.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles })
            }
        }.GetNewClosure())
    }
    
    & $setupSysInfoResize $sysInfoResizeNW "NW"
    & $setupSysInfoResize $sysInfoResizeNE "NE"
    & $setupSysInfoResize $sysInfoResizeSW "SW"
    & $setupSysInfoResize $sysInfoResizeSE "SE"
    
    # System info checkbox changes - refresh preview (deferred to avoid reentrancy)
    $chkSystemInfo.Add_Checked({ $chkSystemInfo.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles }) })
    $chkSystemInfo.Add_Unchecked({ $chkSystemInfo.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles }) })
    $chkHostName.Add_Click({ $chkHostName.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles }) })
    $chkIPAddress.Add_Click({ $chkIPAddress.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles }) })
    $chkUserName.Add_Click({ $chkUserName.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles }) })
    $chkOSVersion.Add_Click({ $chkOSVersion.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles }) })
    $chkCPU.Add_Click({ $chkCPU.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles }) })
    $chkRAM.Add_Click({ $chkRAM.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles }) })
    $chkDisk.Add_Click({ $chkDisk.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles }) })
    $chkUptime.Add_Click({ $chkUptime.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles }) })
    $chkSerial.Add_Click({ $chkSerial.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles }) })
    $cmbSystemInfoPosition.Add_SelectionChanged({ $cmbSystemInfoPosition.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles }) })
    $cmbSysInfoFont.Add_SelectionChanged({ $cmbSysInfoFont.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles }) })
    $txtSysInfoFontSize.Add_LostFocus({ $txtSysInfoFontSize.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles }) })
    
    # System info color pickers
    $btnSysInfoLabelColor.Add_MouseLeftButtonDown({
        $colorDialog = New-Object System.Windows.Forms.ColorDialog
        $colorDialog.FullOpen = $true
        if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:SysInfoLabelColorHex = "{0:X2}{1:X2}{2:X2}" -f $colorDialog.Color.R, $colorDialog.Color.G, $colorDialog.Color.B
            $btnSysInfoLabelColor.Background = New-Object System.Windows.Media.SolidColorBrush(
                [System.Windows.Media.Color]::FromRgb($colorDialog.Color.R, $colorDialog.Color.G, $colorDialog.Color.B)
            )
            & $updateDragHandles
        }
    })
    
    $btnSysInfoValueColor.Add_MouseLeftButtonDown({
        $colorDialog = New-Object System.Windows.Forms.ColorDialog
        $colorDialog.FullOpen = $true
        if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:SysInfoValueColorHex = "{0:X2}{1:X2}{2:X2}" -f $colorDialog.Color.R, $colorDialog.Color.G, $colorDialog.Color.B
            $btnSysInfoValueColor.Background = New-Object System.Windows.Media.SolidColorBrush(
                [System.Windows.Media.Color]::FromRgb($colorDialog.Color.R, $colorDialog.Color.G, $colorDialog.Color.B)
            )
            & $updateDragHandles
        }
    })
    
    # Reset positions button
    $btnResetPositions.Add_Click({
        $script:UseCustomLogoPosition = $false
        $script:UseCustomTextPosition = $false
        $script:UseCustomSysInfoPosition = $false
        $script:CustomLogoX = -1
        $script:CustomLogoY = -1
        $script:CustomLogoW = -1
        $script:CustomLogoH = -1
        $script:CustomTextX = -1
        $script:CustomTextY = -1
        $script:CustomFontSize = -1
        $script:CustomSysInfoX = -1
        $script:CustomSysInfoY = -1
        $script:CustomSysInfoFontSize = -1
        $lblLogoPos.Text = "Preset"
        $lblTextPos.Text = "With Logo"
        & $updatePreview
    })
    
    # Reset crop button
    $btnResetCrop.Add_Click({
        $txtCropLeft.Text = "0"
        $txtCropRight.Text = "0"
        $txtCropTop.Text = "0"
        $txtCropBottom.Text = "0"
        & $updatePreview
    })
    
    # Crop value changes
    $txtCropLeft.Add_LostFocus({ $txtCropLeft.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles }) })
    $txtCropRight.Add_LostFocus({ $txtCropRight.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles }) })
    $txtCropTop.Add_LostFocus({ $txtCropTop.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles }) })
    $txtCropBottom.Add_LostFocus({ $txtCropBottom.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles }) })
    
    # Text styling changes - refresh preview
    $cmbFont.Add_SelectionChanged({ $cmbFont.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles }) })
    $txtFontSize.Add_LostFocus({ $txtFontSize.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles }) })
    $chkBold.Add_Click({ $chkBold.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles }) })
    $chkShadow.Add_Click({ $chkShadow.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles }) })
    # Don't update on every keystroke - wait for LostFocus instead
    $txtPrimaryText.Add_LostFocus({ $txtPrimaryText.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles }) })
    $txtSecondaryText.Add_LostFocus({ $txtSecondaryText.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles }) })
    
    # Update handles when window resizes
    $imgPreview.Add_SizeChanged({
        $imgPreview.Dispatcher.BeginInvoke([Action]{ & $updateDragHandles })
    })
    
    # Event handlers
    $btnBrowseWallpaper.Add_Click({
        $dialog = New-Object Microsoft.Win32.OpenFileDialog
        $dialog.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp"
        if ($dialog.ShowDialog()) {
            $txtWallpaper.Text = $dialog.FileName
            & $updatePreview
        }
    })
    
    $btnBrowseLogo.Add_Click({
        $dialog = New-Object Microsoft.Win32.OpenFileDialog
        $dialog.Filter = "Image Files|*.png;*.jpg;*.jpeg;*.bmp"
        if ($dialog.ShowDialog()) {
            $txtLogo.Text = $dialog.FileName
            # Reset custom size and crop when new logo loaded
            $script:CustomLogoW = -1
            $script:CustomLogoH = -1
            $txtCropLeft.Text = "0"
            $txtCropRight.Text = "0"
            $txtCropTop.Text = "0"
            $txtCropBottom.Text = "0"
            $sliderScale.Value = 1.0
            & $updatePreview
        }
    })
    
    $sliderScale.Add_ValueChanged({
        $lblScaleValue.Text = "{0:N1}x" -f $sliderScale.Value
        # Reset custom size when slider is manually adjusted
        if (-not $script:IsResizingLogo) {
            $script:CustomLogoW = -1
            $script:CustomLogoH = -1
        }
    })
    
    $btnFontColor.Add_MouseLeftButtonDown({
        $colorDialog = New-Object System.Windows.Forms.ColorDialog
        $colorDialog.FullOpen = $true
        if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:FontColorHex = "{0:X2}{1:X2}{2:X2}" -f $colorDialog.Color.R, $colorDialog.Color.G, $colorDialog.Color.B
            $btnFontColor.Background = New-Object System.Windows.Media.SolidColorBrush(
                [System.Windows.Media.Color]::FromRgb($colorDialog.Color.R, $colorDialog.Color.G, $colorDialog.Color.B)
            )
            & $updateDragHandles
        }
    })
    
    $btnBackdropColor.Add_MouseLeftButtonDown({
        $colorDialog = New-Object System.Windows.Forms.ColorDialog
        $colorDialog.FullOpen = $true
        if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:BackdropColorHex = "{0:X2}{1:X2}{2:X2}" -f $colorDialog.Color.R, $colorDialog.Color.G, $colorDialog.Color.B
            $btnBackdropColor.Background = New-Object System.Windows.Media.SolidColorBrush(
                [System.Windows.Media.Color]::FromRgb($colorDialog.Color.R, $colorDialog.Color.G, $colorDialog.Color.B)
            )
        }
    })
    
    $btnRefresh.Add_Click({ & $updatePreview })
    
    $btnSaveImage.Add_Click({
        $dialog = New-Object Microsoft.Win32.SaveFileDialog
        $dialog.Filter = "PNG Image|*.png|JPEG Image|*.jpg"
        $dialog.FileName = "WallBrand_Wallpaper"
        if ($dialog.ShowDialog()) {
            $cfg = & $getConfigFromUI
            $primary = [System.Windows.Forms.Screen]::PrimaryScreen
            $bitmap = Render-BrandedWallpaper -Config $cfg -TargetWidth $primary.Bounds.Width -TargetHeight $primary.Bounds.Height
            
            $format = if ($dialog.FileName -match '\.jpg$') { [System.Drawing.Imaging.ImageFormat]::Jpeg } else { [System.Drawing.Imaging.ImageFormat]::Png }
            $bitmap.Save($dialog.FileName, $format)
            $bitmap.Dispose()
            
            $lblStatus.Text = "Saved: $($dialog.FileName)"
        }
    })
    
    $btnApply.Add_Click({
        try {
            $cfg = & $getConfigFromUI
            $primary = [System.Windows.Forms.Screen]::PrimaryScreen
            $bitmap = Render-BrandedWallpaper -Config $cfg -TargetWidth $primary.Bounds.Width -TargetHeight $primary.Bounds.Height
            
            $tempPath = Join-Path $env:LOCALAPPDATA "WallBrand\wallpaper.png"
            $tempDir = Split-Path $tempPath
            if (-not (Test-Path $tempDir)) { New-Item -Path $tempDir -ItemType Directory -Force | Out-Null }
            
            $bitmap.Save($tempPath, [System.Drawing.Imaging.ImageFormat]::Png)
            $bitmap.Dispose()
            
            Set-DesktopWallpaper -Path $tempPath -Style "Fill"
            
            $lblStatus.Text = "Wallpaper applied successfully!"
            $lblStatus.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::LightGreen)
        } catch {
            $lblStatus.Text = "Error: $($_.Exception.Message)"
            $lblStatus.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Salmon)
        }
    })
    
    $btnSaveConfig.Add_Click({
        $dialog = New-Object Microsoft.Win32.SaveFileDialog
        $dialog.Filter = "JSON Config|*.json"
        $dialog.FileName = "wallbrand-config"
        if ($dialog.ShowDialog()) {
            $cfg = & $getConfigFromUI
            $cfg.ToJson() | Out-File -FilePath $dialog.FileName -Encoding UTF8
            $lblStatus.Text = "Config saved: $($dialog.FileName)"
        }
    })
    
    $btnLoadConfig.Add_Click({
        $dialog = New-Object Microsoft.Win32.OpenFileDialog
        $dialog.Filter = "JSON Config|*.json"
        if ($dialog.ShowDialog()) {
            try {
                $json = Get-Content $dialog.FileName -Raw
                $script:Config = [WallBrandConfig]::FromJson($json)
                
                # Update UI from config
                $txtWallpaper.Text = $script:Config.WallpaperPath
                $txtLogo.Text = $script:Config.LogoPath
                $txtPrimaryText.Text = $script:Config.PrimaryText
                $txtSecondaryText.Text = $script:Config.SecondaryText
                $txtMargin.Text = $script:Config.Margin.ToString()
                $sliderScale.Value = $script:Config.LogoScale
                $txtFontSize.Text = $script:Config.FontSize.ToString()
                $chkShadow.IsChecked = $script:Config.EnableShadow
                $chkBold.IsChecked = $script:Config.FontBold
                $chkBackdrop.IsChecked = $script:Config.EnableBackdrop
                $txtBackdropOpacity.Text = $script:Config.BackdropOpacity.ToString()
                $chkSystemInfo.IsChecked = $script:Config.EnableSystemInfo
                
                $script:FontColorHex = $script:Config.FontColor
                $script:BackdropColorHex = $script:Config.BackdropColor
                
                & $updatePreview
                $lblStatus.Text = "Config loaded: $($dialog.FileName)"
            } catch {
                $lblStatus.Text = "Error loading config: $($_.Exception.Message)"
            }
        }
    })
    
    $btnExportGPO.Add_Click({
        $dialog = New-Object Microsoft.Win32.SaveFileDialog
        $dialog.Filter = "PowerShell Script|*.ps1"
        $dialog.FileName = "WallBrand-GPO-Deploy"
        if ($dialog.ShowDialog()) {
            $cfg = & $getConfigFromUI
            Export-GPODeploymentScript -OutputPath $dialog.FileName -Config $cfg -WallpaperNetworkPath "\\SERVER\Share\wallpaper.png"
            $lblStatus.Text = "GPO script exported: $($dialog.FileName)"
        }
    })
    
    $btnExportIntune.Add_Click({
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = "Select folder for Intune package"
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $cfg = & $getConfigFromUI
            
            # First save the wallpaper
            $primary = [System.Windows.Forms.Screen]::PrimaryScreen
            $bitmap = Render-BrandedWallpaper -Config $cfg -TargetWidth $primary.Bounds.Width -TargetHeight $primary.Bounds.Height
            $tempWallpaper = Join-Path $env:TEMP "wallbrand_export.png"
            $bitmap.Save($tempWallpaper, [System.Drawing.Imaging.ImageFormat]::Png)
            $bitmap.Dispose()
            
            Export-IntunePackage -OutputFolder $dialog.SelectedPath -Config $cfg -WallpaperPath $tempWallpaper
            Remove-Item $tempWallpaper -ErrorAction SilentlyContinue
            
            $lblStatus.Text = "Intune package exported: $($dialog.SelectedPath)"
        }
    })
    
    # Scheduled Task constants
    $taskName = "WallBrandPro-AutoUpdate"
    $configFolder = Join-Path $env:LOCALAPPDATA "WallBrand"
    $autoConfigPath = Join-Path $configFolder "auto-update-config.json"
    $scriptPath = $PSCommandPath
    if (-not $scriptPath) { $scriptPath = $MyInvocation.MyCommand.Path }
    
    # Function to check if task exists
    $checkTaskStatus = {
        try {
            $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
            if ($task) {
                $state = $task.State
                $nextRun = (Get-ScheduledTaskInfo -TaskName $taskName -ErrorAction SilentlyContinue).NextRunTime
                if ($nextRun) {
                    $lblTaskStatus.Text = "Status: Active ($state) - Next: $($nextRun.ToString('g'))"
                } else {
                    $lblTaskStatus.Text = "Status: Active ($state)"
                }
                $lblTaskStatus.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0x3f, 0xb9, 0x50))
            } else {
                $lblTaskStatus.Text = "Status: No scheduled task"
                $lblTaskStatus.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0x6e, 0x76, 0x81))
            }
        } catch {
            $lblTaskStatus.Text = "Status: Unknown"
        }
    }
    
    # Check task status on load
    & $checkTaskStatus
    
    # Create scheduled task button
    $btnCreateTask.Add_Click({
        try {
            # Ensure config folder exists
            if (-not (Test-Path $configFolder)) { 
                New-Item -Path $configFolder -ItemType Directory -Force | Out-Null 
            }
            
            # Save current config for the scheduled task to use
            $cfg = & $getConfigFromUI
            $cfg.ToJson() | Out-File -FilePath $autoConfigPath -Encoding UTF8
            
            # Parse frequency
            $frequencyText = $cmbUpdateFrequency.SelectedItem.Content
            $intervalMinutes = switch ($frequencyText) {
                "Every 5 minutes" { 5 }
                "Every 15 minutes" { 15 }
                "Every 30 minutes" { 30 }
                "Every hour" { 60 }
                "Every 4 hours" { 240 }
                "Daily" { 1440 }
                default { 15 }
            }
            
            # Remove existing task if present
            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
            
            # Build triggers array
            $triggers = @()
            
            # Repetition trigger (main schedule)
            if ($intervalMinutes -lt 1440) {
                # For intervals less than daily, use repetition
                $repetitionDuration = [TimeSpan]::FromDays(1)
                $repetitionInterval = [TimeSpan]::FromMinutes($intervalMinutes)
                $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $repetitionInterval -RepetitionDuration $repetitionDuration
                $triggers += $trigger
            } else {
                # Daily trigger
                $triggers += New-ScheduledTaskTrigger -Daily -At "8:00AM"
            }
            
            # Additional triggers
            if ($chkTriggerLogon.IsChecked) {
                $triggers += New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
            }
            
            # Build the action - run this script with the saved config
            $actionArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -ConfigPath `"$autoConfigPath`" -Apply"
            if ($chkRunHidden.IsChecked) {
                $actionArgs = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`" -ConfigPath `"$autoConfigPath`" -Apply"
            }
            $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $actionArgs
            
            # Settings
            $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 5)
            
            # Principal (run level)
            if ($chkRunAsAdmin.IsChecked) {
                $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest
            } else {
                $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited
            }
            
            # Register the task
            Register-ScheduledTask -TaskName $taskName -Trigger $triggers -Action $action -Settings $settings -Principal $principal -Description "WallBrand Pro - Auto-updates wallpaper with current system information" | Out-Null
            
            # Add unlock trigger via XML modification if requested (not directly supported by New-ScheduledTaskTrigger)
            if ($chkTriggerUnlock.IsChecked -or $chkTriggerResume.IsChecked) {
                $task = Get-ScheduledTask -TaskName $taskName
                $taskXml = Export-ScheduledTask -TaskName $taskName
                $xml = [xml]$taskXml
                $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
                $ns.AddNamespace("t", "http://schemas.microsoft.com/windows/2004/02/mit/task")
                $triggersNode = $xml.SelectSingleNode("//t:Triggers", $ns)
                
                if ($chkTriggerUnlock.IsChecked) {
                    $unlockTrigger = $xml.CreateElement("SessionStateChangeTrigger", "http://schemas.microsoft.com/windows/2004/02/mit/task")
                    $unlockTrigger.InnerXml = "<StateChange>SessionUnlock</StateChange><UserId>$env:USERDOMAIN\$env:USERNAME</UserId>"
                    $triggersNode.AppendChild($unlockTrigger) | Out-Null
                }
                
                if ($chkTriggerResume.IsChecked) {
                    # Add event trigger for resume from sleep (Event ID 1 from Power-Troubleshooter)
                    $resumeTrigger = $xml.CreateElement("EventTrigger", "http://schemas.microsoft.com/windows/2004/02/mit/task")
                    $resumeTrigger.InnerXml = "<Subscription>&lt;QueryList&gt;&lt;Query Id='0' Path='System'&gt;&lt;Select Path='System'&gt;*[System[Provider[@Name='Microsoft-Windows-Power-Troubleshooter'] and EventID=1]]&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>"
                    $triggersNode.AppendChild($resumeTrigger) | Out-Null
                }
                
                # Re-register with updated XML
                Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
                Register-ScheduledTask -TaskName $taskName -Xml $xml.OuterXml | Out-Null
            }
            
            & $checkTaskStatus
            $lblStatus.Text = "Scheduled task created successfully!"
            $lblStatus.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::LightGreen)
            
        } catch {
            $lblStatus.Text = "Error creating task: $($_.Exception.Message)"
            $lblStatus.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Salmon)
        }
    })
    
    # Remove scheduled task button
    $btnRemoveTask.Add_Click({
        try {
            $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
            if ($task) {
                Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
                
                # Also remove the auto-config file
                if (Test-Path $autoConfigPath) {
                    Remove-Item $autoConfigPath -Force -ErrorAction SilentlyContinue
                }
                
                & $checkTaskStatus
                $lblStatus.Text = "Scheduled task removed"
            } else {
                $lblStatus.Text = "No scheduled task to remove"
            }
        } catch {
            $lblStatus.Text = "Error removing task: $($_.Exception.Message)"
            $lblStatus.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Salmon)
        }
    })
    
    # Refresh now button - manually trigger wallpaper update
    $btnRefreshNow.Add_Click({
        try {
            # Clear system info cache to get fresh data
            $script:SysInfoCache = $null
            
            $cfg = & $getConfigFromUI
            $primary = [System.Windows.Forms.Screen]::PrimaryScreen
            $bitmap = Render-BrandedWallpaper -Config $cfg -TargetWidth $primary.Bounds.Width -TargetHeight $primary.Bounds.Height
            
            $tempPath = Join-Path $env:LOCALAPPDATA "WallBrand\wallpaper.png"
            $tempDir = Split-Path $tempPath
            if (-not (Test-Path $tempDir)) { New-Item -Path $tempDir -ItemType Directory -Force | Out-Null }
            
            $bitmap.Save($tempPath, [System.Drawing.Imaging.ImageFormat]::Png)
            $bitmap.Dispose()
            
            Set-DesktopWallpaper -Path $tempPath -Style "Fill"
            
            # Refresh preview
            & $updatePreview
            
            $lblStatus.Text = "Wallpaper refreshed with current system info!"
            $lblStatus.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::LightGreen)
        } catch {
            $lblStatus.Text = "Error: $($_.Exception.Message)"
            $lblStatus.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Salmon)
        }
    })
    
    # Initial preview
    & $updatePreview
    
    # Show window
    $window.ShowDialog() | Out-Null
    
    # Cleanup
    if ($script:PreviewBitmap) { $script:PreviewBitmap.Dispose() }
}

# ============================================================================
# MAIN ENTRY POINT
# ============================================================================

# Load config if specified
$config = [WallBrandConfig]::new()
if ($ConfigPath -and (Test-Path $ConfigPath)) {
    $json = Get-Content $ConfigPath -Raw
    $config = [WallBrandConfig]::FromJson($json)
}

# Handle export modes
if ($ExportGPO) {
    Export-GPODeploymentScript -OutputPath $ExportGPO -Config $config -WallpaperNetworkPath "\\SERVER\Share\wallpaper.png"
    Write-Host "GPO deployment script exported to: $ExportGPO" -ForegroundColor Green
    exit 0
}

if ($ExportIntune) {
    $primary = [System.Windows.Forms.Screen]::PrimaryScreen
    $bitmap = Render-BrandedWallpaper -Config $config -TargetWidth $primary.Bounds.Width -TargetHeight $primary.Bounds.Height
    $tempWallpaper = Join-Path $env:TEMP "wallbrand_intune.png"
    $bitmap.Save($tempWallpaper, [System.Drawing.Imaging.ImageFormat]::Png)
    $bitmap.Dispose()
    
    Export-IntunePackage -OutputFolder $ExportIntune -Config $config -WallpaperPath $tempWallpaper
    Remove-Item $tempWallpaper -ErrorAction SilentlyContinue
    
    Write-Host "Intune package exported to: $ExportIntune" -ForegroundColor Green
    exit 0
}

# Silent mode or Apply mode (for scheduled tasks)
if ($Silent -or ($Apply -and $ConfigPath)) {
    Write-Host "WallBrand Pro - Silent Mode" -ForegroundColor Cyan
    
    $primary = [System.Windows.Forms.Screen]::PrimaryScreen
    $bitmap = Render-BrandedWallpaper -Config $config -TargetWidth $primary.Bounds.Width -TargetHeight $primary.Bounds.Height
    
    if ($OutputPath) {
        $bitmap.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
        Write-Host "  Wallpaper saved to: $OutputPath" -ForegroundColor Green
    }
    
    if ($Apply -or $Silent) {
        $tempPath = Join-Path $env:LOCALAPPDATA "WallBrand\wallpaper.png"
        $tempDir = Split-Path $tempPath
        if (-not (Test-Path $tempDir)) { New-Item -Path $tempDir -ItemType Directory -Force | Out-Null }
        
        $bitmap.Save($tempPath, [System.Drawing.Imaging.ImageFormat]::Png)
        Set-DesktopWallpaper -Path $tempPath -Style "Fill"
        Write-Host "  Wallpaper applied successfully!" -ForegroundColor Green
    }
    
    $bitmap.Dispose()
    exit 0
}

# GUI mode
Show-WallBrandProGUI -InitialConfig $config
