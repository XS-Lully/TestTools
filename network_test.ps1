# ============================================================
# 工具名稱：iperf - IP set & iperf 測試 (Rev.1)
# ============================================================

# --- [ 1. 使用者自定義變數區 ] ---
$DebugMode        = $False        
$EnableLLDP       = $False        
$RefreshInterval  = 10
$DefaultDuration  = 600
$DefaultParallel  = 8
$ReportThreshold  = 0.8

# --- [ 系統預設參數區 ] ---
$IperfFileName    = "iperf-2.1.7-win.exe"
$LogRootBase      = "C:\TestLogs\iperf"
$TimeStamp        = Get-Date -Format "yyyyMMdd_HHmmss"
$LogRootPath      = Join-Path $LogRootBase $TimeStamp
$IperfArgsBase    = "-w 512k -l 64k -i 10"

# --- [ 2. 功能模組區 ] ---

function Initialize-Env {
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "錯誤：請以系統管理員身分執行！" -ForegroundColor Red
        Pause; exit
    }
    if (-not (Test-Path $LogRootPath)) { New-Item $LogRootPath -ItemType Directory -Force | Out-Null }
}

function Show-Header-Static {
    Clear-Host
    $line = "=" * 170
    Write-Host $line -ForegroundColor Cyan
    Write-Host "  iperf 測試工具 Rev.1 | Log路徑: $LogRootPath"
    Write-Host $line -ForegroundColor Cyan
}

function Get-MyAdapters {
    $rawAdapters = Get-NetAdapter
    if (-not $DebugMode) { $rawAdapters = $rawAdapters | Where-Object { $_.Status -eq 'Up' } }
    
    $adapterData = foreach ($adp in $rawAdapters) {
        $slotValue = "N/A"; $sortPriority = 99 
        if ($adp.Name -match "Onboard") { $slotValue = "Onboard"; $sortPriority = 1 }
        elseif ($adp.Name -match "OCP|LOM|Flexible") { $slotValue = "OCP"; $sortPriority = 10 }
        elseif ($adp.Name -match "Slot\s+(\d+)") { $slotValue = $Matches[1]; $sortPriority = 20 + [int]$slotValue }
        else {
            try {
                $hw = $adp | Get-NetAdapterHardwareInfo -ErrorAction SilentlyContinue
                if ($hw -and $hw.SlotNumber -ne 0) { $slotValue = $hw.SlotNumber.ToString(); $sortPriority = 20 + [int]$slotValue }
            } catch { }
        }

        $ipsV4 = Get-NetIPAddress -InterfaceIndex $adp.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.IPAddress -like "192.168.*" }
        $v4Str = if ($ipsV4) { ($ipsV4 | Select-Object -ExpandProperty IPAddress) -join "," } else { "None" }
        $ipsV6 = Get-NetIPAddress -InterfaceIndex $adp.ifIndex -AddressFamily IPv6 -ErrorAction SilentlyContinue | Where-Object { $_.IPAddress -notlike "fe80*" }
        $v6Str = if ($ipsV6) { ($ipsV6 | Select-Object -ExpandProperty IPAddress) -join "," } else { "None" }

        [PSCustomObject]@{
            Priority  = $sortPriority; Index = $adp.ifIndex; Name = $adp.Name; Slot = $slotValue
            Desc      = $adp.InterfaceDescription; Speed = $adp.LinkSpeed; Status = $adp.Status
            IPv4      = $v4Str; IPv6 = $v6Str
        }
    }
    return $adapterData | Sort-Object Priority, Name
}

function Invoke-AdapterSelection {
    param($Adapters)
    if ($null -eq $Adapters -or $Adapters.Count -eq 0) { return $null }
    Write-Host "`n[ 可用網卡列表 ]" -ForegroundColor Yellow
    # 增加寬度以顯示完整描述
    Write-Host ("{0,-4} {1,-8} {2,-22} {3,-65} {4,-12} {5}" -f "No.", "Slot", "名稱", "完整網卡描述", "速度", "IPv4 / IPv6")
    Write-Host ("-" * 170)
    for ($i=0; $i -lt $Adapters.Count; $i++) {
        $a = $Adapters[$i]; $color = if ($a.Status -eq "Up") { "Green" } else { "Gray" }
        $dispSlot = if ($a.Slot -eq "N/A" -or $a.Slot -eq "0") { "---" } else { $a.Slot }
        
        Write-Host ("{0,-5} {1,-9} {2,-23} {3,-66}" -f ($i+1), $dispSlot, $a.Name, $a.Desc) -NoNewline
        Write-Host ("{0,-13} v4:{1} / v6:{2}" -f $a.Speed, $a.IPv4, $a.IPv6) -ForegroundColor $color
    }
    Write-Host "`n>> 請輸入網卡編號 (例如: 1-3 或 1,2)" -ForegroundColor Cyan
    $InputString = Read-Host ">> [直接按 Enter 預設選取除 Onboard 外的所有網卡]，按 Q 返回"
    if ($InputString -eq "Q") { return $null }
    if ([string]::IsNullOrWhiteSpace($InputString)) { return $Adapters | Where-Object { $_.Name -notmatch "Onboard" } }
    $Indices = @()
    try {
        foreach ($part in $InputString.Split(',').Trim()) {
            if ($part -match '-') {
                $range = $part.Split('-'); for ($j=[int]$range[0]; $j -le [int]$range[1]; $j++) { $Indices += ($j-1) }
            } else { $Indices += ([int]$part-1) }
        }
    } catch { return $null }
    return $Adapters[$Indices] | Where-Object { $_ -ne $null }
}

function Set-IPConfiguration {
    Show-Header-Static
    Write-Host "`n [ 設定網卡 IP 流程 ]" -ForegroundColor Yellow
    Write-Host " [1] Server (192.168.x.1) `n [2] Client (192.168.x.2)"
    $mode = Read-Host "請選擇模式"
    if ($mode -notmatch "1|2") { return }
    $AllAdapters = Get-MyAdapters
    $selected = Invoke-AdapterSelection -Adapters $AllAdapters
    if (-not $selected) { return }
    $subnetCounter = 1
    $val = Read-Host "`n請輸入起始 Subnet X (192.168.x.$mode)[預設 1]"; if ($val) { $subnetCounter = [int]$val }
    foreach ($item in $selected) {
        $v4 = "192.168.$subnetCounter.$(if($mode -eq '1'){'1'}else{'2'})"
        $v6 = "2001::${subnetCounter}:$(if($mode -eq '1'){'1'}else{'2'})"
        Write-Host "正在設定 $($item.Name) -> v4: $v4 | v6: $v6" -ForegroundColor Cyan
        Set-NetIPInterface -InterfaceIndex $item.Index -DHCP Disabled -ErrorAction SilentlyContinue
        Remove-NetIPAddress -InterfaceIndex $item.Index -Confirm:$false -ErrorAction SilentlyContinue
        New-NetIPAddress -InterfaceIndex $item.Index -IPAddress $v4 -PrefixLength 24 -Confirm:$false | Out-Null
        New-NetIPAddress -InterfaceIndex $item.Index -IPAddress $v6 -PrefixLength 64 -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        $subnetCounter++
    }
    Write-Host "`nIP 設定完成！" -ForegroundColor Green
}

function Parse-Mbps {
    param($LogPath)
    if (-not (Test-Path $LogPath)) { return 0 }
    $content = Get-Content $LogPath -Tail 5 -ErrorAction SilentlyContinue
    if (!$content) { return 0 }
    for ($i = $content.Count - 1; $i -ge 0; $i--) {
        if ($content[$i] -match "(\d+\.?\d*)\s+Mbits/sec") { return [double]$Matches[1] }
        if ($content[$i] -match "(\d+\.?\d*)\s+Gbits/sec") { return [double]$Matches[1] * 1000 }
    }
    return 0
}

function Start-IperfTest {
    param($SelectedAdapters)
    $IperfPath = Join-Path (Get-Location) $IperfFileName
    if (-not (Test-Path $IperfPath)) { Write-Host "找不到 $IperfFileName ！" -ForegroundColor Red; return }

    $Duration = Read-Host "時長 (秒) [預設 $DefaultDuration]"; if (!$Duration) { $Duration = $DefaultDuration }
    $Parallel = Read-Host "平行數 (P) [預設 $DefaultParallel]"; if (!$Parallel) { $Parallel = $DefaultParallel }

    $TaskList = @()
    foreach ($item in $SelectedAdapters) {
        $ips = Get-NetIPAddress -InterfaceIndex $item.Index -AddressFamily IPv4 -ErrorAction SilentlyContinue | 
               Where-Object { $_.IPAddress -like "192.168.*" } | Select-Object -First 1
        if ($ips -and $ips.IPAddress) {
            $local = $ips.IPAddress.Trim()
            $target = if ($local -match "\.1$") { $local -replace "\.1$", ".2" } else { $local -replace "\.\d+$", ".1" }
            $outLog = Join-Path $LogRootPath "iperf_$($item.Name)_TX.log"; $inLog  = Join-Path $LogRootPath "iperf_$($item.Name)_RX.log"
            $cmdTX = "-c $target -B $local -t $Duration -P $Parallel $IperfArgsBase "
            $cmdRX = "-c $target -B $local -t $Duration -P $Parallel $IperfArgsBase -R "
            $pOut = Start-Process $IperfPath -ArgumentList $cmdTX -NoNewWindow -PassThru -RedirectStandardOutput $outLog
            $pIn  = Start-Process $IperfPath -ArgumentList $cmdRX -NoNewWindow -PassThru -RedirectStandardOutput $inLog
            $TaskList += [PSCustomObject]@{ 
                Task=$item; OutProc=$pOut; InProc=$pIn; OutLog=$outLog; InLog=$inLog; 
                FullTXCmd = "$IperfFileName $cmdTX"; FullRXCmd = "$IperfFileName $cmdRX" 
            }
        }
    }

    while ($TaskList | Where { $_.OutProc.HasExited -eq $false -or $_.InProc.HasExited -eq $false }) {
        Show-Header-Static
        Write-Host ">>> 測試執行中 (門檻: $($ReportThreshold*100)%)" -ForegroundColor Yellow
        Write-Host ("{0,-4} {1,-6} {2,-22} {3,-65} {4,-12} {5,-18} {6,-18}" -f "No.", "Slot", "網卡名稱", "網卡描述", "網卡速度", "TX (上傳)", "RX (下載)")
        Write-Host ("-" * 170)
        foreach ($t in $TaskList) {
            $tx = Parse-Mbps -LogPath $t.OutLog; $rx = Parse-Mbps -LogPath $t.InLog
            $linkMbps = if ($t.Task.Speed -match "Gbps") { [double]($t.Task.Speed -replace " Gbps","") * 1000 } else { [double]($t.Task.Speed -replace " Mbps","") }
            $txColor = if ($tx -lt ($linkMbps * $ReportThreshold)) { "Red" } else { "Green" }
            $rxColor = if ($rx -lt ($linkMbps * $ReportThreshold)) { "Red" } else { "Green" }
            Write-Host ("{0,-5} {1,-7} {2,-23} {3,-66} {4,-13}" -f ($TaskList.IndexOf($t)+1), $t.Task.Slot, $t.Task.Name, $t.Task.Desc, $t.Task.Speed) -NoNewline
            Write-Host ("{0,-19}" -f "$tx Mbps") -ForegroundColor $txColor -NoNewline
            Write-Host ("{0,-19}" -f "$rx Mbps") -ForegroundColor $rxColor
        }
        Start-Sleep -Seconds $RefreshInterval
    }

    # --- [ 3. 產出完整最終報告 ] ---
    $ReportFile = Join-Path $LogRootPath "Final_Report.txt"
    $repHeader = "iperf 測試報告 - $TimeStamp`r`n" + ("=" * 170) + "`r`n"
    $repTitle  = "{0,-4} {1,-8} {2,-20} {3,-65} {4,-12} {5,-12} {6,-12} {7}" -f "No.", "Slot", "網卡名稱", "完整網卡描述", "速度", "TX(Mbps)", "RX(Mbps)", "結果"
    $repHeader + $repTitle | Out-File $ReportFile -Encoding UTF8

    Show-Header-Static
    Write-Host "`n[ 測試報告 ]" -ForegroundColor Cyan
    Write-Host $repTitle; Write-Host ("-" * 170)

    foreach ($t in $TaskList) {
        $tx = Parse-Mbps -LogPath $t.OutLog; $rx = Parse-Mbps -LogPath $t.InLog
        $linkMbps = if ($t.Task.Speed -match "Gbps") { [double]($t.Task.Speed -replace " Gbps","") * 1000 } else { [double]($t.Task.Speed -replace " Mbps","") }
        $res = "PASS"; $resColor = "Green"
        if ($tx -lt ($linkMbps * $ReportThreshold) -or $rx -lt ($linkMbps * $ReportThreshold) -or $tx -eq 0) { $res = "FAIL"; $resColor = "Red" }

        $row = "{0,-5} {1,-9} {2,-21} {3,-66} {4,-13} {5,-13} {6,-13} {7}" -f ($TaskList.IndexOf($t)+1), $t.Task.Slot, $t.Task.Name, $t.Task.Desc, $t.Task.Speed, $tx, $rx, $res
        Write-Host $row -ForegroundColor $resColor
        $row | Out-File $ReportFile -Append -Encoding UTF8
    }

    "`r`n`r`n[ 完整執行指令備份 ]`r`n" + ("-" * 80) | Out-File $ReportFile -Append -Encoding UTF8
    foreach ($t in $TaskList) {
        $cmdInfo = "`r`nNo.$($TaskList.IndexOf($t)+1) [$($t.Task.Name) - $($t.Task.Desc)]`r`nTX: $($t.FullTXCmd)`r`nRX: $($t.FullRXCmd)`r`n"
        $cmdInfo | Out-File $ReportFile -Append -Encoding UTF8
    }
    Write-Host "` log 已儲存: $ReportFile" -ForegroundColor Cyan
}

function Main {
    Initialize-Env
    while ($true) {
        Show-Header-Static
        Write-Host " [1] 設定網卡 IP "
        Write-Host " [2] 進入 iperf 測試模式"
        Write-Host " [Q] 退出"
        $choice = Read-Host "`n請選擇功能"
        if ($choice -eq "Q") { break }
		if ($choice -eq "q") { break }
        if ($choice -eq "1") { Set-IPConfiguration; Pause } 
        elseif ($choice -eq "2") {
            Write-Host "`n [1] Server (-s) / [2] Client (-c)"
            $m = Read-Host "請選擇模式"
            if ($m -eq "1") {
                $AllAdapters = Get-MyAdapters
                $ListenList = @()
                foreach ($adp in $AllAdapters) {
                    if ($adp.IPv4 -like "*192.168.*") {
                        $targetIP = ($adp.IPv4.Split(',') | Where-Object { $_.Trim() -like "192.168.*" } | Select-Object -First 1).Trim()
                        Start-Process (Join-Path (Get-Location) $IperfFileName) -ArgumentList "-s -B $targetIP -f m" -WindowStyle Hidden
                        $ListenList += $adp
                    }
                }
                Write-Host "`n[ iperf Server ]" -ForegroundColor Green
                Write-Host ("{0,-4} {1,-8} {2,-20} {3,-65} {4}" -f "No.", "Slot", "名稱", "完整網卡描述", "監聽 IP")
                Write-Host ("-" * 170)
                foreach ($l in $ListenList) {
                    $dispSlot = if ($l.Slot -eq "N/A" -or $l.Slot -eq "0") { "---" } else { $l.Slot }
                    Write-Host ("{0,-5} {1,-9} {2,-21} {3,-66} {4}" -f ($ListenList.IndexOf($l)+1), $dispSlot, $l.Name, $l.Desc, $l.IPv4)
                }
                Read-Host "`n按 Enter 停止所有 Server 並返回"
                Stop-Process -Name $IperfFileName.Replace(".exe","") -ErrorAction SilentlyContinue
            } else {
                $AllAdapters = Get-MyAdapters
                $selected = Invoke-AdapterSelection -Adapters $AllAdapters
                if ($selected) { Start-IperfTest -SelectedAdapters $selected }
                Pause
            }
        }
    }
}
Main