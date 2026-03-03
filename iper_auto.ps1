# ============================================================
# 工具名稱：iperf - IP set & iperf 測試 (Rev.0)
# 製作日期：2026/03/02
# 製作人  ：Lully
# ============================================================

# --- [ 使用者自定義變數區 ] ---
$RefreshInterval  = 10          # 螢幕刷新時間 (秒)
$DefaultDuration  = 600         # 預設測試總時長 (秒)
$DefaultParallel  = 8          # 預設平行連線數 (P)
$ReportThreshold  = 0.8        # 達標門檻 (80%)

# --- [ 系統預設參數區 ] ---
$Rev             = "iperf 測試工具 Rev.0"
$IperfFileName   = "iperf-2.1.7-win.exe"
$LogRootBase     = "C:\TestLogs\iperf" 
$TimeStamp       = Get-Date -Format "yyyyMMdd_HHmmss"
$LogRootPath     = Join-Path $LogRootBase $TimeStamp
$IperfArgsBase    = "-w 512k -l 64k -i 10"

# --- [ 權限檢查 ] ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "錯誤：請以系統管理員身分執行！" -ForegroundColor Red
    Pause; exit
}

# --- [ 功能模組 ] ---

function Show-Header {
    Clear-Host
    $line = "=" * 115
    Write-Host $line -ForegroundColor Cyan
    Write-Host "  $Rev | log 路徑: $LogRootPath"
    Write-Host $line -ForegroundColor Cyan
}

function Parse-Mbps {
    param($LogPath)
    if (-not (Test-Path $LogPath)) { return 0 }
    $content = Get-Content $LogPath -Tail 5 -ErrorAction SilentlyContinue
    if ($null -eq $content) { return 0 }
    for ($i = $content.Count - 1; $i -ge 0; $i--) {
        if ($content[$i] -match "(\d+\.?\d*)\s+Mbits/sec") { return [double]$Matches[1] }
        if ($content[$i] -match "(\d+\.?\d*)\s+Gbits/sec") { return [double]$Matches[1] * 1000 }
    }
    return 0
}

function Get-SortedAdapters {
    return Get-NetAdapter | Where-Object { $_.Status -eq 'Up' } | 
        Sort-Object { [Regex]::Replace($_.Name, '\d+', { $args[0].Value.PadLeft(10, '0') }) }
}

function Select-Ports {
    param($Adapters)
    Write-Host "`n[ 可用網卡列表 ]" -ForegroundColor Yellow
    Write-Host ("{0,-4} {1,-12} {2,-45} {3,-12} {4}" -f "編號", "名稱", "描述", "速度", "目前的IPv4")
    Write-Host ("{0,-4} {1,-12} {2,-45} {3,-12} {4}" -f "----", "------------", "---------------------------------------------", "------------", "-----------")
    for ($i=0; $i -lt $Adapters.Count; $i++) {
        $adp = $Adapters[$i]
        $ips = Get-NetIPAddress -InterfaceIndex $adp.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue
        $ipStr = if ($ips) { ($ips | Select-Object -ExpandProperty IPAddress) -join ", " } else { "None" }
        Write-Host ("{0,-5}" -f ($i + 1)) -NoNewline
        Write-Host ("{0,-13}" -f $adp.Name) -NoNewline
        Write-Host ("{0,-46}" -f $adp.InterfaceDescription) -NoNewline
        Write-Host ("{0,-13}" -f $adp.LinkSpeed) -NoNewline
        Write-Host $ipStr
    }
    Write-Host ""
    $InputString = ""; while ([string]::IsNullOrWhiteSpace($InputString)) { $InputString = Read-Host ">> 請輸入網卡編號 (例如: 2-5 or 4,5,6)" }
    $Indices = @()
    try {
        foreach ($part in $InputString.Split(',').Trim()) {
            if ($part -match '-') {
                $range = $part.Split('-'); for ($j=[int]$range[0]; $j -le [int]$range[1]; $j++) { $Indices += ($j-1) }
            } else { $Indices += ([int]$part-1) }
        }
    } catch { return $null }
    return $Adapters[$Indices] | Where { $_ -ne $null }
}

# --- [ 主程式邏輯 ] ---
Show-Header
Write-Host " [1] 設定網卡 IP (v4/v6)" 
Write-Host " [2] 進入 iperf 測試模式" 
Write-Host "-----------------------------------------------------------------"
$MainAction = ""; while ($MainAction -notin @("1", "2")) { $MainAction = Read-Host "請選擇功能" }

$AllAdapters = Get-SortedAdapters

if ($MainAction -eq "1") {
    $SelectedAdapters = Select-Ports -Adapters $AllAdapters
    if ($null -eq $SelectedAdapters) { exit }
    Write-Host "`n [1] Server (192.168.x.1)"
	Write-Host "`n [2] Client (192.168.x.2)"
    $Mode = ""; while ($Mode -notin @("1", "2")) { $Mode = Read-Host "請選擇" }
    $HostID = if ($Mode -eq "1") { "1" } else { "2" }
    $StartSubnet = Read-Host "起始 Subnet X [預設 1]"; if ([string]::IsNullOrWhiteSpace($StartSubnet)) { $StartSubnet = 1 }
    $SubID = [int]$StartSubnet
    foreach ($Adp in $SelectedAdapters) {
        $v4 = "192.168.$SubID.$HostID"; $v6 = "2001::${SubID}:${HostID}"
        Write-Host "設定 $($Adp.Name): $v4" -ForegroundColor Cyan
        Set-NetIPInterface -InterfaceIndex $Adp.ifIndex -DHCP Disabled -ErrorAction SilentlyContinue
        Remove-NetIPAddress -InterfaceIndex $Adp.ifIndex -Confirm:$false -ErrorAction SilentlyContinue
        New-NetIPAddress -InterfaceIndex $Adp.ifIndex -IPAddress $v4 -PrefixLength 24 -Confirm:$false | Out-Null
        New-NetIPAddress -InterfaceIndex $Adp.ifIndex -IPAddress $v6 -PrefixLength 64 -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        $SubID++
    }
    Pause
} 
else {
    Write-Host "`n [1] Server 監聽 (-s) / [2] Client 測試 (-c)"
    $ModeAction = ""; while ($ModeAction -notin @("1", "2")) { $ModeAction = Read-Host "請選擇" }
    $IperfPath = Join-Path (Get-Location) $IperfFileName

    if ($ModeAction -eq "1") {
        Get-NetIPAddress -AddressFamily IPv4 | Where { $_.IPAddress -like "192.168.*" } | ForEach-Object {
            $ip = $_.IPAddress; $adpName = (Get-NetAdapter -InterfaceIndex $_.InterfaceIndex).Name
            Start-Process $IperfPath -ArgumentList "-s -B $ip -f m" -WindowStyle Hidden
            Write-Host "  [Listen] $ip ($adpName)" -ForegroundColor Gray
        }
        Read-Host "`nServer 監聽中，按 Enter 停止所有 iperf 進程"
        Stop-Process -Name $IperfFileName.Replace(".exe","") -ErrorAction SilentlyContinue
    } 
    else {
        $SelectedAdapters = Select-Ports -Adapters $AllAdapters
        if ($null -eq $SelectedAdapters) { exit }
        $Duration = Read-Host "時長 (秒) [預設 $DefaultDuration]"; if ([string]::IsNullOrWhiteSpace($Duration)) { $Duration = $DefaultDuration }
        $Parallel = Read-Host "平行數 (P) [預設 $DefaultParallel]"; if ([string]::IsNullOrWhiteSpace($Parallel)) { $Parallel = $DefaultParallel }
        
        if (-not (Test-Path $LogRootPath)) { New-Item $LogRootPath -ItemType Directory -Force | Out-Null }

        $TaskList = @()
        foreach ($Adp in $SelectedAdapters) {
            $ips = Get-NetIPAddress -InterfaceIndex $Adp.ifIndex -AddressFamily IPv4 | Where { $_.IPAddress -like "192.168.*" }
            if ($ips) {
                $local = $ips[0].IPAddress; $target = $local -replace "\.\d+$", ".1"
                $outLog = Join-Path $LogRootPath "iperf_$($Adp.Name)_Out.log"
                $inLog  = Join-Path $LogRootPath "iperf_$($Adp.Name)_In.log"
                $fullCmdTX = "$IperfFileName -c $target -B $local -t $Duration -P $Parallel $IperfArgsBase"
                $fullCmdRX = "$IperfFileName -c $target -B $local -t $Duration -P $Parallel $IperfArgsBase -R"
                
                $pOut = Start-Process $IperfPath -ArgumentList ($fullCmdTX.Replace("$IperfFileName ","")) -NoNewWindow -PassThru -RedirectStandardOutput $outLog
                $pIn  = Start-Process $IperfPath -ArgumentList ($fullCmdRX.Replace("$IperfFileName ","")) -NoNewWindow -PassThru -RedirectStandardOutput $inLog
                
                $TaskList += [PSCustomObject]@{ 
                    ID=($TaskList.Count+1); Name=$Adp.Name; Desc=$Adp.InterfaceDescription; 
                    Link=$Adp.LinkSpeed; OutLog=$outLog; InLog=$inLog; OutProc=$pOut; InProc=$pIn; 
                    TXCmd=$fullCmdTX; RXCmd=$fullCmdRX 
                }
            }
        }

        # --- [ 等待進程真正結束 ] ---
        while ($true) {
            $RunningProcs = $TaskList | Where-Object { $_.OutProc.HasExited -eq $false -or $_.InProc.HasExited -eq $false }
            if ($null -eq $RunningProcs) { break } # 所有進程都結束了，跳出迴圈

            Show-Header
            Write-Host ">>> 測試執行中 " -ForegroundColor Yellow
            Write-Host ("-" * 115)
            Write-Host ("{0,-4} {1,-15} {2,-15} {3,-25} {4,-25}" -f "編號", "網卡名稱", "網卡速度", "TX (上傳) 速度", "RX (下載) 速度")
            
            foreach ($task in $TaskList) {
                $tx = Parse-Mbps -LogPath $task.OutLog; $rx = Parse-Mbps -LogPath $task.InLog
                $linkMbps = if ($task.Link -match "Gbps") { [double]($task.Link -replace " Gbps","") * 1000 } else { [double]($task.Link -replace " Mbps","") }
                
                # 斷線或速度不足判定
                $txColor = if ($tx -lt ($linkMbps * $ReportThreshold) -and $tx -gt 0) { "Red" } else { "Green" }
                $rxColor = if ($rx -lt ($linkMbps * $ReportThreshold) -and $rx -gt 0) { "Red" } else { "Green" }
                if ($tx -eq 0) { $txColor = "Red" }; if ($rx -eq 0) { $rxColor = "Red" }

                Write-Host ("{0,-5} {1,-16} {2,-16}" -f $task.ID, $task.Name, $task.Link) -NoNewline
                Write-Host ("{0,-26}" -f "$tx Mbps") -ForegroundColor $txColor -NoNewline
                Write-Host ("{0,-26}" -f "$rx Mbps") -ForegroundColor $rxColor
            }
            Start-Sleep -Seconds $RefreshInterval
        }

        # --- [ 最終結案與產出報告 ] ---
        Show-Header
        $ReportFile = Join-Path $LogRootPath "Final_Report.txt"
        $header = "測試報告 - $TimeStamp (完整執行結束)"
        $titleLine = "{0,-4} {1,-12} {2,-45} {3,-12} {4,-12} {5,-12} {6}" -f "編號", "名稱", "描述", "網卡速度", "TX(Mbps)", "RX(Mbps)", "結果"
        $sep = "{0,-4} {1,-12} {2,-45} {3,-12} {4,-12} {5,-12} {6}" -f "----", "------------", "---------------------------------------------", "------------", "----------", "----------", "----"
        
        Write-Host "==================================================================================================================="
        Write-Host "                                 $header"
        Write-Host "==================================================================================================================="
        Write-Host $titleLine; Write-Host $sep
        "$header`r`n`r`n$titleLine`r`n$sep" | Out-File $ReportFile -Encoding UTF8

        foreach ($task in $TaskList) {
            $tx = Parse-Mbps -LogPath $task.OutLog; $rx = Parse-Mbps -LogPath $task.InLog
            $linkMbps = if ($task.Link -match "Gbps") { [double]($task.Link -replace " Gbps","") * 1000 } else { [double]($task.Link -replace " Mbps","") }
            
            $res = "PASS"
            if ($tx -eq 0 -or $rx -eq 0) { $res = "FAIL (No Data/Disconnect)" }
            elseif ($tx -lt ($linkMbps * $ReportThreshold) -or $rx -lt ($linkMbps * $ReportThreshold)) { $res = "FAIL (Low Speed)" }
            
            $rowColor = if ($res -eq "PASS") { "Green" } else { "Red" }
            $outLine = "{0,-5} {1,-13} {2,-46} {3,-13} {4,-13} {5,-13} {6}" -f $task.ID, $task.Name, $task.Desc, $task.Link, $tx, $rx, $res
            Write-Host $outLine -ForegroundColor $rowColor
            $outLine | Out-File $ReportFile -Append -Encoding UTF8
        }
        
        # 指令紀錄
        "`r`n" + ("=" * 80) + "`r`n測試指令 (Detailed Commands):`r`n" + ("=" * 80) | Out-File $ReportFile -Append -Encoding UTF8
        foreach ($task in $TaskList) {
            "`r`n[$($task.Name) Command]:`r`nTX: $($task.TXCmd)`r`nRX: $($task.RXCmd)" | Out-File $ReportFile -Append -Encoding UTF8
        }
        
        Write-Host "`n測試結束！完整報告已存至: $ReportFile" -ForegroundColor Cyan
        Pause
    }
}