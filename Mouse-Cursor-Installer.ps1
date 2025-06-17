# 鼠標光標批量安裝工具 v1.0
# 需要管理員權限運行

param(
    [string]$SignalFile,
    [switch]$Debug
)

function Write-DebugLog {
    param([string]$message)
    if ($script:debug) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        "[$timestamp] $message" | Out-File -FilePath "$PWD\mouse-cursor.log" -Append
    }
}

$script:schemeOrder = @(
    "Arrow", "Help", "AppStarting", "Wait", "Crosshair", "IBeam", "NWPen", "No",
    "SizeNS", "SizeWE", "SizeNWSE", "SizeNESW", "SizeAll", "UpArrow", "Hand", "Pin", "Person"
)

# 光标类型中英文对照表
$script:cursorTypeChinese = @{
    "Arrow"       = "正常选择"
    "Help"        = "帮助选择"
    "AppStarting" = "后台运行"
    "Wait"        = "忙"
    "Crosshair"   = "精确选择"
    "IBeam"       = "文本选择"
    "NWPen"       = "手写"
    "No"          = "不可用"
    "SizeNS"      = "垂直调整大小"
    "SizeWE"      = "水平调整大小"
    "SizeNWSE"    = "沿对角线调整大小1"
    "SizeNESW"    = "沿对角线调整大小2"
    "SizeAll"     = "移动"
    "UpArrow"     = "候选"
    "Hand"        = "链接选择"
    "Pin"         = "位置选择"
    "Person"      = "个人选择"
}

# 按类型分组的关键词（按优先级排序，长关键词在前）
    $keywordGroups = @{
        "Arrow" = @(
            "normalselect", "通常の選択", "正常选择",
            "normal", "通常", "arrow", "default", "nomal"
        )
        "Help" = @(
            "helpselect", "ヘルプの選択", "帮助选择",
            "help", "ヘルプ"
        )
        "AppStarting" = @(
            "workinginbackground", "バックグラウンドで作業中", "バックグラウンド", "后台运行",
            "working", "background", "appstarting", "作業中", "loading"
        )
        "Wait" = @(
            "待ち状態", "wait", "busy", "待ち", "忙"
        )
        "Crosshair" = @(
            "領域選択", "精确选择", "crosshair", "precision", "percision", "cross", "領域"
        )
        "IBeam" = @(
            "textselect", "テキスト選択", "文本选择",
            "text", "テキスト", "ibeam", "beam"
        )
        "NWPen" = @(
            "handwriting", "手書き", "手写", "pen", "笔"
        )
        "No" = @(
            "notallowed", "利用不可", "不可用", "unavailable", "unvailable", "unavail", "not", "no"
        )
        "SizeNS" = @(
            "verticalresize", "上下に拡大縮小", "垂直调整大小",
            "ns-resize", "n-resize", "vertical", "上下", "ns", "sn"
        )
        "SizeWE" = @(
            "horizontalresize", "左右に拡大縮小", "水平调整大小",
            "ew-resize", "w-resize", "horizontal", "左右", "ew", "we"
        )
        "SizeNWSE" = @(
            "diagonalresize1", "斜めに拡大縮小1", "斜めに拡大縮小", "沿对角线调整大小1",
            "nwse-resize", "nw-resize", "diagonal1", "resize1", "斜め1", "nwse"
        )
        "SizeNESW" = @(
            "diagonalresize2", "斜めに拡大縮小2", "沿对角线调整大小2",
            "nesw-resize", "ne-resize", "diagonal2", "resize2", "斜め2", "nesw"
        )
        "SizeAll" = @(
            "sizeall", "move", "移動", "移动", "nsew"
        )
        "UpArrow" = @(
            "alternateselec", "代替選択", "候选",
            "alternate", "代替", "uparrow", "up"
        )
        "Hand" = @(
            "linkselect", "リンクの選択", "链接选择",
            "link", "リンク", "hand", "pointer"
        )
        "Pin" = @(
            "位置选择", "pin", "場所の選択", "場所"
        )
        "Person" = @(
            "个人选择", "person", "人の選択", "人“
        )
    }

    # 展开为扁平映射表，按关键词长度降序排序
    $keywordMap = @{}
    foreach ($type in $keywordGroups.Keys) {
        foreach ($keyword in $keywordGroups[$type]) {
            $keywordMap[$keyword] = $type
        }
    }
    
    # 按关键词长度降序排序
    $sortedKeywords = $keywordMap.Keys | Sort-Object Length -Descending

# 定義函數：處理一個光標集合（一個主題）
function ProcessCursorSet {
    param([string]$sourceDir)

    Write-Host "`n處理目錄: $sourceDir"

    # 獲取目錄名作為主題名
    $themeName = [System.IO.Path]::GetFileName($sourceDir)

    # 清理主題名稱中的特殊字符
    $originalThemeName = $themeName
    $themeName = $themeName -replace '\[|\]|\(|\)|&|%', ''

    Write-DebugLog "原始主題名稱: '$originalThemeName'"
    Write-DebugLog "清理後主題名稱: '$themeName'"

    # 收集光標文件 (.ani 和 .cur)
    $cursorFiles = Get-ChildItem -Path $sourceDir -File | Where-Object { $_.Extension -in '.ani', '.cur' } | Sort-Object Name
    Write-Host "主題名稱: $themeName (找到 $($cursorFiles.Count) 個光標文件)"
	
    # 添加统计
    $script:processedFiles += $cursorFiles.Count
    $themeInstallSuccess = $true

    # 獲取所有文件名的基本部分（不含擴展名）
    $baseNames = $cursorFiles | ForEach-Object { 
        [System.IO.Path]::GetFileNameWithoutExtension($_.Name) 
    }

    # 計算公共前綴和後綴
    $prefixSuffix = Get-CommonPrefixSuffix -strings $baseNames
    $commonPrefix = $prefixSuffix[0]
    $commonSuffix = $prefixSuffix[1]
    
    Write-DebugLog "公共前綴: '$commonPrefix'"
    Write-DebugLog "公共後綴: '$commonSuffix'"
    Write-Host ""

    # 創建映射表
    $mapping = @()
    $typeMapping = @{}
    $cursorCount = 0

    foreach ($file in $cursorFiles) {
        $cursorCount++
        $originalName = $file.Name
        $safeName = "${themeName}_${cursorCount}$($file.Extension)"
        
        # 獲取基本文件名（不含擴展名）
        $baseName = [IO.Path]::GetFileNameWithoutExtension($originalName)

        # 检查文件名长度是否小于前缀和后缀长度之和
        $totalPrefixSuffixLength = $commonPrefix.Length + $commonSuffix.Length

        if ($baseName.Length -lt $totalPrefixSuffixLength) {
        $baseName = $commonPrefix + $commonSuffix
        Write-DebugLog "前后缀可能重复，使用组合名称: $baseName"
        }

        # 移除公共前綴和後綴
        $strippedName = $baseName
        if (-not [string]::IsNullOrEmpty($commonPrefix)) {
            # 使用 [regex]::Escape 確保前綴中的特殊字符被正確處理
            $strippedName = $strippedName -replace "^$([regex]::Escape($commonPrefix))", ''
        }
        if (-not [string]::IsNullOrEmpty($commonSuffix)) {
            $strippedName = $strippedName -replace "$([regex]::Escape($commonSuffix))$", ''
        }
        
        # 如果去除後為空，使用arrow
        if ([string]::IsNullOrWhiteSpace($strippedName)) {
            $strippedName = "arrow"
        }

        Write-DebugLog "原始文件名: $originalName"
        Write-DebugLog "處理後文件名: $strippedName"

        # 根據處理後的文件名猜測光標類型
        $cursorType = DetermineCursorType -fileName $strippedName -fileNumber $cursorCount
        Write-Debuglog "文件: $originalName -> 類型: $cursorType"
		
        $typeMapping[$cursorType] = $originalName
        $mapping += [PSCustomObject]@{
            OriginalName = $originalName
            SafeName = $safeName
            CursorType = $cursorType
        }

        # 複製文件到系統光標目錄 - 使用.NET方法
        $destPath = Join-Path $script:systemCursorDir $safeName
        try {
            [System.IO.File]::Copy($file.FullName, $destPath, $true)
            Write-DebugLog "文件复制成功: $($file.Name) -> $safeName"
        } catch [System.UnauthorizedAccessException] {
            Write-Host "    ✗ 权限不足，无法复制文件: $($file.Name)" -ForegroundColor Red
			$themeInstallSuccess = $false
            continue
        } catch [System.IO.IOException] {
            Write-Host "    ✗ 文件被占用或路径问题: $($file.Name)" -ForegroundColor Red
			$themeInstallSuccess = $false
            continue
        } catch {
            Write-Host "    ✗ 复制失败: $($file.Name) - $($_.Exception.Message)" -ForegroundColor Red
			$themeInstallSuccess = $false
            continue
        }
    }

    # 检查光标方案数量是否异常
    $validCounts = @(5, 10, 15, 17)
    if ($cursorFiles.Count -notin $validCounts) {
        $script:abnormalCountThemes += [PSCustomObject]@{
            ThemeName = $originalThemeName
            FileCount = $cursorFiles.Count
        }
        Write-DebugLog "数量异常的方案: $originalThemeName ($($cursorFiles.Count) 个文件)"
    }
    
    # 检查是否存在未匹配的光标文件
    $matchedTypes = ($mapping | Select-Object -Property CursorType -Unique).Count
    if ($matchedTypes -lt $cursorFiles.Count) {
        $script:unmatchedFilesThemes += [PSCustomObject]@{
            ThemeName = $originalThemeName
            TotalFiles = $cursorFiles.Count
            MatchedTypes = $matchedTypes
            UnmatchedCount = $cursorFiles.Count - $matchedTypes
        }
        Write-DebugLog "存在未匹配文件的方案: $originalThemeName (总文件:$($cursorFiles.Count), 匹配类型:$matchedTypes)"
    }

    # 输出标准顺序的匹配结果
    Write-Host "光標類型匹配結果："
    Write-Host ""

    foreach ($type in $script:schemeOrder) {
        $typeLabel = "$type ($($script:cursorTypeChinese[$type]))"
        Write-Host "類型: $typeLabel -> $($typeMapping[$type])"
    }

    # 在嘗試寫入註冊表前，先檢查所有文件是否都已成功複製
    if ($themeInstallSuccess) {
        Write-DebugLog "開始註冊表設置..."
        try {
            CreateRegistryEntries -themeName $themeName -mapping $mapping -typeMapping $typeMapping
            $script:installedThemes++
            Write-Host "  ✓ 完成主題: $themeName ($($cursorFiles.Count) 個光標)" -ForegroundColor Green
        } catch {
            # 這種情況是註冊表寫入失敗
            $script:failedThemes += $originalThemeName
            Write-Host "  ✗ 安裝失敗 (註冊表寫入錯誤): $themeName" -ForegroundColor Red
        }
    } else {
        # 如果有任何文件複製失敗，則將整個主題標記為失敗
        $script:failedThemes += $originalThemeName
        Write-Host "  ✗ 安裝失敗: $themeName (由於文件複製錯誤)" -ForegroundColor Red
    }

    Write-Host ""
}

# 計算公共前後綴
function Get-CommonPrefixSuffix {
    param([string[]]$strings)
    
    if ($strings.Count -eq 0) { return @("", "") }
    
    # 计算前缀
    $prefix = $strings[0]
    foreach ($str in $strings) {
        $i = 0
        while ($i -lt $prefix.Length -and $i -lt $str.Length -and $prefix[$i] -eq $str[$i]) { $i++ }
        $prefix = $prefix.Substring(0, $i)
    }
    
    # 计算后缀
    $suffix = $strings[0]
    foreach ($str in $strings) {
        $i = 0
        $minLen = [Math]::Min($suffix.Length, $str.Length)
        while ($i -lt $minLen -and $suffix[$suffix.Length - $i - 1] -eq $str[$str.Length - $i - 1]) { $i++ }
        $suffix = $suffix.Substring($suffix.Length - $i)
    }
    
    return @($prefix, $suffix)
}

# 定義函數：根據文件名猜測光標類型
function DetermineCursorType {
    param(
        [string]$fileName,
        [int]$fileNumber
    )

    # 定义数字映射表
    $numberMap = @{
        1 = "Arrow"; 2 = "Help"; 3 = "AppStarting"; 4 = "Wait"; 5 = "Crosshair"
        6 = "IBeam"; 7 = "NWPen"; 8 = "No"; 9 = "SizeNS"; 10 = "SizeWE"
        11 = "SizeNWSE"; 12 = "SizeNESW"; 13 = "SizeAll"; 14 = "UpArrow"
        15 = "Hand"; 16 = "Pin"; 17 = "Person"
    }

    # 處理文件名：转换为小写并去除空格
    $fileNameLower = $fileName.ToLower() -replace '[\s・·]', ''
    
    # 关键词匹配 (使用預先計算好的 $script:sortedKeywords)
    foreach ($keyword in $script:sortedKeywords) {
        if ($fileNameLower.Contains($keyword)) {
            return $script:keywordMap[$keyword]
        }
    }

    # 数字序号匹配
    if ($fileNameLower -match '\d+') {
        $extractedNumber = [int]$matches[0]
        if ($numberMap.ContainsKey($extractedNumber)) {
            return $numberMap[$extractedNumber]
        }
    }

    # 按文件序号匹配
    if ($numberMap.ContainsKey($fileNumber)) {
        return $numberMap[$fileNumber]
    }

    # 默认值
    Write-Host "[WARNING] 無法識別光標類型: $fileName (使用默認Arrow)" -ForegroundColor Yellow
    return "Arrow"
}

# 定義函數：創建註冊表項
function CreateRegistryEntries {
    param(
        [string]$themeName,
        [array]$mapping,
        [hashtable]$typeMapping
    )

    Write-DebugLog "构建光标方案顺序: $($schemeOrder -join ', ')"
    
    # 第一遍遍历：记录所有需要的光标路径
    $cursorPaths = @{}
    foreach ($type in $schemeOrder) {
        $cursorFile = $mapping | Where-Object { $_.CursorType -eq $type } | Select-Object -First 1
        if ($cursorFile) {
            $cursorPaths[$type] = Join-Path $script:systemCursorDir $cursorFile.SafeName
        } else {
            $cursorPaths[$type] = $null
        }
    }

    # 构建方案字符串
    $schemeLine = @()

    foreach ($type in $schemeOrder) {
        $cursorPath = $cursorPaths[$type]
        
    # 如果当前类型缺失，尝试替代
    if ([string]::IsNullOrEmpty($cursorPath)) {
        # 特殊规则1：如果缺少AppStarting，使用Wait光标替代
        if ($type -eq "AppStarting" -and $null -ne $cursorPaths["Wait"]) {
            $cursorPath = $cursorPaths["Wait"]
            $waitName = $typeMapping["Wait"]
            Write-Host "[INFO] 使用Wait光标替代缺少的AppStarting: $waitName" -ForegroundColor Cyan
        }
        
        # 特殊规则2：如果缺少Wait，使用AppStarting光标替代
        if ($type -eq "Wait" -and $null -ne $cursorPaths["AppStarting"]) {
            $cursorPath = $cursorPaths["AppStarting"]
            $appStartingName = $typeMapping["AppStarting"]
            Write-Host "[INFO] 使用AppStarting光标替代缺少的Wait: $appStartingName" -ForegroundColor Cyan
        }
        
        # 特殊规则3：如果缺少Hand、Pin或Person，按照优先级替代
        if ($type -in @("Hand", "Pin", "Person")) {
            # 定义替代优先级
            $priorityOrder = @("Hand", "Pin", "Person")
            
            # 按优先级顺序查找可用的替代光标
            foreach ($altType in $priorityOrder) {
                if ($altType -ne $type -and $null -ne $cursorPaths[$altType]) {
                    $cursorPath = $cursorPaths[$altType]
                    $altName = $typeMapping[$altType]
                    Write-Host "[INFO] 使用${altType}光标替代缺少的${type}: $altName" -ForegroundColor Cyan
                    break
                }
            }
        }
    }
    
    $schemeLine += $cursorPath
    Write-DebugLog "光标类型 $type : $cursorPath"
}

    $schemeString = $schemeLine -join ","
    Write-DebugLog "完整方案字符串: $schemeString"

    # 設置註冊表路徑
    $regKeyCursors = "HKCU:\Control Panel\Cursors"
    $regKeySchemes = "HKCU:\Control Panel\Cursors\Schemes"

    # 確保Schemes鍵存在
    if (-not (Test-Path $regKeySchemes)) {
        Write-DebugLog "創建註冊表項: $regKeySchemes"
        New-Item -Path $regKeySchemes -Force | Out-Null
    }

    # 設置方案
    try {
        Write-DebugLog "設置方案: $themeName = $schemeString"
        Set-ItemProperty -Path $regKeySchemes -Name $themeName -Value $schemeString -ErrorAction Stop
        
        # 验证设置是否成功
        $schemeValue = Get-ItemProperty -Path $regKeySchemes -Name $themeName -ErrorAction Stop
        if ([string]::IsNullOrEmpty($schemeValue.$themeName)) {
            throw "注册表验证失败"
        }
        
        Write-Host "`n  ✓ 註冊表設置已應用"
        Write-DebugLog "驗證 - Scheme: $($schemeValue.$themeName)"
    } catch {
        Write-Host "`n  ✗ 註冊表設置失敗: $_" -ForegroundColor Red
        throw $_  # 重新抛出异常以便上层统计
    }
}


# ========== 主程序開始 ========== #

# 請求管理員權限
# 设置工作目录为脚本所在目录
$currentScript = $null
if ($MyInvocation.MyCommand.Path) {
    $currentScript = $MyInvocation.MyCommand.Path
} elseif ($PSCommandPath) {
    $currentScript = $PSCommandPath
} elseif ([System.Reflection.Assembly]::GetExecutingAssembly().Location) {
    $currentScript = [System.Reflection.Assembly]::GetExecutingAssembly().Location
}

$scriptDir = if ($currentScript -and (Test-Path $currentScript)) {
    Split-Path $currentScript -Parent
} else {
    $PWD.Path
}

# 确保目录存在且可访问
if (Test-Path $scriptDir) {
    Set-Location $scriptDir
} else {
    $scriptDir = $PWD.Path
}

# 添加统计变量
$script:installedThemes = 0
$script:processedFiles = 0
$script:failedThemes = @()
$script:abnormalCountThemes = @()  # 数量异常的方案
$script:unmatchedFilesThemes = @()  # 存在未匹配文件的方案

if ($SignalFile) {
    # 这是被提升权限后的新实例，立即创建信号文件
    $null = New-Item -Path $SignalFile -ItemType File -Force -EA 0
}

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "以管理员权限重新启动脚本..." -ForegroundColor Yellow
    
    $signalFile = Join-Path $env:TEMP "mouse_cursor_$(Get-Random).signal"
    
    # 获取当前可执行文件路径
    $scriptPath = try {
        if ($PSCommandPath) { $PSCommandPath }
        elseif ($MyInvocation.MyCommand.Path) { $MyInvocation.MyCommand.Path }
        else { [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName }
    } catch {
        $MyInvocation.InvocationName
    }
    
    if (-not $scriptPath -or -not (Test-Path $scriptPath -EA 0)) {
        Write-Host "无法获取脚本路径，请手动以管理员权限运行" -ForegroundColor Red
        Read-Host "按Enter键退出"
        exit 1
    }
    
    # 构建参数列表
    $argumentList = @("-SignalFile", $signalFile)
    if ($Debug) { $argumentList += "-Debug" }
    
    try {
        if ($scriptPath -match '\.exe$') {
            # 对于exe文件，直接启动
            Start-Process -FilePath $scriptPath -Verb RunAs -ArgumentList $argumentList
        } else {
            # 对于ps1文件，按优先级选择PowerShell版本
            $powershellExe = $null
            foreach ($cmd in @("wt.exe", "pwsh.exe", "powershell.exe")) {
                if (Get-Command $cmd -EA 0) {
                    $powershellExe = $cmd
                    break
                }
            }
            
            if (-not $powershellExe) {
                throw "找不到PowerShell执行程序"
            }
            
            if ($powershellExe -eq "wt.exe") {
                # 使用Windows Terminal启动pwsh
                $baseArgs = @("pwsh.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $scriptPath) + $argumentList
            } elseif ($powershellExe -eq "powershell.exe") {
                # 使用传统PowerShell，设置黑色背景
                $baseArgs = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", 
                    "& {`$Host.UI.RawUI.BackgroundColor='Black'; Clear-Host; & '$scriptPath' $($argumentList -join ' ')}")
            } else {
                # 使用PowerShell Core
                $baseArgs = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $scriptPath) + $argumentList
            }
            
            Start-Process -FilePath $powershellExe -Verb RunAs -ArgumentList $baseArgs
        }
        
        exit  # 成功启动新实例后退出当前实例
        
    } catch {
        Write-Host "提权启动失败: $_" -ForegroundColor Red
        Read-Host "按Enter键退出"
        exit 1
    }
}

# 清理信号文件
if ($SignalFile) { Remove-Item $SignalFile -Force -EA 0 }

try {
    Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -ErrorAction Stop
} catch {
    Write-Host "执行策略已由系统管理，继续执行..." -ForegroundColor Yellow
}


# 設置控制台編碼為 UTF-8
[Console]::OutputEncoding = [Console]::InputEncoding = [System.Text.Encoding]::UTF8

function Show-Menu {
    Clear-Host
	Write-Host "========================================" -ForegroundColor Cyan
	Write-Host "         鼠标光标批量安装工具 v1.0" -ForegroundColor White
	Write-Host "========================================" -ForegroundColor Cyan
	Write-Host ""
	Write-Host ""
    Write-Host "================ 主菜单 ================" -ForegroundColor Yellow
	Write-Host ""
    Write-Host "1. 使用当前目录 ($scriptDir)"
    Write-Host "2. 输入其他目录"
    Write-Host "3. 显示帮助信息"
    Write-Host "0. 退出程序"
	Write-Host ""
    Write-Host "========================================" -ForegroundColor Yellow
   	Write-Host ""

    $choice = Read-Host "请选择操作 (0-3)"
    return $choice
}

# 帮助信息
function Show-Help {
    Clear-Host
    Write-Host "===================== 帮助信息 =====================" -ForegroundColor Yellow
	Write-Host ""
    Write-Host "此工具用于批量安装鼠标光标方案。"
    Write-Host ""
    Write-Host "使用说明:"
	Write-Host ""
    Write-Host "1. 准备光标文件夹: 每个光标方案应放在单独的文件夹中"
    Write-Host "2. 文件夹内应包含.cur或.ani光标文件"
	Write-Host "3. 该工具会自动以文件夹名作为方案名称"
    Write-Host "4. 该工具会根据文件名自动匹配光标类型"
    Write-Host "5. 需要管理员权限运行"
	Write-Host "6. 输入-debug可以启用详细日志"
	Write-Host ""
    Write-Host "===================================================" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "按Enter键返回主菜单"
}

# 处理用户输入的目录路径
function Get-UserDirectory {
    $userInput = Read-Host "请输入目录路径"
    # 移除可能存在的引号
    $userInput = $userInput.Trim('"', "'")
    
    # 检查路径是否存在
    if (-not (Test-Path -Path $userInput -PathType Container)) {
        Write-Host "目录不存在: $userInput" -ForegroundColor Red
        return $null
    }
    return $userInput
}

# 主菜单交互
$rootPath = $scriptDir
$recurse = $false
$validChoice = $false

do {
    $menuChoice = Show-Menu
    
    switch ($menuChoice) {
        "0" { 
            exit
        }
        "1" { 
            Write-Host "`n使用当前目录: $scriptDir" -ForegroundColor Cyan
            $rootPath = $scriptDir
            $validChoice = $true
            break
        }
        "2" {
            $userPath = $null
            while (-not $userPath) {
                $userPath = Get-UserDirectory
            }
            $rootPath = $userPath
            Write-Host "`n使用目录: $rootPath" -ForegroundColor Cyan
            $validChoice = $true
            break
        }
        "3" {
            Show-Help
        }
		"-debug" {
            $script:debug = -not $script:debug
            if ($script:debug) {
                Write-Host "`nDebug模式已启用" -ForegroundColor Green
                Write-Host "日志将保存到: $PWD\mouse-cursor.log" -ForegroundColor Cyan
            } else {
                Write-Host "`nDebug模式已禁用" -ForegroundColor Yellow
            }
            Write-Host ""
            Read-Host "按Enter键继续"
        }
        default {
            Write-Host "`n无效选择，请重新输入" -ForegroundColor Yellow
            Write-Host "请输入 0-3 之间的数字" -ForegroundColor Yellow
            Write-Host ""
            Read-Host "按Enter键继续"
        }
    }
} while (-not $validChoice)

# 询问是否递归检索子文件夹
$subChoice = Read-Host "`n是否递归检索所有子文件夹? (y/N, 默认N)"
if ($subChoice -eq "y" -or $subChoice -eq "Y") {
    $recurse = $true
    Write-Host "启用递归子文件夹检索" -ForegroundColor Cyan
} else {
    Write-Host "仅检索指定目录下的文件夹" -ForegroundColor Cyan
}

# 設置變數 (更新为使用用户选择的路径)
$script:cursorDir = $rootPath
$script:systemCursorDir = Join-Path $env:SystemRoot "Cursors"

Write-DebugLog "當前工作目錄: $script:cursorDir"
Write-DebugLog "系統光標目錄: $script:systemCursorDir"

Write-Host "正在掃描光標文件..."
Write-Host ""

# 记录开始时间
$startTime = Get-Date

# 獲取需要處理的文件夾
function Get-CursorFolders {
    param([string]$BasePath, [bool]$IsRecursive)
    
    # 根據遞歸選項
    $folders = if ($IsRecursive) {
        Write-Host "递归扫描所有子文件夹..." -ForegroundColor Cyan
        Get-ChildItem -Path $BasePath -Directory -Recurse
    } else {
        Write-Host "扫描当前目录下的文件夹..." -ForegroundColor Cyan
        Get-ChildItem -Path $BasePath -Directory
    }
    
    # 过滤出包含光标文件的文件夹
    return $folders | Where-Object {
        # 使用 Test-Path 進行快速檢查
         (Test-Path -Path "$($_.FullName)\*.cur" -PathType Leaf) -or (Test-Path -Path "$($_.FullName)\*.ani" -PathType Leaf)
    }
}

# 處理文件夾
$cursorFolders = Get-CursorFolders -BasePath $script:cursorDir -IsRecursive $recurse
foreach ($folder in $cursorFolders) {
    Write-DebugLog "找到文件夹: $($folder.FullName)"
    ProcessCursorSet -sourceDir $folder.FullName
}

# 處理當前目錄中的光標文件 (這一步很重要，確保根目錄的文件也能被處理)
# 使用 Where-Object 確保過濾正確
$rootCursorFiles = Get-ChildItem -Path $script:cursorDir -File | 
                 Where-Object { $_.Extension -in ('.ani', '.cur') }

if ($rootCursorFiles) {
    Write-DebugLog "找到根目录中的光标文件 ($($rootCursorFiles.Count) 个)"
    ProcessCursorSet -sourceDir $script:cursorDir
}

# 计算耗时
$endTime = Get-Date
$elapsedTime = $endTime - $startTime

# 显示统计信息
Write-Host ""
Write-Host "========================================"
Write-Host "           安裝統計報告" -ForegroundColor White
Write-Host "========================================"
Write-Host ""
Write-Host "✓ 成功安裝光標方案: $script:installedThemes 個" -ForegroundColor Cyan
Write-Host "✓ 處理光標文件總數: $script:processedFiles 個"
Write-Host ""
Write-Host " 總耗時: $($elapsedTime.TotalSeconds.ToString('F2')) 秒" -ForegroundColor Green
Write-Host ""

if ($script:failedThemes.Count -gt 0) {
    Write-Host "✗ 安裝失敗的方案: $($script:failedThemes.Count) 個" -ForegroundColor Red
    foreach ($failedTheme in $script:failedThemes) {
        Write-Host "  - $failedTheme" -ForegroundColor Yellow
    }
}
Write-Host ""
Write-Host "========================================"
Write-Host ""

# 显示异常情况警告
if ($script:abnormalCountThemes.Count -gt 0 -or $script:unmatchedFilesThemes.Count -gt 0) {
    Write-Host ""
    Write-Host "⚠️  檢測到以下異常情況：" -ForegroundColor Yellow
    Write-Host ""
    
    if ($script:abnormalCountThemes.Count -gt 0) {
        Write-Host "📊 光標數量可能異常的方案 (常見數量: 5/10/15/17)：" -ForegroundColor Yellow
        foreach ($theme in $script:abnormalCountThemes) {
            Write-Host "   • $($theme.ThemeName) - $($theme.FileCount) 個光標文件" -ForegroundColor Yellow
        }
        Write-Host ""
    }
    
    if ($script:unmatchedFilesThemes.Count -gt 0) {
        Write-Host "🔍 存在未匹配光標文件的方案：" -ForegroundColor Yellow
        foreach ($theme in $script:unmatchedFilesThemes) {
            Write-Host "   • $($theme.ThemeName) - 總計 $($theme.TotalFiles) 個文件，匹配到 $($theme.MatchedTypes) 種類型，未匹配 $($theme.UnmatchedCount) 個" -ForegroundColor Yellow
        }
        Write-Host ""
    }
    
    Write-Host "建議檢查上述方案的光標文件命名和數量是否正確。" -ForegroundColor Yellow
    Write-Host ""
}

Write-Host ""

if ($script:installedThemes -gt 0) {
    Write-Host "安裝完成！請到控制面板 > 鼠標 > 指針中選擇新的光標方案。" -ForegroundColor Cyan
    Write-Host ""
    $openMouseProps = Read-Host "是否立即打開鼠標屬性? (Y/n, 默认Y)"
    
    if ($openMouseProps -ne "n" -and $openMouseProps -ne "N") {
        $opened = $false
        
        # 使用control.exe
        if (-not $opened) {
            try {
                Start-Process "control.exe" -ArgumentList "main.cpl,,1" -ErrorAction Stop
                $opened = $true
            } catch { }
        }
        
        # 備用方法: 使用rundll32.exe
        if (-not $opened) {
            try {
                Start-Process "rundll32.exe" -ArgumentList "shell32.dll,Control_RunDLL main.cpl,,1" -ErrorAction Stop
                $opened = $true
            } catch { }
        }
    }
} else {
    Write-Host "未找到任何可安裝的光標方案，請檢查目錄中是否包含 .cur 或 .ani 文件。" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "按任意鍵退出..."
[Console]::ReadKey($true) | Out-Null