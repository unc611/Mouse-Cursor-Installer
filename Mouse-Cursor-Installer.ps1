# 鼠標光標批量安裝工具 v1.1.2
# 需要管理員權限運行
# 作者：uncherry (https://github.com/unc611/Mouse-Cursor-Installer)
# 許可證：MIT license

param(
    [string]$SignalFile,
	[string]$launcherScriptPath
)

$script:debug = $args -contains '-log'

# 刪除已存在的日誌文件
function Initialize-Log {
    if ($script:debug) {
        $logPath = Join-Path $scriptDir 'mouse-cursor.log'
        Remove-Item -LiteralPath $logPath -EA 0
    }
}

function Write-DebugLog {
    param([string]$message)
    if ($script:debug) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logPath = Join-Path $scriptDir 'mouse-cursor.log'
        "[$timestamp] $message" | Out-File -LiteralPath $logPath -Append -Encoding utf8
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
            "diagonalresize1", "diagonalresize", "斜めに拡大縮小1", "斜めに拡大縮小", "沿对角线调整大小1", "沿对角线调整大小"
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
            "位置选择", "pin", "場所の選択", "場所", "location"
        )
        "Person" = @(
            "个人选择", "person", "人の選択", "人"
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

# 定义函数：处理一个光标集合（一个主题）
function ProcessCursorSet {
    param([string]$sourceDir)

    Write-Host "`n处理目录: $sourceDir"

    # 获取目录名作为主题名
    $themeName = [System.IO.Path]::GetFileName($sourceDir)

    # 清理主题名称中的特殊字符
    $originalThemeName = $themeName
    $themeName = $themeName -replace '\[|\]|\(|\)|&|%', ''

    Write-DebugLog "原始主题名称: '$originalThemeName'"
    Write-DebugLog "清理后主题名称: '$themeName'"

    # 收集光标文件 (.ani 和 .cur)
    $cursorFiles = Get-ChildItem -LiteralPath $sourceDir -File | Where-Object { $_.Extension -in '.ani', '.cur' } | Sort-Object Name
    Write-Host "主题名称: $themeName (找到 $($cursorFiles.Count) 个光标文件)"
	
    # 添加统计
    $script:processedFiles += $cursorFiles.Count
    $themeInstallSuccess = $true

    # 获取所有文件名的基本部分（不含扩展名）
    $baseNames = $cursorFiles | ForEach-Object { 
        [System.IO.Path]::GetFileNameWithoutExtension($_.Name) 
    }

    # 计算公共前缀和后缀
    $prefixSuffix = Get-CommonPrefixSuffix -strings $baseNames
    $commonPrefix = $prefixSuffix[0]
    $commonSuffix = $prefixSuffix[1]
    
    Write-DebugLog "公共前缀: '$commonPrefix'"
    Write-DebugLog "公共后缀: '$commonSuffix'"
    Write-Host ""

    # 預處理階段：分析所有文件名
    $strippedNames = @()
    foreach ($file in $cursorFiles) {
        $baseName = [IO.Path]::GetFileNameWithoutExtension($file.Name)
        
        # 檢查文件名長度是否小於前綴和後綴長度之和
        $totalPrefixSuffixLength = $commonPrefix.Length + $commonSuffix.Length
        if ($baseName.Length -lt $totalPrefixSuffixLength) {
            $baseName = $commonPrefix + $commonSuffix
            Write-DebugLog "前後綴可能重複，使用組合名稱: $baseName"
        }

        # 移除公共前綴和后缀
        $strippedName = $baseName
        if (-not [string]::IsNullOrEmpty($commonPrefix)) {
            $strippedName = $strippedName -replace "^$([regex]::Escape($commonPrefix))", ''
        }
        if (-not [string]::IsNullOrEmpty($commonSuffix)) {
            $strippedName = $strippedName -replace "$([regex]::Escape($commonSuffix))$", ''
        }
        
        # 如果去除後為空，則明確指定為 "arrow"
        if ([string]::IsNullOrWhiteSpace($strippedName)) {
            $strippedName = "arrow"
        }

        $strippedNames += $strippedName
    }

    # === 核心決策：獲取並執行最終策略 ===
    $matchingStrategy = AnalyzeMatchingStrategy -strippedNames $strippedNames
    
    if ($matchingStrategy -eq "Failed") {
        Write-Host "  ✗ 安裝失敗: $originalThemeName" -ForegroundColor Red
        Write-Host "    原因: 無法確定命名規則 (既不符合關鍵詞特徵，也不符合數字序號特徵)。" -ForegroundColor Yellow
        $script:failedThemes += $originalThemeName
        return # 立即終止此方案的處理
    }

    Write-Host "光标类型匹配结果："
    Write-Host "使用策略: $matchingStrategy"
    Write-Host ""

    # === 根據已確定的策略進行匹配與處理 ===
    $typeMapping = @{}; $alternativeMapping = @{}; $trulyUnmatchedFiles = @()
    $fileClassification = @{}

    $fileIndex = 0
    foreach ($file in $cursorFiles) {
        $fileIndex++
        $result = DetermineCursorType -fileName $strippedNames[$fileIndex-1] -strategy $matchingStrategy
        $cursorType = $result.Type

        if ($cursorType) {
            if ($typeMapping.ContainsKey($cursorType)) {
                if (-not $alternativeMapping.ContainsKey($cursorType)) { $alternativeMapping[$cursorType] = @() }
                $alternativeMapping[$cursorType] += $file.Name
                $fileClassification[$file.Name] = @{ Type = $cursorType; IsAlternative = $true }
            } else {
                $typeMapping[$cursorType] = $file.Name
                $fileClassification[$file.Name] = @{ Type = $cursorType; IsAlternative = $false }
            }
        } else {
            $trulyUnmatchedFiles += $file.Name
            $fileClassification[$file.Name] = @{ Type = "Unmatched"; IsAlternative = $false }
        }
    }

    # === 文件複製與生成最終映射 ===
    $mapping = @()
    foreach ($file in $cursorFiles) {
        $classification = $fileClassification[$file.Name]
        if ($classification.Type -eq "Unmatched") { continue }
        
        $originalName = $file.Name; $cursorType = $classification.Type; $isAlternative = $classification.IsAlternative
        $typeIndex = $script:schemeOrder.IndexOf($cursorType)
        $safeNameNumber = if ($typeIndex -ge 0) { $typeIndex + 1 } else { 99 }
        $safeNameSuffix = ""
        if ($isAlternative) {
            $altIndex = if ($alternativeMapping[$cursorType]) { $alternativeMapping[$cursorType].IndexOf($originalName) + 1 } else { 1 }
            $safeNameSuffix = "_alt$altIndex"
        }
        $safeName = "${themeName}_${safeNameNumber}${safeNameSuffix}$($file.Extension)"

        $mapping += [PSCustomObject]@{ OriginalName = $originalName; SafeName = $safeName; CursorType = $cursorType; IsAlternative = $isAlternative }
        
        try {
            [System.IO.File]::Copy($file.FullName, (Join-Path $script:systemCursorDir $safeName), $true)
        } catch {
            Write-Host "    ✗ 文件复制失败: $($file.Name) - $($_.Exception.Message)" -ForegroundColor Red
            $themeInstallSuccess = $false
        }
    }

    # === 統計與報告 ===
    # 檢查匹配到的光标类型数量是否异常
    if ($typeMapping.Count -notin @(10, 13, 15, 17)) {
        $script:abnormalCountThemes += [PSCustomObject]@{
            ThemeName    = $themeName
            MatchedCount = $typeMapping.Count # 使用匹配到的類型數量
        }
    }

    if ($trulyUnmatchedFiles.Count -gt 0) {
        $script:unmatchedFilesThemes += [PSCustomObject]@{
            ThemeName      = $themeName
            TotalFiles     = $cursorFiles.Count
            MatchedTypes   = $typeMapping.Count
            UnmatchedCount = $trulyUnmatchedFiles.Count
        }
    }

    if ($alternativeMapping.Keys.Count -gt 0) {
        $script:themesWithAlternatives += [PSCustomObject]@{
            ThemeName        = $themeName
            AlternativeCount = ($alternativeMapping.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
            AlternativeTypes = $alternativeMapping.Keys -join ', '
        }
    }

    # === 顯示匹配詳情與註冊表寫入 ===
    foreach ($type in $script:schemeOrder) {
        $typeLabel = "$type ($($script:cursorTypeChinese[$type]))"
        if ($typeMapping[$type]) {
            Write-Host "类型: $typeLabel -> $($typeMapping[$type])"
            if ($alternativeMapping[$type]) { Write-Host "      备选: $($alternativeMapping[$type] -join ', ')" -ForegroundColor Yellow }
        } else {
            Write-Host "类型: $typeLabel -> " -ForegroundColor Magenta
        }
    }

    if ($themeInstallSuccess) {
        try {
            CreateRegistryEntries -themeName $themeName -mapping $mapping -typeMapping $typeMapping
            $script:installedThemes++; Write-Host "  ✓ 完成主题: $themeName" -ForegroundColor Green
        } catch {
            $script:failedThemes += $originalThemeName; Write-Host "  ✗ 安装失败 (注册表写入错误): $themeName" -ForegroundColor Red
        }
    } else {
        $script:failedThemes += $originalThemeName; Write-Host "  ✗ 安装失败: $themeName (由于文件复制错误)" -ForegroundColor Red
    }
	Write-Host ""
}

# 計算公共前後綴
function Get-CommonPrefixSuffix {
    param([string[]]$strings)
    
    # 如果文件少於2個，不可能有“公共”前後綴，直接返回空。
    if ($strings.Count -lt 2) { 
        return @("", "") 
    }
    
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

# 分析匹配策略的函數
function AnalyzeMatchingStrategy {
    param([string[]]$strippedNames)
    
    # 用於策略分析時，需要排除的僅與UpArrow相關的修飾性關鍵詞
    $upArrowKeywords = @("alternateselec", "alternate", "代替", "uparrow", "up")
    
    $substantiveKeywordMatches = 0
    $distinctNumbers = @{} # 使用哈希表來統計不重複的數字

    foreach ($name in $strippedNames) {
        $processedName = $name.ToLower() -replace '[\s・·]', ''

        # 檢查實質性關鍵詞
        $foundSubstantiveKeyword = $false
        foreach ($keyword in $script:sortedKeywords) {
            if ($processedName -match [regex]::Escape($keyword.ToLower())) {
                # 如果匹配到的關鍵詞不在UpArrow的排除列表中，則視為實質性匹配
                if ($keyword.ToLower() -notin $upArrowKeywords) {
                    $foundSubstantiveKeyword = $true
                    break # 找到一個實質關鍵詞就足夠了
                }
            }
        }
        if ($foundSubstantiveKeyword) {
            $substantiveKeywordMatches++
        }

        # 提取並統計不同的數字
        if ($processedName -match '(\d+)') {
            $distinctNumbers[$matches[1]] = $true
        }
    }

    # 只要有任何一個實質性關鍵詞匹配，就使用關鍵詞策略。
    if ($substantiveKeywordMatches -gt 0) {
        Write-DebugLog "決策：檢測到 $substantiveKeywordMatches 個實質性關鍵詞，採用 Keyword 策略。"
        return "Keyword"
    }

    # 如果沒有實質性關鍵詞，但找到了3個或以上不同的數字，則使用數字策略。
    if ($distinctNumbers.Count -ge 3) {
        Write-DebugLog "決策：未檢測到實質性關鍵詞，但找到 $($distinctNumbers.Count) 個不同數字，採用 Number 策略。"
        return "Number"
    }

    # 以上條件均不滿足，則判定為無法識別的失敗方案。
    Write-DebugLog "決策：未找到足夠的關鍵詞或數字特徵，判定為匹配失敗。"
    return "Failed"
}

# 定義函數：根據文件名猜測光標類型
function DetermineCursorType {
    param(
        [string]$fileName,
        [string]$strategy # 只接收策略，不再需要文件序號作為匹配依據
    )
    
    # 處理文件名：转换为小写并去除空格
    $processedFileName = $fileName.ToLower() -replace '[\s・·]', ''
    
    if ($strategy -eq "Keyword") {
        # 關鍵詞策略：嚴格按排序後的關鍵詞列表匹配
        foreach ($keyword in $script:sortedKeywords) {
            if ($processedFileName -match [regex]::Escape($keyword.ToLower())) {
                # 找到第一個（即最長的）匹配項，立即返回其類型
                return @{ Type = $script:keywordMap[$keyword] }
            }
        }
        # 如果遍歷完所有關鍵詞都沒有匹配，則返回 null 表示匹配失敗
        return @{ Type = $null }

    } elseif ($strategy -eq "Number") {
        # 數字序號策略：僅當文件名中明確包含可識別的數字時才匹配
        if ($processedFileName -match '(\d+)') {
            $extractedNumber = [int]$matches[1]
            
            # 光標類型與數字的映射表
            $numberToType = @{
                1 = "Arrow"; 2 = "Help"; 3 = "AppStarting"; 4 = "Wait"; 5 = "Crosshair";
                6 = "IBeam"; 7 = "NWPen"; 8 = "No"; 9 = "SizeNS"; 10 = "SizeWE";
                11 = "SizeNWSE"; 12 = "SizeNESW"; 13 = "SizeAll"; 14 = "UpArrow";
                15 = "Hand"; 16 = "Pin"; 17 = "Person"
            }
            
            $targetType = $numberToType[$extractedNumber]
            # 只有當提取的數字在映射表中時才返回類型
            if ($targetType) {
                return @{ Type = $targetType }
            }
        }
        
        # 如果文件名中沒有數字，或數字不在映射表中，則返回 null 表示匹配失敗
        return @{ Type = $null }
    }
    
    # 如果傳入了未知的策略，同樣返回失敗
    return @{ Type = $null }
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

[ValidateSet('auto', 'wt', 'pwsh', 'powershell')]
$PowerShellMode = "auto" # 修改這裡快速切換啟動方式

# --- 全局變數 ---
$script:installedThemes = 0
$script:processedFiles = 0
$script:failedThemes = @()
$script:abnormalCountThemes = @()
$script:unmatchedFilesThemes = @()
$script:themesWithAlternatives = @()

# ------------------- 初始化與路徑設定 -------------------

# 使用 $PSCommandPath
if ($PSCommandPath) {
    $currentScriptPath = $PSCommandPath
} else {
    $currentScriptPath = ([System.Reflection.Assembly]::GetEntryAssembly()).Location
    $currentScriptPath = (New-Object System.Uri($currentScriptPath)).LocalPath
}

try {
    $scriptDir = Split-Path $currentScriptPath -Parent -ErrorAction Stop
    Set-Location -LiteralPath $scriptDir
} catch {
    Write-Host "警告: 無法確定或訪問腳本所在目錄。腳本將在當前目錄 '$($PWD.Path)' 繼續執行。" -ForegroundColor Yellow
}

Initialize-Log
Write-DebugLog "powershell啟動模式：$PowerShellMode"
Write-DebugLog "腳本路徑：$currentScriptPath"

# ------------------- 管理員權限檢查與提權邏輯 -------------------

if ($SignalFile) {
    $null = New-Item -Path $SignalFile -ItemType File -Force -EA 0
}

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "需要管理員權限，正在嘗試自動提權..." -ForegroundColor Yellow
    
    if (-not (Test-Path -LiteralPath $currentScriptPath)) {
        Write-Host "致命錯誤: 無法獲取主腳本的有效路徑 '$currentScriptPath'。請嘗試手動以管理員身份運行。" -ForegroundColor Red
        Read-Host "按 Enter 鍵退出。"
        exit 1
    }

    try {
        $signalFileForElevated = Join-Path $env:TEMP "mouse_cursor_$(Get-Random).signal"
        $argumentsForElevated = @("-SignalFile", $signalFileForElevated)
        if ($script:debug) { $argumentsForElevated += "-log" }

        # --- 檢查腳本是否為 .exe ---
        if ($currentScriptPath -match '\.exe$') {

            $psi = [System.Diagnostics.ProcessStartInfo]@{
                FileName = $currentScriptPath
                Arguments = $argumentsForElevated -join " "
                Verb = "runas"
                UseShellExecute = $true
            }
            [System.Diagnostics.Process]::Start($psi) | Out-Null

        } else {
            # --- 如果是 .ps1 腳本，則執行 PowerShell 啟動器邏輯 ---
            $powershellExePath = $null
            $exeName = ''
            
            $priorityOrder = @("wt.exe", "pwsh.exe", "powershell.exe")
            $searchOrder = switch ($PowerShellMode.ToLower()) {
                'wt'         { @("wt.exe") + $priorityOrder }
                'pwsh'       { @("pwsh.exe") + $priorityOrder }
                'powershell' { @("powershell.exe") + $priorityOrder }
                default      { $priorityOrder }
            }

            foreach ($cmd in ($searchOrder | Get-Unique)) {
                $foundCommand = Get-Command $cmd -EA 0
                if ($foundCommand) {
                    $powershellExePath = $foundCommand.Source
                    $exeName = $cmd
                    break
                }
            }

            if (-not $powershellExePath) { throw "在系統 PATH 中找不到任何可用的 PowerShell 執行程序。" }

            if ($exeName -in @('wt.exe', 'pwsh.exe')) {
                $startArgs = if ($exeName -eq 'wt.exe') {
                    @("pwsh.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$currentScriptPath`"") + $argumentsForElevated
                } else { # pwsh.exe
                    @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$currentScriptPath`"") + $argumentsForElevated
                }
                Start-Process -FilePath $powershellExePath -Verb RunAs -ArgumentList $startArgs
            } else { # powershell.exe
                $launcherScriptPath = Join-Path $env:TEMP "Launcher_$(Get-Random).ps1"
                $launcherContent = @"
Param([string]`$TargetDirectory, [string]`$TargetName)
try { 
    Set-Location -LiteralPath `$TargetDirectory
	& "`.`\`$TargetName" '$launcherScriptPath' @args 
	}
finally { Remove-Item -LiteralPath `$MyInvocation.MyCommand.Path -Force -EA 0 }
"@
                Set-Content -LiteralPath $launcherScriptPath -Value $launcherContent -Encoding UTF8 -Force

                $finalArgumentList = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$launcherScriptPath`"", "`"$scriptDir`"", "`"$(Split-Path $currentScriptPath -Leaf)`"") + $argumentsForElevated
                Start-Process -FilePath $powershellExePath -Verb RunAs -ArgumentList $finalArgumentList
            }
        }
        
        exit # 成功啟動新實例後，退出當前非管理員實例
    } catch {
		if ($launcherScriptPath) { Remove-Item $launcherScriptPath -Force -EA 0 }
        Write-Host "提權啟動失敗: $_" -ForegroundColor Red
        Read-Host "按 Enter 鍵退出。"
        exit 1
    }
}

# =================== 提權後的主腳本邏輯 ===================

if ($SignalFile) { Remove-Item $SignalFile -Force -EA 0 }
if ($launcherScriptPath) { Remove-Item $launcherScriptPath -Force -EA 0 }

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
	Write-Host "        鼠标光标批量安装工具 v1.1.2" -ForegroundColor White
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
    if ($Debug) {
		Write-Host "调试模式已开启，选择3查看帮助信息。" -ForegroundColor Cyan
		Write-Host ""
	}
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
	Write-Host "6. 输入-log可以启用调试模式，详细日志将位于脚本目录下"
	Write-Host ""
    Write-Host "===================================================" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "按Enter键返回主菜单"
}

# 处理用户输入的目录路径
function Get-UserDirectory {
    $userInput = Read-Host "请输入目录路径"
    # 移除可能存在的引号
    if ($userInput.StartsWith("'") -and $userInput.EndsWith("'")) {
    $userInput = $userInput.Substring(1, $userInput.Length - 2)
    }
    # 检查路径是否存在
    if (-not (Test-Path -LiteralPath $userInput -PathType Container)) {
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
		"-log" {
            $script:debug = -not $script:debug
            if ($script:debug) {
                Write-Host "`n调试模式已启用" -ForegroundColor Green
                Write-Host "日志将保存到: $PWD\mouse-cursor.log" -ForegroundColor Cyan
            } else {
                Write-Host "`n调试模式已禁用" -ForegroundColor Yellow
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

Write-Host "`n正在掃描光標文件..."
Write-Host ""

# 记录开始时间
$startTime = Get-Date

# 获取需要处理的文件夹
function Get-CursorFolders {
    param([string]$BasePath, [bool]$IsRecursive)
    
    # 根据递归选项
    $folders = if ($IsRecursive) {
        Write-Host "递归扫描所有子文件夹..." -ForegroundColor Cyan
        Get-ChildItem -LiteralPath $BasePath -Directory -Recurse
    } else {
        Write-Host "扫描当前目录下的文件夹..." -ForegroundColor Cyan
        Get-ChildItem -LiteralPath $BasePath -Directory
    }
    
    # 过滤出包含光标文件的文件夹
    return $folders | Where-Object {

        $curFiles = Get-ChildItem -LiteralPath $_.FullName -Filter "*.cur" -File -EA 0
        $aniFiles = Get-ChildItem -LiteralPath $_.FullName -Filter "*.ani" -File -EA 0
        
        return ($curFiles.Count -gt 0) -or ($aniFiles.Count -gt 0)
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
$rootCursorFiles = Get-ChildItem -LiteralPath $script:cursorDir -File | 
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
    Write-Host ""
	}

Write-Host "========================================"
Write-Host ""

if ($script:themesWithAlternatives.Count -gt 0) {
    Write-Host "💡 檢測到以下方案可能包含額外的備選光標：" -ForegroundColor Cyan
    Write-Host "   (備選光標指多個文件匹配到同一類型，僅第一個被設為默認)" -ForegroundColor Gray

    foreach ($theme in $script:themesWithAlternatives) {
        Write-Host "`n  - $($theme.ThemeName): $($theme.AlternativeCount) 個備選文件" -ForegroundColor White
        Write-Host "    涉及類型: $($theme.AlternativeTypes)"
    }
    Write-Host ""
    Write-Host "========================================"
    Write-Host ""
}

# 显示异常情况警告
if ($script:abnormalCountThemes.Count -gt 0 -or $script:unmatchedFilesThemes.Count -gt 0) {
    Write-Host ""
    Write-Host "⚠️  檢測到以下異常情況：" -ForegroundColor Yellow
    Write-Host ""
	
    if ($script:abnormalCountThemes.Count -gt 0) {
        Write-Host "📊 光標類型數量可能異常的方案 (常見數量: 10/13/15/17)：" -ForegroundColor Yellow
        foreach ($theme in $script:abnormalCountThemes) {
            Write-Host "   • $($theme.ThemeName) - 已匹配 $($theme.MatchedCount) 種類型" -ForegroundColor Yellow
        }
        Write-Host ""
    }

    if ($script:unmatchedFilesThemes.Count -gt 0) {
        Write-Host "🔍 存在無法識別其類型的光標文件：" -ForegroundColor Yellow
        foreach ($theme in $script:unmatchedFilesThemes) {
            Write-Host "   • $($theme.ThemeName) - 其中 $($theme.UnmatchedCount) 個文件無法識別類型。" -ForegroundColor Yellow
            Write-Host "     (總計 $($theme.TotalFiles) 個文件，成功匹配 $($theme.MatchedTypes) 種類型)" -ForegroundColor Yellow
        }
        Write-Host ""
    }
    
    Write-Host "建議檢查上述方案的光標文件命名和數量是否正確。" -ForegroundColor Yellow
    Write-Host ""
}

if ($script:failedThemes.Count -gt 0) {
    Write-Host ""
    Write-Host "--------- 以下方案安裝失敗，請檢查日誌 ---------" -ForegroundColor Red
    foreach ($failedTheme in $script:failedThemes) {
        Write-Host "  - $failedTheme" -ForegroundColor Red
    }
    Write-Host "------------------------------------------------" -ForegroundColor Red
	Write-Host ""
}

if ($script:installedThemes -gt 0) {
    Write-Host "安裝完成！請到控制面板 > 鼠標 > 指針中選擇新的光標方案。" -ForegroundColor Cyan
    Write-Host ""
    $openMouseProps = Read-Host "是否立即打開鼠標屬性? (Y/n, 默认Y)"
    
    if ($openMouseProps -ne "n" -and $openMouseProps -ne "N") {
        $opened = $false
        
        # 使用control.exe
        if (-not $opened) {
            try {
                Start-Process "control.exe" -ArgumentList "main.cpl,,1" -WorkingDirectory $env:SystemRoot -ErrorAction Stop
                $opened = $true
            } catch { }
        }
        
        # 備用方法: 使用rundll32.exe
        if (-not $opened) {
            try {
                Start-Process "rundll32.exe" -ArgumentList "shell32.dll,Control_RunDLL main.cpl,,1" -WorkingDirectory $env:SystemRoot -ErrorAction Stop
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