# ocr_word.ps1 - Tesseract 5 OCR 取词脚本
# 用法: powershell -File ocr_word.ps1 -ImagePath "xxx.png" -MouseX 150 -MouseY 80 -OutputFile "result.txt"

param(
    [Parameter(Mandatory = $true)][string]$ImagePath,
    [Parameter(Mandatory = $true)][int]$MouseX,
    [Parameter(Mandatory = $true)][int]$MouseY,
    [string]$OutputFile = "",
    [switch]$DebugLog
)

# 输出函数：写入文件（UTF-8）或标准输出
Function WriteResult($text) {
    if ($OutputFile -ne "") {
        [System.IO.File]::WriteAllText($OutputFile, $text, [System.Text.Encoding]::UTF8)
    }
    else {
        Write-Output $text
    }
}

try {
    $tesseractPath = "C:\Program Files\Tesseract-OCR\tesseract.exe"
    if (-not (Test-Path $tesseractPath)) {
        WriteResult '{"found":false,"error":"Tesseract not found"}'
        exit 1
    }

    $absPath = (Resolve-Path $ImagePath).Path

    # 直接使用原始图片进行 OCR，不进行强制灰度和反色预处理
    # Tesseract 5 内部有更好的 Otsu 二值化和处理逻辑，复杂背景下我们的全局反色往往适得其反
    $processedPath = $absPath

    # 设置 UTF-8 编码，避免中文乱码
    $prevEncoding = [Console]::OutputEncoding
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8

    # 调用 Tesseract 输出 TSV 格式（包含每个单词的坐标）
    $tsvOutput = & $tesseractPath $processedPath stdout -l eng+chi_sim --psm 11 tsv 2>$null

    # 恢复编码
    [Console]::OutputEncoding = $prevEncoding
    if (-not $tsvOutput) {
        WriteResult '{"found":false,"error":"Tesseract returned no output"}'
        exit 1
    }

    # 解析 TSV：收集所有单词及其坐标，按行分组
    # TSV 列: level page_num block_num par_num line_num word_num left top width height conf text
    $words = @()
    $lines = @{}

    foreach ($row in $tsvOutput) {
        $cols = $row -split "`t"
        if ($cols.Count -lt 12) { continue }

        $level = $cols[0]
        $lineNum = $cols[4]
        $text = $cols[11]

        # level 5 = 单词级别
        if ($level -ne "5" -or [string]::IsNullOrWhiteSpace($text)) { continue }

        $left = [int]$cols[6]
        $top = [int]$cols[7]
        $width = [int]$cols[8]
        $height = [int]$cols[9]
        $conf = [int]$cols[10]

        # 跳过置信度过低的结果
        if ($conf -lt 10) { continue }

        $wordObj = @{
            Text   = $text
            Left   = $left
            Top    = $top
            Right  = $left + $width
            Bottom = $top + $height
            Line   = $lineNum
        }
        $words += $wordObj

        # 按行号分组，拼接行文本
        if (-not $lines.ContainsKey($lineNum)) {
            $lines[$lineNum] = @()
        }
        $lines[$lineNum] += $text
    }

    # 调试模式：输出所有识别到的单词及坐标到日志文件
    if ($DebugLog) {
        $debugPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "ahk_ocr_debug.log")
        $debugLines = @()
        $debugLines += "=== OCR Debug Log ==="
        $debugLines += "Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        $debugLines += "Image: $ImagePath"
        $debugLines += "Mouse: X=$MouseX, Y=$MouseY"
        $debugLines += "Words found: $($words.Count)"
        $debugLines += "--- All words ---"
        foreach ($w in $words) {
            $debugLines += "  [$($w.Text)] L=$($w.Left) T=$($w.Top) R=$($w.Right) B=$($w.Bottom) Line=$($w.Line)"
        }
        $debugLines += "--- TSV raw (first 50 lines) ---"
        $tsvOutput | Select-Object -First 50 | ForEach-Object { $debugLines += "  $_" }
        [System.IO.File]::WriteAllLines($debugPath, $debugLines, [System.Text.Encoding]::UTF8)
    }

    if ($words.Count -eq 0) {
        WriteResult '{"found":false,"error":"no words recognized"}'
        exit 0
    }

    # 第一遍：精确匹配（鼠标在单词的 bounding box 内）
    $foundWord = $null
    $foundLine = $null
    $padding = 5

    foreach ($w in $words) {
        if ($MouseX -ge ($w.Left - $padding) -and $MouseX -le ($w.Right + $padding) -and
            $MouseY -ge ($w.Top - $padding) -and $MouseY -le ($w.Bottom + $padding)) {
            $foundWord = $w.Text
            $foundLine = ($lines[$w.Line] -join " ")
            break
        }
    }

    # 第二遍：最近距离匹配（同行内最近的单词，垂直容差 20px）
    if (-not $foundWord) {
        $minDist = [double]::MaxValue
        $vPadding = 20
        foreach ($w in $words) {
            if ($MouseY -ge ($w.Top - $vPadding) -and $MouseY -le ($w.Bottom + $vPadding)) {
                $cx = ($w.Left + $w.Right) / 2
                $dist = [Math]::Abs($MouseX - $cx)
                if ($dist -lt $minDist) {
                    $minDist = $dist
                    $foundWord = $w.Text
                    $foundLine = ($lines[$w.Line] -join " ")
                }
            }
        }
    }

    # 第三遍：全局最近词回退（欧几里得距离，限制最大 100px）
    if (-not $foundWord) {
        $minDist = [double]::MaxValue
        $maxGlobalDist = 100
        foreach ($w in $words) {
            $cx = ($w.Left + $w.Right) / 2
            $cy = ($w.Top + $w.Bottom) / 2
            $dist = [Math]::Sqrt(($MouseX - $cx) * ($MouseX - $cx) + ($MouseY - $cy) * ($MouseY - $cy))
            if ($dist -lt $minDist -and $dist -le $maxGlobalDist) {
                $minDist = $dist
                $foundWord = $w.Text
                $foundLine = ($lines[$w.Line] -join " ")
            }
        }
    }

    # 调试模式：输出匹配结果
    if ($DebugLog) {
        $debugPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "ahk_ocr_debug.log")
        $appendLines = @()
        $appendLines += "--- Match result ---"
        if ($foundWord) {
            $appendLines += "  FOUND: [$foundWord] line=[$foundLine]"
        } else {
            $appendLines += "  NOT FOUND: no word at cursor position"
        }
        [System.IO.File]::AppendAllText($debugPath, "`r`n" + ($appendLines -join "`r`n"), [System.Text.Encoding]::UTF8)
    }

    if ($foundWord) {
        $cleanWord = $foundWord -replace '\\', '\\' -replace '"', '\"'
        $cleanLine = $foundLine -replace '\\', '\\' -replace '"', '\"'
        WriteResult "{`"found`":true,`"word`":`"$cleanWord`",`"line`":`"$cleanLine`"}"
    }
    else {
        WriteResult '{"found":false,"error":"no word at cursor position"}'
    }
}
catch {
    $errMsg = $_.Exception.Message -replace '"', '\"'
    WriteResult "{`"found`":false,`"error`":`"$errMsg`"}"
    exit 1
}
