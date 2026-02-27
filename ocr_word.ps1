# ocr_word.ps1 - Windows 内置 OCR 取词脚本
# 用法: powershell -File ocr_word.ps1 -ImagePath "xxx.png" -MouseX 150 -MouseY 80 -OutputFile "result.txt"

param(
    [Parameter(Mandatory = $true)][string]$ImagePath,
    [Parameter(Mandatory = $true)][int]$MouseX,
    [Parameter(Mandatory = $true)][int]$MouseY,
    [string]$OutputFile = ""
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
    # 加载 WinRT 相关程序集
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    
    $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
    $null = [Windows.Storage.FileAccessMode, Windows.Storage, ContentType = WindowsRuntime]
    $null = [Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
    $null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
    $null = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
    $null = [Windows.Media.Ocr.OcrEngine, Windows.Media.Ocr, ContentType = WindowsRuntime]
    $null = [Windows.Media.Ocr.OcrResult, Windows.Media.Ocr, ContentType = WindowsRuntime]
    $null = [Windows.Globalization.Language, Windows.Globalization, ContentType = WindowsRuntime]

    $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() |
        Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
            $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]

    Function Await($WinRtTask, $ResultType) {
        $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
        $netTask = $asTask.Invoke($null, @($WinRtTask))
        $netTask.Wait(-1) | Out-Null
        $netTask.Result
    }

    $absPath = (Resolve-Path $ImagePath).Path

    $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($absPath)) ([Windows.Storage.StorageFile])
    $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])

    $ocrEngine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
    if (-not $ocrEngine) {
        $lang = [Windows.Globalization.Language]::new("en-US")
        $ocrEngine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
    }

    if (-not $ocrEngine) {
        WriteResult '{"found":false,"error":"OCR engine not available"}'
        exit 1
    }

    $result = Await ($ocrEngine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])

    $foundWord = $null
    $foundLine = $null

    foreach ($line in $result.Lines) {
        $lineText = $line.Text
        foreach ($word in $line.Words) {
            $rect = $word.BoundingRect
            $x1 = $rect.X
            $y1 = $rect.Y
            $x2 = $rect.X + $rect.Width
            $y2 = $rect.Y + $rect.Height

            $padding = 5
            if ($MouseX -ge ($x1 - $padding) -and $MouseX -le ($x2 + $padding) -and
                $MouseY -ge ($y1 - $padding) -and $MouseY -le ($y2 + $padding)) {
                $foundWord = $word.Text
                $foundLine = $lineText
                break
            }
        }
        if ($foundWord) { break }
    }

    if (-not $foundWord) {
        $minDist = [double]::MaxValue
        foreach ($line in $result.Lines) {
            $lineText = $line.Text
            foreach ($word in $line.Words) {
                $rect = $word.BoundingRect
                $vPadding = 10
                if ($MouseY -ge ($rect.Y - $vPadding) -and $MouseY -le ($rect.Y + $rect.Height + $vPadding)) {
                    $cx = $rect.X + $rect.Width / 2
                    $dist = [Math]::Abs($MouseX - $cx)
                    if ($dist -lt $minDist) {
                        $minDist = $dist
                        $foundWord = $word.Text
                        $foundLine = $lineText
                    }
                }
            }
        }
    }

    $stream.Dispose()

    if ($foundWord) {
        $cleanWord = $foundWord -replace '[^\p{L}\p{N}''`-]', ''
        $cleanWord = $cleanWord -replace '\\', '\\' -replace '"', '\"'
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
