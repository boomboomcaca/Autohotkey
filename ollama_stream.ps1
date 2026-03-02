# ollama_stream.ps1 - Ollama 流式请求公用脚本
# 用法: powershell -File ollama_stream.ps1 -JsonFile "request.json" -OutputFile "response.txt"

param(
    [Parameter(Mandatory = $true)][string]$JsonFile,
    [Parameter(Mandatory = $true)][string]$OutputFile
)

try {
    $body = Get-Content -Path $JsonFile -Raw -Encoding UTF8
    $utf8 = [System.Text.Encoding]::UTF8
    $bytes = $utf8.GetBytes($body)

    $req = [System.Net.HttpWebRequest]::Create('http://localhost:11434/api/generate')
    $req.Method = 'POST'
    $req.ContentType = 'application/json'
    $req.ContentLength = $bytes.Length

    $reqStream = $req.GetRequestStream()
    $reqStream.Write($bytes, 0, $bytes.Length)
    $reqStream.Close()

    $resp = $req.GetResponse()
    $reader = New-Object System.IO.StreamReader($resp.GetResponseStream())

    $fs = New-Object System.IO.FileStream($OutputFile, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
    $sw = New-Object System.IO.StreamWriter($fs, [System.Text.Encoding]::UTF8)

    while (-not $reader.EndOfStream) {
        $line = $reader.ReadLine()
        $sw.WriteLine($line)
        $sw.Flush()
    }

    $sw.Close()
    $fs.Close()
    $reader.Close()
    $resp.Close()
}
catch {
    # 静默失败，AHK 端通过检查输出文件判断结果
    exit 1
}
