;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; 鼠标取词 + 语境解释 - Alt+W：截取鼠标所在窗口 → Windows OCR → 定位单词 → Ollama 解释
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

#Include "Gdip_All.ahk"

; ===== 全局变量 =====
g_WL_Gui := ""
g_WL_ResultCtrl := ""
g_WL_TitleCtrl := ""
g_WL_TtsIcon := ""
WL_CurrentWord := ""
g_WL_StreamFile := ""
g_WL_StreamPid := 0
g_WL_Pending := false
g_WL_StreamContent := ""

; ===== 快捷键 Alt+F1 =====
!F1::
{
  global g_WL_Gui, g_WL_ResultCtrl, g_WL_TitleCtrl
  global g_WL_StreamFile, g_WL_StreamPid, g_WL_Pending, g_WL_StreamContent

  ; 如果浮窗已存在，先关闭
  if (g_WL_Gui != "") {
    CloseWordGui()
  }

  ; 1. 获取鼠标绝对坐标
  MouseGetPos(&mouseX, &mouseY, &winUnder)

  ; 2. 获取鼠标所在窗口的位置和大小
  try {
    WinGetPos(&winX, &winY, &winW, &winH, "ahk_id " . winUnder)
  } catch {
    ToolTip("无法获取窗口信息")
    SetTimer(ToolTip, -2000)
    return
  }

  ; 3. 限制截图最大尺寸（防止超大窗口导致 OCR 过慢）
  maxW := 1920
  maxH := 1080
  if (winW > maxW || winH > maxH) {
    ; 窗口太大，改为截取鼠标周围区域
    dpi := DllCall("GetDpiForWindow", "Ptr", winUnder, "UInt")
    if (dpi < 96)
      dpi := 96
    scale := dpi / 96
    captureW := Round(600 * scale)
    captureH := Round(400 * scale)
    winX := mouseX - Round(captureW / 2)
    winY := mouseY - Round(captureH / 2)
    if (winX < 0)
      winX := 0
    if (winY < 0)
      winY := 0
    winW := captureW
    winH := captureH
  }

  ; 4. 计算鼠标在截图中的相对坐标
  relMouseX := mouseX - winX
  relMouseY := mouseY - winY

  ; 5. 启动 GDI+
  pToken := Gdip_Startup()
  if (!pToken) {
    ToolTip("GDI+ 启动失败")
    SetTimer(ToolTip, -2000)
    return
  }

  ; 6. 截取屏幕区域
  screenRect := winX . "|" . winY . "|" . winW . "|" . winH
  pBitmap := Gdip_BitmapFromScreen(screenRect)

  if (!pBitmap) {
    Gdip_Shutdown(pToken)
    ToolTip("截图失败")
    SetTimer(ToolTip, -2000)
    return
  }

  ; 7. 保存为临时 PNG
  tempImg := A_Temp . "\ahk_word_capture.png"
  Gdip_SaveBitmapToFile(pBitmap, tempImg)
  Gdip_DisposeImage(pBitmap)
  Gdip_Shutdown(pToken)

  ; 8. 显示"正在识别..."提示
  ToolTip("🔍 正在识别...")

  ; 9. 调用 PowerShell OCR 脚本（同步等待结果）
  ocrScript := A_ScriptDir . "\ocr_word.ps1"
  ocrResultFile := A_Temp . "\ahk_ocr_result.txt"
  try FileDelete(ocrResultFile)

  ; 使用 -OutputFile 参数让 PowerShell 直接以 UTF-8 写入结果文件
  psCmd := 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' . ocrScript . '" -ImagePath "' . tempImg . '" -MouseX ' . relMouseX . ' -MouseY ' . relMouseY . ' -OutputFile "' . ocrResultFile . '"'

  RunWait(psCmd, , "Hide")

  ; 清除提示
  ToolTip()

  ; 10. 读取 OCR 结果
  ocrOutput := ""
  try {
    ocrOutput := Trim(FileRead(ocrResultFile, "UTF-8"))
  }

  if (ocrOutput = "") {
    ToolTip("OCR 无结果（输出文件为空）")
    SetTimer(ToolTip, -3000)
    return
  }

  ; 11. 解析 JSON 结果
  word := ""
  line := ""
  found := false

  if (RegExMatch(ocrOutput, '"found"\s*:\s*true'))
    found := true

  if (found) {
    if (RegExMatch(ocrOutput, '"word"\s*:\s*"((?:[^"\\]|\\.)*)"', &m))
      word := m[1]
    if (RegExMatch(ocrOutput, '"line"\s*:\s*"((?:[^"\\]|\\.)*)"', &m))
      line := m[1]
  }

  if (!found || word = "") {
    ToolTip("未识别到单词")
    SetTimer(ToolTip, -2000)
    return
  }

  ; 12. 显示浮窗并请求 Ollama
  ShowWordPopup(word, line, mouseX, mouseY)
}

; ===== 显示取词浮窗 =====
ShowWordPopup(word, context, posX, posY)
{
  global g_WL_Gui, g_WL_ResultCtrl, g_WL_TitleCtrl, g_WL_TtsIcon, WL_CurrentWord
  WL_CurrentWord := word

  g_WL_Gui := Gui("+AlwaysOnTop -Caption +Border +Owner")
  g_WL_Gui.BackColor := "FFFFFF"
  g_WL_Gui.MarginX := 12
  g_WL_Gui.MarginY := 8

  ; 标题行水平排列
  g_WL_Gui.SetFont("s14 c1a1a2e Bold", "Microsoft YaHei")
  
  ; 在标题右侧添加朗读图标
  g_WL_Gui.AddText("w120 Section", word)
  g_WL_TtsIcon := g_WL_Gui.AddText("x+5 ys c888888", "🔊")
  
  ; 关联朗读事件（点击也可朗读）
  if (g_WL_TtsIcon) {
    g_WL_TtsIcon.OnEvent("Click", (*) => PlayTtsText(word))
  }

  ; 语境行
  if (context != "" && context != word) {
    g_WL_Gui.SetFont("s9 c888888 Norm", "Microsoft YaHei")
    contextCtrl := g_WL_Gui.AddText("xs w320", "📖 " . context)
  }

  ; 分隔线
  g_WL_Gui.SetFont("s1 cCCCCCC", "Microsoft YaHei")
  g_WL_Gui.AddText("xs w320 0x10")  ; SS_ETCHEDHORZ

  ; 结果区域
  g_WL_Gui.SetFont("s10 c333333 Norm", "Microsoft YaHei")
  g_WL_ResultCtrl := g_WL_Gui.AddText("xs w320 h180", "⏳ 正在查询...")

  ; 底部提示
  g_WL_Gui.SetFont("s8 cAAAAAA", "Microsoft YaHei")
  g_WL_Gui.AddText("w320", "Esc 关闭 | 点击外部关闭")

  ; 计算显示位置（鼠标右下方偏移 15px，避免超出屏幕）
  guiW := 350
  guiH := 320
  showX := posX + 15
  showY := posY + 15

  ; 防止超出屏幕右边界
  if (showX + guiW > A_ScreenWidth)
    showX := posX - guiW - 15
  ; 防止超出屏幕下边界
  if (showY + guiH > A_ScreenHeight)
    showY := posY - guiH - 15

  g_WL_Gui.Show("x" . showX . " y" . showY . " NoActivate")

  ; 绑定 Esc 关闭
  HotIfWinActive("ahk_id " g_WL_Gui.Hwnd)
  Hotkey("Escape", WL_HandleEsc, "On")
  HotIfWinActive()

  ; 启动点击外部关闭的检测定时器
  SetTimer(WL_CheckClickOutside, 200)

  ; 发起 Ollama 请求
  StartWordOllamaRequest(word, context)
  
  ; 自动发声读一次单词
  PlayTtsText(word)
}

; ===== Esc 关闭处理 =====
WL_HandleEsc(*)
{
  CloseWordGui()
}

; ===== 检测点击浮窗外部及朗读悬停 =====
WL_CheckClickOutside()
{
  global g_WL_Gui, g_WL_TtsIcon, WL_CurrentWord

  if (g_WL_Gui = "") {
    SetTimer(WL_CheckClickOutside, 0)
    return
  }

  ; 检测是否悬停在朗读图标上
  static WL_lastHoverTts := false
  currentHoverTts := false
  try {
    MouseGetPos(&mx, &my, &winUnder, &ctrlUnder, 2)
    if (g_WL_TtsIcon && ctrlUnder = g_WL_TtsIcon.Hwnd) {
      currentHoverTts := true
    }
  } catch {
  }
  
  if (currentHoverTts && !WL_lastHoverTts) {
    PlayTtsText(WL_CurrentWord)
  }
  WL_lastHoverTts := currentHoverTts

  ; 检测鼠标左键是否按下
  if (GetKeyState("LButton", "P")) {
    MouseGetPos(&mx, &my, &winAtMouse)
    try {
      if (winAtMouse != g_WL_Gui.Hwnd) {
        CloseWordGui()
        return
      }
    }
  }
}

; ===== 发起 Ollama 语境解释请求 =====
StartWordOllamaRequest(word, context)
{
  global g_WL_StreamFile, g_WL_StreamPid, g_WL_Pending, g_WL_StreamContent

  ; 终止之前的请求
  if (g_WL_StreamPid > 0) {
    try ProcessClose(g_WL_StreamPid)
    g_WL_StreamPid := 0
  }

  g_WL_Pending := true
  g_WL_StreamContent := ""

  ; 构建 prompt
  prompt := "你是一个英语词典。解释单词 '" . word . "'"
  if (context != "" && context != word)
    prompt .= " 在以下语境中的含义。\n语境：" . context
  else
    prompt .= " 的含义。"

  prompt .= "\n\n请用以下格式输出（纯文本，不用Markdown）：\n音标：/xxx/\n词性：xxx\n释义：xxx\n语境释义：在这个句子中表示...\n常见搭配：xxx"

  ; 转义 JSON
  prompt := StrReplace(prompt, "\", "\\")
  prompt := StrReplace(prompt, "`"", "\`"")
  prompt := StrReplace(prompt, "`n", "\n")
  prompt := StrReplace(prompt, "`r", "\r")
  prompt := StrReplace(prompt, "`t", "\t")

  ; 设置文件
  g_WL_StreamFile := A_Temp . "\ollama_stream_word.txt"
  jsonFile := A_Temp . "\ollama_request_word.json"

  try FileDelete(g_WL_StreamFile)
  try FileDelete(jsonFile)

  ; 系统提示
  sysPrompt := "纯文本输出，不要用任何符号（如反斜杠、星号、井号）包裹或强调单词。简洁回答。"

  ; JSON
  json := '{"model":"huihui_ai/qwen3-abliterated:8b-v2","system":"' . sysPrompt . '","prompt":"' . prompt . '","stream":true,"options":{"temperature":0,"num_predict":512,"think":true}}'

  try {
    FileAppend(json, jsonFile, "UTF-8-RAW")
  } catch {
    return
  }

  ; PowerShell 流式请求
  psScript := ""
  . "$body = Get-Content -Path '" . jsonFile . "' -Raw -Encoding UTF8;"
  . "$utf8 = [System.Text.Encoding]::UTF8;"
  . "$bytes = $utf8.GetBytes($body);"
  . "$req = [System.Net.HttpWebRequest]::Create('http://localhost:11434/api/generate');"
  . "$req.Method = 'POST';"
  . "$req.ContentType = 'application/json';"
  . "$req.ContentLength = $bytes.Length;"
  . "$reqStream = $req.GetRequestStream();"
  . "$reqStream.Write($bytes, 0, $bytes.Length);"
  . "$reqStream.Close();"
  . "$resp = $req.GetResponse();"
  . "$reader = New-Object System.IO.StreamReader($resp.GetResponseStream());"
  . "$fs = New-Object System.IO.FileStream('" . g_WL_StreamFile . "', [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite);"
  . "$sw = New-Object System.IO.StreamWriter($fs, [System.Text.Encoding]::UTF8);"
  . "while(-not $reader.EndOfStream) {"
  . "  $line = $reader.ReadLine();"
  . "  $sw.WriteLine($line);"
  . "  $sw.Flush();"
  . "}"
  . "$sw.Close();"
  . "$fs.Close();"
  . "$reader.Close();"
  . "$resp.Close();"

  try {
    Run('powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "' . psScript . '"', , "Hide", &outPid)
    g_WL_StreamPid := outPid
  } catch {
    return
  }

  ; 启动轮询
  SetTimer(CheckWordResult, 200)
}

; ===== 轮询 Ollama 结果 =====
CheckWordResult()
{
  global g_WL_Pending, g_WL_StreamFile, g_WL_StreamContent
  global g_WL_ResultCtrl, g_WL_Gui

  if (!g_WL_Pending || g_WL_Gui = "") {
    SetTimer(CheckWordResult, 0)
    return
  }

  if (g_WL_StreamFile != "" && FileExist(g_WL_StreamFile)) {
    ; 实时读取流式内容并更新浮窗
    currentContent := WL_ReadStreamContent(g_WL_StreamFile)
    if (currentContent != "" && currentContent != g_WL_StreamContent) {
      g_WL_StreamContent := currentContent
      if (g_WL_ResultCtrl != "") {
        try g_WL_ResultCtrl.Text := currentContent
      }
    }

    ; 检查是否完成
    if (IsStreamComplete(g_WL_StreamFile)) {
      Sleep(100)
      finalResult := WL_ReadStreamContent(g_WL_StreamFile)
      if (finalResult != "" && g_WL_ResultCtrl != "") {
        try g_WL_ResultCtrl.Text := finalResult
      }
      g_WL_Pending := false
      SetTimer(CheckWordResult, 0)
    }
  }
}

; ===== 读取流式文件内容 =====
WL_ReadStreamContent(filePath)
{
  if (!FileExist(filePath))
    return ""

  try {
    f := FileOpen(filePath, "r", "UTF-8")
    if (!f)
      return ""
    content := f.Read()
    f.Close()
  } catch {
    return ""
  }

  ; 解析流式 JSON
  result := ""
  Loop Parse, content, "`n", "`r"
  {
    line := Trim(A_LoopField)
    if (line = "")
      continue
    if RegExMatch(line, '"response":"((?:[^"\\]|\\.)*)"', &m) {
      token := m[1]
      token := StrReplace(token, "\n", "`n")
      token := StrReplace(token, "\r", "`r")
      token := StrReplace(token, "\t", "`t")
      token := StrReplace(token, '\`"', '`"')
      token := StrReplace(token, "\\", "\")
      result .= token
    }
  }

  ; 清理 think 标签
  result := RegExReplace(result, "s)<think>.*?</think>", "")
  result := StrReplace(result, "<think>", "")
  result := StrReplace(result, "</think>", "")
  result := StrReplace(result, "/think", "")
  result := StrReplace(result, "/no_think", "")
  result := Trim(result)

  return result
}

; ===== 关闭浮窗 =====
CloseWordGui()
{
  global g_WL_Gui, g_WL_ResultCtrl, g_WL_TitleCtrl
  global g_WL_StreamPid, g_WL_Pending

  ; 终止请求
  if (g_WL_StreamPid > 0) {
    try ProcessClose(g_WL_StreamPid)
    g_WL_StreamPid := 0
  }
  g_WL_Pending := false
  SetTimer(CheckWordResult, 0)
  SetTimer(WL_CheckClickOutside, 0)

  ; 销毁 GUI
  if (g_WL_Gui != "") {
    try {
      HotIfWinActive("ahk_id " g_WL_Gui.Hwnd)
      Hotkey("Escape", WL_HandleEsc, "Off")
      HotIfWinActive()
    }
    try g_WL_Gui.Destroy()
    g_WL_Gui := ""
    g_WL_ResultCtrl := ""
    g_WL_TitleCtrl := ""
  }
}
