;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; 鼠标取词 + 语境解释 - Alt+W：截取鼠标所在窗口 → Windows OCR → 定位单词 → Ollama 解释
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

#Include "Gdip_All.ahk"

; ===== 全局变量 =====
g_WL_Gui := ""
g_WL_ResultCtrl := ""
g_WL_TitleCtrl := ""
g_WL_WordEdit := ""
g_WL_ContextEdit := ""
WL_CurrentWord := ""
WL_CurrentContext := ""
g_WL_LangMode := "EN"
try g_WL_LangMode := IniRead(A_ScriptDir . "\ollama_config.ini", "Settings", "WordLookupLang", "EN")
g_WL_LangBtn := ""
g_WL_StreamFile := ""
g_WL_StreamPid := 0
g_WL_Pending := false
g_WL_StreamContent := ""
g_WL_ShowTick := 0
g_WL_InitMouseX := 0
g_WL_InitMouseY := 0
g_WL_MouseMoved := false
g_WL_TtsFile := ""
g_WL_TtsPid := 0
g_WL_TtsWord := ""
g_WL_History := []
g_WL_HistoryIdx := 0

; ===== 快捷键 F2 =====
F2::
{
  global g_WL_Gui, g_WL_ResultCtrl, g_WL_TitleCtrl
  global g_WL_StreamFile, g_WL_StreamPid, g_WL_Pending, g_WL_StreamContent

  ; 1. 获取鼠标全屏绝对坐标（很重要，防止在自身窗口上查词时获取到相对坐标）
  CoordMode("Mouse", "Screen")
  MouseGetPos(&mouseX, &mouseY, &winUnder)

  ; 2. 以鼠标为中心截取固定区域（避免多显示器/DPI 下窗口坐标不准）
  dpi := 96
  try dpi := DllCall("GetDpiForWindow", "Ptr", winUnder, "UInt")
  if (dpi < 96)
    dpi := 96
  scale := dpi / 96
  captureW := Round(600 * scale)
  captureH := Round(400 * scale)
  winX := mouseX - Round(captureW / 2)
  winY := mouseY - Round(captureH / 2)
  winW := captureW
  winH := captureH

  ; 3. 计算鼠标在截图中的相对坐标（始终在中心）
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

  ; 不关闭旧窗口，以便截取到浮窗内部的内容，后续就地复用更新

  if (!pBitmap) {
    Gdip_Shutdown(pToken)
    ToolTip("截图失败")
    SetTimer(ToolTip, -2000)
    return
  }

  ; 7. 放大 2 倍以提升 OCR 准确率
  ocrScale := 2
  origW := Gdip_GetImageWidth(pBitmap)
  origH := Gdip_GetImageHeight(pBitmap)
  newW := origW * ocrScale
  newH := origH * ocrScale
  pBitmapScaled := Gdip_CreateBitmap(newW, newH)
  G := Gdip_GraphicsFromImage(pBitmapScaled)
  Gdip_SetInterpolationMode(G, 7)  ; HighQualityBicubic
  Gdip_DrawImage(G, pBitmap, 0, 0, newW, newH, 0, 0, origW, origH)
  Gdip_DeleteGraphics(G)
  Gdip_DisposeImage(pBitmap)

  ; 8. 保存为临时 PNG
  tempImg := A_Temp . "\ahk_word_capture.png"
  Gdip_SaveBitmapToFile(pBitmapScaled, tempImg)
  Gdip_DisposeImage(pBitmapScaled)
  Gdip_Shutdown(pToken)

  ; 鼠标坐标同步放大
  relMouseX := relMouseX * ocrScale
  relMouseY := relMouseY * ocrScale

  ; 9. 显示"正在识别..."提示
  ToolTip("🔍 正在识别...")

  ; 10. 调用 PowerShell OCR 脚本（同步等待结果）
  ocrScript := A_ScriptDir . "\ocr_word.ps1"
  ocrResultFile := A_Temp . "\ahk_ocr_result.txt"
  try FileDelete(ocrResultFile)

  ; 使用 -OutputFile 参数让 PowerShell 直接以 UTF-8 写入结果文件
  psCmd := 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' . ocrScript . '" -ImagePath "' . tempImg . '" -MouseX ' . relMouseX . ' -MouseY ' . relMouseY . ' -OutputFile "' . ocrResultFile . '"'

  Run(psCmd, , "Hide", &ocrPid)
  SetTimer(CheckOcrResult, 100)

  CheckOcrResult() {
    if ProcessExist(ocrPid)
      return
    SetTimer(CheckOcrResult, 0)

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
      if (RegExMatch(ocrOutput, 's)"word"\s*:\s*"(.*?)"\s*,\s*"line"', &m))
        word := m[1]
      else if (RegExMatch(ocrOutput, '"word"\s*:\s*"((?:[^"\\]|\\.)*)"', &m))
        word := m[1]

      if (RegExMatch(ocrOutput, 's)"line"\s*:\s*"(.*?)"\s*}', &m))
        line := m[1]
      else if (RegExMatch(ocrOutput, '"line"\s*:\s*"((?:[^"\\]|\\.)*)"', &m))
        line := m[1]
        
      ; 清理解析出的转义符号，防止乱码
      word := StrReplace(word, '\"', '"')
      word := StrReplace(word, '\\', '\')
      line := StrReplace(line, '\"', '"')
      line := StrReplace(line, '\\', '\')
    }

    if (!found || word = "") {
      ToolTip("未识别到单词")
      SetTimer(ToolTip, -2000)
      return
    }

    ; 12. 显示浮窗并请求 Ollama
    ShowWordPopup(word, line, mouseX, mouseY)
  }
}

; ===== 显示取词浮窗 =====
ShowWordPopup(word, context, posX, posY)
{
  global g_WL_Gui, g_WL_ResultCtrl, g_WL_TitleCtrl, g_WL_WordEdit, g_WL_ContextEdit, WL_CurrentWord, WL_CurrentContext, g_WL_LangMode, g_WL_LangBtn
  WL_CurrentWord := word
  WL_CurrentContext := context

  if (g_WL_Gui != "") {
    ; 如果窗口已存在，直接更新内容，不重新创建
    g_WL_WordEdit.Value := word
    g_WL_ContextEdit.Value := (context != "" && context != word) ? context : ""
    g_WL_ResultCtrl.Value := "⏳ 正在查询..."

    ; 重置悬停自动关闭的检测状态，防止刚更新完就消失
    global g_WL_InitMouseX, g_WL_InitMouseY, g_WL_MouseMoved, g_WL_ShowTick
    MouseGetPos(&g_WL_InitMouseX, &g_WL_InitMouseY)
    g_WL_MouseMoved := false
    g_WL_ShowTick := A_TickCount

    ; 激活窗口，把焦点放在窗口上，保证快捷键可用
    try WinActivate("ahk_id " . g_WL_Gui.Hwnd)

    ; 预生成 TTS 音频（后台，会自动终止上一个）
    WL_PregenTts(word)
    ; 发起 Ollama 请求（会自动终止上一个）
    StartWordOllamaRequest(word, context)
    return
  }


  g_WL_Gui := Gui("+AlwaysOnTop -Caption +Border +Owner")
  g_WL_Gui.BackColor := "FFFFFF"
  g_WL_Gui.MarginX := 12
  g_WL_Gui.MarginY := 8

  ; 标题行水平排列
  g_WL_Gui.SetFont("s14 c1a1a2e Bold", "Microsoft YaHei")
  
  ; 单词可编辑输入框 + 中英切换按钮
  g_WL_WordEdit := g_WL_Gui.AddEdit("w275 Section -E0x200", word)
  
  g_WL_Gui.SetFont("s9 c333333 Norm", "Microsoft YaHei")
  g_WL_LangBtn := g_WL_Gui.AddButton("x+5 ys w40 h26", g_WL_LangMode = "EN" ? "EN" : "中")
  
  ; 关联事件
  if (g_WL_LangBtn) {
    g_WL_LangBtn.OnEvent("Click", (*) => WL_ToggleLang())
  }

  ; 语境行（可编辑）
  g_WL_Gui.SetFont("s10 c444444 Norm", "Microsoft YaHei")
  g_WL_ContextEdit := g_WL_Gui.AddEdit("xs w320 -E0x200", (context != "" && context != word) ? context : "")

  ; 分隔线
  g_WL_Gui.SetFont("s1 cCCCCCC", "Microsoft YaHei")
  g_WL_Gui.AddText("xs w320 0x10")  ; SS_ETCHEDHORZ

  ; 结果区域（可选中复制）
  g_WL_Gui.SetFont("s10 c333333 Norm", "Microsoft YaHei")
  g_WL_ResultCtrl := g_WL_Gui.AddEdit("xs w320 h180 ReadOnly -E0x200", "⏳ 正在查询...")

  ; 底部提示
  g_WL_Gui.SetFont("s8 cAAAAAA", "Microsoft YaHei")
  g_WL_Gui.AddText("w320", "Enter 重新查询 | Esc 关闭 | 鼠标移出关闭")

  ; 先在屏幕外显示一次，获取窗口的真实尺寸
  g_WL_Gui.Show("x-9999 y-9999 NoActivate")
  WinGetPos(, , &guiW, &guiH, "ahk_id " . g_WL_Gui.Hwnd)

  ; 获取鼠标所在显示器的工作区域（排除任务栏）
  monCount := MonitorGetCount()
  monLeft := 0, monTop := 0, monRight := A_ScreenWidth, monBottom := A_ScreenHeight
  Loop monCount {
    MonitorGetWorkArea(A_Index, &mL, &mT, &mR, &mB)
    if (posX >= mL && posX < mR && posY >= mT && posY < mB) {
      monLeft := mL, monTop := mT, monRight := mR, monBottom := mB
      break
    }
  }

  showX := posX + 15
  showY := posY + 15

  ; 超出右边界 → 弹到鼠标左侧
  if (showX + guiW > monRight)
    showX := posX - guiW - 15
  ; 超出下边界 → 弹到鼠标上方
  if (showY + guiH > monBottom)
    showY := posY - guiH - 15
  ; 最终保底：不能超出左上角
  if (showX < monLeft)
    showX := monLeft
  if (showY < monTop)
    showY := monTop

  ; 移动到正确位置
  g_WL_Gui.Show("x" . showX . " y" . showY . " NoActivate")

  ; 绑定 Esc/Enter、历史记录导航
  HotIfWinActive("ahk_id " g_WL_Gui.Hwnd)
  Hotkey("Escape", WL_HandleEsc, "On")
  Hotkey("Enter", WL_HandleEnter, "On")
  Hotkey("!Left", (*) => WL_NavHistory(-1), "On")
  Hotkey("!Right", (*) => WL_NavHistory(1), "On")
  HotIfWinActive()

  ; 启动鼠标移出关闭的检测定时器
  global g_WL_InitMouseX, g_WL_InitMouseY, g_WL_MouseMoved, g_WL_ShowTick
  MouseGetPos(&g_WL_InitMouseX, &g_WL_InitMouseY)
  g_WL_MouseMoved := false
  g_WL_ShowTick := A_TickCount
  SetTimer(WL_CheckClickOutside, 200)

  ; 预生成 TTS 音频（后台，不阻塞）
  WL_PregenTts(word)

  ; 发起 Ollama 请求
  StartWordOllamaRequest(word, context)
}

; ===== Enter 重新查询 =====
WL_HandleEnter(*)
{
  global g_WL_WordEdit, g_WL_ContextEdit, g_WL_ResultCtrl, WL_CurrentWord, WL_CurrentContext

  if (g_WL_WordEdit = "")
    return

  newWord := Trim(g_WL_WordEdit.Value)
  if (newWord = "")
    return

  newContext := (g_WL_ContextEdit != "") ? Trim(g_WL_ContextEdit.Value) : ""
  WL_CurrentWord := newWord
  WL_CurrentContext := newContext
  if (g_WL_ResultCtrl != "")
    g_WL_ResultCtrl.Value := "⏳ 正在查询..."
  WL_PregenTts(newWord)
  StartWordOllamaRequest(newWord, newContext)
}

; ===== Esc 关闭处理 =====
WL_HandleEsc(*)
{
  CloseWordGui()
}

; ===== 检测点击浮窗外部 =====
WL_CheckClickOutside()
{
  global g_WL_Gui

  if (g_WL_Gui = "") {
    SetTimer(WL_CheckClickOutside, 0)
    return
  }

  ; 窗口显示后 1000ms 内不检测
  global g_WL_ShowTick
  if (A_TickCount - g_WL_ShowTick < 1000)
    return

  ; 鼠标未移动前不检测，避免窗口刚显示就关闭
  global g_WL_InitMouseX, g_WL_InitMouseY, g_WL_MouseMoved
  MouseGetPos(&cx, &cy)
  if (!g_WL_MouseMoved) {
    if (Abs(cx - g_WL_InitMouseX) > 5 || Abs(cy - g_WL_InitMouseY) > 5)
      g_WL_MouseMoved := true
    else
      return
  }

  ; 检测鼠标是否在窗口外，是则自动关闭
  MouseGetPos(&mx, &my, &winAtMouse)
  try {
    if (winAtMouse != g_WL_Gui.Hwnd) {
      ; 检查是否为输入法相关窗口，避免中文输入时误关闭
      try {
        imeClass := WinGetClass("ahk_id " . winAtMouse)
        imePid := WinGetPID("ahk_id " . winAtMouse)
        imeProcName := ProcessGetName(imePid)
        if (imeClass = "ApplicationFrameWindow"
          || imeProcName = "TextInputHost.exe"
          || InStr(imeClass, "IME") || InStr(imeClass, "MSCTFIME") || InStr(imeClass, "Cand")
          || InStr(imeProcName, "IME"))
          return
      }
      CloseWordGui()
      return
    }
  }
}

; ===== 切换语言 =====
WL_ToggleLang()
{
  global g_WL_LangMode, g_WL_LangBtn, WL_CurrentWord, WL_CurrentContext, g_WL_ResultCtrl

  if (g_WL_LangMode = "EN") {
    g_WL_LangMode := "ZH"
    g_WL_LangBtn.Text := "中"
  } else {
    g_WL_LangMode := "EN"
    g_WL_LangBtn.Text := "EN"
  }
  try IniWrite(g_WL_LangMode, A_ScriptDir . "\\ollama_config.ini", "Settings", "WordLookupLang")
  
  if (g_WL_ResultCtrl != "") {
    g_WL_ResultCtrl.Value := "⏳ 正在切换语言并重新查询..."
  }

  ; 重新发起请求
  StartWordOllamaRequest(WL_CurrentWord, WL_CurrentContext)
}

; ===== 发起 Ollama 语境解释请求 =====
StartWordOllamaRequest(word, context, isNavigating := false)
{
  global g_WL_StreamFile, g_WL_StreamPid, g_WL_Pending, g_WL_StreamContent, g_WL_LangMode
  global g_WL_History, g_WL_HistoryIdx

  ; 历史记录处理
  if (!isNavigating) {
    if (g_WL_HistoryIdx == 0 || g_WL_HistoryIdx > g_WL_History.Length || g_WL_History[g_WL_HistoryIdx].word != word || g_WL_History[g_WL_HistoryIdx].context != context) {
      if (g_WL_HistoryIdx > 0 && g_WL_HistoryIdx < g_WL_History.Length) {
        g_WL_History.RemoveAt(g_WL_HistoryIdx + 1, g_WL_History.Length - g_WL_HistoryIdx)
      }
      g_WL_History.Push({word: word, context: context, result: ""})
      
      ; 限制最多只保留 3 个历史记录
      while (g_WL_History.Length > 3) {
        g_WL_History.RemoveAt(1)
      }
      g_WL_HistoryIdx := g_WL_History.Length
    } else {
      ; 触发同样的查询（如切换语言），清空保存的旧结果
      if (g_WL_HistoryIdx > 0 && g_WL_HistoryIdx <= g_WL_History.Length)
        g_WL_History[g_WL_HistoryIdx].result := ""
    }
  }

  ; 终止之前的请求
  if (g_WL_StreamPid > 0) {
    try ProcessClose(g_WL_StreamPid)
    g_WL_StreamPid := 0
  }

  g_WL_Pending := true
  g_WL_StreamContent := ""

  if (g_WL_LangMode = "EN") {
    ; 构建 prompt (英英释义模式)
    prompt := "You are an English-English dictionary. Explain the word '" . word . "' entirely in simple English."
    if (context != "" && context != word)
      prompt .= " Please explain its meaning in the following context:\nContext: " . context

    prompt .= "\n\nPlease output using the following format (plain text only, no Markdown):\n● Word Roots: [Break down prefixes, roots, suffixes with origins and meanings, e.g. 'un-(not) + break-(break) + -able(capable of)']\n● Part of Speech: xxx\n● Definition: [Simple English definition]\n● Context Meaning: [Explanation based on the given context, if any]\n● Collocations: [Common collocations or examples]"
    
    sysPrompt := "Output ONLY in English. Use plain text without Markdown formatting (no asterisks, hashes, etc.). Keep explanations concise and easy to understand."
  } else {
    ; 构建 prompt (英汉释义模式)
    prompt := "你是一个英语词典。解释单词 '" . word . "'"
    if (context != "" && context != word)
      prompt .= " 在以下语境中的含义。\n语境：" . context
    else
      prompt .= " 的含义。"

    prompt .= "\n\n请用以下格式输出（纯文本，不用Markdown）：\n● 词根拆解：拆分前缀、词根、后缀，标注来源和含义（如 un-表否定 + break打破 + -able表能力）\n● 词性：xxx\n● 释义：xxx\n● 语境释义：在这个句子中表示...\n● 常见搭配：xxx"
    
    sysPrompt := "纯文本输出，不要用任何符号（如反斜杠、星号、井号）包裹或强调单词。简洁回答。"
  }

  ; 转义 JSON
  prompt := StrReplace(prompt, "\", "\\")
  prompt := StrReplace(prompt, "`"", "\`"")
  prompt := StrReplace(prompt, "`n", "\n")
  prompt := StrReplace(prompt, "`r", "\r")
  prompt := StrReplace(prompt, "`t", "\t")

  ; 设置文件
  g_WL_StreamFile := A_Temp . "\ahk_wl_stream_word.txt"
  jsonFile := A_Temp . "\ahk_wl_request_word.json"

  try FileDelete(g_WL_StreamFile)
  try FileDelete(jsonFile)

  ; JSON
  json := '{"model":"huihui_ai/qwen3-abliterated:8b-v2","system":"' . sysPrompt . '","prompt":"' . prompt . '","stream":true,"options":{"temperature":0,"num_predict":800}}'

  try {
    FileAppend(json, jsonFile, "UTF-8-RAW")
  } catch {
    return
  }

  ; 弃用耗时的 PowerShell，改用 Windows 内置极速 curl.exe，彻底消除冷启动延迟
  try {
    curlCmd := 'curl.exe -s -N -X POST "http://localhost:11434/api/generate" -H "Content-Type: application/json" -d "@' . jsonFile . '" -o "' . g_WL_StreamFile . '"'
    Run(curlCmd, , "Hide", &outPid)
    g_WL_StreamPid := outPid
  } catch {
    return
  }

  ; 启动极速轮询
  SetTimer(CheckWordResult, 50)
}

; ===== 轮询 Ollama 结果 =====
CheckWordResult()
{
  global g_WL_Pending, g_WL_StreamFile, g_WL_StreamContent, g_WL_StreamPid
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
        try g_WL_ResultCtrl.Value := currentContent
      }
    }

    ; 检查 curl 进程是否结束
    isComplete := false
    if (g_WL_StreamPid > 0 && !ProcessExist(g_WL_StreamPid)) {
      isComplete := true
    }

    if (isComplete) {
      Sleep(30) ; 等待文件最终刷入硬盘
      finalResult := WL_ReadStreamContent(g_WL_StreamFile)
      if (finalResult != "") {
        if (g_WL_ResultCtrl != "") {
          try g_WL_ResultCtrl.Value := finalResult
        }
        global g_WL_History, g_WL_HistoryIdx
        if (g_WL_HistoryIdx > 0 && g_WL_HistoryIdx <= g_WL_History.Length) {
          g_WL_History[g_WL_HistoryIdx].result := finalResult
        }
      }
      
      ; 状态重置与收尾清理
      g_WL_Pending := false
      g_WL_StreamPid := 0
      SetTimer(CheckWordResult, 0)
      
      ; 阅后即焚，清理临时文件
      try FileDelete(g_WL_StreamFile)
      try FileDelete(A_Temp . "\ahk_wl_request_word.json")
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
  global g_WL_StreamPid, g_WL_Pending, g_WL_StreamFile
  global g_WL_WordEdit, g_WL_ContextEdit

  ; 终止请求
  if (g_WL_StreamPid > 0) {
    try ProcessClose(g_WL_StreamPid)
    g_WL_StreamPid := 0
  }
  g_WL_Pending := false
  SetTimer(CheckWordResult, 0)
  SetTimer(WL_CheckClickOutside, 0)
  try SoundPlay("NonExistent.zzz")

  ; 彻底清理临时文件
  try FileDelete(g_WL_StreamFile)
  try FileDelete(A_Temp . "\ahk_wl_request_word.json")

  ; 销毁 GUI
  if (g_WL_Gui != "") {
    try {
      HotIfWinActive("ahk_id " g_WL_Gui.Hwnd)
      Hotkey("Escape", WL_HandleEsc, "Off")
      Hotkey("Enter", WL_HandleEnter, "Off")
      Hotkey("!Left", "Off")
      Hotkey("!Right", "Off")
      HotIfWinActive()
    }
    try g_WL_Gui.Destroy()
    g_WL_Gui := ""
    g_WL_ResultCtrl := ""
    g_WL_TitleCtrl := ""
    g_WL_WordEdit := ""
    g_WL_ContextEdit := ""
  }
}

; ===== 预生成 TTS 音频（后台） =====
WL_PregenTts(word)
{
  global g_WL_TtsFile, g_WL_TtsPid, g_WL_TtsWord

  ; 终止上一次预生成
  if (g_WL_TtsPid > 0) {
    try ProcessClose(g_WL_TtsPid)
    g_WL_TtsPid := 0
  }

  text := Trim(word)
  if (text = "")
    return

  g_WL_TtsWord := text
  g_WL_TtsFile := A_Temp . "\ahk_wl_tts_" . A_TickCount . ".mp3"
  
  ; 顺便清理一下历史遗留的这种临时文件，防止堆积
  try {
    Loop Files, A_Temp . "\ahk_wl_tts_*.mp3"
      if (A_LoopFileFullPath != g_WL_TtsFile)
        FileDelete(A_LoopFileFullPath)
  }

  escapedText := StrReplace(text, '"', '\"')
  try {
    Run('edge-tts --voice en-US-AriaNeural --volume=+500% --text "' . escapedText . '" --write-media "' . g_WL_TtsFile . '"', , "Hide", &outPid)
    g_WL_TtsPid := outPid
  }
}

; ===== 强制单次朗读（右键触发） =====
WL_PlayTtsOnce()
{
  global WL_CurrentWord, g_WL_TtsFile, g_WL_TtsPid, g_WL_TtsWord

  text := Trim(WL_CurrentWord)
  if (text = "")
    return

  if (text != g_WL_TtsWord) {
    WL_PregenTts(text)
  }

  ; 非阻塞等待 edge-tts
  if (g_WL_TtsPid > 0 && ProcessExist(g_WL_TtsPid)) {
    SetTimer(WL_PlayTtsOnce, -100)
    return
  }
  g_WL_TtsPid := 0

  try {
    if FileExist(g_WL_TtsFile) {
      ; 打断上一次朗读
      try SoundPlay("NonExistent.zzz")
      SoundPlay(g_WL_TtsFile)
    }
  } catch {
  }
}

; ===== 基于窗口存在的全局右键拦截防止菜单弹出 =====
#HotIf WL_IsWordGuiShown()
RButton::
{
  global g_WL_InitMouseX, g_WL_InitMouseY, g_WL_MouseMoved, g_WL_ShowTick
  
  ; 1. 朗读（非阻塞）
  WL_PlayTtsOnce()
  
  ; 2. 重置自动关闭防抖动计时器（防止因触发而导致抖动退出）
  MouseGetPos(&g_WL_InitMouseX, &g_WL_InitMouseY)
  g_WL_MouseMoved := false
  g_WL_ShowTick := A_TickCount
}
#HotIf

WL_IsWordGuiShown() {
  global g_WL_Gui
  return (g_WL_Gui != "")
}

; ===== 历史记录导航 =====
WL_NavHistory(dir)
{
  global g_WL_History, g_WL_HistoryIdx, g_WL_WordEdit, g_WL_ContextEdit, g_WL_ResultCtrl
  global WL_CurrentWord, WL_CurrentContext

  if (g_WL_History.Length = 0)
    return

  newIdx := g_WL_HistoryIdx + dir
  if (newIdx < 1)
    newIdx := 1
  if (newIdx > g_WL_History.Length)
    newIdx := g_WL_History.Length

  if (newIdx == g_WL_HistoryIdx)
    return

  g_WL_HistoryIdx := newIdx
  item := g_WL_History[newIdx]

  WL_CurrentWord := item.word
  WL_CurrentContext := item.context

  if (g_WL_WordEdit != "")
    g_WL_WordEdit.Value := item.word
  if (g_WL_ContextEdit != "")
    g_WL_ContextEdit.Value := item.context

  if (item.result != "") {
    if (g_WL_ResultCtrl != "")
      g_WL_ResultCtrl.Value := item.result
    
    ; 终止后台可能还在进行的请求，直接显示保存的结果
    global g_WL_StreamPid, g_WL_Pending
    if (g_WL_StreamPid > 0) {
      try ProcessClose(g_WL_StreamPid)
      g_WL_StreamPid := 0
    }
    g_WL_Pending := false
    SetTimer(CheckWordResult, 0)

  } else {
    if (g_WL_ResultCtrl != "")
      g_WL_ResultCtrl.Value := "⏳ 正在查询..."
    StartWordOllamaRequest(item.word, item.context, true)
  }

  WL_PregenTts(item.word)

  ; 导航操作视同活跃操作，重置自动关闭的计时器
  global g_WL_InitMouseX, g_WL_InitMouseY, g_WL_MouseMoved, g_WL_ShowTick
  MouseGetPos(&g_WL_InitMouseX, &g_WL_InitMouseY)
  g_WL_MouseMoved := false
  g_WL_ShowTick := A_TickCount
}
