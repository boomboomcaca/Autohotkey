;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; Ollama 翻译/纠错 - Ctrl+Alt+Enter: 中文→翻译英文，英文→纠正表达
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

g_OriginalText := ""
g_TranslateResult := ""
g_CorrectResult := ""
g_OldClip := ""
g_MainGui := ""
g_TranslateEditCtrl := ""
g_CorrectEditCtrl := ""
g_CorrectLabelCtrl := ""
g_TranslateLabelCtrl := ""
g_IsChineseMode := false
g_SelectedResult := "translate"
g_OrigEditCtrl := ""

; 异步 HTTP 对象
g_HttpCorrect := ""
g_HttpTranslate := ""
g_CorrectPending := false
g_TranslatePending := false
g_CorrectRequested := false
g_TranslateRequested := false
g_CurrentText := ""
g_TtsPlaying := false
g_HoverTarget := ""
g_PendingShowGui := false
g_GuiHidden := false

; 流式响应相关
g_StreamFileCorrect := ""
g_StreamFileTranslate := ""
g_StreamPidCorrect := 0
g_StreamPidTranslate := 0
g_StreamContentCorrect := ""
g_StreamContentTranslate := ""

; AI 问答相关
g_QuestionEditCtrl := ""
g_AnswerEditCtrl := ""
g_SendBtnCtrl := ""
g_ChatPending := false
g_StreamFileChat := ""
g_StreamPidChat := 0
g_StreamContentChat := ""

OllamaCall(prompt)
{
  ; 构建 JSON
  prompt := StrReplace(prompt, "\", "\\")
  prompt := StrReplace(prompt, "`"", "\`"")
  prompt := StrReplace(prompt, "`n", "\n")
  prompt := StrReplace(prompt, "`r", "\r")
  prompt := StrReplace(prompt, "`t", "\t")
  json := "{`"model`":`"qwen3:latest`",`"prompt`":`"" . prompt . "`",`"stream`":false,`"options`":{`"temperature`":0,`"num_predict`":1024}}"
  
  try {
    http := ComObject("WinHttp.WinHttpRequest.5.1")
    http.Open("POST", "http://localhost:11434/api/generate", false)
    http.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
    http.Send(json)
    http.WaitForResponse()
    
    response := http.ResponseText
    if RegExMatch(response, "`"response`"\s*:\s*`"(.*?)`"(?=\s*,\s*`")", &m)
      result := m[1]
    else
      return "解析失败"
    
    result := StrReplace(result, "\n", "`n")
    result := StrReplace(result, "\r", "`r")
    result := StrReplace(result, "\t", "`t")
    result := StrReplace(result, "\`"", "`"")
    result := StrReplace(result, "\\", "\")
    
    result := RegExReplace(result, "s)<think>.*?</think>", "")
    result := StrReplace(result, "/think")
    result := StrReplace(result, "/no_think")
    
    result := Trim(result)
    return result
  } Catch Error as e {
    return "请求失败: " . e.Message
  }
}

OllamaTranslate(text, isChinese)
{
  if isChinese
    prompt := "/no_think Translate to English. Keep the exact same formatting, including punctuation marks, line breaks, and spacing. Output only the translation:`n" . text
  else
    prompt := "/no_think Translate to Chinese. Keep the exact same formatting, including punctuation marks, line breaks, and spacing. Output only the translation:`n" . text
  return OllamaCall(prompt)
}

OllamaCorrect(text, isChinese)
{
  if isChinese
    prompt := "/no_think You are a Chinese language tutor. Correct and improve the following Chinese text. Fix grammar, punctuation, and improve expression while keeping the original meaning. Output only the corrected text without any explanation:`n" . text
  else
    prompt := "/no_think Correct this English text for a Chinese learner.`n`nRules:`n1. First line: ONLY the corrected sentence, nothing else`n2. Second line: exactly three dashes: ---`n3. Then list errors in Chinese: 错误1: 原文 → 修正 (解释)`n`nExample output:`nI am a real team member.`n---`n错误1: i → I (句首字母需要大写)`n错误2: real team → a real team (需要冠词 a)`n`nNow correct: " . text
  return OllamaCall(prompt)
}

ShowMainGui(original)
{
  global g_OriginalText, g_TranslateResult, g_CorrectResult, g_OldClip, g_MainGui
  global g_TranslateEditCtrl, g_CorrectEditCtrl, g_CorrectLabelCtrl, g_TranslateLabelCtrl, g_OrigEditCtrl, g_IsChineseMode, g_SelectedResult
  global g_TtsOrigCtrl, g_TtsCorrectCtrl, g_TtsTranslateCtrl
  global g_ExplainEditCtrl, g_CorrectedText
  global g_QuestionEditCtrl, g_AnswerEditCtrl, g_SendBtnCtrl
  
    ; 如果已有窗口存在，先关闭
  if (g_MainGui != "") {
    try {
      g_MainGui.Destroy()
    }
    g_MainGui := ""
  }
  
  g_OriginalText := original
  g_TranslateResult := ""
  g_CorrectResult := ""
  g_CorrectedText := ""
  g_SelectedResult := "correct"
  g_ExplainEditCtrl := ""
  
  ; 判断中英文
  g_IsChineseMode := RegExMatch(original, "[\x{4e00}-\x{9fff}]")
  
  if g_IsChineseMode {
    title := "中文处理 - Enter 替换 / Esc 取消"
    correctLabel := "纠错 (中文润色):"
    translateLabel := "翻译 (中→英):"
  } else {
    title := "英文处理 - Enter 替换 / Esc 取消"
    correctLabel := "纠错 (英文润色):"
    translateLabel := "翻译 (英→中):"
  }
  
  g_MainGui := Gui("+AlwaysOnTop", title)
  g_MainGui.SetFont("s10", "Microsoft YaHei")
  
  ; ========== 左侧面板：翻译/纠错 ==========
  ; 英文模式显示朗读图标
  if !g_IsChineseMode {
    g_MainGui.AddText("w120 Section", "原文 (可编辑):")
    g_TtsOrigCtrl := g_MainGui.AddText("x+5 ys cGray", "🔊")
    g_TtsOrigCtrl.OnEvent("Click", Gui_PlayOriginal)
    g_OrigEditCtrl := g_MainGui.AddEdit("xm w500 h60", original)
  } else {
    g_MainGui.AddText("w500", "原文 (可编辑):")
    g_OrigEditCtrl := g_MainGui.AddEdit("xm w500 h60", original)
  }
  
  if g_IsChineseMode {
    ; 中文：翻译在前，添加朗读图标
    g_TranslateLabelCtrl := g_MainGui.AddText("w120 Section", "✓ " . translateLabel)
    g_TtsTranslateCtrl := g_MainGui.AddText("x+5 ys cGray", "🔊")
    g_TtsTranslateCtrl.OnEvent("Click", Gui_PlayTranslate)
    g_TranslateEditCtrl := g_MainGui.AddEdit("xm w500 h120 ReadOnly", "正在处理...")
    g_CorrectLabelCtrl := g_MainGui.AddText("w500", "   " . correctLabel)
    g_CorrectEditCtrl := g_MainGui.AddEdit("w500 h60 ReadOnly", "正在处理...")
    g_SelectedResult := "translate"
  } else {
    ; 英文：纠错在前，添加朗读图标
    g_CorrectLabelCtrl := g_MainGui.AddText("w120 Section", "✓ " . correctLabel)
    g_TtsCorrectCtrl := g_MainGui.AddText("x+5 ys cGray", "🔊")
    g_TtsCorrectCtrl.OnEvent("Click", Gui_PlayCorrect)
    g_CorrectEditCtrl := g_MainGui.AddEdit("xm w500 h60 ReadOnly", "正在纠错...")
    ; 错误解释框
    g_MainGui.AddText("w500", "错误解释:")
    g_ExplainEditCtrl := g_MainGui.AddEdit("w500 h100 ReadOnly", "")
    g_TranslateLabelCtrl := g_MainGui.AddText("w120 Section", "   " . translateLabel)
    g_TtsTranslateCtrl := g_MainGui.AddText("x+5 ys cGray", "🔊")
    g_TtsTranslateCtrl.OnEvent("Click", Gui_PlayTranslate)
    g_TranslateEditCtrl := g_MainGui.AddEdit("xm w500 h40 ReadOnly", "正在处理...")
    g_SelectedResult := "correct"
  }
  
  ; ========== 右侧面板：AI 问答 ==========
  g_MainGui.AddText("x530 y10 w400 Section", "AI 助手:")
  g_QuestionEditCtrl := g_MainGui.AddEdit("xs w330 h60", "")
  g_SendBtnCtrl := g_MainGui.AddButton("x+5 yp h60 w60", "发送")
  g_SendBtnCtrl.OnEvent("Click", Gui_SendQuestion)
  g_MainGui.AddText("xs w400", "回答:")
  g_AnswerEditCtrl := g_MainGui.AddEdit("xs w400 h240 ReadOnly", "")
  
  ; ========== 底部提示 ==========
  g_MainGui.AddText("xm w930 cGray", "Tab 切换输入框 | Ctrl+Tab 切换结果焦点 | Enter 替换/发送 | Ctrl+Enter 强制替换 | Alt+`` 切换窗口 | Esc 取消")
  
  g_MainGui.OnEvent("Close", Gui_Close)
  g_MainGui.OnEvent("Escape", Gui_Close)
  
  ; 窗口创建后暂不显示，等待 AI 响应后再显示
  ; g_MainGui.Show()
  
  HotIfWinActive("ahk_id " g_MainGui.Hwnd)
  Hotkey("Enter", Gui_HandleEnter.Bind(g_MainGui), "On")
  Hotkey("NumpadEnter", Gui_HandleEnter.Bind(g_MainGui), "On")
  Hotkey("^Enter", Gui_Apply.Bind(g_MainGui), "On")
  Hotkey("^NumpadEnter", Gui_Apply.Bind(g_MainGui), "On")
  Hotkey("^Tab", Gui_ToggleSelect, "On")
  Hotkey("Tab", Gui_ToggleFocus, "On")
  HotIfWinActive()
  
  ; 重置请求状态并异步调用 API
  global g_CorrectRequested, g_TranslateRequested, g_PendingShowGui
  g_CorrectRequested := false
  g_TranslateRequested := false
  
  ; 直接显示窗口，不等待 AI 响应
  g_PendingShowGui := false
  if (original = "") {
    g_TranslateEditCtrl.Value := ""
    g_CorrectEditCtrl.Value := ""
  }
  g_MainGui.Show()
  g_QuestionEditCtrl.Focus()
  SetTimer(CheckTtsHover, 200)
  
  ; 有原文时启动异步请求
  if (original != "") {
    StartAsyncRequests(original, g_SelectedResult)
  }
}

StartAsyncRequests(text, requestType := "default")
{
  global g_HttpCorrect, g_CorrectPending, g_TranslatePending, g_IsChineseMode
  global g_CorrectRequested, g_TranslateRequested, g_CurrentText
  
  g_CurrentText := text
  isChinese := RegExMatch(text, "[\x{4e00}-\x{9fff}]")
  g_IsChineseMode := isChinese
  
  ; 一次调用同时完成纠错和翻译
  if (requestType = "default" || !g_CorrectRequested) {
    g_CorrectRequested := true
    g_TranslateRequested := true
    
    if isChinese {
      ; 中文：润色 + 翻译成英文
      combinedPrompt := "/no_think 请对以下中文进行润色和翻译。不要使用Markdown格式。`n`n输出格式(严格遵守):`n===CORRECT===`n润色后的中文`n===TRANSLATE===`n英文翻译`n`n原文: " . text
    } else {
      ; 英文：纠错+解释 + 翻译成中文
      combinedPrompt := "/no_think Correct and translate this English text for a Chinese learner. Do not use Markdown formatting.`n`nOutput format (strict, no backslashes, plain text only):`n===CORRECT===`nCorrected sentence`n---`n错误1: 原文 → 修正 (中文解释)`n===TRANSLATE===`n中文翻译`n`nExample:`n===CORRECT===`nI am going home.`n---`n错误1: i → I (句首字母大写)`n错误2: gohome → going home (需要空格)`n===TRANSLATE===`n我要回家了。`n`nText: " . text
    }
    
    g_HttpCorrect := StartAsyncHttp(combinedPrompt, "correct")
    g_CorrectPending := true
    g_TranslatePending := true
  }
  
  ; 启动轮询定时器
  SetTimer(CheckAsyncResults, 100)
}

StartAsyncHttp(prompt, requestType)
{
  global g_StreamFileCorrect, g_StreamFileTranslate, g_StreamPidCorrect, g_StreamPidTranslate
  global g_StreamContentCorrect, g_StreamContentTranslate
  
  ; 转义 prompt 用于 JSON
  prompt := StrReplace(prompt, "\", "\\")
  prompt := StrReplace(prompt, "`"", "\`"")
  prompt := StrReplace(prompt, "`n", "\n")
  prompt := StrReplace(prompt, "`r", "\r")
  prompt := StrReplace(prompt, "`t", "\t")
  
  ; 设置临时文件
  if (requestType = "correct") {
    g_StreamFileCorrect := A_Temp . "\ollama_stream_correct.txt"
    g_StreamContentCorrect := ""
    streamFile := g_StreamFileCorrect
    jsonFile := A_Temp . "\ollama_request_correct.json"
  } else {
    g_StreamFileTranslate := A_Temp . "\ollama_stream_translate.txt"
    g_StreamContentTranslate := ""
    streamFile := g_StreamFileTranslate
    jsonFile := A_Temp . "\ollama_request_translate.json"
  }
  
  ; 删除旧文件
  try FileDelete(streamFile)
  try FileDelete(jsonFile)
  
  ; 构建 JSON (使用流式)
  json := '{"model":"qwen3:latest","prompt":"' . prompt . '","stream":true,"options":{"temperature":0,"num_predict":1024}}'
  
  ; 将 JSON 写入临时文件
  try {
    FileAppend(json, jsonFile, "UTF-8")
  } catch {
    return 0
  }
  
  ; 使用 PowerShell 发起流式请求并写入文件（使用共享写入模式）
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
  . "$fs = New-Object System.IO.FileStream('" . streamFile . "', [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite);"
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
  
  ; 启动 PowerShell 进程
  try {
    Run('powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "' . psScript . '"', , "Hide", &outPid)
    if (requestType = "correct")
      g_StreamPidCorrect := outPid
    else
      g_StreamPidTranslate := outPid
    return outPid
  } catch {
    return 0
  }
}

CheckAsyncResults()
{
  global g_CorrectPending, g_TranslatePending
  global g_IsChineseMode, g_CorrectRequested, g_TranslateRequested, g_CurrentText
  global g_CorrectEditCtrl, g_TranslateEditCtrl
  global g_StreamFileCorrect, g_StreamFileTranslate
  global g_StreamContentCorrect, g_StreamContentTranslate
  global g_StreamPidCorrect, g_StreamPidTranslate
  
  ; 检查组合结果（一次调用同时返回纠错和翻译）
  if (g_CorrectPending && g_StreamFileCorrect != "") {
    if (IsStreamComplete(g_StreamFileCorrect)) {
      Sleep(200)
      result := ReadStreamFile(g_StreamFileCorrect, &g_StreamContentCorrect)
      if (result != "") {
        ; 解析组合结果
        ParseCombinedResult(result)
      }
      g_CorrectPending := false
      g_TranslatePending := false
    }
  }
  
  ; 如果都完成了，停止定时器
  if (!g_CorrectPending && !g_TranslatePending) {
    SetTimer(CheckAsyncResults, 0)
    
    ; 收到响应后显示窗口
    global g_PendingShowGui, g_MainGui, g_QuestionEditCtrl
    if (g_PendingShowGui && g_MainGui != "") {
      g_PendingShowGui := false
      g_MainGui.Show()
      g_QuestionEditCtrl.Focus()
      ; 启动悬停检测定时器
      SetTimer(CheckTtsHover, 200)
    }
  }
}

ParseCombinedResult(result)
{
  global g_CorrectEditCtrl, g_TranslateEditCtrl, g_ExplainEditCtrl
  
  correctPart := ""
  translatePart := ""
  
  ; 先将字面 \n 转换为真正的换行符
  result := StrReplace(result, "\n", "`n")
  
  ; 解析 ===CORRECT=== 和 ===TRANSLATE=== 分隔的内容
  if (InStr(result, "===CORRECT===") && InStr(result, "===TRANSLATE===")) {
    ; 提取纠错部分
    correctStart := InStr(result, "===CORRECT===") + StrLen("===CORRECT===")
    translateStart := InStr(result, "===TRANSLATE===")
    correctPart := Trim(SubStr(result, correctStart, translateStart - correctStart), " `t`n`r")
    
    ; 提取翻译部分
    translatePart := Trim(SubStr(result, translateStart + StrLen("===TRANSLATE===")), " `t`n`r")
  } else {
    ; 无法解析，整个作为纠错结果
    correctPart := result
  }
  
  ; 更新纠错结果
  if (correctPart != "") {
    UpdateCorrectResult(correctPart)
  }
  
  ; 更新翻译结果
  if (translatePart != "") {
    UpdateTranslateResult(translatePart)
  }
}

IsStreamComplete(filePath)
{
  if (!FileExist(filePath))
    return false
  try {
    f := FileOpen(filePath, "r", "UTF-8")
    if (!f)
      return false
    content := f.Read()
    f.Close()
    return InStr(content, '"done":true')
  } catch {
    return false
  }
}

ReadStreamFile(filePath, &accumulatedContent)
{
  if (!FileExist(filePath))
    return ""
  
  try {
    ; 使用共享读取模式打开文件
    f := FileOpen(filePath, "r", "UTF-8")
    if (!f)
      return accumulatedContent
    content := f.Read()
    f.Close()
  } catch {
    return accumulatedContent
  }
  
  ; 解析流式 JSON 行 - 使用正则表达式
  result := ""
  Loop Parse, content, "`n", "`r"
  {
    line := Trim(A_LoopField)
    if (line = "")
      continue
    ; 使用正则提取 response 字段
    if RegExMatch(line, '"response":"([^"]*)"', &m) {
      token := m[1]
      ; 反转义
      token := StrReplace(token, "\\n", "`n")
      token := StrReplace(token, "\\r", "`r")
      token := StrReplace(token, "\\t", "`t")
      token := StrReplace(token, "\`"", "`"")
      token := StrReplace(token, "\\\\", "\")
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
  
  if (result != "")
    accumulatedContent := result
  
  return accumulatedContent
}

Gui_Retry(*)
{
  global g_OrigEditCtrl, g_TranslateEditCtrl, g_CorrectEditCtrl, g_IsChineseMode
  global g_CorrectLabelCtrl, g_TranslateLabelCtrl, g_SelectedResult
  global g_CorrectRequested, g_TranslateRequested, g_TranslateResult, g_CorrectResult
  
  ; 重置请求状态
  g_CorrectRequested := false
  g_TranslateRequested := false
  g_TranslateResult := ""
  g_CorrectResult := ""
  
  newText := Trim(g_OrigEditCtrl.Value)
  if (newText = "")
    return
  
  g_IsChineseMode := RegExMatch(newText, "[\x{4e00}-\x{9fff}]")
  
  if g_IsChineseMode {
    correctLabel := "纠错 (中文润色):"
    translateLabel := "翻译 (中→英):"
    g_TranslateLabelCtrl.Text := "✓ " . translateLabel
    g_CorrectLabelCtrl.Text := "   " . correctLabel
    g_SelectedResult := "translate"
  } else {
    correctLabel := "纠错 (英文润色):"
    translateLabel := "翻译 (英→中):"
    g_CorrectLabelCtrl.Text := "✓ " . correctLabel
    g_TranslateLabelCtrl.Text := "   " . translateLabel
    g_SelectedResult := "correct"
  }
  
  ; 只显示当前选中的加载状态
  if (g_SelectedResult = "translate") {
    g_TranslateEditCtrl.Value := "正在翻译..."
    g_CorrectEditCtrl.Value := "(切换后加载)"
  } else {
    g_CorrectEditCtrl.Value := "正在纠错..."
    g_TranslateEditCtrl.Value := "(切换后加载)"
  }
  
  StartAsyncRequests(newText, g_SelectedResult)
}

Gui_ToggleSelect(*)
{
  global g_SelectedResult, g_CorrectLabelCtrl, g_TranslateLabelCtrl, g_IsChineseMode
  global g_CorrectRequested, g_TranslateRequested, g_CurrentText
  global g_CorrectEditCtrl, g_TranslateEditCtrl, g_AnswerEditCtrl
  
  if (g_IsChineseMode) {
    correctLabel := "纠错 (中文润色):"
    translateLabel := "翻译 (中→英):"
  } else {
    correctLabel := "纠错 (英文润色):"
    translateLabel := "翻译 (英→中):"
  }
  
  ; 获取当前焦点控件
  focusedHwnd := ControlGetFocus("A")
  
  ; 在翻译、纠错、AI回答三个结果框之间循环切换焦点
  if (focusedHwnd = g_TranslateEditCtrl.Hwnd) {
    g_CorrectEditCtrl.Focus()
    g_SelectedResult := "correct"
    g_CorrectLabelCtrl.Text := "✓ " . correctLabel
    g_TranslateLabelCtrl.Text := "   " . translateLabel
  } else if (focusedHwnd = g_CorrectEditCtrl.Hwnd) {
    g_AnswerEditCtrl.Focus()
  } else {
    g_TranslateEditCtrl.Focus()
    g_SelectedResult := "translate"
    g_CorrectLabelCtrl.Text := "   " . correctLabel
    g_TranslateLabelCtrl.Text := "✓ " . translateLabel
  }
}

Gui_ToggleFocus(*)
{
  global g_OrigEditCtrl, g_QuestionEditCtrl
  
  ; 获取当前焦点控件
  focusedHwnd := ControlGetFocus("A")
  
  ; 在原文输入框和 AI 问题输入框之间切换
  if (focusedHwnd = g_OrigEditCtrl.Hwnd) {
    g_QuestionEditCtrl.Focus()
  } else {
    g_OrigEditCtrl.Focus()
  }
}

Gui_HandleEnter(guiObj, *)
{
  global g_OrigEditCtrl, g_QuestionEditCtrl
  
  ; 获取当前焦点控件
  focusedHwnd := ControlGetFocus("A")
  
  ; 根据焦点位置决定操作
  if (focusedHwnd = g_QuestionEditCtrl.Hwnd) {
    Gui_SendQuestion()
  } else {
    Gui_Apply(guiObj)
  }
}

UpdateTranslateResult(result)
{
  global g_TranslateResult, g_TranslateEditCtrl, g_MainGui
  ; 去除翻译结果中的反斜杠
  result := StrReplace(result, "\", "")
  g_TranslateResult := result
  if (g_TranslateEditCtrl != "") {
    try {
      g_TranslateEditCtrl.Value := result
    } catch {
      g_TranslateEditCtrl := ""
    }
  }
}

UpdateCorrectResult(result)
{
  global g_CorrectResult, g_CorrectEditCtrl, g_ExplainEditCtrl, g_CorrectedText, g_IsChineseMode
  g_CorrectResult := result
  
  ; 英文模式：解析纠正文本和解释
  if (!g_IsChineseMode) {
    corrected := ""
    explanation := ""
    
    ; 先将字面 \n 转换为真正的换行符，并清理标记
    result := StrReplace(result, "\n", "`n")
    result := StrReplace(result, "===CORRECT===", "")
    result := StrReplace(result, "===TRANSLATE===", "")
    result := Trim(result)
    
    if (InStr(result, "---")) {
      ; 有分隔符：按 --- 分割
      parts := StrSplit(result, "---", , 2)
      corrected := Trim(parts[1], " `t`n`r")
      explanation := (parts.Length > 1) ? Trim(parts[2], " `t`n`r") : ""
    } else if (RegExMatch(result, "^(.+?)\s*(错误|1\.|1、)", &m)) {
      ; 无分隔符：尝试找到第一个中文解释的开始位置
      corrected := Trim(m[1])
      explanation := Trim(SubStr(result, StrLen(m[1]) + 1))
    } else {
      ; 无法分割：整个作为纠正文本
      corrected := result
    }
    
    ; 去除解释中的反斜杠
    explanation := StrReplace(explanation, "\", "")
    
    ; 清理纠正文本（只保留第一行英文句子）
    if (InStr(corrected, "`n")) {
      firstLine := Trim(StrSplit(corrected, "`n")[1])
      if (firstLine != "" && !RegExMatch(firstLine, "[\x{4e00}-\x{9fff}]"))
        corrected := firstLine
    }
    corrected := Trim(corrected)
    ; 去除纠错文本中的反斜杠
    corrected := StrReplace(corrected, "\", "")
    g_CorrectedText := corrected
    
    if (g_CorrectEditCtrl != "") {
      try {
        g_CorrectEditCtrl.Value := corrected
      } catch {
        g_CorrectEditCtrl := ""
      }
    }
    if (g_ExplainEditCtrl != "") {
      try {
        g_ExplainEditCtrl.Value := explanation
      } catch {
        g_ExplainEditCtrl := ""
      }
    }
  } else {
    ; 中文模式或无分隔符：直接显示
    ; 去除反斜杠
    result := StrReplace(result, "\", "")
    g_CorrectedText := result
    if (g_CorrectEditCtrl != "") {
      try {
        g_CorrectEditCtrl.Value := result
      } catch {
        g_CorrectEditCtrl := ""
      }
    }
  }
}

Gui_Apply(guiObj, *)
{
  global g_TranslateResult, g_CorrectResult, g_OldClip, g_SelectedResult
  global g_MainGui, g_TranslateEditCtrl, g_CorrectEditCtrl, g_OrigEditCtrl
  global g_CorrectedText, g_IsChineseMode
  guiObj.Destroy()
  g_MainGui := ""
  g_TranslateEditCtrl := ""
  g_CorrectEditCtrl := ""
  g_OrigEditCtrl := ""
  
  ; 英文纠错时使用分离后的纠正文本（不含解释）
  if (g_SelectedResult = "translate")
    result := g_TranslateResult
  else if (!g_IsChineseMode && g_CorrectedText != "")
    result := g_CorrectedText
  else
    result := g_CorrectResult
  
  if (result != "" && !InStr(result, "失败")) {
    A_Clipboard := result
    Sleep(30)
    Send("^a")
    Sleep(30)
    Send("^v")
    Sleep(100)
    A_Clipboard := g_OldClip
  } else {
    A_Clipboard := g_OldClip
  }
}

Gui_PlayOriginal(*)
{
  global g_OrigEditCtrl, g_IsChineseMode
  static tempFile := ""
  
  ; 只在英文模式下朗读
  if (g_IsChineseMode)
    return
  
  text := Trim(g_OrigEditCtrl.Value)
  if (text = "")
    return
  
  ; 停止之前的播放
  try {
    SoundPlay("NonExistent.zzz")
  }
  Sleep(50)
  
  ; 删除旧文件
  if (tempFile != "" && FileExist(tempFile)) {
    try {
      FileDelete(tempFile)
    }
  }
  
  ; 使用 Google TTS API
  try {
    ; 构建 Google TTS URL
    encodedText := ""
    Loop Parse, text
    {
      char := A_LoopField
      if RegExMatch(char, "[a-zA-Z0-9\-_.~]")
        encodedText .= char
      else
        encodedText .= "%" . Format("{:02X}", Ord(char))
    }
    
    ttsUrl := "https://translate.google.com/translate_tts?ie=UTF-8&tl=en-US&client=tw-ob&q=" . encodedText
    
    ; 使用固定文件名
    tempFile := A_Temp . "\ahk_tts_audio.mp3"
    
    http := ComObject("WinHttp.WinHttpRequest.5.1")
    http.Open("GET", ttsUrl, false)
    http.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
    http.SetRequestHeader("Referer", "https://translate.google.com/")
    http.Send()
    http.WaitForResponse()
    
    if (http.Status = 200) {
      ; 保存音频文件
      adoStream := ComObject("ADODB.Stream")
      adoStream.Type := 1  ; Binary
      adoStream.Open()
      adoStream.Write(http.ResponseBody)
      adoStream.SaveToFile(tempFile, 2)  ; 2 = overwrite
      adoStream.Close()
      
      ; 播放音频
      SoundPlay(tempFile)
    }
  } catch Error as e {
    ; 静默失败
  }
}

Gui_PlayCorrect(*)
{
  global g_CorrectEditCtrl
  PlayTtsText(g_CorrectEditCtrl.Value)
}

Gui_PlayTranslate(*)
{
  global g_TranslateEditCtrl
  PlayTtsText(g_TranslateEditCtrl.Value)
}

PlayTtsText(text)
{
  static tempFile := ""
  
  text := Trim(text)
  if (text = "" || InStr(text, "正在") || InStr(text, "切换后"))
    return
  
  ; 停止之前的播放
  try {
    SoundPlay("NonExistent.zzz")
  }
  Sleep(50)
  
  ; 删除旧文件
  if (tempFile != "" && FileExist(tempFile)) {
    try {
      FileDelete(tempFile)
    }
  }
  
  ; 使用 Google TTS API
  try {
    encodedText := ""
    Loop Parse, text
    {
      char := A_LoopField
      if RegExMatch(char, "[a-zA-Z0-9\-_.~]")
        encodedText .= char
      else
        encodedText .= "%" . Format("{:02X}", Ord(char))
    }
    
    ttsUrl := "https://translate.google.com/translate_tts?ie=UTF-8&tl=en-US&client=tw-ob&q=" . encodedText
    tempFile := A_Temp . "\ahk_tts_audio.mp3"
    
    http := ComObject("WinHttp.WinHttpRequest.5.1")
    http.Open("GET", ttsUrl, false)
    http.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
    http.SetRequestHeader("Referer", "https://translate.google.com/")
    http.Send()
    http.WaitForResponse()
    
    if (http.Status = 200) {
      adoStream := ComObject("ADODB.Stream")
      adoStream.Type := 1
      adoStream.Open()
      adoStream.Write(http.ResponseBody)
      adoStream.SaveToFile(tempFile, 2)
      adoStream.Close()
      
      SoundPlay(tempFile)
    }
  } catch {
  }
}

CheckTtsHover()
{
  global g_TtsOrigCtrl, g_TtsCorrectCtrl, g_TtsTranslateCtrl
  global g_MainGui, g_TtsPlaying, g_IsChineseMode, g_HoverTarget
  static lastHoverCtrl := ""

  ; 如果窗口已关闭，停止定时器
  if (g_MainGui = "") {
    SetTimer(CheckTtsHover, 0)
    return
  }

  ; 检测鼠标在哪个朗读图标上
  currentHover := ""
  try {
    MouseGetPos(&mx, &my, &winUnder, &ctrlUnder, 2)
    ; 中文模式只有翻译图标，英文模式有三个图标
    if (!g_IsChineseMode && ctrlUnder = g_TtsOrigCtrl.Hwnd)
      currentHover := "orig"
    else if (!g_IsChineseMode && ctrlUnder = g_TtsCorrectCtrl.Hwnd)
      currentHover := "correct"
    else if (ctrlUnder = g_TtsTranslateCtrl.Hwnd)
      currentHover := "translate"
  } catch {
  }

  if (currentHover != "" && currentHover != lastHoverCtrl) {
    ; 进入新图标，开始播放
    g_TtsPlaying := true
    g_HoverTarget := currentHover
    PlayTtsLoop()
  } else if (currentHover = "" && lastHoverCtrl != "") {
    ; 离开图标，停止播放
    g_TtsPlaying := false
    g_HoverTarget := ""
    try {
      SoundPlay("NonExistent.zzz")
    }
  }

  lastHoverCtrl := currentHover
}

PlayTtsLoop()
{
  global g_TtsPlaying, g_IsChineseMode, g_HoverTarget
  global g_OrigEditCtrl, g_CorrectEditCtrl, g_TranslateEditCtrl
  static tempFile := ""

  if (!g_TtsPlaying || g_HoverTarget = "")
    return

  ; 根据悬停目标获取文本
  if (g_HoverTarget = "orig")
    text := Trim(g_OrigEditCtrl.Value)
  else if (g_HoverTarget = "correct")
    text := Trim(g_CorrectEditCtrl.Value)
  else if (g_HoverTarget = "translate")
    text := Trim(g_TranslateEditCtrl.Value)
  else
    return

  if (text = "" || InStr(text, "正在") || InStr(text, "切换后"))
    return

  ; 使用 Google TTS API
  try {
    encodedText := ""
    Loop Parse, text
    {
      char := A_LoopField
      if RegExMatch(char, "[a-zA-Z0-9\-_.~]")
        encodedText .= char
      else
        encodedText .= "%" . Format("{:02X}", Ord(char))
    }

    ttsUrl := "https://translate.google.com/translate_tts?ie=UTF-8&tl=en-US&client=tw-ob&q=" . encodedText
    tempFile := A_Temp . "\ahk_tts_audio.mp3"

    http := ComObject("WinHttp.WinHttpRequest.5.1")
    http.Open("GET", ttsUrl, false)
    http.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
    http.SetRequestHeader("Referer", "https://translate.google.com/")
    http.Send()
    http.WaitForResponse()

    if (http.Status = 200) {
      adoStream := ComObject("ADODB.Stream")
      adoStream.Type := 1
      adoStream.Open()
      adoStream.Write(http.ResponseBody)
      adoStream.SaveToFile(tempFile, 2)
      adoStream.Close()

      ; 播放并等待完成
      SoundPlay(tempFile, "Wait")

      ; 播放完毕后如果还在悬停，继续播放
      if (g_TtsPlaying)
        SetTimer(PlayTtsLoop, -100)
    }
  } catch {
  }
}

Gui_SendQuestion(*)
{
  global g_QuestionEditCtrl, g_AnswerEditCtrl, g_SendBtnCtrl
  global g_ChatPending, g_StreamFileChat, g_StreamPidChat, g_StreamContentChat
  
  question := Trim(g_QuestionEditCtrl.Value)
  if (question = "")
    return
  
  ; 如果有正在进行的请求，先停止
  if (g_ChatPending && g_StreamPidChat > 0) {
    SetTimer(CheckChatResult, 0)
    try ProcessClose(g_StreamPidChat)
    g_StreamPidChat := 0
  }
  
  ; 禁用发送按钮
  g_SendBtnCtrl.Enabled := false
  g_AnswerEditCtrl.Value := "正在思考..."
  
  ; 启动异步请求
  g_ChatPending := true
  g_StreamContentChat := ""
  StartChatAsync(question)
}

StartChatAsync(question)
{
  global g_StreamFileChat, g_StreamPidChat, g_StreamContentChat
  
  ; 转义 prompt 用于 JSON
  prompt := "/no_think 请用纯文本回答，不要使用 Markdown 格式（如 **加粗** 或 * 列表）。问题: " . question
  prompt := StrReplace(prompt, "\", "\\")
  prompt := StrReplace(prompt, "`"", "\`"")
  prompt := StrReplace(prompt, "`n", "\n")
  prompt := StrReplace(prompt, "`r", "\r")
  prompt := StrReplace(prompt, "`t", "\t")
  
  ; 设置临时文件
  g_StreamFileChat := A_Temp . "\ollama_stream_chat.txt"
  g_StreamContentChat := ""
  jsonFile := A_Temp . "\ollama_request_chat.json"
  
  ; 删除旧文件
  try FileDelete(g_StreamFileChat)
  try FileDelete(jsonFile)
  
  ; 构建 JSON (使用流式)
  json := '{"model":"qwen3:latest","prompt":"' . prompt . '","stream":true,"options":{"temperature":0.7,"num_predict":2048}}'
  
  ; 将 JSON 写入临时文件
  try {
    FileAppend(json, jsonFile, "UTF-8")
  } catch {
    return
  }
  
  ; 使用 PowerShell 发起流式请求
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
  . "$fs = New-Object System.IO.FileStream('" . g_StreamFileChat . "', [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite);"
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
  
  ; 启动 PowerShell 进程
  try {
    Run('powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "' . psScript . '"', , "Hide", &outPid)
    g_StreamPidChat := outPid
  } catch {
    return
  }
  
  ; 启动轮询定时器
  SetTimer(CheckChatResult, 100)
}

CheckChatResult()
{
  global g_ChatPending, g_StreamFileChat, g_StreamContentChat, g_StreamPidChat
  global g_AnswerEditCtrl, g_SendBtnCtrl
  
  if (!g_ChatPending)
    return
  
  ; 检查控件是否已被销毁
  if (g_AnswerEditCtrl = "" || g_SendBtnCtrl = "") {
    SetTimer(CheckChatResult, 0)
    return
  }
  
  ; 检查是否完成
  if (g_StreamFileChat != "" && FileExist(g_StreamFileChat)) {
    if (IsStreamComplete(g_StreamFileChat)) {
      Sleep(200)
      result := ReadStreamFile(g_StreamFileChat, &g_StreamContentChat)
      if (result != "" && g_AnswerEditCtrl != "") {
        ; 转换换行符并去除反斜杠
        result := StrReplace(result, "\n", "`n")
        result := StrReplace(result, "\", "")
        try g_AnswerEditCtrl.Value := result
      }
      g_ChatPending := false
      if (g_SendBtnCtrl != "")
        try g_SendBtnCtrl.Enabled := true
      SetTimer(CheckChatResult, 0)
    }
  }
}

Gui_Close(guiObj, *)
{
  global g_OldClip, g_TtsPlaying, g_HoverTarget
  global g_StreamPidCorrect, g_StreamPidTranslate, g_CorrectPending, g_TranslatePending
  global g_MainGui, g_TranslateEditCtrl, g_CorrectEditCtrl, g_OrigEditCtrl
  global g_StreamPidChat, g_ChatPending, g_QuestionEditCtrl, g_AnswerEditCtrl, g_SendBtnCtrl
  
  ; 终止正在运行的 PowerShell 进程
  if (g_StreamPidCorrect > 0) {
    try ProcessClose(g_StreamPidCorrect)
    g_StreamPidCorrect := 0
  }
  if (g_StreamPidTranslate > 0) {
    try ProcessClose(g_StreamPidTranslate)
    g_StreamPidTranslate := 0
  }
  if (g_StreamPidChat > 0) {
    try ProcessClose(g_StreamPidChat)
    g_StreamPidChat := 0
  }
  g_CorrectPending := false
  g_TranslatePending := false
  g_ChatPending := false
  SetTimer(CheckAsyncResults, 0)
  SetTimer(CheckChatResult, 0)
  
  g_TtsPlaying := false
  g_HoverTarget := ""
  SetTimer(CheckTtsHover, 0)
  guiObj.Destroy()
  g_MainGui := ""
  g_TranslateEditCtrl := ""
  g_CorrectEditCtrl := ""
  g_OrigEditCtrl := ""
  g_QuestionEditCtrl := ""
  g_AnswerEditCtrl := ""
  g_SendBtnCtrl := ""
  A_Clipboard := g_OldClip
}

!SC029::
{
  global g_MainGui, g_GuiHidden
  if (g_MainGui = "")
    return
  
  if (g_GuiHidden) {
    ; 窗口已隐藏，恢复显示
    g_MainGui.Show()
    WinActivate("ahk_id " g_MainGui.Hwnd)
    g_GuiHidden := false
  } else {
    ; 隐藏窗口到托盘
    g_MainGui.Hide()
    g_GuiHidden := true
  }
}

^!Enter::
^!NumpadEnter::
{
  global g_OldClip, g_IsChineseMode
  g_OldClip := ClipboardAll()
  A_Clipboard := ""

  ; 先尝试复制当前选中的文字
  Send("^c")
  ClipWait(0.3)
  text := Trim(A_Clipboard)

  ; 如果没有选中文字，则全选
  if (text = "") {
    Send("^a")
    Sleep(50)
    Send("^c")
    ClipWait(0.5)
    text := Trim(A_Clipboard)
  }

  ; 即使文本为空也显示窗口（可使用 AI 助手）
  ShowMainGui(text)
}
