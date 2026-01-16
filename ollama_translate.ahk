;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; Ollama 翻译/纠错 - Alt+`: 中文→翻译英文，英文→纠正表达 (再按隐藏/显示)
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

; Prompt 模板相关
g_ConfigFile := A_ScriptDir . "\ollama_config.ini"
g_PromptList := []
g_PromptNames := []
g_SelectedPrompt := ""
g_PromptDropdown := ""
g_PromptManageBtn := ""

; 初始化 Prompt 模板
InitPrompts()

InitPrompts()
{
  global g_ConfigFile, g_PromptList, g_PromptNames, g_SelectedPrompt
  
  ; 如果文件不存在，创建默认配置
  if (!FileExist(g_ConfigFile)) {
    defaultConfig := "
(
[Settings]
SelectedPrompt=无

[Prompt_通用助手]
prompt=你是一个有帮助的助手。请用简洁的中文回答问题。

[Prompt_代码解释]
prompt=请解释以下代码的功能和工作原理，用中文回答：

[Prompt_翻译助手]
prompt=请将以下内容翻译成中文，保持原意：

[Prompt_写作润色]
prompt=请帮我润色以下文字，使其更加流畅自然：

[Prompt_总结摘要]
prompt=请用简洁的语言总结以下内容的要点：
)"
    FileAppend(defaultConfig, g_ConfigFile, "UTF-8-RAW")
  }
  
  ; 读取所有 prompt
  LoadPrompts()
  
  ; 从配置文件读取上次选中的模板
  savedPrompt := ""
  try savedPrompt := IniRead(g_ConfigFile, "Settings", "SelectedPrompt", "")
  
  ; 如果保存的模板存在，使用它；否则使用第一个
  if (savedPrompt != "" && HasPromptName(savedPrompt))
    g_SelectedPrompt := savedPrompt
  else if (g_PromptNames.Length > 0)
    g_SelectedPrompt := g_PromptNames[1]
}

LoadPrompts()
{
  global g_ConfigFile, g_PromptList, g_PromptNames
  
  g_PromptList := []
  g_PromptNames := []
  
  if (!FileExist(g_ConfigFile))
    return
  
  content := FileRead(g_ConfigFile, "UTF-8")
  currentName := ""
  currentPrompt := ""
  
  Loop Parse, content, "`n", "`r"
  {
    line := Trim(A_LoopField)
    if (line = "")
      continue
    
    ; 检测 section 名称 [Prompt_xxx]，跳过 [Settings]
    if (RegExMatch(line, "^\[Prompt_(.+)\]$", &m)) {
      ; 保存上一个（允许空 prompt）
      if (currentName != "") {
        g_PromptNames.Push(currentName)
        g_PromptList.Push({name: currentName, prompt: currentPrompt})
      }
      currentName := m[1]
      currentPrompt := ""
    } else if (RegExMatch(line, "^prompt=(.*)$", &m) && currentName != "") {
      currentPrompt := m[1]
    }
  }
  
  ; 保存最后一个（允许空 prompt）
  if (currentName != "") {
    g_PromptNames.Push(currentName)
    g_PromptList.Push({name: currentName, prompt: currentPrompt})
  }
  
  ; 在列表开头插入"无"选项（如果不存在）
  if (g_PromptNames.Length = 0 || g_PromptNames[1] != "无") {
    g_PromptNames.InsertAt(1, "无")
    g_PromptList.InsertAt(1, {name: "无", prompt: ""})
  }
}

SavePrompts()
{
  global g_ConfigFile, g_PromptList, g_SelectedPrompt
  
  ; 构建配置文件内容：Settings + Prompts
  content := "[Settings]`n"
  content .= "SelectedPrompt=" . g_SelectedPrompt . "`n`n"
  
  for item in g_PromptList {
    ; 跳过"无"选项（动态添加的，不需要保存）
    if (item.name = "无")
      continue
    content .= "[Prompt_" . item.name . "]`n"
    content .= "prompt=" . item.prompt . "`n`n"
  }
  
  try FileDelete(g_ConfigFile)
  FileAppend(content, g_ConfigFile, "UTF-8-RAW")
}

GetPromptByName(name)
{
  global g_PromptList
  
  for item in g_PromptList {
    if (item.name = name)
      return item.prompt
  }
  return ""
}

OllamaCall(prompt)
{
  ; 构建 JSON
  prompt := StrReplace(prompt, "\", "\\")
  prompt := StrReplace(prompt, "`"", "\`"")
  prompt := StrReplace(prompt, "`n", "\n")
  prompt := StrReplace(prompt, "`r", "\r")
  prompt := StrReplace(prompt, "`t", "\t")
  
  ; 系统提示：强制禁用 Markdown 和符号
  sysPrompt := "纯文本输出，不要用任何符号（如反斜杠、星号、井号）包裹或强调单词。"
  
  json := "{`"model`":`"huihui_ai/qwen3-abliterated:8b-v2`",`"system`":`"" . sysPrompt . "`",`"prompt`":`"" . prompt . "`",`"stream`":false,`"options`":{`"temperature`":0,`"num_predict`":1024,`"think`":true}}"
  
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
    prompt := "Translate to English. Keep the exact same formatting, including punctuation marks, line breaks, and spacing. Output only the translation:`n" . text
  else
    prompt := "Translate to Chinese. Keep the exact same formatting, including punctuation marks, line breaks, and spacing. Output only the translation:`n" . text
  return OllamaCall(prompt)
}

OllamaCorrect(text, isChinese)
{
  if isChinese
    prompt := "You are a Chinese language tutor. Correct and improve the following Chinese text. Fix grammar, punctuation, and improve expression while keeping the original meaning. Output only the corrected text without any explanation:`n" . text
  else
    prompt := "Correct this English text for a Chinese learner.`n`nRules:`n1. First line: ONLY the corrected sentence, nothing else`n2. Second line: exactly three dashes: ---`n3. Then list errors in Chinese: 错误1: 原文 → 修正 (解释)`n`nExample output:`nI am a real team member.`n---`n错误1: i → I (句首字母需要大写)`n错误2: real team → a real team (需要冠词 a)`n`nNow correct: " . text
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
    g_CorrectEditCtrl := g_MainGui.AddEdit("xm w500 h60 ReadOnly", "正在处理...")
    ; 错误解释框
    g_MainGui.AddText("w500", "错误解释:")
    g_ExplainEditCtrl := g_MainGui.AddEdit("w500 h100 ReadOnly", "正在处理...")
    g_TranslateLabelCtrl := g_MainGui.AddText("w120 Section", "   " . translateLabel)
    g_TtsTranslateCtrl := g_MainGui.AddText("x+5 ys cGray", "🔊")
    g_TtsTranslateCtrl.OnEvent("Click", Gui_PlayTranslate)
    g_TranslateEditCtrl := g_MainGui.AddEdit("xm w500 h40 ReadOnly", "正在处理...")
    g_SelectedResult := "correct"
  }
  
  ; ========== 右侧面板：AI 问答 ==========
  g_MainGui.AddText("x530 y10 w60 Section", "Prompt:")
  promptList := ""
  for name in g_PromptNames {
    promptList .= (promptList = "" ? "" : "|") . name
  }
  g_PromptDropdown := g_MainGui.AddDropDownList("x+5 yp w280", StrSplit(promptList, "|"))
  if (g_SelectedPrompt != "")
    g_PromptDropdown.Text := g_SelectedPrompt
  g_PromptDropdown.OnEvent("Change", Gui_PromptChanged)
  g_PromptManageBtn := g_MainGui.AddButton("x+5 yp w50", "管理")
  g_PromptManageBtn.OnEvent("Click", Gui_ManagePrompts)
  
  g_QuestionEditCtrl := g_MainGui.AddEdit("xs w330 h50", "")
  g_SendBtnCtrl := g_MainGui.AddButton("x+5 yp h50 w60", "发送")
  g_SendBtnCtrl.OnEvent("Click", Gui_SendQuestion)
  g_MainGui.AddText("xs w400", "回答:")
  g_AnswerEditCtrl := g_MainGui.AddEdit("xs w400 h220 ReadOnly", "")
  
  ; ========== 底部提示 ==========
  g_MainGui.AddText("xm w930 cGray", "Tab 切换输入框 | Ctrl+Tab 切换结果焦点 | Enter 替换/发送 | Ctrl+Enter 强制替换 | Alt+`` 隐藏/显示")
  
  g_MainGui.OnEvent("Close", Gui_Hide)
  
  ; 窗口创建后暂不显示，等待 AI 响应后再显示
  ; g_MainGui.Show()
  
  HotIfWinActive("ahk_id " g_MainGui.Hwnd)
  Hotkey("Enter", Gui_HandleEnter.Bind(g_MainGui), "On")
  Hotkey("NumpadEnter", Gui_HandleEnter.Bind(g_MainGui), "On")
  Hotkey("^Enter", Gui_Apply.Bind(g_MainGui), "On")
  Hotkey("^NumpadEnter", Gui_Apply.Bind(g_MainGui), "On")
  Hotkey("^Tab", Gui_ToggleSelect, "On")
  Hotkey("Tab", Gui_ToggleFocus, "On")
  Hotkey("^v", Gui_PasteAsText, "On")
  Hotkey("Escape", Gui_Close.Bind(g_MainGui), "On")
  Hotkey("^Backspace", Gui_DeleteWord, "On")
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
    if (g_ExplainEditCtrl != "")
      g_ExplainEditCtrl.Value := ""
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
  global g_StreamPidCorrect, g_StreamPidTranslate
  
  ; 终止之前正在运行的请求
  if (g_StreamPidCorrect > 0) {
    try ProcessClose(g_StreamPidCorrect)
    g_StreamPidCorrect := 0
  }
  if (g_StreamPidTranslate > 0) {
    try ProcessClose(g_StreamPidTranslate)
    g_StreamPidTranslate := 0
  }
  
  g_CurrentText := text
  isChinese := RegExMatch(text, "[\x{4e00}-\x{9fff}]")
  g_IsChineseMode := isChinese
  
  ; 一次调用同时完成纠错和翻译
  if (requestType = "default" || !g_CorrectRequested) {
    g_CorrectRequested := true
    g_TranslateRequested := true
    
    if isChinese {
      ; 中文：润色 + 翻译成英文
      combinedPrompt := "请对以下中文进行润色和翻译。不要使用Markdown格式。`n`n输出格式(严格遵守):`n===CORRECT===`n润色后的中文`n===TRANSLATE===`n英文翻译`n`n原文: " . text
    } else {
      ; 英文：纠错+解释 + 翻译成中文
      combinedPrompt := "纠正并翻译以下英文。纯文本输出，不要用任何符号包裹单词。`n`n格式：`n===CORRECT===`n纠正后的英文`n---`n错误: 原文 → 修正 (解释)`n===TRANSLATE===`n中文翻译`n`n英文: " . text
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
  
  ; 系统提示：强制禁用 Markdown 和符号
  sysPrompt := "纯文本输出，不要用任何符号（如反斜杠、星号、井号）包裹或强调单词。"
  
  ; 构建 JSON (使用流式，添加 system 参数)
  json := '{"model":"huihui_ai/qwen3-abliterated:8b-v2","system":"' . sysPrompt . '","prompt":"' . prompt . '","stream":true,"options":{"temperature":0,"num_predict":1024,"think":true}}'
  
  ; 将 JSON 写入临时文件
  try {
    FileAppend(json, jsonFile, "UTF-8-RAW")
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
    ; 使用正则提取 response 字段（支持转义字符）
    if RegExMatch(line, '"response":"((?:[^"\\]|\\.)*)"', &m) {
      token := m[1]
      ; 反转义 JSON 字符串
      token := StrReplace(token, "\n", "`n")
      token := StrReplace(token, "\r", "`r")
      token := StrReplace(token, "\t", "`t")
      token := StrReplace(token, '\"', '"')
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
  
  if (result != "")
    accumulatedContent := result
  
  return accumulatedContent
}

Gui_Retry(*)
{
  global g_OrigEditCtrl, g_TranslateEditCtrl, g_CorrectEditCtrl, g_IsChineseMode
  global g_CorrectLabelCtrl, g_TranslateLabelCtrl, g_SelectedResult, g_MainGui
  global g_CorrectRequested, g_TranslateRequested, g_TranslateResult, g_CorrectResult
  global g_ExplainEditCtrl
  
  newText := Trim(g_OrigEditCtrl.Value)
  if (newText = "")
    return
  
  ; 检测语言是否改变
  newIsChinese := RegExMatch(newText, "[\x{4e00}-\x{9fff}]")
  if (newIsChinese != g_IsChineseMode) {
    ; 语言模式改变，需要重新创建窗口
    try g_MainGui.Destroy()
    g_MainGui := ""
    ShowMainGui(newText)
    return
  }
  
  ; 重置请求状态
  g_CorrectRequested := false
  g_TranslateRequested := false
  g_TranslateResult := ""
  g_CorrectResult := ""
  
  ; 所有框都显示正在处理
  if (g_ExplainEditCtrl != "")
    try g_ExplainEditCtrl.Value := "正在处理..."
  g_TranslateEditCtrl.Value := "正在处理..."
  g_CorrectEditCtrl.Value := "正在处理..."
  
  StartAsyncRequests(newText, "default")
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

Gui_PasteAsText(*)
{
  global g_OrigEditCtrl, g_QuestionEditCtrl
  
  ; 使用 Windows API 直接获取剪贴板文本（解决 PixPin OCR 延迟渲染问题）
  clipText := GetClipboardText()
  
  if (clipText != "") {
    ; 获取当前焦点控件
    focusedHwnd := ControlGetFocus("A")
    
    ; 只在可编辑的输入框中粘贴（在光标位置插入，不覆盖全部内容）
    if (focusedHwnd = g_OrigEditCtrl.Hwnd || focusedHwnd = g_QuestionEditCtrl.Hwnd) {
      EditPaste(clipText, focusedHwnd)
    }
  }
}

Gui_DeleteWord(*)
{
  ; 发送 Ctrl+Shift+Left 选中前一个单词，然后删除
  Send("^+{Left}{Delete}")
}

GetClipboardText()
{
  ; 使用 Windows API 直接获取剪贴板文本
  ; 这可以触发延迟渲染，解决 PixPin OCR 等软件的兼容性问题
  
  CF_UNICODETEXT := 13
  
  ; 打开剪贴板
  if !DllCall("OpenClipboard", "Ptr", 0)
    return A_Clipboard  ; 回退到 AHK 方式
  
  ; 获取 Unicode 文本数据
  hData := DllCall("GetClipboardData", "UInt", CF_UNICODETEXT, "Ptr")
  if (!hData) {
    DllCall("CloseClipboard")
    return A_Clipboard  ; 回退到 AHK 方式
  }
  
  ; 锁定内存并获取指针
  pData := DllCall("GlobalLock", "Ptr", hData, "Ptr")
  if (!pData) {
    DllCall("CloseClipboard")
    return A_Clipboard
  }
  
  ; 读取字符串
  text := StrGet(pData, "UTF-16")
  
  ; 解锁并关闭
  DllCall("GlobalUnlock", "Ptr", hData)
  DllCall("CloseClipboard")
  
  return text
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
  
  ; 检测输入法是否处于组合状态（正在输入中文）
  if (IsImeComposing()) {
    ; 让回车键正常传递给输入法确认候选词
    Send("{Enter}")
    return
  }
  
  ; 获取当前焦点控件
  focusedHwnd := ControlGetFocus("A")
  
  ; 根据焦点位置决定操作
  if (focusedHwnd = g_QuestionEditCtrl.Hwnd) {
    Gui_SendQuestion()
  } else if (focusedHwnd = g_OrigEditCtrl.Hwnd) {
    Gui_Retry()  ; 重新翻译
  }
  ; 其他情况不做处理
}

IsImeComposing()
{
  ; 检测输入法是否处于组合/候选状态
  ; 返回 true 表示正在输入中文（有候选词）
  
  focusedHwnd := ControlGetFocus("A")
  if (!focusedHwnd)
    focusedHwnd := WinGetID("A")
  
  ; 获取输入法上下文
  hImc := DllCall("imm32\ImmGetContext", "Ptr", focusedHwnd, "Ptr")
  if (!hImc)
    return false
  
  ; 检查组合字符串长度 (GCS_COMPSTR = 0x8)
  compLen := DllCall("imm32\ImmGetCompositionStringW", "Ptr", hImc, "UInt", 0x8, "Ptr", 0, "UInt", 0)
  
  ; 释放上下文
  DllCall("imm32\ImmReleaseContext", "Ptr", focusedHwnd, "Ptr", hImc)
  
  ; 如果组合字符串长度 > 0，说明正在输入中
  return (compLen > 0)
}

UpdateTranslateResult(result)
{
  global g_TranslateResult, g_TranslateEditCtrl, g_MainGui
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
    
    ; 清理纠正文本（只保留第一行英文句子）
    if (InStr(corrected, "`n")) {
      firstLine := Trim(StrSplit(corrected, "`n")[1])
      if (firstLine != "" && !RegExMatch(firstLine, "[\x{4e00}-\x{9fff}]"))
        corrected := firstLine
    }
    corrected := Trim(corrected)
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

Gui_PromptChanged(ctrl, *)
{
  global g_SelectedPrompt
  g_SelectedPrompt := ctrl.Text
  ; 保存选中的模板到配置文件
  SavePrompts()
}

Gui_ManagePrompts(*)
{
  global g_PromptList, g_PromptNames, g_PromptDropdown, g_SelectedPrompt
  
  manageGui := Gui("+AlwaysOnTop", "管理 Prompt 模板")
  manageGui.SetFont("s10", "Microsoft YaHei")
  
  manageGui.AddText("w400", "选择模板:")
  listBox := manageGui.AddListBox("w400 h150", g_PromptNames)
  if (g_PromptNames.Length > 0)
    listBox.Choose(1)
  
  manageGui.AddText("w400", "模板名称:")
  nameEdit := manageGui.AddEdit("w400", "")
  
  manageGui.AddText("w400", "Prompt 内容:")
  promptEdit := manageGui.AddEdit("w400 h80", "")
  
  ; 选择变化时更新编辑框
  listBox.OnEvent("Change", (*) => UpdatePromptEdit(listBox, nameEdit, promptEdit))
  
  ; 按钮行
  btnAdd := manageGui.AddButton("w95", "新增")
  btnSave := manageGui.AddButton("x+10 w95", "保存")
  btnDelete := manageGui.AddButton("x+10 w95", "删除")
  btnClose := manageGui.AddButton("x+10 w95", "关闭")
  
  btnAdd.OnEvent("Click", (*) => AddPrompt(listBox, nameEdit, promptEdit))
  btnSave.OnEvent("Click", (*) => SavePromptItem(listBox, nameEdit, promptEdit))
  btnDelete.OnEvent("Click", (*) => DeletePrompt(listBox, nameEdit, promptEdit))
  btnClose.OnEvent("Click", (*) => CloseManageGui(manageGui))
  
  ; 初始加载第一个
  if (g_PromptNames.Length > 0)
    UpdatePromptEdit(listBox, nameEdit, promptEdit)
  
  manageGui.Show()
}

UpdatePromptEdit(listBox, nameEdit, promptEdit)
{
  global g_PromptList
  
  idx := listBox.Value
  if (idx > 0 && idx <= g_PromptList.Length) {
    nameEdit.Value := g_PromptList[idx].name
    promptEdit.Value := g_PromptList[idx].prompt
  }
}

AddPrompt(listBox, nameEdit, promptEdit)
{
  global g_PromptList, g_PromptNames, g_PromptDropdown
  
  newName := "新模板"
  newPrompt := ""
  
  g_PromptNames.Push(newName)
  g_PromptList.Push({name: newName, prompt: newPrompt})
  
  ; 更新列表
  listBox.Delete()
  listBox.Add(g_PromptNames)
  listBox.Choose(g_PromptNames.Length)
  
  nameEdit.Value := newName
  promptEdit.Value := newPrompt
  
  SavePrompts()
  RefreshPromptDropdown()
}

SavePromptItem(listBox, nameEdit, promptEdit)
{
  global g_PromptList, g_PromptNames, g_PromptDropdown, g_SelectedPrompt
  
  idx := listBox.Value
  if (idx <= 0 || idx > g_PromptList.Length)
    return
  
  newName := Trim(nameEdit.Value)
  newPrompt := Trim(promptEdit.Value)
  
  if (newName = "")
    return
  
  ; 如果修改的是当前选中的，同步更新
  oldName := g_PromptList[idx].name
  if (g_SelectedPrompt = oldName)
    g_SelectedPrompt := newName
  
  g_PromptList[idx].name := newName
  g_PromptList[idx].prompt := newPrompt
  g_PromptNames[idx] := newName
  
  ; 更新列表
  listBox.Delete()
  listBox.Add(g_PromptNames)
  listBox.Choose(idx)
  
  SavePrompts()
  RefreshPromptDropdown()
}

DeletePrompt(listBox, nameEdit, promptEdit)
{
  global g_PromptList, g_PromptNames, g_PromptDropdown, g_SelectedPrompt
  
  idx := listBox.Value
  if (idx <= 0 || idx > g_PromptList.Length)
    return
  
  ; 不能删除"无"选项
  if (g_PromptList[idx].name = "无") {
    MsgBox("不能删除[无]选项", "提示", "Icon!")
    return
  }
  
  ; 至少保留一个（除"无"外）
  if (g_PromptList.Length <= 2) {
    MsgBox("至少需要保留一个模板", "提示", "Icon!")
    return
  }
  
  g_PromptList.RemoveAt(idx)
  g_PromptNames.RemoveAt(idx)
  
  ; 更新列表
  listBox.Delete()
  listBox.Add(g_PromptNames)
  if (idx > g_PromptNames.Length)
    idx := g_PromptNames.Length
  listBox.Choose(idx)
  
  UpdatePromptEdit(listBox, nameEdit, promptEdit)
  
  ; 如果删除的是当前选中的，重置选中
  if (g_SelectedPrompt != "" && !HasPromptName(g_SelectedPrompt))
    g_SelectedPrompt := g_PromptNames[1]
  
  SavePrompts()
  RefreshPromptDropdown()
}

CloseManageGui(manageGui)
{
  global g_MainGui, g_OrigEditCtrl
  
  ; 保存当前原文
  currentText := ""
  if (g_OrigEditCtrl != "")
    try currentText := g_OrigEditCtrl.Value
  
  manageGui.Destroy()
  ; 重新加载配置
  LoadPrompts()
  
  ; 销毁并重建主界面（最可靠的方法）
  if (g_MainGui != "") {
    try g_MainGui.Destroy()
    g_MainGui := ""
    ShowMainGui(currentText)
  }
}

HasPromptName(name)
{
  global g_PromptNames
  for n in g_PromptNames {
    if (n = name)
      return true
  }
  return false
}

RefreshPromptDropdown()
{
  global g_PromptDropdown, g_PromptNames, g_SelectedPrompt, g_MainGui
  
  if (g_PromptDropdown = "" || g_MainGui = "")
    return
  
  try {
    ; 检查控件是否有效
    if (!IsObject(g_PromptDropdown) || !g_PromptDropdown.Hwnd)
      return
    
    ; 使用控件原生方法清空并添加
    g_PromptDropdown.Delete()
    g_PromptDropdown.Add(g_PromptNames)
    
    ; 设置选中项
    if (g_SelectedPrompt != "" && HasPromptName(g_SelectedPrompt)) {
      g_PromptDropdown.Choose(g_SelectedPrompt)
    } else if (g_PromptNames.Length > 0) {
      g_SelectedPrompt := g_PromptNames[1]
      g_PromptDropdown.Choose(1)
    }
  }
}

Gui_SendQuestion(*)
{
  global g_QuestionEditCtrl, g_AnswerEditCtrl, g_SendBtnCtrl
  global g_ChatPending, g_StreamFileChat, g_StreamPidChat, g_StreamContentChat
  
  question := Trim(g_QuestionEditCtrl.Value)
  if (question = "")
    return
  
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
  global g_StreamFileChat, g_StreamPidChat, g_StreamContentChat, g_ChatPending
  global g_SelectedPrompt
  
  ; 终止之前正在运行的 Chat 请求
  if (g_StreamPidChat > 0) {
    try ProcessClose(g_StreamPidChat)
    g_StreamPidChat := 0
  }
  
  ; 获取选中的 prompt 模板
  selectedPromptText := GetPromptByName(g_SelectedPrompt)
  
  ; 替换模板中的变量
  if (selectedPromptText != "") {
    ; 获取原文框内容
    global g_OrigEditCtrl
    originalText := ""
    if (g_OrigEditCtrl != "")
      try originalText := g_OrigEditCtrl.Value
    
    ; 替换 {原文} 变量
    selectedPromptText := StrReplace(selectedPromptText, "{原文}", originalText)
    
    fullQuestion := selectedPromptText . "`n`n" . question
  } else {
    fullQuestion := question
  }
  
  ; 转义 prompt 用于 JSON
  prompt := fullQuestion
  prompt := StrReplace(prompt, "\", "\\")
  prompt := StrReplace(prompt, "`"", "\`"")
  prompt := StrReplace(prompt, "`n", "\n")
  prompt := StrReplace(prompt, "`r", "\r")
  prompt := StrReplace(prompt, "`t", "\t")
  
  ; 系统提示：强制禁用 Markdown 和符号
  sysPrompt := "纯文本输出，不要用任何符号（如反斜杠、星号、井号）包裹或强调单词。"
  
  ; 设置临时文件
  g_StreamFileChat := A_Temp . "\ollama_stream_chat.txt"
  g_StreamContentChat := ""
  jsonFile := A_Temp . "\ollama_request_chat.json"
  
  ; 删除旧文件
  try FileDelete(g_StreamFileChat)
  try FileDelete(jsonFile)
  
  ; 构建 JSON (使用流式，添加 system 参数)
  json := '{"model":"huihui_ai/qwen3-abliterated:8b-v2","system":"' . sysPrompt . '","prompt":"' . prompt . '","stream":true,"options":{"temperature":0.7,"num_predict":2048,"think":true}}'
  
  ; 将 JSON 写入临时文件
  try {
    FileAppend(json, jsonFile, "UTF-8-RAW")
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
        ; 转换换行符
        result := StrReplace(result, "\n", "`n")
        try g_AnswerEditCtrl.Value := result
      }
      g_ChatPending := false
      if (g_SendBtnCtrl != "")
        try g_SendBtnCtrl.Enabled := true
      SetTimer(CheckChatResult, 0)
    }
  }
}

Gui_Hide(guiObj, *)
{
  global g_MainGui, g_GuiHidden, g_OldClip
  
  ; 隐藏窗口到后台
  if (g_MainGui != "") {
    g_MainGui.Hide()
    g_GuiHidden := true
  }
  
  ; 不再恢复剪贴板，避免覆盖用户的截图等内容
  ; A_Clipboard := g_OldClip
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
  g_PromptDropdown := ""
  ; 不再恢复剪贴板，避免覆盖用户的截图等内容
  ; A_Clipboard := g_OldClip
}

!SC029::
{
  global g_MainGui, g_GuiHidden, g_OldClip, g_OrigEditCtrl, g_OriginalText, g_SelectedResult
  
  ; 窗口已显示 → 隐藏到后台
  if (g_MainGui != "" && !g_GuiHidden) {
    g_MainGui.Hide()
    g_GuiHidden := true
    return
  }
  
  ; 窗口已隐藏 → 复制文本 + 恢复显示
  if (g_MainGui != "" && g_GuiHidden) {
    ; 复制选中文本（不自动全选）
    g_OldClip := ClipboardAll()
    A_Clipboard := ""
    Send("^c")
    ClipWait(0.3)
    text := Trim(A_Clipboard)
    
    ; 更新原文并重新请求
    if (text != "") {
      ; 检测语言是否改变
      newIsChinese := RegExMatch(text, "[\x{4e00}-\x{9fff}]")
      if (newIsChinese != g_IsChineseMode) {
        ; 语言模式改变，需要重新创建窗口
        try g_MainGui.Destroy()
        g_MainGui := ""
        ShowMainGui(text)
        return
      }
      
      ; 语言模式未变，只更新内容
      global g_CorrectRequested, g_TranslateRequested, g_TranslateEditCtrl, g_CorrectEditCtrl
      global g_ExplainEditCtrl, g_AnswerEditCtrl
      g_CorrectRequested := false
      g_TranslateRequested := false
      g_OrigEditCtrl.Value := text
      g_OriginalText := text
      
      ; 清空旧结果（包括 AI 回答框）
      if (g_ExplainEditCtrl != "")
        try g_ExplainEditCtrl.Value := ""
      if (g_AnswerEditCtrl != "")
        try g_AnswerEditCtrl.Value := ""
      g_TranslateEditCtrl.Value := "正在处理..."
      g_CorrectEditCtrl.Value := "正在处理..."
      StartAsyncRequests(text, "default")
    }
    
    g_MainGui.Show()
    WinActivate("ahk_id " g_MainGui.Hwnd)
    g_GuiHidden := false
    return
  }
  
  ; 窗口不存在 → 复制文本 + 创建窗口（不自动全选）
  g_OldClip := ClipboardAll()
  A_Clipboard := ""

  ; 只复制选中的文字，不自动全选
  Send("^c")
  ClipWait(0.3)
  text := Trim(A_Clipboard)

  ; 即使文本为空也显示窗口（可使用 AI 助手）
  ShowMainGui(text)
}

; Ctrl+Alt+Enter: 自动全选 + 翻译
^!Enter::
^!NumpadEnter::
{
  global g_OldClip
  g_OldClip := ClipboardAll()
  A_Clipboard := ""

  ; 自动全选并复制
  Send("^a")
  Sleep(50)
  Send("^c")
  ClipWait(0.5)
  text := Trim(A_Clipboard)

  ; 即使文本为空也显示窗口
  ShowMainGui(text)
}
