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
g_HttpChat := ""
g_CorrectPending := false
g_TranslatePending := false
g_CorrectRequested := false
g_TranslateRequested := false
g_CurrentText := ""
g_TtsPlaying := false
g_HoverTarget := ""
g_PendingShowGui := false
g_PrevForegroundHwnd := 0
g_GuiHidden := false

; 流式响应缓存
g_StreamContentCorrect := ""
g_StreamContentTranslate := ""
g_StreamContentChat := ""

; AI 问答相关
g_QuestionEditCtrl := ""
g_AnswerEditCtrl := ""
g_SendBtnCtrl := ""
g_ChatPending := false

; Prompt 模板相关
g_ConfigFile := A_ScriptDir . "\ollama_config.ini"
g_PromptList := []
g_PromptNames := []
g_SelectedPrompt := ""
g_PromptDropdown := ""
g_PromptManageBtn := ""

; 引入拆分模块
#Include "ollama_tts.ahk"
#Include "ollama_prompt_chat.ahk"

; 初始化 Prompt 模板
InitPrompts()

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
  global g_TtsOrigCtrl, g_TtsCorrectCtrl, g_TtsTranslateCtrl, g_TtsQuestionCtrl
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
  
  g_MainGui.AddText("xs w40 Section", "问题:")
  g_TtsQuestionCtrl := g_MainGui.AddText("x+5 ys cGray", "🔊")
  g_TtsQuestionCtrl.OnEvent("Click", Gui_PlayQuestion)
  g_QuestionEditCtrl := g_MainGui.AddEdit("xs w330 h50", original)
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
  
  ; 记录当前前台窗口，用于朗读时恢复
  global g_PrevForegroundHwnd
  try {
    fgHwnd := WinGetID("A")
    if (fgHwnd != g_MainGui.Hwnd)
      g_PrevForegroundHwnd := fgHwnd
  }
  
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
  if (IsObject(g_HttpCorrect)) {
    try g_HttpCorrect.Abort()
    g_HttpCorrect := ""
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
  ; 系统提示：强制禁用 Markdown 和符号
  sysPrompt := "纯文本输出，不要用任何符号（如反斜杠、星号、井号）包裹或强调单词。"
  
  ; 转义 prompt 用于 JSON
  prompt := StrReplace(prompt, "\", "\\")
  prompt := StrReplace(prompt, "`"", "\`"")
  prompt := StrReplace(prompt, "`n", "\n")
  prompt := StrReplace(prompt, "`r", "\r")
  prompt := StrReplace(prompt, "`t", "\t")
  
  ; 构建 JSON (使用流式)
  json := '{"model":"huihui_ai/qwen3-abliterated:8b-v2","system":"' . sysPrompt . '","prompt":"' . prompt . '","stream":true,"options":{"temperature":0,"num_predict":1024,"think":true}}'
  
  try {
    ; 使用 Msxml2.XMLHTTP 支持在接收过程中读取数据 (readyState=3)
    http := ComObject("Msxml2.XMLHTTP")
    http.Open("POST", "http://localhost:11434/api/generate", true)
    http.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
    http.Send(json)
    return http
  } catch Error as e {
    TrayTip("请求启动失败", e.Message, "Icon!")
    return 0
  }
}

CheckAsyncResults()
{
  global g_CorrectPending, g_TranslatePending, g_HttpCorrect
  global g_IsChineseMode, g_CorrectRequested, g_TranslateRequested, g_CurrentText
  global g_CorrectEditCtrl, g_TranslateEditCtrl
  global g_StreamFileCorrect, g_StreamFileTranslate
  global g_StreamContentCorrect, g_StreamContentTranslate
  global g_StreamPidCorrect, g_StreamPidTranslate
  
  ; 检查组合结果（一次调用同时返回纠错和翻译）
  if (g_CorrectPending && IsObject(g_HttpCorrect)) {
    ; 检查是否已经开始返回或已完成 (3=Receiving, 4=Complete)
    if (g_HttpCorrect.readyState >= 3) {
      try {
        result := ParseStreamData(g_HttpCorrect.responseText, &g_StreamContentCorrect)
        
        ; 实时更新 GUI（可选，如果需要实时效果）
        if (result != "") {
          ; 简单的预解析，或者等完成后再一次性解析
          ; 这里我们先尝试局部解析来获得更好的反馈感
          UpdateCorrectResult("正在生成输出...") 
        }
      } catch {
        ; 0x8000000A 报错说明数据暂时不可用，跳过本次轮询等待下一次
      }
      
      ; 检查是否完全结束
      if (g_HttpCorrect.readyState == 4) {
        try {
          finalRes := ParseStreamData(g_HttpCorrect.responseText, &g_StreamContentCorrect)
          if (finalRes != "") {
            ParseCombinedResult(finalRes)
          }
        }
        g_CorrectPending := false
        g_TranslatePending := false
        g_HttpCorrect := ""
      }
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

; IsStreamComplete 和 ReadStreamFile 已移至 ollama_prompt_chat.ahk

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
  global g_PrevForegroundHwnd
  
  ; 记录当前前台窗口，用于朗读时恢复焦点
  try {
    fgHwnd := WinGetID("A")
    g_PrevForegroundHwnd := fgHwnd
  }
  
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
      if (g_QuestionEditCtrl != "")
        try g_QuestionEditCtrl.Value := text
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
