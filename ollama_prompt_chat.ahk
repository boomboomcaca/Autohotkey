;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; Prompt 模板管理 + AI 问答 + GUI 事件处理
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

; ===== Prompt 模板初始化/加载/保存 =====

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
    } else if (currentName != "" && currentPrompt != "" && !RegExMatch(line, "^\[")) {
      ; 多行 prompt：非 section 头的后续行追加到当前 prompt
      currentPrompt .= "`n" . line
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

HasPromptName(name)
{
  global g_PromptNames
  for n in g_PromptNames {
    if (n = name)
      return true
  }
  return false
}

; ===== Prompt 模板 GUI 管理 =====

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
  manageGui.Destroy()
  ; 重新加载配置
  LoadPrompts()
  ; 刷新主界面的下拉列表
  RefreshPromptDropdown()
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

; ===== AI 问答 =====

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
  global g_SelectedPrompt, g_HttpChat
  global g_GroqApiKey, g_GroqModel, g_GroqEndpoint
  
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
  
  g_StreamContentChat := ""
  
  ; 构建 JSON (使用流式，OpenAI 格式)
  json := '{"model":"' . g_GroqModel . '","messages":[{"role":"system","content":"' . sysPrompt . '"},{"role":"user","content":"' . prompt . '"}],"temperature":0.7,"max_tokens":2048,"stream":true}'
  
  try {
    ; 使用 Msxml2.XMLHTTP 启动异步流式请求
    http := ComObject("Msxml2.XMLHTTP")
    http.Open("POST", g_GroqEndpoint, true)
    http.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
    http.SetRequestHeader("Authorization", "Bearer " . g_GroqApiKey)
    http.Send(json)
    g_HttpChat := http ; 记录对象用于轮询
    g_ChatPending := true
  } catch Error as e {
    g_AnswerEditCtrl.Value := "请求启动失败: " . e.Message
    g_SendBtnCtrl.Enabled := true
    return
  }
  
  ; 启动轮询定时器
  SetTimer(CheckChatResult, 100)
}

CheckChatResult()
{
  global g_ChatPending, g_HttpChat, g_AnswerEditCtrl, g_SendBtnCtrl, g_StreamContentChat
  
  if (!g_ChatPending || !IsObject(g_HttpChat))
    return
  
  ; 检查控件是否已被销毁
  if (g_AnswerEditCtrl = "" || g_SendBtnCtrl = "") {
    SetTimer(CheckChatResult, 0)
    return
  }
  
  ; 检查状态 (3=Receiving, 4=Complete)
  if (g_HttpChat.readyState >= 3) {
    try {
      result := ParseStreamData(g_HttpChat.responseText, &g_StreamContentChat)
      if (result != "") {
        ; 实时更新回答框内容
        g_AnswerEditCtrl.Value := result
      }
    }
    
    ; 如果完成，清理
    if (g_HttpChat.readyState == 4) {
      g_ChatPending := false
      g_SendBtnCtrl.Enabled := true
      g_HttpChat := ""
      SetTimer(CheckChatResult, 0)
    }
  }
}

; ===== GUI 事件处理 =====

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
  
  ; 如果翻译/纠错控件不存在（如在查词窗口中），直接返回
  if (!IsObject(g_TranslateEditCtrl) || !IsObject(g_CorrectEditCtrl))
    return
  
  ; 在翻译、纠错、AI回答三个结果框之间循环切换焦点
  if (focusedHwnd = g_TranslateEditCtrl.Hwnd) {
    g_CorrectEditCtrl.Focus()
    g_SelectedResult := "correct"
    g_CorrectLabelCtrl.Text := "✓ " . correctLabel
    g_TranslateLabelCtrl.Text := "   " . translateLabel
  } else if (focusedHwnd = g_CorrectEditCtrl.Hwnd) {
    if (IsObject(g_AnswerEditCtrl))
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
    if (focusedHwnd = g_OrigEditCtrl.Hwnd || focusedHwnd = g_QuestionEditCtrl.Hwnd || (IsSet(g_WL_WordEdit) && g_WL_WordEdit != "" && focusedHwnd = g_WL_WordEdit.Hwnd)) {
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
  
  ; 在单词输入框、原文输入框和 AI 问题输入框之间切换
  if (IsSet(g_WL_WordEdit) && g_WL_WordEdit != "" && focusedHwnd = g_WL_WordEdit.Hwnd) {
    g_OrigEditCtrl.Focus()
  } else if (focusedHwnd = g_OrigEditCtrl.Hwnd) {
    g_QuestionEditCtrl.Focus()
  } else if (IsSet(g_WL_WordEdit) && g_WL_WordEdit != "") {
    g_WL_WordEdit.Focus()
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

ParseStreamData(rawContent, &accumulatedContent)
{
  if (rawContent = "")
    return accumulatedContent
  
  result := ""
  Loop Parse, rawContent, "`n", "`r"
  {
    line := Trim(A_LoopField)
    if (line = "")
      continue
    
    ; OpenAI SSE 格式: 每行以 "data: " 开头
    if (SubStr(line, 1, 6) = "data: ") {
      line := SubStr(line, 7)
    }
    
    ; 跳过 [DONE] 标记
    if (line = "[DONE]")
      continue
    
    if (!InStr(line, "{"))
      continue
    
    ; 增加对错误的检测
    if RegExMatch(line, '"error"\s*:\s*\{[^}]*"message"\s*:\s*"((?:[^"\\]|\\.)*)"', &m) {
      errorMsg := m[1]
      errorMsg := StrReplace(errorMsg, "\n", "`n")
      errorMsg := StrReplace(errorMsg, '\"', '"')
      return "错误: " . errorMsg
    }
    
    ; 也检测简单的 error 字符串格式
    if RegExMatch(line, '"error"\s*:\s*"((?:[^"\\]|\\.)*)"', &m) {
      errorMsg := m[1]
      errorMsg := StrReplace(errorMsg, "\n", "`n")
      errorMsg := StrReplace(errorMsg, '\"', '"')
      return "错误: " . errorMsg
    }
    
    ; OpenAI 格式: 提取 choices[0].delta.content 或 choices[0].message.content
    if RegExMatch(line, '"content"\s*:\s*"((?:[^"\\]|\\.)*)"', &m) {
      token := m[1]
      ; 基础转义还原
      token := StrReplace(token, "\n", "`n")
      token := StrReplace(token, "\r", "`r")
      token := StrReplace(token, "\t", "`t")
      token := StrReplace(token, '\"', '"')
      token := StrReplace(token, "\\", "\")
      result .= token
    }
  }
  
  if (result != "")
    accumulatedContent := StripEmoji(result)
  
  return accumulatedContent
}

; ===== 过滤 Emoji 和不可渲染的 Unicode 字符 =====
StripEmoji(text)
{
  ; 移除零宽字符、变体选择符、对象替换字符、装饰符号
  text := RegExReplace(text, "[\x{200B}-\x{200F}\x{200D}\x{2060}-\x{206F}\x{FEFF}\x{FFFC}\x{FFFD}\x{FE00}-\x{FE0F}\x{2600}-\x{27BF}\x{2B50}-\x{2B55}]", "")
  ; 移除补充平面字符（Emoji 等）：过滤 UTF-16 代理对
  result := ""
  Loop Parse, text {
    cp := Ord(A_LoopField)
    if (cp >= 0xD800 && cp <= 0xDFFF)
      continue
    result .= A_LoopField
  }
  return result
}

; 保留旧函数名作为兼容性代理，但逻辑改为 ParseStreamData
IsStreamComplete(filePath) => FileExist(filePath) && InStr(FileRead(filePath), '"done"')
ReadStreamFile(filePath, &accumulatedContent) => ParseStreamData(FileRead(filePath), &accumulatedContent)

