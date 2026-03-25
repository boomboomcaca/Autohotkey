; ===== OCR 字符混淆修正 =====
; 等宽/代码字体下 OCR 容易把 l↔1、o↔0 混淆
; 策略：根据相邻字母的大小写语境决定修正目标
FixOcrConfusion(word)
{
    ; 纯数字不修正
    if RegExMatch(word, "^\d+$")
        return word

    result := ""
    len := StrLen(word)
    Loop len {
        ch := SubStr(word, A_Index, 1)
        prev := A_Index > 1 ? SubStr(word, A_Index - 1, 1) : ""
        next := A_Index < len ? SubStr(word, A_Index + 1, 1) : ""
        prevIsLetter := RegExMatch(prev, "[a-zA-Z]")
        nextIsLetter := RegExMatch(next, "[a-zA-Z]")

        if (ch = "1" && prevIsLetter && nextIsLetter) {
            ; 1 夹在字母间 → l 或 I
            ; 前后都是大写 → I (如 F1LE → FILE)，否则 → l (如 fi1e → file)
            prevIsUpper := RegExMatch(prev, "[A-Z]")
            nextIsUpper := RegExMatch(next, "[A-Z]")
            result .= (prevIsUpper && nextIsUpper) ? "I" : "l"
        } else if (ch = "0" && prevIsLetter && nextIsLetter) {
            ; 0 夹在字母间 → o 或 O
            ; 前后都是大写 → O (如 B0OK → BOOK)，否则 → o (如 b0ok → book)
            prevIsUpper := RegExMatch(prev, "[A-Z]")
            nextIsUpper := RegExMatch(next, "[A-Z]")
            result .= (prevIsUpper && nextIsUpper) ? "O" : "o"
        } else {
            result .= ch
        }
    }

    return result
}

; ===== 切换图钉状态 =====
WL_TogglePin()
{
    global g_WL_IsPinned, g_WL_PinBtn
    g_WL_IsPinned := !g_WL_IsPinned
    if (g_WL_PinBtn) {
        g_WL_PinBtn.Text := g_WL_IsPinned ? "📍" : "📌"
        g_WL_PinBtn.SetFont("s9 " . (g_WL_IsPinned ? "cCC0000 Bold" : "c333333 Norm"), "Microsoft YaHei")
    }
}

; ===== 支持拖动窗口 =====
WL_WM_LBUTTONDOWN(wParam, lParam, msg, hwnd)
{
    global g_WL_Gui
    if (g_WL_Gui != "" && hwnd == g_WL_Gui.Hwnd) {
        PostMessage(0xA1, 2, 0, , "ahk_id " . hwnd)
    }
}

; ===== 拦截上下文菜单（阻止默认的系统文本框右键菜单） =====
WL_WM_CONTEXTMENU(wParam, lParam, msg, hwnd)
{
    global g_WL_Gui
    if (g_WL_Gui != "") {
        isOurGui := false
        if (hwnd == g_WL_Gui.Hwnd || wParam == g_WL_Gui.Hwnd) {
            isOurGui := true
        } else {
            try {
                ctrl := GuiCtrlFromHwnd(hwnd)
                if (ctrl && ctrl.Gui && ctrl.Gui.Hwnd == g_WL_Gui.Hwnd)
                    isOurGui := true
                ctrlW := GuiCtrlFromHwnd(wParam)
                if (ctrlW && ctrlW.Gui && ctrlW.Gui.Hwnd == g_WL_Gui.Hwnd)
                    isOurGui := true
            }
        }
        if (!isOurGui && WinActive("ahk_id " . g_WL_Gui.Hwnd)) {
            isOurGui := true
        }

        if (isOurGui) {
            global g_WL_InitMouseX, g_WL_InitMouseY, g_WL_MouseMoved, g_WL_ShowTick
            WL_PlayTtsOnce()
            CoordMode("Mouse", "Screen")
            MouseGetPos(&g_WL_InitMouseX, &g_WL_InitMouseY)
            g_WL_MouseMoved := false
            g_WL_ShowTick := A_TickCount
            return 0
        }
    }
}

; ===== 拦截底层 Edit 控件的右键按下和抬起，防止其私自呼出菜单 =====
WL_WM_RBUTTON(wParam, lParam, msg, hwnd)
{
    global g_WL_Gui
    if (g_WL_Gui != "") {
        try {
            ctrl := GuiCtrlFromHwnd(hwnd)
            ; 若目标是我们查词窗口内的控件
            if (ctrl && ctrl.Gui && ctrl.Gui.Hwnd == g_WL_Gui.Hwnd) {
                if (msg == 0x0205) { ; WM_RBUTTONUP
                    global g_WL_InitMouseX, g_WL_InitMouseY, g_WL_MouseMoved, g_WL_ShowTick
                    WL_PlayTtsOnce()
                    CoordMode("Mouse", "Screen")
                    MouseGetPos(&g_WL_InitMouseX, &g_WL_InitMouseY)
                    g_WL_MouseMoved := false
                    g_WL_ShowTick := A_TickCount
                }
                return 0 ; 让控件彻底忽略右键事件
            }
        }
    }
}

; ===== 关闭浮窗 =====
CloseWordGui()
{
  global g_WL_Gui, g_WL_ResultCtrl, g_WL_TitleCtrl
  global g_WL_StreamPid, g_WL_Pending, g_WL_StreamFile
  global g_WL_WordEdit, g_WL_ContextEdit
  global g_StreamPidChat, g_ChatPending, g_QuestionEditCtrl, g_AnswerEditCtrl, g_SendBtnCtrl
  global g_TtsPlaying, g_HoverTarget, g_MainGui

  ; 终止请求
  if (g_WL_StreamPid > 0) {
    try ProcessClose(g_WL_StreamPid)
    g_WL_StreamPid := 0
  }
  if (g_StreamPidChat > 0) {
    try ProcessClose(g_StreamPidChat)
    g_StreamPidChat := 0
  }
  g_WL_Pending := false
  g_ChatPending := false
  SetTimer(CheckWordResult, 0)
  SetTimer(CheckChatResult, 0)
  SetTimer(WL_CheckClickOutside, 0)
  SetTimer(CheckTtsHover, 0)
  try SoundPlay("NonExistent.zzz")
  g_TtsPlaying := false
  g_HoverTarget := ""

  ; 彻底清理临时文件
  try FileDelete(g_WL_StreamFile)
  try FileDelete(A_Temp . "\ahk_wl_request_word.json")

  ; 销毁 GUI
  if (g_WL_Gui != "") {
    try {
      HotIfWinActive("ahk_id " g_WL_Gui.Hwnd)
      Hotkey("Escape", WL_HandleEsc, "Off")
      Hotkey("Enter", WL_HandleEnter, "Off")
      Hotkey("NumpadEnter", WL_HandleEnter, "Off")
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
    g_QuestionEditCtrl := ""
    g_AnswerEditCtrl := ""
    g_SendBtnCtrl := ""
    g_PromptDropdown := ""
    g_MainGui := ""
    g_WL_AnkiBtn := ""
  }
}

; ===== 兼容性辅助函数 (供 ollama_prompt_chat.ahk 使用) =====

; IsStreamComplete 已移至共享库


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
  ; 性能优化: 直接删除上一个文件，避免 glob 遍历 TEMP 目录
  static prevTtsFile := ""
  if (prevTtsFile != "" && prevTtsFile != g_WL_TtsFile)
    try FileDelete(prevTtsFile)
  g_WL_TtsFile := A_Temp . "\ahk_wl_tts_" . A_TickCount . ".mp3"
  prevTtsFile := g_WL_TtsFile

  escapedText := StrReplace(text, '"', '\"')
  try {
    Run('edge-tts --voice en-US-AriaNeural --text "' . escapedText . '" --write-media "' . g_WL_TtsFile . '"', , "Hide", &outPid)
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
  CoordMode("Mouse", "Screen")
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
  
  global g_QuestionEditCtrl
  if (g_QuestionEditCtrl != "") {
    g_QuestionEditCtrl.Value := item.word
    g_QuestionEditCtrl.Focus()
    SendMessage(0x00B1, -1, -1, g_QuestionEditCtrl.Hwnd)
  }

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
      g_WL_ResultCtrl.Value := (g_WL_LangMode = "EN" ? "⏳ Querying..." : "⏳ 正在查询...")
    StartWordOllamaRequest(item.word, item.context, true)
  }

  WL_PregenTts(item.word)

  ; 导航操作视同活跃操作，重置自动关闭的计时器
  global g_WL_InitMouseX, g_WL_InitMouseY, g_WL_MouseMoved, g_WL_ShowTick
  CoordMode("Mouse", "Screen")
  MouseGetPos(&g_WL_InitMouseX, &g_WL_InitMouseY)
  g_WL_MouseMoved := false
  g_WL_ShowTick := A_TickCount
}

; ===== 发送至 Anki 的核心通信模块 =====
WL_SendToAnki(*)
{
    global g_WL_WordEdit, g_WL_ContextEdit, g_WL_ResultCtrl
    global g_WL_TtsFile, g_WL_AnkiBtn

    if (!g_WL_WordEdit || !g_WL_ResultCtrl || !g_WL_AnkiBtn)
        return

    word := Trim(g_WL_WordEdit.Value)
    context := Trim(g_WL_ContextEdit.Value)
    explanation := Trim(g_WL_ResultCtrl.Value)

    if (word = "")
        return

    isAdd := InStr(g_WL_AnkiBtn.Text, "➕") || InStr(g_WL_AnkiBtn.Text, "添加")

    ; 从配置文件动态读取 Anki 卡片类型映射关系
    deckName := "英语生词"
    try deckName := IniRead(A_ScriptDir . "\ollama_config.ini", "Anki", "DeckName")
    
    modelName := "问答题"
    try modelName := IniRead(A_ScriptDir . "\ollama_config.ini", "Anki", "ModelName")
    
    frontField := "正面"
    try frontField := IniRead(A_ScriptDir . "\ollama_config.ini", "Anki", "FrontField")
    
    backField := "背面"
    try backField := IniRead(A_ScriptDir . "\ollama_config.ini", "Anki", "BackField")
    
    ; 兜底防乱码
    if (InStr(modelName, "闁") || InStr(modelName, "瓟") || InStr(modelName, "ue1be") || InStr(modelName, "u95c2")) {
        deckName := "英语生词"
        modelName := "问答题"
        frontField := "正面"
        backField := "背面"
    }

    try {
        http := ComObject("WinHttp.WinHttpRequest.5.1")
        
        if (isAdd) {
            if (explanation = "" || InStr(explanation, "Querying") || InStr(explanation, "正在查询")) {
                return
            }

            ; 格式化卡片文本
            frontText := "<h2>" . word . "</h2>"
            if (context != "" && context != word)
                frontText .= "<br><br><span style='color:grey;'>" . StrReplace(context, "`n", "<br>") . "</span>"
            
            backText := StrReplace(explanation, "`n", "<br>")

            ; 文本转 JSON 安全字符串闭包
            EscapeJSON := (str) => StrReplace(StrReplace(StrReplace(StrReplace(str, "\", "\\"), "`n", "\n"), "`r", ""), "`"", "\`"")

            frontJson := EscapeJSON(frontText)
            backJson := EscapeJSON(backText)

            ; 处理音频
            audioJson := ""
            if (g_WL_TtsFile != "" && FileExist(g_WL_TtsFile)) {
                absPath := StrReplace(g_WL_TtsFile, "\", "\\")
                audioJson := ',"audio": [{"path": "' . absPath . '", "filename": "ahk_tts_' . word . '.mp3", "fields": ["' . frontField . '"]}]'
            }

            ; 构建 JSON 报文
            payload := '{"action": "addNote", "version": 6, "params": {"note": {"deckName": "' . deckName . '", "modelName": "' . modelName . '", "fields": {"' . frontField . '": "' . frontJson . '", "' . backField . '": "' . backJson . '"}, "options": {"allowDuplicate": false}, "tags": ["AHK抓取"]' . audioJson . '}}}'

            try {
                deckPayload := '{"action": "createDeck", "version": 6, "params": {"deck": "' . deckName . '"}}'
                http.Open("POST", "http://127.0.0.1:8765", false)
                http.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
                http.Send(deckPayload)
                http.WaitForResponse()
            }

            http.Open("POST", "http://127.0.0.1:8765", false)
            http.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
            http.Send(payload)
            http.WaitForResponse()
            res := http.ResponseText
            
            if (InStr(res, '"error": null')) {
                g_WL_AnkiBtn.Text := "➖ Anki"
                g_WL_AnkiBtn.SetFont("c008800 Norm")
            } else if (InStr(res, "cannot create note because it is a duplicate")) {
                g_WL_AnkiBtn.Text := "➖ Anki"
                g_WL_AnkiBtn.SetFont("c008800 Norm")
            } else {
            }
        } else {
            ; 删除逻辑（仅搜索正面字段，避免释义误匹配）
            escapeWord := StrReplace(StrReplace(word, "\", "\\"), "`"", "\`"")
            query := 'deck:"' . deckName . '" ' . frontField . ':re:<h2>' . escapeWord . '</h2>'
            jsonQuery := StrReplace(query, '"', '\"')
            payload := '{"action": "findNotes", "version": 6, "params": {"query": "' . jsonQuery . '"}}'
            
            http.Open("POST", "http://127.0.0.1:8765", false)
            http.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
            http.Send(payload)
            http.WaitForResponse()
            res := http.ResponseText
            
            if (RegExMatch(res, '"result":\s*\[(.*?)\]', &m)) {
                ids := Trim(m[1])
                if (ids != "") {
                    delPayload := '{"action": "deleteNotes", "version": 6, "params": {"notes": [' . ids . ']}}'
                    http.Open("POST", "http://127.0.0.1:8765", false)
                    http.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
                    http.Send(delPayload)
                    http.WaitForResponse()
                    
                    g_WL_AnkiBtn.Text := "➕ Anki"
                    g_WL_AnkiBtn.SetFont("c333333 Norm")
                    return
                }
            }
            g_WL_AnkiBtn.Text := "➕ Anki"
            g_WL_AnkiBtn.SetFont("c333333 Norm")
        }
        http := "" ; 使用完毕，释放 COM 对象
    } catch Error as e {
        http := ""
    }
}

WL_CheckAnkiStatus(word) {
    global g_WL_AnkiBtn, g_WL_Gui
    if (!g_WL_AnkiBtn || !g_WL_Gui) {
        return
    }
        
    deckName := "英语生词"
    try deckName := IniRead(A_ScriptDir . "\ollama_config.ini", "Anki", "DeckName")
    if (InStr(deckName, "闁") || InStr(deckName, "ue1be")) {
        deckName := "英语生词"
    }

    frontField := "正面"
    try frontField := IniRead(A_ScriptDir . "\ollama_config.ini", "Anki", "FrontField")

    escapeWord := StrReplace(StrReplace(word, "\", "\\"), "`"", "\`"")
    query := 'deck:"' . deckName . '" ' . frontField . ':re:<h2>' . escapeWord . '</h2>'
    jsonQuery := StrReplace(query, '"', '\"')
    payload := '{"action": "findNotes", "version": 6, "params": {"query": "' . jsonQuery . '"}}'
    
    try {
        http := ComObject("WinHttp.WinHttpRequest.5.1")
        http.Open("POST", "http://127.0.0.1:8765", true) ; 改用异步模式（true），不阻塞主线程
        http.SetTimeouts(1000, 1000, 1000, 1000)
        http.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
        http.Send(payload)
        
        ; 使用闭包和定时器来轮询异步请求状态，不卡界面
        CheckAnkiState() {
            try {
                if (http.ReadyState != 4)
                    return ; 还没请求完，下次继续查
                
                SetTimer(, 0) ; 请求结束，关掉当前定时器
                
                global g_WL_AnkiBtn, g_WL_Gui
                if (!g_WL_AnkiBtn || !g_WL_Gui) {
                    http := ""
                    return
                }
                    
                res := http.ResponseText
                http := "" ; 释放 COM 对象
                
                if (RegExMatch(res, '"result":\s*\[(.*?)\]', &m)) {
                    ids := Trim(m[1])
                    if (ids != "") {
                        try g_WL_AnkiBtn.Text := "➖ Anki"
                        try g_WL_AnkiBtn.SetFont("c008800 Norm")
                        return
                    }
                }
                try g_WL_AnkiBtn.Text := "➕ Anki"
                try g_WL_AnkiBtn.SetFont("c333333 Norm")
            } catch {
                try SetTimer(, 0)
                global g_WL_AnkiBtn
                try g_WL_AnkiBtn.Text := "➕ Anki"
                try g_WL_AnkiBtn.SetFont("c333333 Norm")
                http := "" ; 发生异常时也释放
            }
        }
        
        SetTimer(CheckAnkiState, 50)
    } catch {
        ; 忽略连接失败或控件已销毁
        return
    }
    try g_WL_AnkiBtn.Text := "➕ Anki"
    try g_WL_AnkiBtn.SetFont("c333333 Norm")
}

