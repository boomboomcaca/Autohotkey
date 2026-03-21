;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; 鼠标取词 + 语境解释 - Alt+W：截取鼠标所在窗口 → Windows OCR → 定位单词 → Ollama 解释
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

#Include "UIA.ahk"
#Include "OCR.ahk"
; 共享模块已在入口处引入（emacs.ahk）

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
g_WL_QuestionLabel := ""
g_WL_AnswerLabel := ""
g_WL_PromptLabel := ""
g_WL_BottomHint := ""
g_WL_AnkiBtn := "" ; 新增 Anki 按钮全局变量
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

; AI 问答相关 (与 ollama_translate.ahk 保持一致，以便复用逻辑)
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

; 其他依赖变量 (供 ollama_tts.ahk 和 ollama_prompt_chat.ahk 使用)
g_MainGui := ""
g_OrigEditCtrl := ""
g_IsChineseMode := false
g_TtsOrigCtrl := ""
g_TtsCorrectCtrl := ""
g_TtsTranslateCtrl := ""
g_TtsQuestionCtrl := ""
g_TtsPlaying := false
g_HoverTarget := ""
g_PrevForegroundHwnd := 0

; 初始化 Prompt 模板
InitPrompts()

; ===== 快捷键 F2 =====
F2::
{
  global g_WL_Gui, g_WL_ResultCtrl, g_WL_TitleCtrl
  global g_WL_StreamFile, g_WL_StreamPid, g_WL_Pending, g_WL_StreamContent

  CoordMode("Mouse", "Screen")
  MouseGetPos(&mouseX, &mouseY, &winUnder, &ctlUnder)

  word := ""
  line := ""
  found := false

  ; ==========================================================
  ; 【优先级 1】: UIA (UI Automation) - 内存直读，0延迟，100% 准确
  ; ==========================================================
  try {
    el := UIA.ElementFromPoint(mouseX, mouseY)
    if (el) {
      if (el.IsTextPatternAvailable) {
        textPattern := el.TextPattern
        range := textPattern.RangeFromPoint(mouseX, mouseY)
        if (range) {
          lineRange := range.Clone()
          range.ExpandToEnclosingUnit(UIA.TextUnit.Word)
          rawWord := Trim(range.GetText())
          
          lineRange.ExpandToEnclosingUnit(UIA.TextUnit.Paragraph)
          rawLine := RegExReplace(Trim(lineRange.GetText()), "s)[\r\n]+", " ") ; 基础行

          ; --- 优化：尝试获取 UIA 上下文（增加上下各一行） ---
          try {
              p_prev := lineRange.Clone()
              p_prev.Move(UIA.TextUnit.Paragraph, -1)
              txt_prev := Trim(p_prev.GetText())
              if (txt_prev != "" && txt_prev != rawLine)
                  rawLine := txt_prev . " " . rawLine
                  
              p_next := lineRange.Clone()
              p_next.Move(UIA.TextUnit.Paragraph, 1)
              txt_next := Trim(p_next.GetText())
              if (txt_next != "" && txt_next != rawLine && !InStr(rawLine, txt_next))
                  rawLine := rawLine . " " . txt_next
          }
          ; ----------------------------------------------

          if (rawWord != "" && !RegExMatch(rawWord, "\s")) {
            cleanedWord := RegExReplace(rawWord, "^[^\w\x{4e00}-\x{9fa5}\-]+|[^\w\x{4e00}-\x{9fa5}\-]+$", "")
            if (cleanedWord != "") {
              word := cleanedWord
              line := rawLine
              found := true
            }
          }
        }
      }
      if (!found && el.Name != "") {
        rawName := Trim(el.Name)
        ; [Fix] 过滤浏览器自带的图片辅助说明或通用占位文本 (如 "要获取缺失的图片说明...", "Logo", "Icon" 等)
        ; 只要 UIA 的 TextPattern 没抓到真实文本，且 Name 是这类通用标签，就跳过并回退到 OCR
        genericLabels := "Logo|Icon|Image|Picture|Graphic|Illustration|Avatar|Banner|SVG|Brand"
        if (InStr(rawName, "获取缺失的图片说明") || InStr(rawName, "missing image descriptions") || RegExMatch(rawName, "i)^(" . genericLabels . ")$")) {
           ; 跳过，这通常是浏览器或网页产生的无效 UIA 标签，触发优先级 2 的 OCR 识别画面真实内容
        } else if (rawName != "" && !RegExMatch(rawName, "\s") && StrLen(rawName) < 50) {
          cleanedWord := RegExReplace(rawName, "^[^\w\x{4e00}-\x{9fa5}\-]+|[^\w\x{4e00}-\x{9fa5}\-]+$", "")
          if (cleanedWord != "") {
            word := cleanedWord
            line := rawName
            found := true
          }
        }
      }
    }
  } catch {
  }

  ; ==========================================================
  ; 【优先级 2】: Windows 10/11 原生 WinRT OCR API - 屏幕极速截取
  ; ==========================================================
  if (!found) {
    ToolTip("Analyzing Screen...")
    try {
      dpi := 96
      try dpi := DllCall("GetDpiForWindow", "Ptr", winUnder, "UInt")
      if (dpi < 96)
        dpi := 96
      scale := 1 ; WinRT OCR 不需要强制放大，原生支持得很好
      captureW := Round(3840 * scale)
      captureH := Round(2160 * scale)
      winX := mouseX - Round(captureW / 2)
      winY := mouseY - Round(captureH / 2)
      
      ; 调用 WinRT OCR 库局部截屏取词 (优先尝试英语，若系统没装英文包则自动回退到系统可用语言并提示)
      try {
        ocrResult := OCR.FromRect(winX, winY, captureW, captureH, {Language: "en-US"})
      } catch {
        ToolTip("Warning: OCR Language Package Missing. Please install English (United States) in Windows Settings.")
        SetTimer(ToolTip, -5000)
        ocrResult := OCR.FromRect(winX, winY, captureW, captureH)
      }
      
      if (ocrResult) {
        bestDist := 999999
        bestWord := ""
        bestLine := ""
        
        for lineIndex, ocrLine in ocrResult.Lines {
          for index, ocrWord in ocrLine.Words {
            cx := ocrWord.x + ocrWord.w / 2
            cy := ocrWord.y + ocrWord.h / 2
            dist := Sqrt((mouseX - cx)**2 + (mouseY - cy)**2)
            
            if (mouseX >= ocrWord.x && mouseX <= ocrWord.x + ocrWord.w && mouseY >= ocrWord.y && mouseY <= ocrWord.y + ocrWord.h) {
              bestWord := ocrWord.Text
              bestDist := 0
              bestIndex := index
              bestLineIndex := lineIndex
              bestLineObj := ocrLine
              break
            }
            
            if (dist < bestDist) {
              bestDist := dist
              bestWord := ocrWord.Text
              bestIndex := index
              bestLineIndex := lineIndex
              bestLineObj := ocrLine
            }
          }
          if (bestDist == 0)
            break
        }
        
        ; 收集上下文行 (向上最多取3行，向下最多取3行，增加范围)
        bestLine := ""
        if (IsSet(bestLineIndex)) {
          startLineIdx := Max(1, bestLineIndex - 3)
          endLineIdx := Min(ocrResult.Lines.Length, bestLineIndex + 3)
          for i, lObj in ocrResult.Lines {
            if (i >= startLineIdx && i <= endLineIdx) {
              bestLine .= (bestLine=""?"":" ") . lObj.Text
            }
          }
        }
        
        ; 尝试合并紧邻的单词结块 (譬如 OpenClaw 被 OCR 分拆为了 Open 和 Claw)
        if (bestWord != "" && IsSet(bestLineObj)) {
            ; 往前合并
            tempIndex := bestIndex - 1
            while (tempIndex > 0) {
                prevWord := bestLineObj.Words[tempIndex]
                currWord := bestLineObj.Words[tempIndex + 1]
                ; 判断间距阈值调大 (由 4 像素增加到 15 像素)，支持大字体识别
                if (currWord.x - (prevWord.x + prevWord.w) <= 4) {
                    bestWord := prevWord.Text . bestWord
                    tempIndex--
                } else {
                    break
                }
            }
            ; 往后合并
            tempIndex := bestIndex + 1
            while (tempIndex <= bestLineObj.Words.Length) {
                nextWord := bestLineObj.Words[tempIndex]
                currWord := bestLineObj.Words[tempIndex - 1]
                ; 判断间距阈值调大
                if (nextWord.x - (currWord.x + currWord.w) <= 4) {
                    bestWord := bestWord . nextWord.Text
                    tempIndex++
                } else {
                    break
                }
            }
        }
        
        if (bestWord != "") {
          bestWord := Trim(bestWord)
          cleanedWord := RegExReplace(bestWord, "^[^\w\x{4e00}-\x{9fa5}\-]+|[^\w\x{4e00}-\x{9fa5}\-]+$", "")
          if (cleanedWord != "") {
            word := cleanedWord
            line := bestLine
            found := true
          }
        }
      }
    } catch as err {
      ToolTip("OCR 失败: " err.Message)
      SetTimer(ToolTip, -2000)
      return
    }
    ToolTip()
  }

  if (!found || word = "") {
    ToolTip("未识别到单词")
    SetTimer(ToolTip, -2000)
    return
  }

  ShowWordPopup(word, line, mouseX, mouseY)
}

; ===== 显示取词浮窗 =====
ShowWordPopup(word, context, posX, posY)
{
  global g_WL_Gui, g_WL_ResultCtrl, g_WL_TitleCtrl, g_WL_WordEdit, g_WL_ContextEdit, WL_CurrentWord, WL_CurrentContext, g_WL_LangMode, g_WL_LangBtn, g_WL_AnkiBtn
  global g_IsChineseMode, g_QuestionEditCtrl, g_AnswerEditCtrl, g_SendBtnCtrl, g_PromptDropdown
  global g_MainGui, g_OrigEditCtrl, g_PromptNames, g_SelectedPrompt, g_PromptManageBtn, g_TtsQuestionCtrl
  global g_WL_QuestionLabel, g_WL_AnswerLabel, g_WL_PromptLabel, g_WL_BottomHint
  word := StripEmoji(word)
  context := StripEmoji(context)
  WL_CurrentWord := word
  WL_CurrentContext := context
  g_IsChineseMode := RegExMatch(word, "[\x{4e00}-\x{9fff}]")

  if (g_WL_Gui != "") {
    ; 如果窗口已存在，直接更新内容，不重新创建
    g_WL_WordEdit.Value := word
    g_WL_ContextEdit.Value := (context != "" && context != word) ? context : ""
    g_QuestionEditCtrl.Value := word
    g_WL_ResultCtrl.Value := (g_WL_LangMode = "EN" ? "Querying..." : "正在查询...")

    ; 重置悬停自动关闭的检测状态，防止刚更新完就消失
    global g_WL_InitMouseX, g_WL_InitMouseY, g_WL_MouseMoved, g_WL_ShowTick
    CoordMode("Mouse", "Screen")
    MouseGetPos(&g_WL_InitMouseX, &g_WL_InitMouseY)
    g_WL_MouseMoved := false
    g_WL_ShowTick := A_TickCount

    ; 激活窗口并聚焦问题框
    try {
      WinActivate("ahk_id " . g_WL_Gui.Hwnd)
      g_QuestionEditCtrl.Focus()
      SendMessage(0x00B1, -1, -1, g_QuestionEditCtrl.Hwnd)
    }

    ; 预生成 TTS 音频（后台，会自动终止上一个）
    WL_PregenTts(word)
    ; 发起 Ollama 请求（会自动终止上一个）
    StartWordOllamaRequest(word, context)
    return
  }

  g_WL_Gui := Gui("+AlwaysOnTop -Caption +Border +Owner")
  g_MainGui := g_WL_Gui  ; 兼容 ollama_prompt_chat.ahk
  g_WL_Gui.BackColor := "FFFFFF"
  g_WL_Gui.MarginX := 12
  g_WL_Gui.MarginY := 8

  ; 标题行水平排列
  g_WL_Gui.SetFont("s14 c1a1a2e Bold", "Microsoft YaHei")
  
  ; 单词可编辑输入框 + 中英切换按钮
  g_WL_WordEdit := g_WL_Gui.AddEdit("w225 Section -E0x200", word)
  
  g_WL_Gui.SetFont("s9 c333333 Norm", "Microsoft YaHei")
  g_WL_LangBtn := g_WL_Gui.AddButton("x+5 ys w40 h26", g_WL_LangMode = "EN" ? "EN" : "中")
  g_WL_AnkiBtn := g_WL_Gui.AddButton("x+5 ys w60 h26", "➕ Anki")
  
  ; 关联事件
  if (g_WL_LangBtn) {
    g_WL_LangBtn.OnEvent("Click", (*) => WL_ToggleLang())
  }
  if (g_WL_AnkiBtn) {
    g_WL_AnkiBtn.OnEvent("Click", (*) => WL_SendToAnki())
  }


  ; 语境行（可编辑）
  g_WL_Gui.SetFont("s10 c444444 Norm", "Microsoft YaHei")
  g_WL_ContextEdit := g_WL_Gui.AddEdit("xs w320 h50 Multi -E0x200", (context != "" && context != word) ? context : "")
  g_OrigEditCtrl := g_WL_ContextEdit ; 兼容 ollama_prompt_chat.ahk

  ; 分隔线
  g_WL_Gui.SetFont("s1 cCCCCCC", "Microsoft YaHei")
  g_WL_Gui.AddText("xs w320 0x10")  ; SS_ETCHEDHORZ

  ; 结果区域（可选中复制）
  g_WL_Gui.SetFont("s10 c333333 Norm", "Microsoft YaHei")
  g_WL_ResultCtrl := g_WL_Gui.AddEdit("xs w320 h265 ReadOnly -E0x200", g_WL_LangMode = "EN" ? "Querying..." : "正在查询...")

  ; ==========================================================
  ; AI 问答区域 (右侧面板)
  ; ==========================================================
  g_WL_Gui.SetFont("s9 c666666", "Microsoft YaHei")
  g_WL_PromptLabel := g_WL_Gui.AddText("x350 y12 w50 Section", g_WL_LangMode = "EN" ? "Prompt:" : "提示词:")
  
  promptList := ""
  for name in g_PromptNames {
    promptList .= (promptList = "" ? "" : "|") . name
  }
  g_PromptDropdown := g_WL_Gui.AddDropDownList("x+2 yp-3 w200", StrSplit(promptList, "|"))
  if (g_SelectedPrompt != "")
    g_PromptDropdown.Text := g_SelectedPrompt
  g_PromptDropdown.OnEvent("Change", Gui_PromptChanged)
  
  g_PromptManageBtn := g_WL_Gui.AddButton("x+5 yp w58 h24", g_WL_LangMode = "EN" ? "Manage" : "管理")
  g_PromptManageBtn.OnEvent("Click", Gui_ManagePrompts)

  g_WL_Gui.SetFont("s9 c333333", "Microsoft YaHei")
  g_WL_QuestionLabel := g_WL_Gui.AddText("xs Section", g_WL_LangMode = "EN" ? "Question:" : "问题:")
  g_WL_Gui.SetFont("s16") ; 放大图标
  g_TtsQuestionCtrl := g_WL_Gui.AddText("x+5 ys-6 cGray", "🔊")
  g_WL_Gui.SetFont("s9")  ; 恢复字体
  g_TtsQuestionCtrl.OnEvent("Click", Gui_PlayQuestion)
  
  g_QuestionEditCtrl := g_WL_Gui.AddEdit("xs w255 h50 -E0x200", word)
  g_SendBtnCtrl := g_WL_Gui.AddButton("x+5 yp w60 h50", g_WL_LangMode = "EN" ? "Send" : "发送")
  g_SendBtnCtrl.OnEvent("Click", Gui_SendQuestion)

  g_WL_AnswerLabel := g_WL_Gui.AddText("xs", g_WL_LangMode = "EN" ? "Answer:" : "回答:")
  g_AnswerEditCtrl := g_WL_Gui.AddEdit("xs w320 h197 ReadOnly -E0x200", "")

  ; 底部提示
  g_WL_Gui.SetFont("s8 cAAAAAA", "Microsoft YaHei")
  g_WL_BottomHint := g_WL_Gui.AddText("xm w650", g_WL_LangMode = "EN" ? "Enter to Re-query | Mouse out to Close | Esc to Close | Tab to Switch Focus" : "Enter 重新查询 | 鼠标移出关闭 | Esc 关闭 | Tab 切换焦点")

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

  ; 移动到正确位置并聚焦
  g_WL_Gui.Show("x" . showX . " y" . showY)
  g_QuestionEditCtrl.Focus()
  SendMessage(0x00B1, -1, -1, g_QuestionEditCtrl.Hwnd)

  ; 绑定 Esc/Enter、历史记录导航、以及 Emacs 文本操作 (参照主窗口处理)
  HotIfWinActive("ahk_id " g_WL_Gui.Hwnd)
  Hotkey("Escape", WL_HandleEsc, "On")
  Hotkey("Enter", WL_HandleEnter, "On")
  Hotkey("NumpadEnter", WL_HandleEnter, "On")
  Hotkey("!Left", (*) => WL_NavHistory(-1), "On")
  Hotkey("!Right", (*) => WL_NavHistory(1), "On")
  ; 文本操作增强
  Hotkey("Tab", Gui_ToggleFocus, "On")
  Hotkey("^Tab", Gui_ToggleSelect, "On")
  Hotkey("^v", Gui_PasteAsText, "On")
  Hotkey("^Backspace", Gui_DeleteWord, "On")
  Hotkey("^s", (*) => WL_SendToAnki(), "On") ; 新增 Ctrl+S 快捷键发送至 Anki
  HotIfWinActive()

  ; 启动鼠标移出关闭的检测定时器
  global g_WL_InitMouseX, g_WL_InitMouseY, g_WL_MouseMoved, g_WL_ShowTick
  CoordMode("Mouse", "Screen")
  MouseGetPos(&g_WL_InitMouseX, &g_WL_InitMouseY)
  g_WL_MouseMoved := false
  g_WL_ShowTick := A_TickCount
  SetTimer(WL_CheckClickOutside, 200)
  SetTimer(CheckTtsHover, 200)

  ; 预生成 TTS 音频（后台，不阻塞）
  WL_PregenTts(word)

  ; 发起 Ollama 请求
  StartWordOllamaRequest(word, context)
}

; ===== Enter 重新查询 =====
WL_HandleEnter(*)
{
  global g_WL_WordEdit, g_WL_ContextEdit, g_WL_ResultCtrl, WL_CurrentWord, WL_CurrentContext
  global g_QuestionEditCtrl, g_AnswerEditCtrl

  ; 检测输入法是否处于组合状态（正在输入中文）
  if (IsImeComposing()) {
    ; 让回车键正常传递给输入法确认候选词
    Send("{Enter}")
    return
  }

  if (g_WL_WordEdit = "")
    return

  focusedHwnd := ControlGetFocus("A")
  if (g_QuestionEditCtrl != "" && focusedHwnd = g_QuestionEditCtrl.Hwnd) {
    Gui_SendQuestion()
    return
  }

  newWord := Trim(g_WL_WordEdit.Value)
  if (newWord = "")
    return

  newContext := (g_WL_ContextEdit != "") ? Trim(g_WL_ContextEdit.Value) : ""
  WL_CurrentWord := newWord
  WL_CurrentContext := newContext
  if (g_WL_ResultCtrl != "")
    g_WL_ResultCtrl.Value := (g_WL_LangMode = "EN" ? "Querying..." : "正在查询...")
  if (g_AnswerEditCtrl != "")
    g_AnswerEditCtrl.Value := ""
  
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
  CoordMode("Mouse", "Screen")
  MouseGetPos(&cx, &cy, &winAtMouse)
  if (!g_WL_MouseMoved) {
    if (Abs(cx - g_WL_InitMouseX) > 5 || Abs(cy - g_WL_InitMouseY) > 5)
      g_WL_MouseMoved := true
    else
      return
  }

  ; 如果鼠标还在查词主窗口内，直接返回
  if (winAtMouse == g_WL_Gui.Hwnd)
    return

  ; 检查鼠标所在窗口的特征（白名单机制）
  try {
    if (winAtMouse = 0)
      return ; 瞬时获取失败，暂时忽略

    curClass := WinGetClass("ahk_id " . winAtMouse)
    curPid := WinGetPID("ahk_id " . winAtMouse)
    curProc := ProcessGetName(curPid)
    ourPid := ProcessExist()

    ; 判定是否豁免（不关闭）：
    ; 1. 窗口属于当前脚本进程 (包含 DropDownList 的 ComboLBox 弹出层、管理子窗口等)
    ; 2. 属于 Windows 系统通用组件 (菜单、阴影等)
    ; 3. 属于正在工作中的输入法 (IME/TextInputHost)
    if (curPid == ourPid 
        || curClass == "ComboLBox" || curClass == "#32768" || curClass == "SysShadow" || InStr(curClass, "Combo")
        || InStr(curClass, "IME") || InStr(curClass, "Cand") || InStr(curProc, "IME") || curProc == "TextInputHost.exe" || curClass == "ApplicationFrameWindow") {
      return
    }
    
    ; 确认处于外部窗口且非豁免窗口，执行关闭
    CloseWordGui()
  } catch {
    ; 报错（往往是因为 winAtMouse 句柄正好无效了）则暂时忽略，不执行关闭
    return
  }
}

; ===== 切换语言 =====
WL_ToggleLang()
{
    global g_WL_LangMode, g_WL_LangBtn, WL_CurrentWord, WL_CurrentContext, g_WL_ResultCtrl
    global g_WL_QuestionLabel, g_WL_AnswerLabel, g_WL_PromptLabel, g_WL_BottomHint, g_SendBtnCtrl, g_PromptManageBtn

    if (g_WL_LangMode = "EN")
    {
        g_WL_LangMode := "ZH"
        g_WL_LangBtn.Text := "中"
        if (g_WL_QuestionLabel) 
            g_WL_QuestionLabel.Value := "问题:"
        if (g_WL_AnswerLabel) 
            g_WL_AnswerLabel.Value := "回答:"
        if (g_WL_PromptLabel) 
            g_WL_PromptLabel.Value := "提示词:"
        if (g_WL_BottomHint) 
            g_WL_BottomHint.Value := "Enter 重新查询 | 鼠标移出关闭 | Esc 关闭 | Tab 切换焦点"
        if (g_SendBtnCtrl) 
            g_SendBtnCtrl.Text := "发送"
        if (g_PromptManageBtn) 
            g_PromptManageBtn.Text := "管理"
        if (g_WL_ResultCtrl) 
            g_WL_ResultCtrl.Value := "正在切换语言并重新查询..."
    }
    else
    {
        g_WL_LangMode := "EN"
        g_WL_LangBtn.Text := "EN"
        if (g_WL_QuestionLabel) 
            g_WL_QuestionLabel.Value := "Question:"
        if (g_WL_AnswerLabel) 
            g_WL_AnswerLabel.Value := "Answer:"
        if (g_WL_PromptLabel) 
            g_WL_PromptLabel.Value := "Prompt:"
        if (g_WL_BottomHint) 
            g_WL_BottomHint.Value := "Enter to Re-query | Mouse out to Close | Esc to Close | Tab to Switch Focus"
        if (g_SendBtnCtrl) 
            g_SendBtnCtrl.Text := "Send"
        if (g_PromptManageBtn) 
            g_PromptManageBtn.Text := "Manage"
        if (g_WL_ResultCtrl) 
            g_WL_ResultCtrl.Value := "Switching language and re-querying..."
    }
    try IniWrite(g_WL_LangMode, A_ScriptDir . "\ollama_config.ini", "Settings", "WordLookupLang")
    
    ; 重新发起请求
    StartWordOllamaRequest(WL_CurrentWord, WL_CurrentContext)
}

; ===== 发起 Ollama 语境解释请求 =====
StartWordOllamaRequest(word, context, isNavigating := false)
{
  global g_WL_StreamFile, g_WL_StreamPid, g_WL_Pending, g_WL_StreamContent, g_WL_LangMode
  global g_WL_History, g_WL_HistoryIdx

  WL_CheckAnkiStatus(word)

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

    prompt .= "\n\nPlease output using the following format (plain text only):\n● Part of Speech: xxx /phonetics/\n● Word Roots: [Break down prefixes, roots, suffixes with origins and meanings]\n● Definition: [Simple English definition]\n● Context Meaning: [Explanation based on the given context]\n● Collocations: [Common collocations or examples]"
    
    sysPrompt := "Output ONLY in English. Use plain text without Markdown formatting. Keep explanations concise."
  } else {
    ; 构建 prompt (英汉释义模式)
    prompt := "你是一个英语词典。解释单词 '" . word . "'"
    if (context != "" && context != word)
      prompt .= " 在以下语境中的含义。\n语境：" . context
    else
      prompt .= " 的含义。"

    prompt .= "\n\n请用以下格式输出（纯文本）：\n● 词性：xxx /音标/\n● 词根拆解：拆分前缀、词根、后缀，标注来源和含义\n● 释义：xxx\n● 语境释义：在这个句子中表示...\n● 常见搭配：xxx"
    
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

  ; JSON (OpenAI 格式)
  global g_GroqApiKey, g_GroqModel, g_GroqEndpoint
  json := '{"model":"' . g_GroqModel . '","messages":[{"role":"system","content":"' . sysPrompt . '"},{"role":"user","content":"' . prompt . '"}],"temperature":0,"max_tokens":800,"stream":true}'

  try {
    FileAppend(json, jsonFile, "UTF-8-RAW")
  } catch {
    return
  }

  ; 使用 curl.exe 调用 Groq API
  try {
    curlCmd := 'curl.exe -s -N -X POST "' . g_GroqEndpoint . '" -H "Content-Type: application/json" -H "Authorization: Bearer ' . g_GroqApiKey . '" -d "@' . jsonFile . '" -o "' . g_WL_StreamFile . '"'
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

  ; 解析 OpenAI SSE 流式 JSON
  result := ""
  Loop Parse, content, "`n", "`r"
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
    
    if RegExMatch(line, '"content"\s*:\s*"((?:[^"\\]|\\.)*)"', &m) {
      token := m[1]
      token := StrReplace(token, "\n", "`n")
      token := StrReplace(token, "\r", "`r")
      token := StrReplace(token, "\t", "`t")
      token := StrReplace(token, '\`"', '`"')
      token := StrReplace(token, "\\", "\")
      result .= token
    }
  }

  result := Trim(result)

  return StripEmoji(result)
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
  g_WL_TtsFile := A_Temp . "\ahk_wl_tts_" . A_TickCount . ".mp3"
  
  ; 顺便清理一下历史遗留的这种临时文件，防止堆积
  try {
    Loop Files, A_Temp . "\ahk_wl_tts_*.mp3"
      if (A_LoopFileFullPath != g_WL_TtsFile)
        FileDelete(A_LoopFileFullPath)
  }

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
                ToolTip("⚠️ 单词或释义为空/未完成，无法添加到 Anki")
                SetTimer(ToolTip, -2000)
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
                ToolTip("⭐ 已成功添加到 Anki！")
                g_WL_AnkiBtn.Text := "➖ Anki"
                SetTimer(ToolTip, -2000)
            } else if (InStr(res, "cannot create note because it is a duplicate")) {
                ToolTip("💡 Anki 中已存在该单词，无需重复添加")
                g_WL_AnkiBtn.Text := "➖ Anki"
                SetTimer(ToolTip, -2000)
            } else {
                ToolTip("❌ Anki 数据格式错误（字段名不匹配）!`n" . res)
                SetTimer(ToolTip, -3000)
            }
        } else {
            ; 删除逻辑
            escapeWord := StrReplace(StrReplace(word, "\", "\\"), "`"", "\`"")
            query := 'deck:"' . deckName . '" "' . escapeWord . '"'
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
                    
                    ToolTip("🗑️ 已从 Anki 移除该单词")
                    g_WL_AnkiBtn.Text := "➕ Anki"
                    SetTimer(ToolTip, -2000)
                    return
                }
            }
            ToolTip("⚠️ 未找到对应的 Anki 笔记")
            g_WL_AnkiBtn.Text := "➕ Anki"
            SetTimer(ToolTip, -2000)
        }
    } catch Error as e {
        ToolTip("❌ 无法连接到 Anki`n请确保 Anki 客户端已启动且安装了 AnkiConnect 插件")
        SetTimer(ToolTip, -3000)
    }
}

WL_CheckAnkiStatus(word) {
    global g_WL_AnkiBtn
    if (!g_WL_AnkiBtn) {
        return
    }
        
    deckName := "英语生词"
    try deckName := IniRead(A_ScriptDir . "\ollama_config.ini", "Anki", "DeckName")
    if (InStr(deckName, "闁") || InStr(deckName, "ue1be")) {
        deckName := "英语生词"
    }

    escapeWord := StrReplace(StrReplace(word, "\", "\\"), "`"", "\`"")
    query := 'deck:"' . deckName . '" "' . escapeWord . '"'
    jsonQuery := StrReplace(query, '"', '\"')
    payload := '{"action": "findNotes", "version": 6, "params": {"query": "' . jsonQuery . '"}}'
    
    try {
        http := ComObject("WinHttp.WinHttpRequest.5.1")
        http.Open("POST", "http://127.0.0.1:8765", false)
        http.SetTimeouts(1000, 1000, 1000, 1000)
        http.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
        http.Send(payload)
        http.WaitForResponse()
        res := http.ResponseText
        
        if (RegExMatch(res, '"result":\s*\[(.*?)\]', &m)) {
            ids := Trim(m[1])
            if (ids != "") {
                g_WL_AnkiBtn.Text := "➖ Anki"
                return
            }
        }
    } catch {
        ; 忽略连接失败
    }
    g_WL_AnkiBtn.Text := "➕ Anki"
}

