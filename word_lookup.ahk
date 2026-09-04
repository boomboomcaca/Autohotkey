;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; 鼠标取词 + 语境解释 - Alt+W：截取鼠标所在窗口 → Windows OCR → 定位单词 → Ollama 解释
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

#Include "lib/UIA.ahk"
#Include "lib/OCR.ahk"
; ollama_api.ahk 已在 emacs.ahk 中引入，不要重复引入！

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
g_WL_IsPinned := false
g_WL_PinBtn := ""
g_WL_StreamFile := ""
g_WL_StreamPid := 0
g_WL_Pending := false
g_WL_StreamContent := ""
g_WL_StreamFileSize := 0
g_WL_ShowTick := 0
g_WL_InitMouseX := 0
g_WL_InitMouseY := 0
g_WL_MouseMoved := false
g_WL_TtsFile := ""
g_WL_TtsPid := 0
g_WL_TtsWord := ""
g_WL_TtsGenTick := 0
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

; 注册窗口拖动事件
OnMessage(0x0201, WL_WM_LBUTTONDOWN)
; 注册上下文菜单事件拦截（防止Edit控件在触摸板或触屏等情况下弹出右键菜单）
OnMessage(0x007B, WL_WM_CONTEXTMENU)
; 拦截控件私自处理右键事件（彻底切断右键触发系统菜单的可能）
OnMessage(0x0204, WL_WM_RBUTTON) ; WM_RBUTTONDOWN
OnMessage(0x0205, WL_WM_RBUTTON) ; WM_RBUTTONUP

GetWordAndLineAtMouse(&word, &line)
{
  CoordMode("Mouse", "Screen")
  MouseGetPos(&mouseX, &mouseY, &winUnder, &ctlUnder)

  word := ""
  line := ""
  found := false
  fromOcr := false

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
              if (p_prev.Move(UIA.TextUnit.Paragraph, -1) != 0) {
                  p_prev.ExpandToEnclosingUnit(UIA.TextUnit.Paragraph)
                  txt_prev := Trim(p_prev.GetText())
                  if (txt_prev != "" && txt_prev != rawLine)
                      rawLine := txt_prev . " " . rawLine
              }
                  
              p_next := lineRange.Clone()
              if (p_next.Move(UIA.TextUnit.Paragraph, 1) != 0) {
                  p_next.ExpandToEnclosingUnit(UIA.TextUnit.Paragraph)
                  txt_next := Trim(p_next.GetText())
                  if (txt_next != "" && txt_next != rawLine && !InStr(rawLine, txt_next))
                      rawLine := rawLine . " " . txt_next
              }
          }
          ; ----------------------------------------------

          ; --- 如果段落上下文不足（仅含单词本身），从父元素回溯获取完整文本 ---
          if (rawLine = rawWord || StrLen(rawLine) <= StrLen(rawWord) + 5) {
              try {
                  ancestor := el
                  Loop 5 {
                      ancestor := ancestor.Parent
                      if (!ancestor)
                          break
                      ancestorText := ""
                      try ancestorText := Trim(ancestor.Name)
                      if (ancestorText != "" && StrLen(ancestorText) > StrLen(rawWord) + 5 && InStr(ancestorText, rawWord)) {
                          rawLine := RegExReplace(ancestorText, "s)[\r\n]+", " ")
                          break
                      }
                  }
              }
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
        
        ; [核心修复] 基于 UIA ControlType 判断元素是否为文本类型
        ; 非文本类型 (Image/Group/Pane/Custom/Button 等) 的 Name 通常是无障碍辅助标签，不是用户看到的真实文字
        ; 应该跳过，让程序回退到 OCR 识别实际可见内容
        elType := 0
        try elType := el.Type
        
        ; 定义"可信文本类型"白名单 (只有这些类型的 Name 才可能是用户看到的真实文字)
        ; Text=50020, Edit=50004, Hyperlink/Link=50005, Document=50030, ListItem=50007, TreeItem=50024, MenuItem=50011, TabItem=50019
        isTextElement := (elType == 50020 || elType == 50004 || elType == 50005 || elType == 50030 || elType == 50007 || elType == 50024 || elType == 50011 || elType == 50019)
        
        ; [安全网] 即使类型匹配，仍保留通用标签黑名单作为二级过滤
        genericLabels := "Logo|Icon|Image|Picture|Graphic|Illustration|Avatar|Banner|SVG|Brand"
        isGenericLabel := (InStr(rawName, "获取缺失的图片说明") || InStr(rawName, "missing image descriptions") || RegExMatch(rawName, "i)^(" . genericLabels . ")$"))
        
        if (isTextElement && !isGenericLabel && rawName != "" && !RegExMatch(rawName, "\s") && StrLen(rawName) < 50) {
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
    try {
      ; 按目标窗口 DPI 缩放截取尺寸，保证高缩放屏幕上覆盖等量的文本行数
      dpi := 96
      try dpi := DllCall("GetDpiForWindow", "Ptr", winUnder, "UInt")
      ; GetDpiForWindow 在部分窗口/多屏环境下会返回 0，此时若不兜底，
      ; captureW/captureH 会算成 0，截出 0×0 空图导致 OCR 永远识别不到
      if (!dpi || dpi < 96)
        dpi := 96
      ; 只截取鼠标周围一小块区域：足够容纳目标词 + 上下各 3 行上下文。
      ; WinRT OCR 同步执行会阻塞脚本主线程，区域越小识别越快；
      ; 同时也避免距离兜底吸附到屏幕远处无关的词
      captureW := Round(800 * dpi / 96)
      captureH := Round(400 * dpi / 96)
      winX := mouseX - Round(captureW / 2)
      winY := mouseY - Round(captureH / 2)

      ; OCR 识别语言可通过 ollama_config.ini 的 [Settings] OcrLanguage 配置（如 zh-CN）。
      ; 默认 en-US；若系统没装对应语言包则自动回退到系统可用语言
      ocrLang := "en-US"
      try ocrLang := IniRead(A_ScriptDir . "\ollama_config.ini", "Settings", "OcrLanguage", "en-US")
      try {
        ocrResult := OCR.FromRect(winX, winY, captureW, captureH, {Language: ocrLang})
      } catch {
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
        
        ; 距离兜底设上限：鼠标未直接命中词框时，只允许吸附到约一行高范围内的词，
        ; 鼠标指着空白处时宁可取不到，也不抓取截取区域内远处无关的词
        maxSnapDist := 40 * dpi / 96
        if (bestDist > maxSnapDist)
          bestWord := ""

        ; 收集上下文行 (向上最多取3行，向下最多取3行，增加范围)
        bestLine := ""
        if (bestWord != "" && IsSet(bestLineIndex)) {
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
            ; 间距阈值按字高比例计算（约 1/5 字高，至少 3 像素）：
            ; 固定像素阈值在大字体/高 DPI 下会漏合并，调太大又会把正常空格分隔的词错误拼接
            mergeGap := Max(3, Round(bestLineObj.Words[bestIndex].h * 0.2))
            ; 往前合并
            tempIndex := bestIndex - 1
            while (tempIndex > 0) {
                prevWord := bestLineObj.Words[tempIndex]
                currWord := bestLineObj.Words[tempIndex + 1]
                if (currWord.x - (prevWord.x + prevWord.w) <= mergeGap) {
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
                if (nextWord.x - (currWord.x + currWord.w) <= mergeGap) {
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
            fromOcr := true
          }
        }
      }
    } catch {
      ; OCR 失败（截屏被拒、引擎初始化失败等）静默放弃，调用方按"未取到词"处理
      return false
    }
  }

  if (found && word != "") {
    ; 字符混淆修正只针对 OCR 结果：UIA 是内存直读，文本本身就是准确的，
    ; 强行修正反而会改错真实含数字的词（如 a1b、b0t）
    if (fromOcr)
      word := FixOcrConfusion(word)
    return true
  }
  return false
}

; ===== 快捷键 F2 =====
; 已停止原 F2 鼠标取词浮窗功能，该键现已改由 Gemini 联动功能使用
; F2::
; {
;   global g_WL_Gui, g_WL_ResultCtrl, g_WL_TitleCtrl
;   global g_WL_StreamFile, g_WL_StreamPid, g_WL_Pending, g_WL_StreamContent
; 
;   word := ""
;   line := ""
;   if (GetWordAndLineAtMouse(&word, &line)) {
;     CoordMode("Mouse", "Screen")
;     MouseGetPos(&mouseX, &mouseY)
;     ShowWordPopup(word, line, mouseX, mouseY)
;   }
; }

; ===== 显示取词浮窗 =====
; 【注意】当前无任何调用入口（原 F2 入口已改为 Gemini 联动，见 emacs.ahk）。
; 整套浮窗功能（GUI/查询/历史/Anki）保留以备将来恢复。
; 若恢复入口，注意本浮窗与翻译主窗口共享 g_QuestionEditCtrl/g_AnswerEditCtrl/
; g_SendBtnCtrl/g_MainGui 等全局变量，两窗口同时存在时会互相覆盖控件引用。
ShowWordPopup(word, context, posX, posY)
{
  global g_WL_Gui, g_WL_ResultCtrl, g_WL_TitleCtrl, g_WL_WordEdit, g_WL_ContextEdit, WL_CurrentWord, WL_CurrentContext, g_WL_LangMode, g_WL_LangBtn, g_WL_AnkiBtn
  global g_IsChineseMode, g_QuestionEditCtrl, g_AnswerEditCtrl, g_SendBtnCtrl, g_PromptDropdown
  global g_MainGui, g_OrigEditCtrl, g_PromptNames, g_SelectedPrompt, g_PromptManageBtn, g_TtsQuestionCtrl
  global g_WL_QuestionLabel, g_WL_AnswerLabel, g_WL_PromptLabel, g_WL_BottomHint, g_WL_IsPinned, g_WL_PinBtn
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
  g_WL_Gui.BackColor := "F5F6F8"
  g_WL_Gui.MarginX := 16
  g_WL_Gui.MarginY := 12

  ; Windows 11 圆角 + 阴影
  try {
    DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", g_WL_Gui.Hwnd, "Int", 33, "Int*", 2, "Int", 4)  ; 圆角
    DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", g_WL_Gui.Hwnd, "Int", 2, "Int*", 1, "Int", 4)   ; 阴影
  }

  ; 标题行水平排列
  g_WL_Gui.SetFont("s14 c2D3142 Bold", "Microsoft YaHei")
  
  ; 单词可编辑输入框 + 中英切换按钮
  g_WL_WordEdit := g_WL_Gui.AddEdit("w225 Section -E0x200", word)
  
  g_WL_Gui.SetFont("s9 c333333 Norm", "Microsoft YaHei")
  g_WL_LangBtn := g_WL_Gui.AddText("x+5 ys w35 h26 Center 0x200 Border BackgroundE8EAF0", g_WL_LangMode = "EN" ? "EN" : "中")
  g_WL_Gui.SetFont("s9 c333333 Norm", "Microsoft YaHei")
  g_WL_AnkiBtn := g_WL_Gui.AddText("x+5 ys w50 h26 Center 0x200 Border BackgroundE8EAF0", "➕ Anki")
  g_WL_Gui.SetFont("s9 c333333 Norm", "Microsoft YaHei")
  
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
  g_WL_Gui.SetFont("s1 cE0E2E8", "Microsoft YaHei")
  g_WL_Gui.AddText("xs w320 0x10")  ; SS_ETCHEDHORZ

  ; 结果区域（可选中复制）
  g_WL_Gui.SetFont("s10 c333333 Norm", "Microsoft YaHei")
  g_WL_ResultCtrl := g_WL_Gui.AddEdit("xs w320 h265 ReadOnly -E0x200", g_WL_LangMode = "EN" ? "Querying..." : "正在查询...")

  ; ==========================================================
  ; AI 问答区域 (右侧面板)
  ; ==========================================================
  g_WL_Gui.SetFont("s9 c333333", "Microsoft YaHei")
  g_WL_PromptLabel := g_WL_Gui.AddText("x358 y16 w50 Section", g_WL_LangMode = "EN" ? "Prompt:" : "提示词:")
  
  ; 直接传数组：先用 | 拼接再 StrSplit 会把名称里含 | 的模板拆成多项
  g_PromptDropdown := g_WL_Gui.AddDropDownList("x+2 yp-1 w170", g_PromptNames)
  if (g_SelectedPrompt != "")
    g_PromptDropdown.Text := g_SelectedPrompt
  g_PromptDropdown.OnEvent("Change", Gui_PromptChanged)
  
  g_PromptManageBtn := g_WL_Gui.AddText("x+5 yp-2 w58 h24 Center 0x200 Border BackgroundE8EAF0", g_WL_LangMode = "EN" ? "Manage" : "管理")
  g_PromptManageBtn.OnEvent("Click", Gui_ManagePrompts)

  ; 图钉按钮（紧跟管理按钮，高度对齐）
  g_WL_Gui.SetFont("s9 " . (g_WL_IsPinned ? "cCC0000 Bold" : "c333333 Norm"), "Microsoft YaHei")
  g_WL_PinBtn := g_WL_Gui.AddText("x+5 yp w24 h24 Center 0x200 Border BackgroundE8EAF0", g_WL_IsPinned ? "📍" : "📌")
  g_WL_PinBtn.OnEvent("Click", (*) => WL_TogglePin())
  g_WL_Gui.SetFont("s9 c333333 Norm", "Microsoft YaHei")

  g_WL_Gui.SetFont("s9 c333333", "Microsoft YaHei")
  g_WL_QuestionLabel := g_WL_Gui.AddText("xs Section", g_WL_LangMode = "EN" ? "Question:" : "问题:")
  g_WL_Gui.SetFont("s16") ; 放大图标
  g_TtsQuestionCtrl := g_WL_Gui.AddText("x+5 ys-6 cGray", "🔊")
  g_WL_Gui.SetFont("s9")  ; 恢复字体
  g_TtsQuestionCtrl.OnEvent("Click", Gui_PlayQuestion)
  
  g_QuestionEditCtrl := g_WL_Gui.AddEdit("xs w255 h50 -E0x200", word)
  g_SendBtnCtrl := g_WL_Gui.AddText("x+5 yp w60 h50 Center 0x200 Border BackgroundE8EAF0", g_WL_LangMode = "EN" ? "Send" : "发送")
  g_SendBtnCtrl.OnEvent("Click", Gui_SendQuestion)

  g_WL_Gui.SetFont("s9 c333333", "Microsoft YaHei")
  g_WL_AnswerLabel := g_WL_Gui.AddText("xs", g_WL_LangMode = "EN" ? "Answer:" : "回答:")
  g_AnswerEditCtrl := g_WL_Gui.AddEdit("xs w320 h197 ReadOnly -E0x200", "")

  ; 底部提示
  g_WL_Gui.SetFont("s8 cA0A4B0", "Microsoft YaHei")
  g_WL_BottomHint := g_WL_Gui.AddText("xm w660", g_WL_LangMode = "EN" ? "Enter Re-query │ Mouse-out Close │ Esc Close │ Tab Switch" : "Enter 重新查询 │ 鼠标移出关闭 │ Esc 关闭 │ Tab 切换焦点")

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

  ; 全局注册右键拦截（不限窗口，查词期间全局生效）
  Hotkey("RButton", WL_RButtonHandler, "On")

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
  global g_QuestionEditCtrl, g_AnswerEditCtrl, g_WL_LangMode

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
  global g_WL_Gui, g_WL_IsPinned

  if (g_WL_Gui = "") {
    SetTimer(WL_CheckClickOutside, 0)
    return
  }

  if (g_WL_IsPinned)
    return

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

; ===== 请求启动失败时的统一复位 =====
; StartWordOllamaRequest 在启动 curl 之前就把 g_WL_Pending 置成 true，
; 而轮询定时器 CheckWordResult 要到函数末尾才启动。中途 return 却不复位，
; 浮窗会永远停在"正在查询..."，除了关窗没有任何恢复途径。
; 顺带删掉可能已落盘的 curl 配置文件——里面明文写着 API key。
WL_AbortPendingRequest(msg)
{
  global g_WL_Pending, g_WL_StreamPid, g_WL_ResultCtrl
  g_WL_Pending := false
  g_WL_StreamPid := 0
  SetTimer(CheckWordResult, 0)
  if (g_WL_ResultCtrl != "")
    try g_WL_ResultCtrl.Value := msg
  try FileDelete(A_Temp . "\ahk_wl_request_word.json")
  try FileDelete(A_Temp . "\ahk_wl_curl.cfg")
}

; ===== 发起 Ollama 语境解释请求 =====
StartWordOllamaRequest(word, context, isNavigating := false, isRetry := false)
{
  global g_WL_StreamFile, g_WL_StreamPid, g_WL_Pending, g_WL_StreamContent, g_WL_LangMode
  global g_WL_History, g_WL_HistoryIdx, g_WL_RetryCount, g_WL_StreamFileSize
  global g_WL_ResultCtrl  ; 启动失败时要往结果框写提示

  if (!isRetry)
    g_WL_RetryCount := 0

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
  g_WL_StreamFileSize := 0

  ; prompt 模板统一维护在 shared/ollama_api.ahk
  if (g_WL_LangMode = "EN") {
    prompt := GetWordLookupPromptEn(word, context)
    sysPrompt := "Output ONLY in English. Use plain text without Markdown formatting. Keep explanations concise."
  } else {
    prompt := GetWordLookupPromptZh(word, context)
    sysPrompt := "你必须全程使用中文进行解释说明（包括词根的含义也必须翻译为中文，不要夹杂英文解释）。纯文本输出，不要用任何符号（如反斜杠、星号、井号）包裹或强调单词。简洁回答。"
  }

  ; 转义 JSON
  prompt := EscapeJsonForApi(prompt)

  ; 设置文件
  g_WL_StreamFile := A_Temp . "\ahk_wl_stream_word.txt"
  jsonFile := A_Temp . "\ahk_wl_request_word.json"
  curlCfg := A_Temp . "\ahk_wl_curl.cfg"

  try FileDelete(g_WL_StreamFile)
  try FileDelete(jsonFile)
  try FileDelete(curlCfg)

  ; JSON (OpenAI 格式)
  global g_MistralApiKey, g_MistralModel, g_MistralEndpoint
  json := '{"model":"' . g_MistralModel . '","messages":[{"role":"system","content":"' . sysPrompt . '"},{"role":"user","content":"' . prompt . '"}],"temperature":0,"max_tokens":800,"stream":true}'

  try {
    FileAppend(json, jsonFile, "UTF-8-RAW")
    ; Authorization 头写入 curl 配置文件而非命令行，避免 API key 暴露在进程命令行中
    FileAppend('header = "Authorization: Bearer ' . g_MistralApiKey . '"`n', curlCfg, "UTF-8-RAW")
  } catch {
    ; 上面已把 g_WL_Pending 置 true，而轮询定时器要到函数末尾才启动。
    ; 直接 return 会让浮窗永远停在"正在查询..."，除了关窗没有任何恢复途径
    WL_AbortPendingRequest(g_WL_LangMode = "EN" ? "⚠ Failed to write temp file (disk full or locked)." : "⚠ 无法写入临时文件，请检查磁盘空间或杀毒软件拦截。")
    return
  }

  ; 使用 curl.exe 调用 Mistral API
  try {
    curlCmd := 'curl.exe -s -N --connect-timeout 10 -m 60 -X POST "' . g_MistralEndpoint . '" -H "Content-Type: application/json" -K "' . curlCfg . '" -d "@' . jsonFile . '" -o "' . g_WL_StreamFile . '"'
    Run(curlCmd, , "Hide", &outPid)
    g_WL_StreamPid := outPid
    global g_WL_StartTick, g_WL_LastDataTick
    g_WL_StartTick := A_TickCount
    g_WL_LastDataTick := A_TickCount
  } catch {
    ; 同上：curl.exe 不存在或被拦截时，不复位就会永远卡在"正在查询..."
    WL_AbortPendingRequest(g_WL_LangMode = "EN" ? "⚠ Failed to launch curl.exe." : "⚠ 无法启动 curl.exe，请确认系统已安装 curl。")
    return
  }

  ; 启动极速轮询(性能优化：增加间隔至 100ms)
  SetTimer(CheckWordResult, 100)
}

; ===== 轮询 Ollama 结果 =====
CheckWordResult()
{
  global g_WL_Pending, g_WL_StreamFile, g_WL_StreamContent, g_WL_StreamPid, g_WL_StreamFileSize
  global g_WL_ResultCtrl, g_WL_Gui, g_WL_StartTick, g_WL_LastDataTick, g_WL_LangMode

  if (!g_WL_Pending || g_WL_Gui = "") {
    SetTimer(CheckWordResult, 0)
    return
  }

  isComplete := false
  isTimeout := false

  ; 检查 curl 进程是否结束
  if (g_WL_StreamPid > 0 && !ProcessExist(g_WL_StreamPid)) {
    isComplete := true
  }

  ; 安全兜底：10 秒内没有收到任何新数据则判定超时（自动重试一次）。
  ; 按"空闲时间"而不是总时长判定：流式输出正常进行时总时长常超过 10 秒，
  ; 按总时长会把正在输出的回答中途截断、当作完整结果存进历史
  if (A_TickCount - g_WL_LastDataTick > 10000) {
    isComplete := true
    isTimeout := true
  }

  if (!isComplete && g_WL_StreamFile != "" && FileExist(g_WL_StreamFile)) {
    curSize := 0
    try curSize := FileGetSize(g_WL_StreamFile)
    
    ; 性能优化: 仅在文件被 curl 追加了新内容时，才触发磁盘读取和高昂的 JSON 字符串解析操作
    if (curSize != g_WL_StreamFileSize) {
      g_WL_StreamFileSize := curSize
      g_WL_LastDataTick := A_TickCount

      ; 实时读取流式内容并更新浮窗
      currentContent := WL_ReadStreamContent(g_WL_StreamFile)
      if (currentContent != "" && currentContent != g_WL_StreamContent) {
        g_WL_StreamContent := currentContent
        if (g_WL_ResultCtrl != "") {
          try g_WL_ResultCtrl.Value := currentContent
        }
      }
    }
  }

  if (isComplete) {
    if (isTimeout && g_WL_StreamPid > 0) {
      try ProcessClose(g_WL_StreamPid)
    }

    Sleep(30) ; 等待文件最终刷入硬盘
    finalResult := ""
    if (g_WL_StreamFile != "" && FileExist(g_WL_StreamFile)) {
      finalResult := WL_ReadStreamContent(g_WL_StreamFile)
    }

    if (finalResult != "") {
      finalResult := StripEmoji(finalResult) ; 仅最终结果时过滤一次 Emoji
      if (g_WL_ResultCtrl != "") {
        try g_WL_ResultCtrl.Value := finalResult
      }
      global g_WL_History, g_WL_HistoryIdx
      if (g_WL_HistoryIdx > 0 && g_WL_HistoryIdx <= g_WL_History.Length) {
        g_WL_History[g_WL_HistoryIdx].result := finalResult
      }
    } else {
      global g_WL_RetryCount, WL_CurrentWord, WL_CurrentContext
      ; 解析不出内容时，先看看是不是接口本身报错：curl 的 -s 不带 -f，HTTP 4xx/5xx
      ; 退出码同样是 0，错误 JSON 就躺在结果文件里。不做这层判断的话，
      ; 403（模型不在套餐内）/401/429 会被一律说成"请检查网络连接"
      apiErr := ExtractApiError(ReadTextFileUtf8(g_WL_StreamFile))

      ; AI 无响应或请求超时，如果在 10 秒超时并且是初次超时，执行自动重试一次。
      ; 接口已明确报错时不重试：403/401 这类错误重试多少次结果都一样，只会白等 10 秒
      if (apiErr = "" && isTimeout && g_WL_RetryCount < 1) {
        g_WL_RetryCount++
        if (g_WL_ResultCtrl != "") {
          msg := (g_WL_LangMode = "EN" ? "⏱ Timeout... Retrying (" . g_WL_RetryCount . "/1)..." : "⏱ 查询缓慢... 正在自动重试 (" . g_WL_RetryCount . "/1)...")
          try g_WL_ResultCtrl.Value := msg
        }
        
        g_WL_Pending := false
        g_WL_StreamPid := 0
        SetTimer(CheckWordResult, 0)
        try FileDelete(g_WL_StreamFile)
        try FileDelete(A_Temp . "\ahk_wl_request_word.json")
        try FileDelete(A_Temp . "\ahk_wl_curl.cfg")

        StartWordOllamaRequest(WL_CurrentWord, WL_CurrentContext, false, true)
        return
      }

      ; 如果不是超时或者重试依然失败，则显示彻底失败提示
      if (g_WL_ResultCtrl != "") {
        if (apiErr != "")
          msg := (g_WL_LangMode = "EN" ? "⚠ API error: " : "⚠ 接口报错：") . apiErr
        else if (g_WL_LangMode = "EN")
          msg := isTimeout ? "⚠ Request timed out (>10s)." : "⚠ Connection failed or empty response."
        else
          msg := isTimeout ? "⚠ 请求连续超时，请检查网络。" : "⚠ 请求失败或响应为空，请检查网络连接。"
        try g_WL_ResultCtrl.Value := msg
      }
    }
    
    ; 状态重置与收尾清理
    g_WL_Pending := false
    g_WL_StreamPid := 0
    SetTimer(CheckWordResult, 0)
    
    ; 阅后即焚，清理临时文件
    try FileDelete(g_WL_StreamFile)
    try FileDelete(A_Temp . "\ahk_wl_request_word.json")
    try FileDelete(A_Temp . "\ahk_wl_curl.cfg")
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
      token := UnescapeApiJson(m[1])
      result .= token
    }
  }

  result := Trim(result)
  result := RegExReplace(result, "(\r?\n\s*){2,}", "`n")

  return result ; 性能优化: StripEmoji 移至最终结果时统一调用，避免流式热路径重复执行
}

#Include "word_lookup_utils.ahk"
