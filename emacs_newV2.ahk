#Requires AutoHotkey v2.0

;;;;;;;;;使用管理员权限;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
if !A_IsAdmin {
    try {
        Run '*RunAs "' A_AhkPath '" "' A_ScriptFullPath '"'
    }
    ExitApp()
}
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;;
;; An autohotkey script that provides emacs-like keybinding on Windows
;;
#SingleInstance force
InstallKeybdHook()
#UseHook

; The following line is a contribution of NTEmacs wiki http://www49.atwiki.jp/ntemacs/pages/20.html
SetKeyDelay(0)

; turns to be 1 when ctrl-x is pressed
is_pre_x := 0
; turns to be 1 when ctrl-space is pressed
is_pre_spc := 0

; Applications you want to disable emacs-like keybindings
; (Please comment out applications you don't use)
is_target()
{
  ; IfWinActive,ahk_class ConsoleWindowClass ; Cygwin
  ;  Return 1
  if WinActive("ahk_class MEADOW") ; Meadow
    Return 1
  if WinActive("ahk_class cygwin/x X rl-xterm-XTerm-0")
  Return 1
  if WinActive("ahk_class MozillaUIWindowClass") ; keysnail on Firefox
    Return 1
  ; Avoid VMwareUnity with AutoHotkey
  if WinActive("ahk_class VMwareUnityHostWndClass")
    Return 1
  if WinActive("ahk_class Vim") ; GVIM
    Return 1
  if WinActive("ahk_class TMobaXtermForm") ; Eclipse
    Return 1
  ;IfWinActive,ahk_class CASCADIA_HOSTING_WINDOW_CLASS
  ;  Return 1
  if WinActive("ahk_class PotPlayer64")
    Return 1
  if WinActive("ahk_class Emacs") ; NTEmacs
    Return 1
  if WinActive("ahk_class XEmacs") ; XEmacs on Cygwin
    Return 1
  Return 0
}

SelectAll()
{
  Send("^a")
  global is_pre_spc := ""
  return
}
delete_char()
{
  Send("{Del}")
  global is_pre_spc := ""
  Return
}
delete_word()
{
  Send("^+{Right}")
  Send("{Del}")
  global is_pre_spc := ""
  Return
}
delete_backward_char()
{
  Send("{BS}")
  global is_pre_spc := ""
  Return
}
delete_backward_word()
{
  Send("^+{Left}")
  Send("{BS}")
  global is_pre_spc := ""
  Return
}
kill_line()
{
  Send("{ShiftDown}{END}{ShiftUp}")
  ;Sleep 50 ;[ms] this value depends on your environment
  Send("{Del}")
  global is_pre_spc := ""
  Return
}
open_line()
{
  Send("{END}{Enter}")
  global is_pre_spc := ""
  Return
}
quit()
{
  Send("{ESC}")
  global is_pre_spc := ""
  Return
}
indent_for_tab_command()
{
  Send("{Tab}")
  global is_pre_spc := ""
  Return
}
newline_and_indent()
{
  Send("{Enter}{Tab}")
  global is_pre_spc := ""
  Return
}
isearch_current_file()
{
  Send("^f")
  global is_pre_spc := ""
  Return
}
isearch_all_files()
{
  Send("+^f")
  global is_pre_spc := ""
  Return
}
kill_region()
{
  Send("^x")
  global is_pre_spc := ""
  Return
}
kill_ring_save()
{
  Send("^c")
  global is_pre_spc := ""
  Return
}
yank()
{
  Send("^v")
  global is_pre_spc := ""
  Return
}
undo()
{
  Send("^z")
  global is_pre_spc := ""
  Return
}
redo()
{
  Send("+^z")
  global is_pre_spc := ""
  Return
}
find_file()
{
  Send("^o")
  global is_pre_x := ""
  Return
}
save_buffer()
{
  Send("^s")
  global is_pre_x := ""
  Return
}
kill_emacs()
{
  Send("!{F4}")
  global is_pre_x := ""
  Return
}
move_beginning_of_line()
{
  global
  if is_pre_spc
    Send("+{HOME}")
  Else
    Send("{HOME}")
  Return
}
move_end_of_line()
{
  global
  if is_pre_spc
    Send("+{END}")
  Else
    Send("{END}")
  Return
}
beginning_of_all()
{
  global
  if is_pre_spc
    Send("+^{HOME}")
  else
    Send("^{HOME}")
  return
}
end_of_all()
{
  global
  if is_pre_spc
    Send("+^{END}")
  else
    Send("^{END}")
  return
}
previous_line()
{
  global
  if is_pre_spc
    Send("+{Up}")
  Else
    Send("{Up}")
  Return
}
next_line()
{
  global
  if is_pre_spc
    Send("+{Down}")
  Else
    Send("{Down}")
  Return
}
forward_char()
{
  global
  if is_pre_spc
    Send("+{Right}")
  Else
    Send("{Right}")
  Return
}
forward_word()
{
  global
  if is_pre_spc
    Send("+^{Right}")
  Else
    Send("^{Right}")
  Return
}
backward_char()
{
  global
  if is_pre_spc
    Send("+{Left}")
  Else
    Send("{Left}")
  Return
}
backward_word()
{
  global
  If is_pre_spc
    Send("+^{Left}")
  Else
    Send("^{Left}")
  Return
}
scroll_up()
{
  global
  if is_pre_spc
    Send("+{PgUp}")
  Else
    Send("{PgUp}")
  Return
}
scroll_down()
{
  global
  if is_pre_spc
    Send("+{PgDn}")
  Else
    Send("{PgDn}")
  Return
}

!k::
{ ; V1toV2: Added opening brace for [!k]
global ; V1toV2: Made function global
  if is_target()
    Send(A_ThisHotkey)
  else
    Send("^{Tab}")
return
} ; V1toV2: Added closing brace for [!k]
+!k::
{ ; V1toV2: Added opening brace for [+!k]
global ; V1toV2: Made function global
  if is_target()
    Send(A_ThisHotkey)
  else
    Send("^+{Tab}")
Return
} ; V1toV2: Added closing brace for [+!k]
^q::
{ ; V1toV2: Added opening brace for [^q]
global ; V1toV2: Made function global
  if is_target()
    Send(A_ThisHotkey)
  else
    Send("^{F4}")
Return
} ; V1toV2: Added closing brace for [^q]
!q::
{ ; V1toV2: Added opening brace for [!q]
global ; V1toV2: Made function global
  if is_target()
    Send(A_ThisHotkey)
  else
    Send("!{F4}")
Return
} ; V1toV2: Added closing brace for [!q]
!s::
{ ; V1toV2: Added opening brace for [!s]
global ; V1toV2: Made function global
  if is_target()
    Send(A_ThisHotkey)
  else
    Send("^{s}")
Return
} ; V1toV2: Added closing brace for [!s]

!a::
{ ; V1toV2: Added opening brace for [!a]
global ; V1toV2: Made function global
  if is_target()
    Send(A_ThisHotkey)
  else
    SelectAll()
return
} ; V1toV2: Added closing brace for [!a]
!f::
{ ; V1toV2: Added opening brace for [!f]
global ; V1toV2: Made function global
  if is_target()
    Send(A_ThisHotkey)
  Else
    forward_word()
Return
} ; V1toV2: Added closing brace for [!f]
!b::
{ ; V1toV2: Added opening brace for [!b]
global ; V1toV2: Made function global
  if is_target()
    Send(A_ThisHotkey)
  Else
    backward_word()
Return
} ; V1toV2: Added closing brace for [!b]
!d::
{ ; V1toV2: Added opening brace for [!d]
global ; V1toV2: Made function global
  If is_target()
    Send(A_ThisHotkey)
  Else
    delete_word()
Return
} ; V1toV2: Added closing brace for [!d]
^f::
{ ; V1toV2: Added opening brace for [^f]
global ; V1toV2: Made function global
  If is_target()
    Send(A_ThisHotkey)
  Else
  {
    If is_pre_x
      find_file()
    Else
      forward_char()
  }
Return
} ; V1toV2: Added closing brace for [^f]
^d::
{ ; V1toV2: Added opening brace for [^d]
global ; V1toV2: Made function global
  If is_target()
    Send(A_ThisHotkey)
  Else
    delete_char()
Return
} ; V1toV2: Added closing brace for [^d]
^h::
{ ; V1toV2: Added opening brace for [^h]
global ; V1toV2: Made function global
  If is_target()
    Send(A_ThisHotkey)
  Else
    delete_backward_char()
Return
} ; V1toV2: Added closing brace for [^h]
^k::
{ ; V1toV2: Added opening brace for [^k]
global ; V1toV2: Made function global
  If is_target()
    Send(A_ThisHotkey)
  Else
    kill_line()
Return
} ; V1toV2: Added closing brace for [^k]
^o::
{ ; V1toV2: Added opening brace for [^o]
global ; V1toV2: Made function global
  If is_target()
    Send(A_ThisHotkey)
  Else
    open_line()
Return
} ; V1toV2: Added closing brace for [^o]
^g::
{ ; V1toV2: Added opening brace for [^g]
global ; V1toV2: Made function global
  If is_target()
    Send(A_ThisHotkey)
  Else
    quit()
Return
} ; V1toV2: Added closing brace for [^g]
!h::
{ ; V1toV2: Added opening brace for [!h]
global ; V1toV2: Made function global
  if is_target()
    Send(A_ThisHotkey)
  else
    delete_backward_word()
return
} ; V1toV2: Added closing brace for [!h]
^s::
{ ; V1toV2: Added opening brace for [^s]
global ; V1toV2: Made function global
  If is_target()
    Send(A_ThisHotkey)
  Else
  {
    If is_pre_x
      save_buffer()
    Else
      isearch_current_file()
  }
Return
} ; V1toV2: Added closing brace for [^s]
^+s::
{ ; V1toV2: Added opening brace for [^+s]
global ; V1toV2: Made function global
  If is_target()
    Send(A_ThisHotkey)
  Else
    isearch_all_files()
Return
} ; V1toV2: Added closing brace for [^+s]
^w::
{ ; V1toV2: Added opening brace for [^w]
global ; V1toV2: Made function global
  If is_target()
    Send(A_ThisHotkey)
  Else
    kill_region()
Return
} ; V1toV2: Added closing brace for [^w]
!w::
{ ; V1toV2: Added opening brace for [!w]
global ; V1toV2: Made function global
  If is_target()
    Send(A_ThisHotkey)
  Else
    kill_ring_save()
Return
} ; V1toV2: Added closing brace for [!w]

$^Space::
{ ; V1toV2: Added opening brace for [$^Space]
global ; V1toV2: Made function global
  If is_target()
    Send("{CtrlDown}{Space}{CtrlUp}")
  Else
  {
    If is_pre_spc
      is_pre_spc := 0
    Else
      is_pre_spc := 1
  }
Return
} ; V1toV2: Added closing brace for [$^Space]

^@::
{ ; V1toV2: Added opening brace for [^@]
global ; V1toV2: Made function global
  If is_target()
    Send(A_ThisHotkey)
  Else
  {
    If is_pre_spc
      is_pre_spc := 0
    Else
      is_pre_spc := 1
  }
Return
} ; V1toV2: Added closing brace for [^@]
^a::
{ ; V1toV2: Added opening brace for [^a]
global ; V1toV2: Made function global
  If is_target()
    Send(A_ThisHotkey)
  Else
    move_beginning_of_line()
Return
} ; V1toV2: Added closing brace for [^a]
^e::
{ ; V1toV2: Added opening brace for [^e]
global ; V1toV2: Made function global
  If is_target()
    Send(A_ThisHotkey)
  Else
    move_end_of_line()
Return
} ; V1toV2: Added closing brace for [^e]
^p::
{ ; V1toV2: Added opening brace for [^p]
global ; V1toV2: Made function global
  If is_target()
    Send(A_ThisHotkey)
  Else
    previous_line()
Return
} ; V1toV2: Added closing brace for [^p]
^n::
{ ; V1toV2: Added opening brace for [^n]
global ; V1toV2: Made function global
  If is_target()
    Send(A_ThisHotkey)
  Else
    next_line()
Return
} ; V1toV2: Added closing brace for [^n]
^b::
{ ; V1toV2: Added opening brace for [^b]
global ; V1toV2: Made function global
  If is_target()
    Send(A_ThisHotkey)
  Else
    backward_char()
Return
} ; V1toV2: Added closing brace for [^b]
!n::
{ ; V1toV2: Added opening brace for [!n]
global ; V1toV2: Made function global
  If is_target()
    Send(A_ThisHotkey)
  Else
    scroll_down()
Return
} ; V1toV2: Added closing brace for [!n]
!p::
{ ; V1toV2: Added opening brace for [!p]
global ; V1toV2: Made function global
  If is_target()
    Send(A_ThisHotkey)
  Else
    scroll_up()
Return
} ; V1toV2: Added closing brace for [!p]
!<::
{ ; V1toV2: Added opening brace for [!<]
global ; V1toV2: Made function global
  if is_target()
    Send(A_ThisHotkey)
  else
    beginning_of_all()
return
} ; V1toV2: Added closing brace for [!<]
!>::
{ ; V1toV2: Added opening brace for [!>]
global ; V1toV2: Made function global
  if is_target()
    Send(A_ThisHotkey)
  else
    end_of_all()
return
} ; V1toV2: Added closing brace for [!>]

V1toV2_GblCode_001:
XButton1::Send("!{Left}")  ; 将鼠标的前进按钮映射为Alt + Left
XButton2::Send("!{Right}") ; 将鼠标的后退按钮映射为Alt + Right
#Space::Send("{Ctrl down}{Space}{Ctrl up}")


;居中显示
; V1toV2: Removed #NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode("Input") ; Recommended for new scripts due to its superior spee·d and reliability.
SetWorkingDir(A_ScriptDir) ; Ensures a consistent starting directory.
^!down::
{ ; V1toV2: Added opening brace for [^!down]
global ; V1toV2: Made function global
  MonitorGet(, &MonitorLeft, &MonitorTop, &MonitorRight, &MonitorBottom)
  MonitorWidth := MonitorRight-MonitorLeft
  MonitorHeight := MonitorBottom-MonitorTop
  MonitorGetWorkArea(, &MonitorWorkAreaLeft, &MonitorWorkAreaTop, &MonitorWorkAreaRight, &MonitorWorkAreaBottom)
  MonitorWorkAreaWidth := MonitorWorkAreaRight-MonitorWorkAreaLeft
  MonitorWorkAreaHeight := MonitorWorkAreaBottom-MonitorWorkAreaTop
  If (MonitorWidth=MonitorWorkAreaWidth)
    TrayWidth := MonitorWidth
  Else
    TrayWidth := MonitorWidth-MonitorWorkAreaWidth
  If (MonitorHeight=MonitorWorkAreaHeight)
    TrayHeight := MonitorHeight
  Else
    TrayHeight := MonitorHeight-MonitorWorkAreaHeight
  ActiveWindowTitle := WinGetTitle("A") ; Get the active window's title for "targetting" it/acting on it.
  WinGetPos(, , &Width, &Height, ActiveWindowTitle) ; Get the active window's position, used for our calculations.
  TargetX := (A_ScreenWidth/2)-(Width/2) ; Calculate the horizontal target where we'll move the window.
  TargetY := (A_ScreenHeight/2)-(Height/2)-20 ; Calculate the vertical placement of the window.
  WinMove(TargetX, TargetY, , , ActiveWindowTitle) ; Move the window to the calculated coordinates.
return
} ; V1toV2: Added closing brace for [^!down]
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

; 流式响应相关
g_StreamFileCorrect := ""
g_StreamFileTranslate := ""
g_StreamPidCorrect := 0
g_StreamPidTranslate := 0
g_StreamContentCorrect := ""
g_StreamContentTranslate := ""

OllamaCall(prompt)
{
  ; 构建 JSON
  prompt := StrReplace(prompt, "\", "\\")
  prompt := StrReplace(prompt, "`"", "\`"")
  prompt := StrReplace(prompt, "`n", "\n")
  prompt := StrReplace(prompt, "`r", "\r")
  prompt := StrReplace(prompt, "`t", "\t")
  json := "{`"model`":`"qwen3:latest`",`"prompt`":`"" . prompt . "`",`"stream`":false,`"options`":{`"temperature`":0,`"num_predict`":2048}}"
  
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
    prompt := "/no_think You are an English language tutor. Correct and improve the following English text. Fix grammar, spelling, punctuation, and improve expression while keeping the original meaning. Output only the corrected text without any explanation:`n" . text
  return OllamaCall(prompt)
}

ShowMainGui(original)
{
  global g_OriginalText, g_TranslateResult, g_CorrectResult, g_OldClip, g_MainGui
  global g_TranslateEditCtrl, g_CorrectEditCtrl, g_CorrectLabelCtrl, g_TranslateLabelCtrl, g_OrigEditCtrl, g_IsChineseMode, g_SelectedResult
  global g_TtsOrigCtrl, g_TtsCorrectCtrl, g_TtsTranslateCtrl
  
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
  g_SelectedResult := "correct"
  
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
  
  g_MainGui := Gui("+AlwaysOnTop -MinimizeBox", title)
  g_MainGui.SetFont("s10", "Microsoft YaHei")
  
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
    g_TranslateEditCtrl := g_MainGui.AddEdit("xm w500 h60 ReadOnly", "正在翻译...")
    g_CorrectLabelCtrl := g_MainGui.AddText("w500", "   " . correctLabel)
    g_CorrectEditCtrl := g_MainGui.AddEdit("w500 h60 ReadOnly", "(切换后加载)")
    g_SelectedResult := "translate"
  } else {
    ; 英文：纠错在前，添加朗读图标
    g_CorrectLabelCtrl := g_MainGui.AddText("w120 Section", "✓ " . correctLabel)
    g_TtsCorrectCtrl := g_MainGui.AddText("x+5 ys cGray", "🔊")
    g_TtsCorrectCtrl.OnEvent("Click", Gui_PlayCorrect)
    g_CorrectEditCtrl := g_MainGui.AddEdit("xm w500 h60 ReadOnly", "正在纠错...")
    g_TranslateLabelCtrl := g_MainGui.AddText("w120 Section", "   " . translateLabel)
    g_TtsTranslateCtrl := g_MainGui.AddText("x+5 ys cGray", "🔊")
    g_TtsTranslateCtrl.OnEvent("Click", Gui_PlayTranslate)
    g_TranslateEditCtrl := g_MainGui.AddEdit("xm w500 h60 ReadOnly", "(切换后加载)")
    g_SelectedResult := "correct"
  }
  
  g_MainGui.AddText("w500 cGray", "Ctrl+Tab 切换 | Enter 替换 | Ctrl+Enter 重新处理 | Esc 取消")
  
  g_MainGui.OnEvent("Close", Gui_Close)
  g_MainGui.OnEvent("Escape", Gui_Close)
  
  g_MainGui.Show()
  
  
  HotIfWinActive("ahk_id " g_MainGui.Hwnd)
  Hotkey("Enter", Gui_Apply.Bind(g_MainGui), "On")
  Hotkey("NumpadEnter", Gui_Apply.Bind(g_MainGui), "On")
  Hotkey("^Enter", Gui_Retry, "On")
  Hotkey("^NumpadEnter", Gui_Retry, "On")
  Hotkey("^Tab", Gui_ToggleSelect, "On")
  HotIfWinActive()
  
  ; 重置请求状态并异步调用 API
  global g_CorrectRequested, g_TranslateRequested
  g_CorrectRequested := false
  g_TranslateRequested := false
  StartAsyncRequests(original, g_SelectedResult)
  
  ; 启动悬停检测定时器（两种模式都需要）
  SetTimer(CheckTtsHover, 200)
}

StartAsyncRequests(text, requestType := "default")
{
  global g_HttpCorrect, g_HttpTranslate, g_CorrectPending, g_TranslatePending, g_IsChineseMode
  global g_CorrectRequested, g_TranslateRequested, g_CurrentText
  
  g_CurrentText := text
  isChinese := RegExMatch(text, "[\x{4e00}-\x{9fff}]")
  
  ; 初始请求：按界面顺序，中文先翻译，英文先纠错
  if (requestType = "default") {
    g_CorrectRequested := false
    g_TranslateRequested := false
    if isChinese
      requestType := "translate"
    else
      requestType := "correct"
  }
  
  if (requestType = "correct" && !g_CorrectRequested) {
    if isChinese
      correctPrompt := "/no_think You are a Chinese language tutor. Correct and improve the following Chinese text. Fix grammar, punctuation, and improve expression while keeping the original meaning. Output only the corrected text without any explanation:`n" . text
    else
      correctPrompt := "/no_think You are an English language tutor. Correct and improve the following English text. Fix grammar, spelling, punctuation, and improve expression while keeping the original meaning. Output only the corrected text without any explanation:`n" . text
    g_HttpCorrect := StartAsyncHttp(correctPrompt, "correct")
    g_CorrectPending := true
    g_CorrectRequested := true
  }
  
  if (requestType = "translate" && !g_TranslateRequested) {
    if isChinese
      translatePrompt := "/no_think Translate to English. Keep the exact same formatting, including punctuation marks, line breaks, and spacing. Output only the translation:`n" . text
    else
      translatePrompt := "/no_think Translate to Chinese. Keep the exact same formatting, including punctuation marks, line breaks, and spacing. Output only the translation:`n" . text
    g_HttpTranslate := StartAsyncHttp(translatePrompt, "translate")
    g_TranslatePending := true
    g_TranslateRequested := true
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
  json := '{"model":"qwen3:latest","prompt":"' . prompt . '","stream":true,"options":{"temperature":0,"num_predict":2048}}'
  
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
  
  ; 检查纠错结果（检测 done:true）
  if (g_CorrectPending && g_StreamFileCorrect != "") {
    if (IsStreamComplete(g_StreamFileCorrect)) {
      Sleep(200)
      result := ReadStreamFile(g_StreamFileCorrect, &g_StreamContentCorrect)
      if (result != "") {
        UpdateCorrectResult(result)
      }
      g_CorrectPending := false
      ; 
      if (!g_IsChineseMode && !g_TranslateRequested && g_TranslateEditCtrl != "") {
        try {
          g_TranslateEditCtrl.Value := "正在翻译..."
        }
        StartAsyncRequests(g_CurrentText, "translate")
      }
    }
  }
  
  ; 
  if (g_TranslatePending && g_StreamFileTranslate != "") {
    if (IsStreamComplete(g_StreamFileTranslate)) {
      Sleep(200)
      result := ReadStreamFile(g_StreamFileTranslate, &g_StreamContentTranslate)
      if (result != "") {
        UpdateTranslateResult(result)
      }
      g_TranslatePending := false
      ; 中文模式：翻译完成后自动开始纠错
      if (g_IsChineseMode && !g_CorrectRequested && g_CorrectEditCtrl != "") {
        try {
          g_CorrectEditCtrl.Value := "正在纠错..."
        }
        StartAsyncRequests(g_CurrentText, "correct")
      }
    }
  }
  
  ; 如果都完成了，停止定时器
  if (!g_CorrectPending && !g_TranslatePending) {
    SetTimer(CheckAsyncResults, 0)
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
  global g_CorrectEditCtrl, g_TranslateEditCtrl
  
  if (g_IsChineseMode) {
    correctLabel := "纠错 (中文润色):"
    translateLabel := "翻译 (中→英):"
  } else {
    correctLabel := "纠错 (英文润色):"
    translateLabel := "翻译 (英→中):"
  }
  
  if (g_SelectedResult = "correct") {
    g_SelectedResult := "translate"
    g_CorrectLabelCtrl.Text := "   " . correctLabel
    g_TranslateLabelCtrl.Text := "✓ " . translateLabel
    ; 按需请求翻译
    if (!g_TranslateRequested) {
      g_TranslateEditCtrl.Value := "正在翻译..."
      StartAsyncRequests(g_CurrentText, "translate")
    }
  } else {
    g_SelectedResult := "correct"
    g_CorrectLabelCtrl.Text := "✓ " . correctLabel
    g_TranslateLabelCtrl.Text := "   " . translateLabel
    ; 按需请求纠错
    if (!g_CorrectRequested) {
      g_CorrectEditCtrl.Value := "正在纠错..."
      StartAsyncRequests(g_CurrentText, "correct")
    }
  }
}

UpdateTranslateResult(result)
{
  global g_TranslateResult, g_TranslateEditCtrl
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
  global g_CorrectResult, g_CorrectEditCtrl
  g_CorrectResult := result
  if (g_CorrectEditCtrl != "") {
    try {
      g_CorrectEditCtrl.Value := result
    } catch {
      g_CorrectEditCtrl := ""
    }
  }
}

Gui_Apply(guiObj, *)
{
  global g_TranslateResult, g_CorrectResult, g_OldClip, g_SelectedResult
  global g_MainGui, g_TranslateEditCtrl, g_CorrectEditCtrl, g_OrigEditCtrl
  guiObj.Destroy()
  g_MainGui := ""
  g_TranslateEditCtrl := ""
  g_CorrectEditCtrl := ""
  g_OrigEditCtrl := ""
  
  result := (g_SelectedResult = "translate") ? g_TranslateResult : g_CorrectResult
  
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

Gui_Close(guiObj, *)
{
  global g_OldClip, g_TtsPlaying, g_HoverTarget
  global g_StreamPidCorrect, g_StreamPidTranslate, g_CorrectPending, g_TranslatePending
  global g_MainGui, g_TranslateEditCtrl, g_CorrectEditCtrl, g_OrigEditCtrl
  
  ; 终止正在运行的 PowerShell 进程
  if (g_StreamPidCorrect > 0) {
    try ProcessClose(g_StreamPidCorrect)
    g_StreamPidCorrect := 0
  }
  if (g_StreamPidTranslate > 0) {
    try ProcessClose(g_StreamPidTranslate)
    g_StreamPidTranslate := 0
  }
  g_CorrectPending := false
  g_TranslatePending := false
  SetTimer(CheckAsyncResults, 0)
  
  g_TtsPlaying := false
  g_HoverTarget := ""
  SetTimer(CheckTtsHover, 0)
  guiObj.Destroy()
  g_MainGui := ""
  g_TranslateEditCtrl := ""
  g_CorrectEditCtrl := ""
  g_OrigEditCtrl := ""
  A_Clipboard := g_OldClip
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
    Errorlevel := !ClipWait(2)
    if ErrorLevel {
      A_Clipboard := g_OldClip
      return
    }
    text := Trim(A_Clipboard)
  }

  if (text = "") {
    A_Clipboard := g_OldClip
    return
  }

  ShowMainGui(text)
}