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
g_PendingText := ""

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
  global g_TranslateEditCtrl, g_CorrectEditCtrl, g_CorrectLabelCtrl, g_TranslateLabelCtrl, g_IsChineseMode, g_SelectedResult
  
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
  
  g_MainGui.AddText("w500", "原文:")
  origEdit := g_MainGui.AddEdit("w500 h60 ReadOnly", original)
  
  if g_IsChineseMode {
    ; 中文：翻译在前
    g_TranslateLabelCtrl := g_MainGui.AddText("w500", "✓ " . translateLabel)
    g_TranslateEditCtrl := g_MainGui.AddEdit("w500 h60 ReadOnly", "正在翻译...")
    g_CorrectLabelCtrl := g_MainGui.AddText("w500", "   " . correctLabel)
    g_CorrectEditCtrl := g_MainGui.AddEdit("w500 h60 ReadOnly", "正在纠错...")
    g_SelectedResult := "translate"
  } else {
    ; 英文：纠错在前
    g_CorrectLabelCtrl := g_MainGui.AddText("w500", "✓ " . correctLabel)
    g_CorrectEditCtrl := g_MainGui.AddEdit("w500 h60 ReadOnly", "正在纠错...")
    g_TranslateLabelCtrl := g_MainGui.AddText("w500", "   " . translateLabel)
    g_TranslateEditCtrl := g_MainGui.AddEdit("w500 h60 ReadOnly", "正在翻译...")
    g_SelectedResult := "correct"
  }
  
  g_MainGui.AddText("w500 cGray", "Tab 切换 | Enter 替换 | Esc 取消")
  
  g_MainGui.OnEvent("Close", Gui_Close)
  g_MainGui.OnEvent("Escape", Gui_Close)
  
  g_MainGui.Show()
  SendMessage(0xB1, -1, 0, origEdit.Hwnd)
  
  HotIfWinActive("ahk_id " g_MainGui.Hwnd)
  Hotkey("Enter", Gui_Apply.Bind(g_MainGui), "On")
  Hotkey("Tab", Gui_ToggleSelect, "On")
  HotIfWinActive()
  
  ; 直接调用 API
  text := original
  isChinese := RegExMatch(text, "[\x{4e00}-\x{9fff}]")
  
  ; 纠错
  correctResult := OllamaCorrect(text, isChinese)
  if (correctResult != "" && correctResult != "解析失败" && !InStr(correctResult, "请求失败"))
    UpdateCorrectResult(correctResult)
  else
    UpdateCorrectResult("纠错失败: " . correctResult)
  
  ; 翻译
  translateResult := OllamaTranslate(text, isChinese)
  if (translateResult != "" && translateResult != "解析失败" && !InStr(translateResult, "请求失败"))
    UpdateTranslateResult(translateResult)
  else
    UpdateTranslateResult("翻译失败: " . translateResult)
}

Gui_ToggleSelect(*)
{
  global g_SelectedResult, g_CorrectLabelCtrl, g_TranslateLabelCtrl, g_IsChineseMode
  
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
  } else {
    g_SelectedResult := "correct"
    g_CorrectLabelCtrl.Text := "✓ " . correctLabel
    g_TranslateLabelCtrl.Text := "   " . translateLabel
  }
}

UpdateTranslateResult(result)
{
  global g_TranslateResult, g_TranslateEditCtrl
  g_TranslateResult := result
  if (g_TranslateEditCtrl != "")
    g_TranslateEditCtrl.Value := result
}

UpdateCorrectResult(result)
{
  global g_CorrectResult, g_CorrectEditCtrl
  g_CorrectResult := result
  if (g_CorrectEditCtrl != "")
    g_CorrectEditCtrl.Value := result
}

Gui_Apply(guiObj, *)
{
  global g_TranslateResult, g_CorrectResult, g_OldClip, g_SelectedResult
  guiObj.Destroy()
  
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

Gui_Close(guiObj, *)
{
  global g_OldClip
  guiObj.Destroy()
  A_Clipboard := g_OldClip
}

^!Enter::
{
  global g_OldClip, g_IsChineseMode
  g_OldClip := ClipboardAll()
  A_Clipboard := ""
  
  Send("^a")
  Sleep(50)
  Send("^c")
  Errorlevel := !ClipWait(2)
  if ErrorLevel {
    A_Clipboard := g_OldClip
    return
  }
  
  text := Trim(A_Clipboard)
  if (text = "") {
    A_Clipboard := g_OldClip
    return
  }
  
  ShowMainGui(text)
}