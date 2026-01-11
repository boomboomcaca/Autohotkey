;;;;;;;;;使用管理员权限;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Loop, %0% ; For each parameter:
{
  param := %A_Index% ; Fetch the contents of the variable whose name is contained in A_Index.
  params .= A_Space . param
}
ShellExecute := A_IsUnicode ? "shell32\ShellExecute":"shell32\ShellExecuteA"

if not A_IsAdmin
{
  If A_IsCompiled
    DllCall(ShellExecute, uint, 0, str, "RunAs", str, A_ScriptFullPath, str, params , str, A_WorkingDir, int, 1)
  Else
    DllCall(ShellExecute, uint, 0, str, "RunAs", str, A_AhkPath, str, """" . A_ScriptFullPath . """" . A_Space . params, str, A_WorkingDir, int, 1)
  ExitApp
}
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;;
;; An autohotkey script that provides emacs-like keybinding on Windows
;;
#SingleInstance force
#InstallKeybdHook
#UseHook

; The following line is a contribution of NTEmacs wiki http://www49.atwiki.jp/ntemacs/pages/20.html
SetKeyDelay 0

; turns to be 1 when ctrl-x is pressed
is_pre_x = 0
; turns to be 1 when ctrl-space is pressed
is_pre_spc = 0

; Applications you want to disable emacs-like keybindings
; (Please comment out applications you don't use)
is_target()
{
  ; IfWinActive,ahk_class ConsoleWindowClass ; Cygwin
  ;  Return 1
  IfWinActive,ahk_class MEADOW ; Meadow
    Return 1
  IfWinActive,ahk_class cygwin/x X rl-xterm-XTerm-0
  Return 1
  IfWinActive,ahk_class MozillaUIWindowClass ; keysnail on Firefox
    Return 1
  ; Avoid VMwareUnity with AutoHotkey
  IfWinActive,ahk_class VMwareUnityHostWndClass
    Return 1
  IfWinActive,ahk_class Vim ; GVIM
    Return 1
  IfWinActive,ahk_class TMobaXtermForm ; Eclipse
    Return 1
  ;IfWinActive,ahk_class CASCADIA_HOSTING_WINDOW_CLASS
  ;  Return 1
  IfWinActive,ahk_class PotPlayer64
    Return 1
  IfWinActive,ahk_class Emacs ; NTEmacs
    Return 1
  IfWinActive,ahk_class XEmacs ; XEmacs on Cygwin
    Return 1
  Return 0
}

SelectAll()
{
  Send ^a
  global is_pre_spc = 0
  return
}
delete_char()
{
  Send {Del}
  global is_pre_spc = 0
  Return
}
delete_word()
{
  Send ^+{Right}
  Send {Del}
  global is_pre_spc = 0
  Return
}
delete_backward_char()
{
  Send {BS}
  global is_pre_spc = 0
  Return
}
delete_backward_word()
{
  Send ^+{Left}
  Send {BS}
  global is_pre_spc = 0
  Return
}
kill_line()
{
  Send {ShiftDown}{END}{ShiftUp}
  ;Sleep 50 ;[ms] this value depends on your environment
  Send {Del}
  global is_pre_spc = 0
  Return
}
open_line()
{
  Send {END}{Enter}
  global is_pre_spc = 0
  Return
}
quit()
{
  Send {ESC}
  global is_pre_spc = 0
  Return
}
indent_for_tab_command()
{
  Send {Tab}
  global is_pre_spc = 0
  Return
}
newline_and_indent()
{
  Send {Enter}{Tab}
  global is_pre_spc = 0
  Return
}
isearch_current_file()
{
  Send ^f
  global is_pre_spc = 0
  Return
}
isearch_all_files()
{
  Send +^f
  global is_pre_spc = 0
  Return
}
kill_region()
{
  Send ^x
  global is_pre_spc = 0
  Return
}
kill_ring_save()
{
  Send ^c
  global is_pre_spc = 0
  Return
}
yank()
{
  Send ^v
  global is_pre_spc = 0
  Return
}
undo()
{
  Send ^z
  global is_pre_spc = 0
  Return
}
redo()
{
  Send +^z
  global is_pre_spc = 0
  Return
}
find_file()
{
  Send ^o
  global is_pre_x = 0
  Return
}
save_buffer()
{
  Send, ^s
  global is_pre_x = 0
  Return
}
kill_emacs()
{
  Send !{F4}
  global is_pre_x = 0
  Return
}
move_beginning_of_line()
{
  global
  if is_pre_spc
    Send +{HOME}
  Else
    Send {HOME}
  Return
}
move_end_of_line()
{
  global
  if is_pre_spc
    Send +{END}
  Else
    Send {END}
  Return
}
beginning_of_all()
{
  global
  if is_pre_spc
    Send +^{HOME}
  else
    Send ^{HOME}
  return
}
end_of_all()
{
  global
  if is_pre_spc
    Send +^{END}
  else
    Send ^{END}
  return
}
previous_line()
{
  global
  if is_pre_spc
    Send +{Up}
  Else
    Send {Up}
  Return
}
next_line()
{
  global
  if is_pre_spc
    Send +{Down}
  Else
    Send {Down}
  Return
}
forward_char()
{
  global
  if is_pre_spc
    Send +{Right}
  Else
    Send {Right}
  Return
}
forward_word()
{
  global
  if is_pre_spc
    Send +^{Right}
  Else
    Send ^{Right}
  Return
}
backward_char()
{
  global
  if is_pre_spc
    Send +{Left}
  Else
    Send {Left}
  Return
}
backward_word()
{
  global
  If is_pre_spc
    Send +^{Left}
  Else
    Send ^{Left}
  Return
}
scroll_up()
{
  global
  if is_pre_spc
    Send +{PgUp}
  Else
    Send {PgUp}
  Return
}
scroll_down()
{
  global
  if is_pre_spc
    Send +{PgDn}
  Else
    Send {PgDn}
  Return
}

!k::
  if is_target()
    Send %A_ThisHotkey%
  else
    Send ^{Tab}
return
+!k::
  if is_target()
    Send %A_ThisHotkey%
  else
    Send ^+{Tab}
Return
^q::
  if is_target()
    Send %A_ThisHotkey%
  else
    Send ^{F4}
Return
!q::
  if is_target()
    Send %A_ThisHotkey%
  else
    Send !{F4}
Return
!s::
  if is_target()
    Send %A_ThisHotkey%
  else
    Send ^{s}
Return

!a::
  if is_target()
    Send %A_ThisHotkey%
  else
    SelectAll()
return
!f::
  if is_target()
    Send %A_ThisHotkey%
  Else
    forward_word()
Return
!b::
  if is_target()
    Send %A_ThisHotkey%
  Else
    backward_word()
Return
!d::
  If is_target()
    Send %A_ThisHotkey%
  Else
    delete_word()
Return
^f::
  If is_target()
    Send %A_ThisHotkey%
  Else
  {
    If is_pre_x
      find_file()
    Else
      forward_char()
  }
Return
^d::
  If is_target()
    Send %A_ThisHotkey%
  Else
    delete_char()
Return
^h::
  If is_target()
    Send %A_ThisHotkey%
  Else
    delete_backward_char()
Return
^k::
  If is_target()
    Send %A_ThisHotkey%
  Else
    kill_line()
Return
^o::
  If is_target()
    Send %A_ThisHotkey%
  Else
    open_line()
Return
^g::
  If is_target()
    Send %A_ThisHotkey%
  Else
    quit()
Return
!h::
  if is_target()
    Send %A_ThisHotkey%
  else
    delete_backward_word()
return
^s::
  If is_target()
    Send %A_ThisHotkey%
  Else
  {
    If is_pre_x
      save_buffer()
    Else
      isearch_current_file()
  }
Return
^+s::
  If is_target()
    Send %A_ThisHotkey%
  Else
    isearch_all_files()
Return
^w::
  If is_target()
    Send %A_ThisHotkey%
  Else
    kill_region()
Return
!w::
  If is_target()
    Send %A_ThisHotkey%
  Else
    kill_ring_save()
Return

$^Space::
  If is_target()
    Send {CtrlDown}{Space}{CtrlUp}
  Else
  {
    If is_pre_spc
      is_pre_spc = 0
    Else
      is_pre_spc = 1
  }
Return

^@::
  If is_target()
    Send %A_ThisHotkey%
  Else
  {
    If is_pre_spc
      is_pre_spc = 0
    Else
      is_pre_spc = 1
  }
Return
^a::
  If is_target()
    Send %A_ThisHotkey%
  Else
    move_beginning_of_line()
Return
^e::
  If is_target()
    Send %A_ThisHotkey%
  Else
    move_end_of_line()
Return
^p::
  If is_target()
    Send %A_ThisHotkey%
  Else
    previous_line()
Return
^n::
  If is_target()
    Send %A_ThisHotkey%
  Else
    next_line()
Return
^b::
  If is_target()
    Send %A_ThisHotkey%
  Else
    backward_char()
Return
!n::
  If is_target()
    Send %A_ThisHotkey%
  Else
    scroll_down()
Return
!p::
  If is_target()
    Send %A_ThisHotkey%
  Else
    scroll_up()
Return
!<::
  if is_target()
    Send %A_ThisHotkey%
  else
    beginning_of_all()
return
!>::
  if is_target()
    Send %A_ThisHotkey%
  else
    end_of_all()
return

XButton1::Send !{Left}  ; 将鼠标的前进按钮映射为Alt + Left
XButton2::Send !{Right} ; 将鼠标的后退按钮映射为Alt + Right
#Space::Send, {Ctrl down}{Space}{Ctrl up}


;居中显示
#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior spee·d and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
^!down::
  SysGet, Monitor, Monitor
  MonitorWidth := MonitorRight-MonitorLeft
  MonitorHeight := MonitorBottom-MonitorTop
  SysGet, MonitorWorkArea, MonitorWorkArea
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
  WinGetTitle, ActiveWindowTitle, A ; Get the active window's title for "targetting" it/acting on it.
  WinGetPos,,, Width, Height, %ActiveWindowTitle% ; Get the active window's position, used for our calculations.
  TargetX := (A_ScreenWidth/2)-(Width/2) ; Calculate the horizontal target where we'll move the window.
  TargetY := (A_ScreenHeight/2)-(Height/2)-20 ; Calculate the vertical placement of the window.
  WinMove, %ActiveWindowTitle%,, %TargetX%, %TargetY% ; Move the window to the calculated coordinates.
return

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; Ollama 翻译 - Ctrl+Alt+Enter: 全选并翻译替换
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

OllamaTranslate(text)
{
  ; 判断中英文
  if RegExMatch(text, "[\x{4e00}-\x{9fff}]")
    prompt := "Translate to English. Keep the exact same formatting. Output only the translation:`n" . text
  else
    prompt := "Translate to Chinese. Keep the exact same formatting. Output only the translation:`n" . text
  
  ; 构建 JSON
  StringReplace, prompt, prompt, \, \\, All
  StringReplace, prompt, prompt, ", \", All
  StringReplace, prompt, prompt, `n, \n, All
  StringReplace, prompt, prompt, `r, \r, All
  StringReplace, prompt, prompt, `t, \t, All
  json := "{""model"":""qwen3:latest"",""prompt"":""" . prompt . """,""stream"":false}"
  
  ; 调用 Ollama API
  try {
    http := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    http.Open("POST", "http://localhost:11434/api/generate", false)
    http.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
    http.Send(json)
    http.WaitForResponse()
    
    response := http.ResponseText
    ; 提取 response 字段
    if RegExMatch(response, """response""\s*:\s*""(.*?)""(?=\s*,\s*"")", m)
      result := m1
    else
      return "解析失败"
    
    ; 还原转义
    StringReplace, result, result, \n, `n, All
    StringReplace, result, result, \r, `r, All
    StringReplace, result, result, \t, `t, All
    StringReplace, result, result, \", ", All
    StringReplace, result, result, \\, \, All
    
    ; 清理 think 标签
    result := RegExReplace(result, "s)<think>.*?</think>", "")
    StringReplace, result, result, /think,, All
    StringReplace, result, result, /no_think,, All
    
    result := Trim(result)
    return result
  } catch e {
    return "请求失败: " . e.Message
  }
}

^!Enter::
  oldClip := ClipboardAll
  Clipboard := ""
  
  ; 先全选再复制
  Send ^a
  Sleep 100
  Send ^c
  ClipWait, 2
  if ErrorLevel {
    Clipboard := oldClip
    return
  }
  
  text := Trim(Clipboard)
  if (text = "") {
    Clipboard := oldClip
    return
  }
  
  result := OllamaTranslate(text)
  
  if (result != "") {
    Clipboard := result
    Sleep 50
    Send ^v
    Sleep 200
    Clipboard := oldClip
  } else {
    Clipboard := oldClip
  }
return