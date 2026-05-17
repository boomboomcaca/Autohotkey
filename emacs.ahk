#Requires AutoHotkey v2.0

;;;;;;;;;使用管理员权限;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
if !A_IsAdmin {
    try {
        Run '*RunAs "' A_AhkPath '" "' A_ScriptFullPath '"'
    }
    ExitApp()
}
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

; 设置 Per-Monitor DPI Aware v2，确保坐标始终使用物理像素（解决高DPI下截图偏移问题）
DllCall("SetProcessDpiAwarenessContext", "ptr", -4)

;;
;; An autohotkey script that provides emacs-like keybinding on Windows
;;
#SingleInstance force
InstallKeybdHook()
InstallMouseHook()
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

XButton1::Send("!{Left}")  ; 将鼠标的前进按钮映射为Alt + Left
XButton2::Send("!{Right}") ; 将鼠标的后退按钮映射为Alt + Right
#Space::Send("{Ctrl down}{Space}{Ctrl up}")


;居中显示
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
  ActiveWindowID := WinGetID("A") ; Get the active window's ID for "targetting" it/acting on it.
  WinGetPos(, , &Width, &Height, "ahk_id " . ActiveWindowID) ; Get the active window's position, used for our calculations.
  TargetX := (A_ScreenWidth/2)-(Width/2) ; Calculate the horizontal target where we'll move the window.
  TargetY := MonitorWorkAreaTop ; 放在顶端（如果任务栏在上方也不会被遮挡）
  WinMove(TargetX, TargetY, , , "ahk_id " . ActiveWindowID) ; Move the window to the calculated coordinates.
return
} ; V1toV2: Added closing brace for [^!down]

; 缓存获取到的 Gemini 窗口句柄
global GeminiAutoHwnd := 0

; 自动寻找 Gemini 窗口的函数
GetGeminiWindow()
{
    global GeminiAutoHwnd
    
    ; 如果之前找到过并且窗口还在，就直接用之前的句柄（避免最小化后找不到）
    if (GeminiAutoHwnd && WinExist("ahk_id " . GeminiAutoHwnd))
        return GeminiAutoHwnd

    hwnds := WinGetList("ahk_class Chrome_WidgetWin_1 ahk_exe chrome.exe")
    for hwnd in hwnds
    {
        ; 必须是可见窗口
        if !(WinGetStyle(hwnd) & 0x10000000)
            continue
            
        title := WinGetTitle(hwnd)
        
        ; 你的 Gemini 作为 Chrome PWA 运行时，系统获取到的窗口标题正好为空字符串 ""
        if (title == "")
        {
            GeminiAutoHwnd := hwnd
            return hwnd
        }
    }
    return 0
}

; F1 自动寻找并切换 Gemini 窗口的显示/隐藏（最小化/激活）
; 当从外部切换到 Gemini 时，会自动抓取当前鼠标下的单词和句子并粘贴到输入框中
F1::
{
    GeminiHwnd := GetGeminiWindow()
    if (!GeminiHwnd)
    {
        ; 如果没找到空标题的，也可以尝试找找名字里带 Gemini 的
        if WinExist("Gemini ahk_exe chrome.exe")
            GeminiHwnd := WinGetID("Gemini ahk_exe chrome.exe")
        else
        {
            MsgBox("未检测到 Gemini 窗口，请确保它已经打开！", "提示", "T3")
            return
        }
    }

    ; 判断窗口是否处于最小化状态（-1 表示最小化）
    isMin := (WinGetMinMax("ahk_id " . GeminiHwnd) == -1)

    if (!isMin)
    {
        ; 只要窗口在屏幕上（不管是不是活动窗口），按 F1 一律直接隐藏（最小化）
        WinMinimize("ahk_id " . GeminiHwnd)
    }
    else
    {
        ; 如果窗口当前被隐藏了（处于最小化状态），则：
        ; 1. 先抓取当前鼠标下的词句（必须在激活窗口前抓取，否则会失去原界面的焦点）
        word := ""
        line := ""
        hasWord := GetWordAndLineAtMouse(&word, &line)
        
        ; 2. 恢复并激活 Gemini 窗口
        WinActivate("ahk_id " . GeminiHwnd)
        
        ; 3. 如果成功抓取到词句，则将其处理干净（过滤表情、对象占位符，且将所有换行和连续空格压缩为单行单空格）
        if (hasWord)
        {
            word := Trim(RegExReplace(StripEmoji(word), "[\r\n\s]+", " "))
            line := Trim(RegExReplace(StripEmoji(line), "[\r\n\s]+", " "))
            
            ; 缓存为当前生词并预生成 TTS 语音，供右键朗读使用
            global WL_CurrentWord
            WL_CurrentWord := word
            WL_PregenTts(word)
            
            textToSend := "单词: " . word . "`n句子: " . line
            
            ClipSaved := ClipboardAll()
            A_Clipboard := textToSend
            
            ; 等待窗口激活后执行清除并粘贴
            if WinWaitActive("ahk_id " . GeminiHwnd, , 2)
            {
                Sleep(200)
                Send("^a") ; 全选已有内容
                Sleep(50)
                Send("^v") ; 粘贴新内容覆盖
            }
            
            Sleep(100)
            A_Clipboard := ClipSaved
        }
    }
}

; 当 Google Gemini 窗口处于活动状态时，拦截鼠标右键，点击后朗读当前查询的单词
#HotIf (GeminiAutoHwnd && WinActive("ahk_id " . GeminiAutoHwnd))
RButton::
{
    WL_PlayTtsOnce()
}
#HotIf

; 共享模块（只引入一次）
#Include "ollama_tts.ahk"
#Include "ollama_prompt_chat.ahk"

; 引入 Ollama 翻译/纠错模块
#Include "ollama_translate.ahk"

; 引入鼠标取词模块
#Include "word_lookup.ahk"