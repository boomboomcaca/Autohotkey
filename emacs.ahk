#Requires AutoHotkey v2.0

;;;;;;;;;使用管理员权限;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
if !A_IsAdmin {
    try {
        Run '*RunAs "' A_AhkPath '" "' A_ScriptFullPath '"'
    } catch {
        ; UAC 提权被拒绝：给出提示再退出，避免双击脚本后无声无息地消失
        MsgBox("脚本需要管理员权限运行（UAC 提权被拒绝），即将退出。", "AI Assistant", "Icon! T5")
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
    forward_char()
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
    isearch_current_file()
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
; 注意：这里故意不注册 ^x:: 热键，Ctrl+X 保持系统默认的剪切功能。
; （曾尝试把 C-x 实现为 Emacs 前缀键以支持 C-x C-f / C-x C-s，
;   但那会吞掉所有普通窗口的 Ctrl+X 剪切，得不偿失，已回退，
;   相关的 is_pre_x 前缀机制及 find_file/save_buffer 函数已一并移除。）
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
  ActiveWindowID := WinGetID("A") ; 活动窗口句柄
  if (!ActiveWindowID)
    return
  WinGetPos(, , &Width, &Height, "ahk_id " . ActiveWindowID)

  ; 取“活动窗口所在显示器”的工作区（而非主显示器），修正多屏下窗口被拉回主屏的问题
  hMon := DllCall("MonitorFromWindow", "Ptr", ActiveWindowID, "UInt", 2, "Ptr") ; MONITOR_DEFAULTTONEAREST
  mi := Buffer(40, 0)
  NumPut("UInt", 40, mi, 0) ; cbSize
  DllCall("GetMonitorInfoW", "Ptr", hMon, "Ptr", mi)
  WorkLeft := NumGet(mi, 20, "Int"), WorkTop := NumGet(mi, 24, "Int")
  WorkRight := NumGet(mi, 28, "Int"), WorkBottom := NumGet(mi, 32, "Int")
  WorkWidth := WorkRight - WorkLeft
  WorkHeight := WorkBottom - WorkTop

  TargetX := WorkLeft + (WorkWidth/2)-(Width/2) ; 水平居中于本显示器工作区
  ; Gemini 窗口：置顶居中（工作区顶部）；其他窗口：垂直居中（工作区内）
  GeminiHwnd := GetGeminiWindow()
  if (GeminiHwnd && ActiveWindowID = GeminiHwnd)
    TargetY := WorkTop
  Else
    TargetY := WorkTop + (WorkHeight/2) - (Height/2)
  WinMove(TargetX, TargetY, , , "ahk_id " . ActiveWindowID) ; Move the window to the calculated coordinates.
return
} ; V1toV2: Added closing brace for [^!down]

; 缓存获取到的 Gemini 窗口句柄
global GeminiAutoHwnd := 0

; 自动寻找 Gemini 窗口的函数
GetGeminiWindow()
{
    global GeminiAutoHwnd
    
    ; 如果之前找到过并且窗口还在，就直接用之前的句柄（避免隐藏后找不到）
    if (GeminiAutoHwnd)
    {
        DetectHiddenWindows(true)
        exists := WinExist("ahk_id " . GeminiAutoHwnd)
        DetectHiddenWindows(false)
        if (exists)
            return GeminiAutoHwnd
    }

    hwnds := WinGetList("ahk_class Chrome_WidgetWin_1 ahk_exe chrome.exe")
    for hwnd in hwnds
    {
        style := WinGetStyle(hwnd)
        ; 必须是可见窗口
        if !(style & 0x10000000)
            continue

        ; 必须带标准标题栏（WS_CAPTION = 0xC00000）且尺寸像正常应用窗口：
        ; Chrome 的拖拽预览、气泡提示等辅助窗口也是可见+空标题的 Chrome_WidgetWin_1，
        ; 不过滤会被误认成 Gemini，导致 F2 隐藏错误的窗口
        if ((style & 0xC00000) != 0xC00000)
            continue
        WinGetPos(, , &w, &h, hwnd)
        if (w < 200 || h < 200)
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

; 窗口完全（或大半）落在当前虚拟屏幕之外时，把它移回主屏工作区中央。
; RDP 连接会把会话分辨率/布局改成客户端的，而隐藏中的窗口不会被系统重新摆放，
; WinShow 后可能整个落在已不存在的屏幕区域上，表现为"按 F2 没反应"
EnsureWindowOnScreen(hwnd)
{
    try {
        WinGetPos(&x, &y, &w, &h, "ahk_id " . hwnd)
        vx := SysGet(76), vy := SysGet(77), vw := SysGet(78), vh := SysGet(79)
        ix := Max(0, Min(x + w, vx + vw) - Max(x, vx))
        iy := Max(0, Min(y + h, vy + vh) - Max(y, vy))
        if (ix * iy >= w * h / 2)
            return
        MonitorGetWorkArea(MonitorGetPrimary(), &wl, &wt, &wr, &wb)
        newW := Min(w, wr - wl), newH := Min(h, wb - wt)
        WinMove(wl + ((wr - wl) - newW) // 2, wt + ((wb - wt) - newH) // 2, newW, newH, "ahk_id " . hwnd)
    }
}

; 通过 UIA 定位 Gemini 输入框（窗口底部最靠下的可聚焦编辑框）并点击聚焦。
; 坐标来自元素实际位置，RDP 会话下 DPI/分辨率变化时依然命中；
; 轮询等待输入框出现，覆盖 RDP 位图远传导致 Chrome 渲染变慢的情况
FocusGeminiInput(hwnd, timeoutMs := 2500)
{
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline)
    {
        try {
            WinGetPos(&wx, &wy, &ww, &wh, "ahk_id " . hwnd)
            root := UIA.ElementFromHandle(hwnd)
            best := ""
            bestLoc := ""
            bestBottom := -2147483648
            for e in root.FindElements({Type:"Edit"})
            {
                try {
                    if (!e.IsKeyboardFocusable || e.IsOffscreen)
                        continue
                    loc := e.Location
                    if (loc.w < 50 || loc.h < 10)
                        continue
                    ; 只接受"输入框形状"的候选：高度有限且贴近窗口底部。
                    ; Gemini 的 Canvas/文档面板也是可聚焦 Edit，但接近全窗口高，
                    ; 若误选中它，后面的 ^a/^v 会覆盖用户文档内容
                    if (loc.h > wh * 0.4 || loc.y + loc.h < wy + wh / 2)
                        continue
                    if (loc.y + loc.h > bestBottom)
                    {
                        bestBottom := loc.y + loc.h
                        best := e
                        bestLoc := loc
                    }
                }
            }
            if (best)
            {
                try best.SetFocus()
                CoordMode("Mouse", "Screen")
                Click(bestLoc.x + (bestLoc.w // 2), bestLoc.y + (bestLoc.h // 2))
                Sleep(100)
                return true
            }
        }
        Sleep(100)
    }
    return false
}

; 轮询等待 Gemini 输入框注册到粘贴内容后再回车。
; Gemini 是 contenteditable + React，粘贴到"发送按钮可用"之间有延迟；
; 通过 UIA 读取底部输入框的文本值，非空即认为可以安全发送。
; UIA 不可用时返回 false，由调用方按固定等待兜底（不影响原有流程）。
WaitGeminiInputReady(hwnd, timeoutMs := 1200)
{
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline)
    {
        try {
            WinGetPos(&wx, &wy, &ww, &wh, "ahk_id " . hwnd)
            root := UIA.ElementFromHandle(hwnd)
            for e in root.FindElements({Type:"Edit"})
            {
                try {
                    if (!e.IsKeyboardFocusable || e.IsOffscreen)
                        continue
                    loc := e.Location
                    if (loc.w < 50 || loc.h < 10)
                        continue
                    if (loc.h > wh * 0.4 || loc.y + loc.h < wy + wh / 2)
                        continue
                    if (Trim(e.Value) != "")
                        return true
                }
            }
        }
        Sleep(60)
    }
    return false
}

; F2 自动寻找并切换 Gemini 窗口的显示/隐藏（最小化/激活）
; 当从外部切换到 Gemini 时，会自动抓取当前鼠标下的单词和句子并粘贴到输入框中
F2::
{
    global GeminiAutoHwnd
    GeminiHwnd := GetGeminiWindow()
    if (!GeminiHwnd)
    {
        ; 如果没找到空标题的，也可以尝试找找名字里带 Gemini 的
        if WinExist("Gemini ahk_exe chrome.exe")
            GeminiHwnd := WinGetID("Gemini ahk_exe chrome.exe")
        else
        {
            ; 未检测到 Gemini 窗口，自动通过 Alt+G 打开并 Pop-out chat
            if WinExist("ahk_exe chrome.exe")
            {
                try {
                    WinActivate("ahk_exe chrome.exe")
                    if !WinWaitActive("ahk_exe chrome.exe", , 2)
                        return
                    Sleep(300)
                    Send("!g")  ; Alt+G 打开 Gemini 侧边栏
                }
                catch TargetError {
                    ; 忽略 Chrome 不存在的错误
                }
            }
            else
            {
                MsgBox("未检测到 Chrome 浏览器，请先打开 Chrome！", "提示", "T3")
            }
            return
        }
    }

    ; 判断窗口是否可见（WS_VISIBLE = 0x10000000）
    try {
        DetectHiddenWindows(true)
        isVisible := (WinGetStyle("ahk_id " . GeminiHwnd) & 0x10000000)
        DetectHiddenWindows(false)
    }
    catch TargetError {
        GeminiAutoHwnd := 0
        return
    }

    if (isVisible)
    {
        ; 只要窗口在屏幕上（不管是不是活动窗口），按 F2 一律直接隐藏（从任务栏也消失）
        try {
            WinHide("ahk_id " . GeminiHwnd)
        }
        catch TargetError {
            GeminiAutoHwnd := 0
        }
    }
    else
    {
        ; 如果窗口当前被隐藏了，则：
        ; 1. 先抓取当前鼠标下的词句（必须在激活窗口前抓取，否则会失去原界面的焦点）
        word := ""
        line := ""
        hasWord := GetWordAndLineAtMouse(&word, &line)
        
        ; 2. 恢复并激活 Gemini 窗口
        try {
            DetectHiddenWindows(true)
            WinShow("ahk_id " . GeminiHwnd)
            DetectHiddenWindows(false)
            EnsureWindowOnScreen(GeminiHwnd)
            Sleep(150)
            WinActivate("ahk_id " . GeminiHwnd)
        }
        catch TargetError {
            GeminiAutoHwnd := 0
            return
        }
        
        ; 3. 如果成功抓取到词句，则将其处理干净（过滤表情、对象占位符，且将所有换行和连续空格压缩为单行单空格）
        if (hasWord)
        {
            word := Trim(RegExReplace(StripEmoji(word), "[\r\n\s]+", " "))
            line := Trim(RegExReplace(StripEmoji(line), "[\r\n\s]+", " "))
            
            ; 缓存为当前生词并预生成 TTS 语音，供右键朗读使用
            global WL_CurrentWord, WL_CurrentContext
            WL_CurrentWord := word
            WL_CurrentContext := line
            WL_PregenTts(word)
            
            textToSend := "单词: " . word . "`n句子: " . line
            
            ClipSaved := ClipboardAll()
            A_Clipboard := textToSend
            if !ClipWait(2)
            {
                A_Clipboard := ClipSaved
                return
            }
            
            ; 等待窗口激活后执行清除并粘贴
            try {
                if WinWaitActive("ahk_id " . GeminiHwnd, , 3)
                {
                    ; 保存鼠标原位置，操作完成后恢复
                    MouseGetPos(&origX, &origY)

                    ; 优先用 UIA 定位输入框：固定坐标偏移在 RDP 会话下会因
                    ; DPI/分辨率改变而点偏，导致 ^a/^v 落到错误的元素上
                    if (!FocusGeminiInput(GeminiHwnd))
                    {
                        ; UIA 不可用时回退原方案：点击底部中央上移 85px 处
                        Sleep(300)
                        WinGetPos(&gx, &gy, &gw, &gh, "ahk_id " . GeminiHwnd)
                        CoordMode("Mouse", "Screen")
                        Click(gx + (gw // 2), gy + gh - 85)
                        Sleep(150)
                    }
                    try {
                        Suspend(true)  ; 暂时挂起热键，防止 ^a 被 emacs 绑定拦截
                        Send("^a") ; 全选输入框内已有内容
                        Sleep(80)
                        Send("^v") ; 粘贴新内容覆盖
                        ; Gemini 输入框是 contenteditable，粘贴后需要等 React 重渲染、
                        ; 发送按钮从禁用变可用；等待过短会让 {Enter} 被当成插入换行而非发送。
                        ; 轮询确认输入框已有内容再回车，最多等约 1.2s（RDP/低配机更慢）
                        Sleep(150)
                        WaitGeminiInputReady(GeminiHwnd, 1200)
                        Send("{Enter}") ; 回车发送
                    } finally {
                        Suspend(false)  ; 无论是否抛错都必须恢复，否则所有 emacs 热键会永久失效
                    }

                    ; 恢复鼠标到原位置
                    MouseMove(origX, origY)
                }
            }
            catch TargetError {
                GeminiAutoHwnd := 0
            }
            finally {
                Sleep(100)
                A_Clipboard := ClipSaved  ; 无论成功失败都恢复剪贴板，避免污染用户剪贴板
            }
        }

    }
}

; 当 Google Gemini 窗口处于活动状态时，拦截鼠标右键，支持单击朗读单词，长按朗读句子
#HotIf (GeminiAutoHwnd && WinActive("ahk_id " . GeminiAutoHwnd))
RButton::
{
    WL_HandleRightClick()
}
#HotIf

; 共享模块（按依赖顺序）
#Include "shared/ollama_api.ahk"        ; 基础工具函数（StripEmoji 等）
#Include "shared/ollama_prompt_chat.ahk" ; GUI 事件处理（依赖 ollama_api）
#Include "shared/ollama_tts.ahk"         ; TTS 朗读

; 引入 Ollama 翻译/纠错模块
#Include "ollama_translate.ahk"

; 引入鼠标取词模块
#Include "word_lookup.ahk"