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
; Alt+H 删除左边一个单词，与 Alt+D（删右边）对称。
; mousekeys 曾短暂占用过 H（vim 的 hjkl 布局），改成倒 T 形 ijkl 后这个键空了出来。
!h::
{ ; V1toV2: Added opening brace for [!h]
global ; V1toV2: Made function global
  if is_target()
    Send(A_ThisHotkey)
  else
    delete_backward_word()
return
} ; V1toV2: Added closing brace for [!h]
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
; Ctrl+Y = Emacs 的 yank（粘贴），补齐 kill / kill-ring-save / yank 这一组。
; 代价：Windows 上 Ctrl+Y 普遍是【重做】，绑了之后全局失去它
; （Office、资源管理器、多数编辑器；重做还可用 Ctrl+Shift+Z，但并非处处支持）。
; 上游当初留了 yank() 函数却没接热键，多半就是不想付这个代价——
; 毕竟原生 Ctrl+V 一直没被动过，粘贴本来就能用。这里是明知代价后的选择。
; is_target() 列出的应用照例原样透传。
^y::
{
global
  If is_target()
    Send(A_ThisHotkey)
  Else
    yank()
Return
}

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

XButton1::Send("!{Left}")   ; 后退键（靠后那个）-> Alt+Left，即浏览器后退
XButton2::Send("!{Right}")  ; 前进键（靠前那个）-> Alt+Right，即浏览器前进
; 注：2026-09-14 曾把前进键改成 Gemini 窗口开关（XButton2::GeminiToggle()），
;     用着不顺手，同月 19 日改回浏览器前进。Gemini 开关仍在 F2 上。
#Space::Send("{Ctrl down}{Space}{Ctrl up}")

; 雷蛇侧键 6（映射为 Right Ctrl + T）：单按保持原 Right Ctrl + T，双击触发 F5
global g_RCtrlTTapState := {isDown: false, downTick: 0, lastTapTick: 0, lastDoubleTick: 0, pendingSingle: false, snapX: 0, snapY: 0, snapWin: 0}

$>^t::HandleRCtrlTDown()

#HotIf g_RCtrlTTapState.isDown
~*t up::HandleRCtrlTUp()
#HotIf

HandleRCtrlTDown()
{
    global g_RCtrlTTapState
    action := F2_OnKeyDown(g_RCtrlTTapState, A_TickCount)
    if (action = "double") {
        SetTimer(RunRCtrlTSingleTap, 0)
        Send("{RCtrl up}{F5}")
    } else if (action = "wait_single") {
        SetTimer(RunRCtrlTSingleTap, -350)
    }
}

HandleRCtrlTUp()
{
    global g_RCtrlTTapState
    g_RCtrlTTapState.isDown := false
}

RunRCtrlTSingleTap()
{
    global g_RCtrlTTapState
    if (!g_RCtrlTTapState.pendingSingle)
        return
    g_RCtrlTTapState.pendingSingle := false
    Send("{RCtrl down}t{RCtrl up}")
}


;居中显示
^!down::
{ ; V1toV2: Added opening brace for [^!down]
  ; 注意：这里不用 global 模式——否则块内所有临时变量（GeminiHwnd/Width/mi 等）都会变成全局变量，
  ; 与 F2 里的局部 GeminiHwnd 同名（#Warn LocalSameAsGlobal 会报警）
  ;
  ; 抑制键盘自动重复：按住不放时本热键会一次不落地连发（实测按住约 300ms 触发 9 次，
  ; 窗口被反复居中）。尤其是按住 Alt+↓ 用键盘移鼠标时再补按 Ctrl，就会落进本热键
  ; （实测 mousekeys 的 *!Down 会自动让位给它）。
  ; 这里不能用 KeyWait("Down") 等松开：热键已经把按键吃掉了，被吃掉的键不写入系统按键状态，
  ; 函数里读 GetKeyState("Down") 恒为 0（逻辑、物理都是），KeyWait 会立刻返回、拦不住。
  ; 自动重复间隔约 30ms，真人有意再按一次不会快于 400ms，用时间去抖最省事。
  static lastCenterTick := 0
  if (lastCenterTick && A_TickCount - lastCenterTick < 400)
    return
  lastCenterTick := A_TickCount

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

; 严格判断 hwnd 是否为独立的 Gemini 浮窗/PWA（绝不匹配 Chrome 主浏览器窗口）
IsGeminiWindowMatch(hwnd, requireVisible := true)
{
    if (!hwnd)
        return false
    prevDetect := DetectHiddenWindows(!requireVisible)
    try {
        if (WinGetClass(hwnd) != "Chrome_WidgetWin_1")
            return false
        style := WinGetStyle(hwnd)
        if (requireVisible && !(style & 0x10000000))
            return false
        if ((style & 0xC00000) != 0xC00000)
            return false
        title := WinGetTitle(hwnd)
        ; Chrome 主浏览器窗口标题始终含 "Google Chrome"，严格排除；
        ; 独立 Gemini 浮窗标题为 ""，独立 PWA 窗口标题精确为 "Gemini"
        if (InStr(title, "Google Chrome") || (title != "" && title != "Gemini"))
            return false
        if (WinGetMinMax(hwnd) != -1) {
            WinGetPos(, , &w, &h, hwnd)
            if (w < 200 || h < 200)
                return false
        }
        if (StrLower(WinGetProcessName(hwnd)) != "chrome.exe")
            return false
        return true
    } catch {
        return false
    } finally {
        DetectHiddenWindows(prevDetect)
    }
}

; 自动寻找 Gemini 窗口的函数
GetGeminiWindow()
{
    global GeminiAutoHwnd

    ; 如果之前找到过并且仍确认为独立 Gemini 窗口，就直接复用（避免隐藏后找不到）
    if (GeminiAutoHwnd)
    {
        if (IsGeminiWindowMatch(GeminiAutoHwnd, false))
            return GeminiAutoHwnd
        GeminiAutoHwnd := 0
    }

    for hwnd in WinGetList("ahk_class Chrome_WidgetWin_1 ahk_exe chrome.exe")
    {
        if (IsGeminiWindowMatch(hwnd, true))
        {
            GeminiAutoHwnd := hwnd
            return hwnd
        }
    }
    return 0
}

; 找 Chrome 的主浏览器窗口（用于在上面定位工具栏按钮）。
; 不能图省事用 WinGetID("ahk_exe chrome.exe")：它返回的是最靠前的匹配窗口，
; 实测会拿到 708x53 的气泡/拖拽预览窗口，UIA 在那上面一个按钮也找不到。
GetMainChromeWindow()
{
    for hwnd in WinGetList("ahk_class Chrome_WidgetWin_1 ahk_exe chrome.exe")
    {
        try {
            style := WinGetStyle(hwnd)
            if !(style & 0x10000000)          ; 必须可见
                continue
            if ((style & 0xC00000) != 0xC00000) ; 必须带标准标题栏
                continue
            WinGetPos(, , &w, &h, hwnd)
            ; 带标题的大窗口才是浏览器主窗口；
            ; 空标题的是已经弹出的 Gemini 窗口，必须排除，否则会在它上面找按钮
            if (w > 600 && h > 400 && WinGetTitle(hwnd) != "")
                return hwnd
        }
    }
    return 0
}

; 找"已经弹出、但当前被隐藏"的 Gemini 窗口。
; GetGeminiWindow 只枚举可见窗口，隐藏中的窗口全靠 GeminiAutoHwnd 缓存找回；
; 一旦按 F2 隐藏后脚本重启（缓存清零），那个窗口就谁也找不到了：
; 此时 "Ask Gemini"（面板已开）和 "Pop-out chat"（已经是弹出状态）两个按钮都不存在，
; 走点按钮的流程只会白等十秒且毫无反应。所以自动打开前先认领它。
FindHiddenGeminiWindow()
{
    global GeminiAutoHwnd
    result := 0
    prevDetect := DetectHiddenWindows(true)
    try {
        for hwnd in WinGetList("ahk_class Chrome_WidgetWin_1 ahk_exe chrome.exe")
        {
            try {
                if (WinGetStyle(hwnd) & 0x10000000) ; 只捡隐藏的，可见的归 GetGeminiWindow 管
                    continue
                if (IsGeminiWindowMatch(hwnd, false))
                {
                    result := hwnd
                    break
                }
            }
        }
    } finally {
        DetectHiddenWindows(prevDetect)
    }
    if (result)
        GeminiAutoHwnd := result
    return result
}

; 点标签栏右侧那个 Gemini 开关按钮。wantClose=false 找"打开"那一面，true 找"关闭"那一面。
; 点到了返回 true。
;
; 不按精确名字匹配是有教训的：Chrome 会改这个按钮的无障碍名称，而且会改回去。
; 2026-09-14 一天之内实测到三种："Ask Gemini" → "Open Gemini in Chrome" → 又变回
; "Ask Gemini"；位置尺寸一点没变（标签栏右侧 280x90）。写死名字的话，Chrome 一更新
; 脚本就瞎，表现为按 F2 弹"未能自动打开 Gemini 窗口，请手动打开一次"。
;
; 改用特征匹配：名字含 Gemini、位于窗口顶部 150px 内。后一条用来排除面板/网页里同样
; 含 Gemini 的按钮（实测有 "Skip to the latest Gemini result" 在 y≈212、
; 收藏栏上还有个 "Unnamed bookmark for https://gemini.google.com/app" 在 y≈208）。
; 开与关是同一个按钮，靠名字里有没有 Close 区分。
InvokeGeminiToolbarButton(chromeHwnd, wantClose := false)
{
    want := wantClose ? 1 : 0
    try {
        WinGetPos(, &wy, , , "ahk_id " . chromeHwnd)
        for el in UIA.ElementFromHandle(chromeHwnd).FindAll({Type: "Button"})
        {
            try {
                nm := el.Name
                if (nm = "" || !InStr(nm, "Gemini"))
                    continue
                if ((InStr(nm, "Close") ? 1 : 0) != want)
                    continue
                if (el.GetPos("screen").y - wy > 150)   ; 只认顶部工具栏那一条
                    continue
                el.Invoke()
                return true
            }
        }
    }
    return false
}

; 点面板里的"弹出成独立窗口"按钮。点到了返回 true。
;
; Chrome 改版记录：
; 1) 老版 Chrome：面板顶部直接有一个 "Pop-out chat" 或 "Pop-out" 独立按钮。
; 2) 新版 Chrome：Google 将 "Pop-out chat" 收进了侧边栏右上角的 "More options"（更多选项）
;    菜单中。因此必须先触发 "More options" 按钮打开下拉菜单，再在弹出的菜单中
;    点击 "Pop-out chat"（MenuItem）项。
InvokeGeminiPopOutButton(chromeHwnd)
{
    try {
        root := UIA.ElementFromHandle(chromeHwnd)
        WinGetPos(&wx, &wy, &ww, &wh, "ahk_id " . chromeHwnd)

        ; 1) 老版 Chrome：面板上直接有 Pop-out 按钮
        for el in root.FindAll({Type: "Button"})
        {
            try {
                nm := el.Name
                if (nm != "" && (InStr(nm, "Pop-out") || InStr(nm, "Pop out") || InStr(nm, "弹出"))) {
                    el.Invoke()
                    return true
                }
            }
        }

        ; 2) 新版 Chrome：在侧边栏右上角的 "More options" 下拉菜单里
        moreBtn := 0
        for b in root.FindAll({Type: "Button"})
        {
            try {
                pos := b.GetPos("screen")
                rely := pos.y - wy
                ; 位于侧边栏顶部区域（y 在 150~400，x 靠右侧）
                if (rely >= 150 && rely <= 400 && pos.x >= wx + ww - 1000) {
                    if (b.Name = "More options" || InStr(b.Name, "options") || InStr(b.Name, "选项")) {
                        moreBtn := b
                        break
                    }
                }
            }
        }

        if (moreBtn)
        {
            moreBtn.Invoke()
            Sleep(300)

            for h in WinGetList("ahk_class Chrome_WidgetWin_1 ahk_exe chrome.exe")
            {
                try {
                    r := UIA.ElementFromHandle(h)
                    for el in r.FindAll()
                    {
                        try {
                            nm := el.Name
                            if (nm != "" && (InStr(nm, "Pop-out") || InStr(nm, "Pop out") || InStr(nm, "弹出"))) {
                                el.Invoke()
                                return true
                            }
                        }
                    }
                }
            }
        }
    }
    return false
}

; 自动打开 Gemini 独立窗口，成功返回窗口句柄，失败返回 0。
;
; 为什么不用快捷键：Alt+G 在 Chrome 上并没有绑定（实测激活 Chrome 后发 !g，
; 6 秒内无任何窗口变化），原先那行 Send("!g") 是空操作，这正是一直要手动开窗口的原因。
; 工具栏那个按钮是标准 UIA Button（支持 Invoke），是目前唯一稳定的入口。
;
; 流程原本写死成"点工具栏开面板 → 点面板里的弹出按钮 → 等窗口"三步，新版 Chrome 把
; 中间一步取消了（直接出独立窗口，约 2~3 秒），于是第二步雷打不动空转满 8 秒，
; 再加第三步的 5 秒，最后照样报错。现在改成：点一下开关，然后在一个循环里边等窗口
; 边顺手试弹出按钮，谁先成谁算——新老 Chrome 都能走通，也不再白等。
OpenGeminiWindow()
{
    ; 窗口其实已经存在、只是被隐藏了，直接复用，不要再去点按钮
    hidden := FindHiddenGeminiWindow()
    if (hidden)
        return hidden

    chromeHwnd := GetMainChromeWindow()
    if (!chromeHwnd)
    {
        ; 没有浏览器主窗口。可能 Chrome 压根没启动，也可能它还活着但只剩下
        ; 弹出的 Gemini 窗口（把主窗口关掉、Gemini 窗口留着，实测就是这个状态）。
        ; 两种情况都是再 Run 一次 chrome.exe：没启动就拉起来，已经在跑就开个新窗口。
        try {
            Run("chrome.exe")
        } catch {
            MsgBox("未找到 Chrome，无法自动打开 Gemini。", "提示", "T3")
            return 0
        }
        deadline := A_TickCount + 20000
        while (A_TickCount < deadline)
        {
            Sleep(300)
            chromeHwnd := GetMainChromeWindow()
            if (chromeHwnd)
                break
        }
        if (!chromeHwnd)
            return 0
        Sleep(1000)  ; 冷启动后等标签栏渲染完，否则 UIA 还枚举不到工具栏按钮
    }

    try WinActivate("ahk_id " . chromeHwnd)
    if !WinWaitActive("ahk_id " . chromeHwnd, , 3)
        return 0
    Sleep(200)

    ; 1) 让开关复位。走到这里说明上面既没找到可见窗口也没找到隐藏窗口，
    ;    可按钮要是写着 "Close Gemini in Chrome"，就说明 Chrome 自己还认为开着
    ;    （用户把弹出窗口点叉关掉时就会这样）。此时直接点它只会再关一次，
    ;    什么都不会出来，所以先点一下让它回到"关"，再点一下打开。
    if (InvokeGeminiToolbarButton(chromeHwnd, true))
        Sleep(800)

    why := ""
    if (!InvokeGeminiToolbarButton(chromeHwnd, false))
        why := "标签栏上找不到 Gemini 按钮"

    ; 2) 等侧边栏展开并弹出独立窗口。
    ;    老版 Chrome 在面板里直接有 Pop-out；新版 Chrome 将 Pop-out chat 收进了 More options 菜单。
    ;    边轮询等窗口，边尝试通过 InvokeGeminiPopOutButton 触发弹出。
    deadline := A_TickCount + 12000
    nextPop := A_TickCount + 600
    while (A_TickCount < deadline)
    {
        hwnd := GetGeminiWindow()
        if (hwnd)
            return hwnd
        if (A_TickCount >= nextPop)
        {
            InvokeGeminiPopOutButton(chromeHwnd)
            nextPop := A_TickCount + 700
        }
        Sleep(100)
    }
    if (why = "")
        why := "按钮点到了，但 12 秒内没出现独立窗口"

    ; 走到这里说明 Chrome 的 Gemini 入口和预期不一样（改版、换了界面语言等）。
    ; 不提示的话 F2 就是按下去毫无反应，没法判断是脚本坏了还是没按到。
    ; 把具体卡在哪一步也带上，下次一看就知道要去查什么。
    MsgBox("未能自动打开 Gemini 窗口，请手动打开一次。`n原因：" . why, "提示", "T5")
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

; 通过 UIA 定位 Gemini 输入框（窗口底部最靠下的可聚焦编辑框）并无感聚焦（不移动物理鼠标）。
; 坐标来自元素实际位置，RDP 会话下 DPI/分辨率变化时依然命中；
; 轮询等待输入框出现，覆盖 RDP 位图远传导致 Chrome 渲染变慢的情况
; 在 Gemini 窗口里筛出"像输入框"的可聚焦 Edit 元素，按 UIA 枚举顺序返回数组
; （每项为 {el, loc}），UIA 不可用或一个都没有时返回空数组。
;
; 只负责筛选、不做取舍：两个调用方要的不一样——FocusGeminiInput 取最靠下的那个无感聚焦，
; WaitGeminiInputReady 只要任意一个有文本就算就绪。
; 筛选条件此前在这两处各抄了一份，其中一份还把"为什么这样筛"的注释弄丢了，
; 以后 Gemini 改版要调阈值，很容易只改一处。
GetGeminiInputCandidates(hwnd)
{
    out := []
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
                ; 只接受"输入框形状"的候选：高度有限且贴近窗口底部。
                ; Gemini 的 Canvas/文档面板也是可聚焦 Edit，但接近全窗口高，
                ; 若误选中它，后面的 ^a/^v 会覆盖用户文档内容
                if (loc.h > wh * 0.4 || loc.y + loc.h < wy + wh / 2)
                    continue
                out.Push({el: e, loc: loc})
            }
        }
    }
    return out
}

FocusGeminiInput(hwnd, timeoutMs := 2500)
{
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline)
    {
        ; 取最靠下的候选：Gemini 的输入框永远在窗口底部
        best := ""
        bestLoc := ""
        bestBottom := -2147483648
        for c in GetGeminiInputCandidates(hwnd)
        {
            if (c.loc.y + c.loc.h > bestBottom)
            {
                bestBottom := c.loc.y + c.loc.h
                best := c.el
                bestLoc := c.loc
            }
        }
        ; 优先用纯 UIA 接口聚焦（0 鼠标移动）；仅在均未获焦时才用后台消息 ControlClick 兜底
        try {
            if (best)
            {
                try if (best.HasKeyboardFocus)
                    return true

                try best.SetFocus()
                Loop 4 {
                    Sleep(20)
                    try if (best.HasKeyboardFocus)
                        return true
                }

                try best.LegacyIAccessiblePattern.Select(1) ; SELFLAG_TAKEFOCUS = 1
                try best.SetFocus()
                Loop 3 {
                    Sleep(20)
                    try if (best.HasKeyboardFocus)
                        return true
                }

                WinGetPos(&wx, &wy, , , "ahk_id " . hwnd)
                relX := (bestLoc.x + (bestLoc.w // 2)) - wx
                relY := (bestLoc.y + (bestLoc.h // 2)) - wy
                ControlClick("x" . relX . " y" . relY, "ahk_id " . hwnd, , "Left", 1, "NA")
                Loop 3 {
                    Sleep(20)
                    try if (best.HasKeyboardFocus)
                        break
                }
                return true
            }
        }
        Sleep(100)
    }
    return false
}

; 优先通过 UIA 无感聚焦底部输入框，不可用时回退后台消息点击（均不移动物理鼠标）
ActivateGeminiInputBox(hwnd)
{
    if (!FocusGeminiInput(hwnd))
    {
        ; UIA 不可用时回退后台消息点击底部中央上移 85px 处（不移动物理鼠标）
        Sleep(200)
        WinGetPos(, , &gw, &gh, "ahk_id " . hwnd)
        ControlClick("x" . (gw // 2) . " y" . Max(10, gh - 85), "ahk_id " . hwnd, , "Left", 1, "NA")
        Sleep(120)
    }
}

; 轮询等待 Gemini 输入框注册到粘贴内容后再回车。
; Gemini 是 contenteditable + React，粘贴到"发送按钮可用"之间有延迟；
; 通过 UIA 读取底部输入框的文本值，确认已出现预期内容即认为可安全发送。
; UIA 不可用时返回 false，由调用方按固定等待兜底（不影响原有流程）。
WaitGeminiInputReady(hwnd, timeoutMs := 1200, expectedSub := "")
{
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline)
    {
        ; 任意一个候选已包含目标文本即认为就绪（不像 FocusGeminiInput 那样只认最靠下的）
        for c in GetGeminiInputCandidates(hwnd)
        {
            try {
                val := Trim(c.el.Value)
                if (val != "" && (expectedSub = "" || InStr(val, expectedSub)))
                    return true
            }
        }
        Sleep(25)
    }
    return false
}

; 写入临时文本到剪贴板并附带 Win10/11 屏蔽历史与云同步标记，防止污染 Win+V 剪贴板历史
SetClipboardTextNoHistory(text)
{
    static fmtExclude := DllCall("RegisterClipboardFormatW", "Str", "ExcludeClipboardContentFromMonitorProcessing", "UInt")
    static fmtNoHist := DllCall("RegisterClipboardFormatW", "Str", "CanIncludeInClipboardHistory", "UInt")
    static fmtNoCloud := DllCall("RegisterClipboardFormatW", "Str", "CanUploadToCloudClipboard", "UInt")

    opened := false
    Loop 5 {
        if DllCall("OpenClipboard", "Ptr", A_ScriptHwnd) {
            opened := true
            break
        }
        Sleep(10)
    }
    if (!opened) {
        A_Clipboard := text
        return ClipWait(1)
    }

    ok := false
    try {
        if !DllCall("EmptyClipboard")
            return false

        bytes := StrPut(text, "UTF-16")
        hMem := DllCall("GlobalAlloc", "UInt", 0x0042, "UPtr", bytes, "Ptr") ; GMEM_MOVEABLE | GMEM_ZEROINIT
        if (!hMem)
            return false
        pMem := DllCall("GlobalLock", "Ptr", hMem, "Ptr")
        if (!pMem) {
            DllCall("GlobalFree", "Ptr", hMem)
            return false
        }
        StrPut(text, pMem, "UTF-16")
        DllCall("GlobalUnlock", "Ptr", hMem)

        if !DllCall("SetClipboardData", "UInt", 13, "Ptr", hMem, "Ptr") { ; CF_UNICODETEXT = 13
            DllCall("GlobalFree", "Ptr", hMem)
            return false
        }
        ok := true

        for fmt in [fmtExclude, fmtNoHist, fmtNoCloud] {
            if (!fmt)
                continue
            hFlag := DllCall("GlobalAlloc", "UInt", 0x0042, "UPtr", 4, "Ptr")
            if (!hFlag)
                continue
            ; GMEM_ZEROINIT 已将 4 字节 DWORD 置 0，表示禁止记入历史与云剪贴板
            if !DllCall("SetClipboardData", "UInt", fmt, "Ptr", hFlag, "Ptr")
                DllCall("GlobalFree", "Ptr", hFlag)
        }
    } finally {
        DllCall("CloseClipboard")
    }

    if (!ok) {
        A_Clipboard := text
        return ClipWait(1)
    }
    return true
}

; F2 单击切换 Gemini 显示/隐藏并在打开时抓词发送；双击 F2 仅打开/激活 Gemini 窗口不抓词。
; 被热键吞掉的键在部分环境下 GetKeyState 恒为 0，不能用 KeyWait 判松开，故用 F2 / F2 up 配合定时器判定。
global g_F2TapState := {isDown: false, downTick: 0, lastTapTick: 0, lastDoubleTick: 0, pendingSingle: false, snapX: 0, snapY: 0, snapWin: 0}

F2::HandleF2Down()
F2 up::HandleF2Up()

HandleF2Down()
{
    global g_F2TapState
    CoordMode("Mouse", "Screen")
    MouseGetPos(&mx, &my, &mwin)
    action := F2_OnKeyDown(g_F2TapState, A_TickCount, mx, my, mwin)
    if (action = "double") {
        SetTimer(RunF2SingleTap, 0)
        GeminiToggle(true)
    } else if (action = "wait_single") {
        SetTimer(RunF2SingleTap, -350)
    }
}

HandleF2Up()
{
    global g_F2TapState
    g_F2TapState.isDown := false
}

RunF2SingleTap()
{
    global g_F2TapState
    if (!g_F2TapState.pendingSingle)
        return
    g_F2TapState.pendingSingle := false
    GeminiToggle(false, g_F2TapState.snapX, g_F2TapState.snapY, g_F2TapState.snapWin)
}

GeminiToggle(onlyOpen := false, snapX?, snapY?, snapWin?)
{
    global GeminiAutoHwnd
    GeminiHwnd := GetGeminiWindow()
    justOpened := false
    word := ""
    line := ""
    hasWord := false
    if (!GeminiHwnd)
    {
        ; 优先认领"已弹出但被隐藏"的真 Gemini 窗口，禁止再用标题子串匹配主 Chrome 窗口
        GeminiHwnd := FindHiddenGeminiWindow()
        if (GeminiHwnd)
        {
            GeminiAutoHwnd := GeminiHwnd
        }
        else
        {
            ; 未检测到 Gemini 窗口：单按先抓词（双击只开窗时跳过），再自动打开。
            ; 抓词必须在 OpenGeminiWindow 之前——它会激活 Chrome，
            ; 之后鼠标底下就是 Chrome 的内容，取到的词不再是原界面上的那个。
            if (!onlyOpen)
                hasWord := GetWordAndLineAtMouse(&word, &line, snapX?, snapY?, snapWin?)
            GeminiHwnd := OpenGeminiWindow()
            if (!GeminiHwnd)
                return
            GeminiAutoHwnd := GeminiHwnd
            justOpened := true
        }
    }
    else
    {
        GeminiAutoHwnd := GeminiHwnd
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

    ; 只有单按且窗口在屏幕上（非刚弹出）时才隐藏；双击 F2 一律走显示/激活分支
    if (!onlyOpen && isVisible && !justOpened)
    {
        if (!IsGeminiWindowMatch(GeminiHwnd, true))
        {
            GeminiAutoHwnd := 0
            return
        }
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
        ; 如果窗口当前被隐藏（或双击 F2 强制打开/激活）：
        ; 1. 单按时先抓取当前鼠标下的词句（必须在激活窗口前抓取，否则会失去原界面的焦点）
        ;    justOpened 的情况已经在打开窗口之前抓过了，这里不能再抓一次
        if (!onlyOpen && !justOpened)
            hasWord := GetWordAndLineAtMouse(&word, &line, snapX?, snapY?, snapWin?)

        ; 2. 恢复并激活 Gemini 窗口
        try {
            DetectHiddenWindows(true)
            if (!isVisible)
                WinShow("ahk_id " . GeminiHwnd)
            if (WinGetMinMax("ahk_id " . GeminiHwnd) == -1)
                WinRestore("ahk_id " . GeminiHwnd)
            DetectHiddenWindows(false)
            EnsureWindowOnScreen(GeminiHwnd)
            WinActivate("ahk_id " . GeminiHwnd)
            if !WinWaitActive("ahk_id " . GeminiHwnd, , 0.15)
                WinActivate("ahk_id " . GeminiHwnd)
        }
        catch TargetError {
            GeminiAutoHwnd := 0
            return
        }

        if (onlyOpen || !hasWord)
        {
            try {
                if WinWaitActive("ahk_id " . GeminiHwnd, , 3)
                    ActivateGeminiInputBox(GeminiHwnd)
            }
            catch TargetError {
                GeminiAutoHwnd := 0
            }
            return
        }

        ; 3. 单按且成功抓取到词句：将其处理干净（过滤表情、对象占位符，且将所有换行和连续空格压缩为单行单空格）
        word := Trim(RegExReplace(StripEmoji(word), "[\r\n\s]+", " "))
        line := Trim(RegExReplace(StripEmoji(line), "[\r\n\s]+", " "))

        ; 缓存为当前生词并预生成 TTS 语音，供右键朗读使用
        global WL_CurrentWord, WL_CurrentContext
        WL_CurrentWord := word
        WL_CurrentContext := line
        WL_PregenTts(word)

        textToSend := "单词: " . word . "`n句子: " . line

        ClipSaved := ClipboardAll()
        if !SetClipboardTextNoHistory(textToSend)
        {
            A_Clipboard := ClipSaved
            return
        }

        pasteConfirmed := false
        ; 等待窗口激活后无感聚焦输入框、清除并粘贴
        try {
            if WinWaitActive("ahk_id " . GeminiHwnd, , 3)
            {
                ActivateGeminiInputBox(GeminiHwnd)
                try {
                    Suspend(true)  ; 暂时挂起热键，防止 ^a 被 emacs 绑定拦截
                    Send("^a") ; 全选输入框内已有内容
                    Sleep(30)
                    Send("^v") ; 粘贴新内容覆盖
                    ; Gemini 输入框是 contenteditable，粘贴后需要等 React 重渲染、
                    ; 发送按钮从禁用变可用；轮询确认目标词已写入输入框后再回车
                    Sleep(30)
                    pasteConfirmed := WaitGeminiInputReady(GeminiHwnd, 1200, word)
                    Sleep(pasteConfirmed ? 35 : 120)
                    Send("{Enter}") ; 回车发送
                } finally {
                    Suspend(false)  ; 无论是否抛错都必须恢复，否则所有 emacs 热键会永久失效
                }
            }
        }
        catch TargetError {
            GeminiAutoHwnd := 0
        }
        finally {
            Sleep(pasteConfirmed ? 20 : 100)
            A_Clipboard := ClipSaved  ; 无论成功失败都恢复剪贴板，避免污染用户剪贴板
        }

    }
}

; 判断指定句柄是否为可见的 Gemini 窗口（先走缓存快路径，失效或脚本刚重载时自动重新认领）
IsGeminiHwnd(hwnd)
{
    global GeminiAutoHwnd
    if (!hwnd)
        return false

    if (GeminiAutoHwnd && hwnd == GeminiAutoHwnd) {
        if (IsGeminiWindowMatch(hwnd, true))
            return true
        GeminiAutoHwnd := 0
        return false
    }

    if (IsGeminiWindowMatch(hwnd, true)) {
        GeminiAutoHwnd := hwnd
        return true
    }
    return false
}

; 右键仅在 Gemini 为当前活动窗口或鼠标正悬停于其上时拦截，兼顾置顶未激活与刚重载脚本场景
IsGeminiRButtonContext()
{
    activeHwnd := WinExist("A")
    if (activeHwnd && IsGeminiHwnd(activeHwnd))
        return true
    MouseGetPos(, , &hoverHwnd)
    return !!(hoverHwnd && hoverHwnd != activeHwnd && IsGeminiHwnd(hoverHwnd))
}

; 当 Google Gemini 窗口处于活动状态或鼠标位于其上方时，拦截鼠标右键，支持单击朗读单词，长按朗读句子
; 加 #SuspendExempt 避免 F2 粘贴期间 Suspend(true) 放行右键；吞掉 RButton up 防止 Chrome 弹出上下文菜单
#SuspendExempt
#HotIf IsGeminiRButtonContext()
RButton::
{
    WL_HandleRightClick()
}
RButton up::return
#HotIf
#SuspendExempt false

; 共享模块（按依赖顺序）
#Include "shared/ollama_api.ahk"        ; 基础工具函数（StripEmoji 等）
#Include "shared/ollama_prompt_chat.ahk" ; GUI 事件处理（依赖 ollama_api）
#Include "shared/ollama_tts.ahk"         ; TTS 朗读

; 引入 Ollama 翻译/纠错模块
#Include "ollama_translate.ahk"

; 引入鼠标取词模块
#Include "word_lookup.ahk"

; 引入键盘控制鼠标模块（按住 Alt + 方向键移动，无模式切换）
; 与 mousemaster 抢同样的按键，两者只能启用其一
#Include "mousekeys.ahk"