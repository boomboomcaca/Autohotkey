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

XButton1::Send("!{Left}")  ; 后退键（靠后那个）-> Alt+Left，即浏览器后退
; 前进键（靠前那个）-> 直接调用 Gemini 窗口开关，与 F2 同一份代码
; 不用 Send("{F2}")：AHK 默认忽略脚本自己发出的按键，那样触发不了本脚本的 F2 热键
; 2026-09-14 改（原为 Alt+Right 浏览器前进）
XButton2::GeminiToggle()
#Space::Send("{Ctrl down}{Space}{Ctrl up}")


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
        ; 枚举之后窗口可能已经关闭（如 Chrome 的临时气泡窗口），WinGet* 会抛 TargetError，
        ; 不捕获会让 F2 / Ctrl+Alt+Down 整个热键报错中断
        try {
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
        } catch {
            continue
        }

        ; 你的 Gemini 作为 Chrome PWA 运行时，系统获取到的窗口标题正好为空字符串 ""
        if (title == "")
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
    DetectHiddenWindows(true)
    for hwnd in WinGetList("ahk_class Chrome_WidgetWin_1 ahk_exe chrome.exe")
    {
        try {
            style := WinGetStyle(hwnd)
            if (style & 0x10000000)             ; 只捡隐藏的，可见的归 GetGeminiWindow 管
                continue
            if ((style & 0xC00000) != 0xC00000) ; 同样要求标准标题栏，滤掉 Chrome 的辅助窗口
                continue
            WinGetPos(, , &w, &h, hwnd)
            if (w < 200 || h < 200)
                continue
            if (WinGetTitle(hwnd) == "")
            {
                result := hwnd
                break
            }
        }
    }
    DetectHiddenWindows(false)
    if (result)
        GeminiAutoHwnd := result
    return result
}

; 自动打开 Gemini 独立窗口，成功返回窗口句柄，失败返回 0。
; 两步：点标签栏的 "Ask Gemini" 打开面板 → 点面板里的 "Pop-out chat" 弹成独立窗口。
;
; 为什么不用快捷键：Alt+G 在 Chrome 上并没有绑定（实测激活 Chrome 后发 !g，
; 6 秒内无任何窗口变化），原先那行 Send("!g") 是空操作，这正是一直要手动开窗口的原因。
; 这两个按钮都是标准 UIA Button（支持 Invoke），是目前唯一稳定的入口。
OpenGeminiWindow()
{
    ; 窗口其实已经存在、只是被隐藏了，直接复用，不要再去点按钮
    hidden := FindHiddenGeminiWindow()
    if (hidden)
        return hidden

    chromeHwnd := GetMainChromeWindow()
    if (!chromeHwnd)
    {
        ; Chrome 没启动：拉起来再等主窗口出现
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

    ; 1) 打开 Gemini 面板。该按钮是开关：面板关着时叫 "Ask Gemini"，
    ;    开着时变成 "Close Gemini in Chrome"。只在找得到 "Ask Gemini" 时点，
    ;    否则会把已经开着的面板关掉。找不到 = 面板已开，直接进入第 2 步。
    try {
        UIA.ElementFromHandle(chromeHwnd).FindElement({Name: "Ask Gemini", Type: "Button"}).Invoke()
    }

    ; 2) 弹成独立窗口。面板要渲染一会儿才会出现 "Pop-out chat"（实测约 1 秒，
    ;    Chrome 冷启动时更久），所以轮询；每轮重新取一次 UIA 根元素，
    ;    避免拿到面板出现之前的旧树。
    deadline := A_TickCount + 8000
    while (A_TickCount < deadline)
    {
        try {
            UIA.ElementFromHandle(chromeHwnd).FindElement({Name: "Pop-out chat", Type: "Button"}).Invoke()
            break
        }
        Sleep(200)
    }

    ; 3) 等独立窗口出现（实测点完约 100~500ms 就能被 GetGeminiWindow 找到）
    deadline := A_TickCount + 5000
    while (A_TickCount < deadline)
    {
        hwnd := GetGeminiWindow()
        if (hwnd)
            return hwnd
        Sleep(200)
    }

    ; 走到这里说明 Chrome 的 Gemini 入口和预期不一样（改版、换了界面语言等）。
    ; 不提示的话 F2 就是按下去毫无反应，没法判断是脚本坏了还是没按到。
    MsgBox("未能自动打开 Gemini 窗口，请手动打开一次。", "提示", "T3")
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
; 在 Gemini 窗口里筛出"像输入框"的可聚焦 Edit 元素，按 UIA 枚举顺序返回数组
; （每项为 {el, loc}），UIA 不可用或一个都没有时返回空数组。
;
; 只负责筛选、不做取舍：两个调用方要的不一样——FocusGeminiInput 取最靠下的那个去点击，
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
        ; 整段包在 try 里：与重构前一致，SetFocus/Click 抛异常时不往外冒，
        ; 而是落到下面的 Sleep 后重试
        try {
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
        ; 任意一个候选有文本即认为就绪（不像 FocusGeminiInput 那样只认最靠下的），
        ; 与重构前逐个元素判断的行为一致
        for c in GetGeminiInputCandidates(hwnd)
        {
            try {
                if (Trim(c.el.Value) != "")
                    return true
            }
        }
        Sleep(60)
    }
    return false
}

; F2 自动寻找并切换 Gemini 窗口的显示/隐藏（最小化/激活）
; 当从外部切换到 Gemini 时，会自动抓取当前鼠标下的单词和句子并粘贴到输入框中
F2::GeminiToggle()

GeminiToggle()
{
    global GeminiAutoHwnd
    GeminiHwnd := GetGeminiWindow()
    justOpened := false
    word := ""
    line := ""
    hasWord := false
    if (!GeminiHwnd)
    {
        ; 如果没找到空标题的，也可以尝试找找名字里带 Gemini 的
        if WinExist("Gemini ahk_exe chrome.exe")
            GeminiHwnd := WinGetID("Gemini ahk_exe chrome.exe")
        else
        {
            ; 未检测到 Gemini 窗口：先抓词，再自动打开（Chrome 没启动时会一并拉起）。
            ; 抓词必须在 OpenGeminiWindow 之前——它会激活 Chrome，
            ; 之后鼠标底下就是 Chrome 的内容，取到的词不再是原界面上的那个。
            hasWord := GetWordAndLineAtMouse(&word, &line)
            GeminiHwnd := OpenGeminiWindow()
            if (!GeminiHwnd)
                return
            justOpened := true
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

    ; justOpened 时窗口刚被弹出来，本来就是可见的；
    ; 不排除掉就会立刻走进隐藏分支，表现为"按 F2 闪一下又没了"
    if (isVisible && !justOpened)
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
        ;    justOpened 的情况已经在打开窗口之前抓过了，这里不能再抓一次
        if (!justOpened)
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

; 引入键盘控制鼠标模块（按住 Alt + 方向键移动，无模式切换）
; 与 mousemaster 抢同样的按键，两者只能启用其一
#Include "mousekeys.ahk"