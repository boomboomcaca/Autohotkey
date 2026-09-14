;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;; 键盘控制鼠标：按住 Alt + 方向键
;;
;;   Alt + i j k l        移动光标（倒 T 形：i=上 j=左 k=下 l=右）
;;                        按住逐渐加速；前 0.2 秒温和便于微调，之后切入高速档奔袭
;;   Alt + Shift + ijkl   拖动：自动压住左键并移动；Alt 或 Shift 任一松开即放下。
;;                        中途补按 Shift 也能起拖（不必一开始就按着）
;;
;;   Alt + Enter          左键单击
;;   Alt + /              右键单击
;;   Alt + p / n          滚轮上 / 下（滚的是鼠标指针下面那块区域，
;;                        不必先让焦点落进去；按住连发）
;;
;; 无模式切换，按住 Alt 即用。
;;
;; 已知冲突（都是全局接管按键的必然结果，不是 bug）：
;;   1) 按住 Alt 时 i j k l p n 被本模块吃掉，打不出这六个字母。
;;      为此移除的 emacs 绑定：Alt+K / Alt+Shift+K（切标签页）、Alt+P / Alt+N（整页翻）。
;;      切标签仍可用原生 Ctrl+Tab / Ctrl+Shift+Tab。i / j / l 本来就是空的。
;;   2) Alt+Enter 被本模块吃掉，Chrome 地址栏"新标签打开"、Excel 单元格内换行、
;;      资源管理器"属性"等原生功能在全局失效。
;;   注：移动键用字母而非方向键，所以 Alt+←/→ 保持为浏览器的后退/前进；
;;      滚轮用 p/n 而非 PgUp/PgDn，Alt+PgUp/PgDn 也留给应用。
;;
;; 实现要点：方向键必须显式跟踪。热键会「吃掉」按键，被吃掉的键不会写入
;; 系统按键状态，GetKeyState 永远读不到，所以按下与松开各注册一个热键。
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

; ---------------- 可调参数 ----------------
global MK_ModAlt       := "Alt"   ; "Alt"=左右 Alt 都可；"LAlt" 仅左；"RAlt" 仅右

; 速度模型：加速度 = 阻力系数 ×（终端速度 − 当前速度），速度平滑收敛到终端速度。
; 两段式：同一条加速曲线没法既让短按精确、又让长按迅猛，所以分两档。
; 前 MK_BoostDelay 毫秒走低速档，专为短距离微调：
; 点一下 40ms ≈ 9px、80ms ≈ 24px；之后切到高速档奔袭，3840 宽的屏跨全屏约 0.9 秒。
; 注：低速档只影响头 200ms，中长距离几乎不受它影响
; （起步 80 与 300 相比，0.5s 处只差 11%、跨全屏只差 2%）。
; 嫌短按走得远/近 → 只动 MK_InitialSpeed / MK_Resistance；
; 嫌起飞晚 → 调小 MK_BoostDelay；嫌飞起来不够猛 → 调大 MK_TermBoost。
global MK_Resistance   := 2.0     ; 低速档阻力系数（1/秒）。调大=更快到顶更跟手；调小=更绵长
global MK_TermNormal   := 2200    ; 低速档终端速度（像素/秒）
global MK_BoostDelay   := 200     ; 连续移动超过这么久（毫秒）切入高速档
global MK_ResistBoost  := 4.0     ; 高速档阻力系数
global MK_TermBoost    := 8000    ; 高速档终端速度（像素/秒）
global MK_InitialSpeed := 150     ; 起步速度，决定「点一下走多远」
global MK_Tick         := 8       ; 定时器周期（毫秒）
global MK_WheelRepeat  := 70      ; 滚轮连发间隔（毫秒）
global MK_MaxHoldMs    := 6000    ; 单个方向键最长持续时间，超时即判定松开事件丢失

; ---------------- 运行期状态 ----------------
global MK_Vel        := 0.0
global MK_AccX       := 0.0       ; 亚像素累积，低速时才不会因取整而原地不动
global MK_AccY       := 0.0
global MK_Dir        := Map("U", false, "D", false, "L", false, "R", false)
global MK_DirSince   := Map("U", 0, "D", 0, "L", 0, "R", 0)
global MK_Dragging   := false     ; Shift+Alt+方向 触发的拖动是否进行中
global MK_BurstSince := 0         ; 本次连续移动的起始时刻，0 = 当前没在移动
global MK_WheelDir   := 0         ; 滚轮连发方向：+1 上 / -1 下 / 0 没在滚
global MK_WheelSince := 0         ; 本次滚轮连发的起始时刻，用于超时兜底
global MK_Masked     := false     ; 本次 Alt 按住期间是否已经发过菜单屏蔽键

; 判定"用户是不是按着 Alt"。必须用物理状态 "P"，也是本模块所有热键改用 #HotIf
; 自行判定、而不是写成 !Up / !Enter 的原因：
;
; MK_Click 点击前会临时松开 Alt（否则应用收到的是 Alt+点击，Chrome 里点链接会变成下载）。
; 松开之后逻辑 Alt = 0 而物理 Alt = 1。实测 AHK 不会从物理状态重新同步，
; 于是 !Up 之类带 Alt 修饰的热键立刻失配，方向键直接漏给应用。
; 以前靠"点完立刻把 Alt 按回去"来兜，但那一下会掐死点击刚打开的菜单——
; 收藏栏文件夹只闪一下高亮、下拉永远出不来。
; 改成物理判定后就不必再把 Alt 按回去，两个问题一起解决。
MK_AltHeld() {
    global MK_ModAlt
    if (MK_ModAlt = "LAlt")
        return GetKeyState("LAlt", "P")
    if (MK_ModAlt = "RAlt")
        return GetKeyState("RAlt", "P")
    return GetKeyState("LAlt", "P") || GetKeyState("RAlt", "P")
}

MK_Press(k) {
    global MK_Dir, MK_DirSince, MK_Tick, MK_Masked
    ; 本次 Alt 按住期间发一次菜单屏蔽键。
    ; i/j/k/l 被热键吃掉、MouseMove 又不算「按键」，应用在整段 Alt 期间一个输入都
    ; 没收到，用户松开 Alt 时 Windows 就判成激活菜单栏（应用的菜单条会亮起来）。
    ; Alt+p/n 没这毛病，是因为滚轮事件本身就把这个判定解掉了。
    ; 只发一次：键盘自动重复会每 30ms 左右反复触发本函数，每次都发就成了狂敲 Ctrl。
    if (!MK_Masked) {
        Send("{Blind}{" . A_MenuMaskKey . "}")
        MK_Masked := true
    }
    ; 拖动的判定不在这里，而在 MK_Step 每一拍做——见那边的注释
    MK_Dir[k] := true
    MK_DirSince[k] := A_TickCount
    SetTimer(MK_Step, MK_Tick)
}

MK_Release(k) {
    global MK_Dir
    MK_Dir[k] := false
}

MK_Stop() {
    global MK_Vel, MK_AccX, MK_AccY, MK_Dir, MK_Dragging, MK_BurstSince, MK_Masked
    SetTimer(MK_Step, 0)
    for k in ["U", "D", "L", "R"]
        MK_Dir[k] := false
    MK_BurstSince := 0
    MK_Masked := false   ; Alt 已经松开，下次按住要重新屏蔽一次
    ; 松开 Alt 视为本次拖动结束，在此放开左键
    if (MK_Dragging) {
        Click("Left Up")
        MK_Dragging := false
    }
    MK_Vel := 0.0
    MK_AccX := 0.0
    MK_AccY := 0.0
}

; 是否已进入高速档。判据是「本次连续移动持续了多久」，不是单个方向键按了多久——
; 中途改方向（比如先 ↑ 再补 →）不该把档位打回起步。
MK_Boosted() {
    global MK_BurstSince, MK_BoostDelay
    return MK_BurstSince && (A_TickCount - MK_BurstSince) >= MK_BoostDelay
}

MK_Step() {
    global MK_Vel, MK_AccX, MK_AccY, MK_Tick, MK_Dir, MK_DirSince
    global MK_Resistance, MK_InitialSpeed, MK_MaxHoldMs
    global MK_TermNormal, MK_TermBoost, MK_ResistBoost, MK_BurstSince
    global MK_Dragging   ; 少了这行，下面的赋值会写进局部变量，
                         ; 全局 MK_Dragging 恒为 false → 每 8ms 压一次左键且永不放开

    if (!MK_AltHeld()) {
        MK_Stop()
        return
    }

    ; 拖动判定放在每一拍，而不是方向键按下那一瞬间：
    ; 否则"先按住 Alt+方向键开始移动、中途再补按 Shift"就进不了拖动，
    ; 而这正是常见的用法——先把光标挪到起点附近，再按 Shift 开拖。
    ; 补按 Shift 的那一拍压下左键，本拍的移动发生在其后，所以不会漏掉起点。
    ;
    ; Alt 和 Shift 任一松开都结束拖动：Shift 在这里判（松手即放下），
    ; Alt 在 MK_Stop 里判（整个模块都停）。
    if (GetKeyState("Shift")) {
        if (!MK_Dragging) {
            Click("Left Down")
            MK_Dragging := true
        }
    } else if (MK_Dragging) {
        Click("Left Up")
        MK_Dragging := false
    }

    ; 保险：松开事件偶尔会丢失，导致光标一直跑。超过上限就强制视为已松开。
    for k, since in MK_DirSince
        if (MK_Dir[k] && since && (A_TickCount - since) > MK_MaxHoldMs)
            MK_Dir[k] := false

    dx := 0, dy := 0
    if MK_Dir["U"]
        dy -= 1
    if MK_Dir["D"]
        dy += 1
    if MK_Dir["L"]
        dx -= 1
    if MK_Dir["R"]
        dx += 1

    if (dx = 0 && dy = 0) {
        ; 方向键松开但 Alt 仍按着：速度归零等待下次起步，档位也回到低速。
        ; 拖动状态保持不变，中途停顿不会误放开左键。
        MK_Vel := 0.0
        MK_AccX := 0.0
        MK_AccY := 0.0
        MK_BurstSince := 0
        return
    }

    if (!MK_BurstSince)
        MK_BurstSince := A_TickCount

    dt := MK_Tick / 1000.0
    if (MK_Boosted()) {
        term := MK_TermBoost
        res := MK_ResistBoost
    } else {
        term := MK_TermNormal
        res := MK_Resistance
    }

    if (MK_Vel = 0.0)
        MK_Vel := (MK_InitialSpeed < term) ? MK_InitialSpeed : term

    MK_Vel += res * (term - MK_Vel) * dt
    if (MK_Vel > term)
        MK_Vel := term

    ; 斜向归一化，否则对角线会比直线快约 41%
    if (dx != 0 && dy != 0) {
        fx := dx * 0.7071, fy := dy * 0.7071
    } else {
        fx := dx, fy := dy
    }

    step := MK_Vel * dt
    MK_AccX += fx * step
    MK_AccY += fy * step

    mx := (MK_AccX >= 0) ? Floor(MK_AccX) : Ceil(MK_AccX)
    my := (MK_AccY >= 0) ? Floor(MK_AccY) : Ceil(MK_AccY)
    if (mx != 0 || my != 0) {
        MK_AccX -= mx
        MK_AccY -= my
        MouseMove(mx, my, 0, "R")
    }
}

; 单击。点之前把 Alt 逻辑松开，点完【不】按回去。
;
; 不松开的话应用收到的是 Alt+点击：Chrome 里点收藏栏链接会变成下载而不是打开。
; 点完不按回，是因为那一下会把点击刚打开的菜单掐掉——收藏栏文件夹只闪一下高亮、
; 下拉永远出不来。两个都是实际遇到过的问题。
;
; 不按回也不影响方向键，前提是热键用 #HotIf MK_AltHeld() 自行判定物理 Alt，
; 而不是写成 !Up（那种写法只认 AHK 跟踪的逻辑状态，松开后立刻失配）。
;
; 松开前发一次 A_MenuMaskKey：否则"这次 Alt 按下期间应用没见过任何按键"，
; 用户之后物理松开 Alt 时 Windows 会判成激活菜单栏。
; NoTimers：这几毫秒里不让 MK_Step / MK_WheelTick 插队。
MK_Click(btn) {
    global MK_Masked
    Thread("NoTimers", true)
    Send("{Blind}{" . A_MenuMaskKey . "}{LAlt up}{RAlt up}")
    MK_Masked := true    ; 这一下也算屏蔽过了，接着按方向键不用再发
    Click(btn)
    Thread("NoTimers", false)
}

; 退出兜底：拖动进行中就放开左键。
; 代码压下的左键不会随进程结束自动弹起，漏了这一步左键会一直卡在按下状态
MK_ReleaseAll() {
    global MK_Dragging
    if (MK_Dragging) {
        Click("Left Up")
        MK_Dragging := false
    }
}

OnExit((*) => MK_ReleaseAll())

; 滚轮连发。两处教训都来自"松开后还在不停滚"：
;  1) 先松 Alt 再松滚轮键时，带 Alt 的松开热键不会触发，光靠它停不下定时器，
;     所以像 MK_Step 一样每一拍自查 Alt 还在不在，并加最长按住时间兜底；
;  2) 键盘自动重复会每 30ms 左右反复触发按下热键，若每次都重设 70ms 的定时器，
;     倒计时会被一直推迟、按住期间几乎不响，攒到松开才连发——同方向重复按下直接忽略。
; 按下即滚一格，再按周期连发；原来第一格要等 70ms，轻点一下什么都不发生。
MK_WheelStart(dir) {
    global MK_WheelDir, MK_WheelSince, MK_WheelRepeat
    if (MK_WheelDir = dir)
        return
    MK_WheelDir := dir
    MK_WheelSince := A_TickCount
    Send(dir > 0 ? "{WheelUp}" : "{WheelDown}")
    SetTimer(MK_WheelTick, MK_WheelRepeat)
}

MK_WheelStop() {
    global MK_WheelDir
    MK_WheelDir := 0
    SetTimer(MK_WheelTick, 0)
}

MK_WheelTick() {
    global MK_WheelDir, MK_WheelSince, MK_MaxHoldMs
    if (!MK_WheelDir || !MK_AltHeld() || (A_TickCount - MK_WheelSince) > MK_MaxHoldMs) {
        MK_WheelStop()
        return
    }
    Send(MK_WheelDir > 0 ? "{WheelUp}" : "{WheelDown}")
}

; ---------------- 热键 ----------------
; * 前缀：允许同时按着 Shift / Ctrl 等其他修饰键，不影响触发。
; 保留它是为了别让 Alt+Shift+方向 之类的组合突然变成"什么都不做"

; 全部走 #HotIf MK_AltHeld()，而不是写成 !Up / !Enter：
; 后者只在 AHK 跟踪的【逻辑】Alt 为按下时匹配，而 MK_Click 会把逻辑 Alt 松开，
; 之后手指虽然还按着，方向键也会全部失配、直接漏给应用（已实测）。
; 条件为假时这些热键根本不存在，所以不按 Alt 时方向键 / Enter / 斜杠一切如常。
#HotIf MK_AltHeld()

; 倒 T 形布局：i=上 j=左 k=下 l=右。键位形状就是方向键的形状，不用记约定。
; 不用 vim 的 hjkl，是因为按标准指法 h 和 j 都归右手食指，
; 左下（j+h）要一根食指同时压两个相邻键，按不出来；
; 倒 T 形下四个斜向分别是 中指+食指 / 中指+无名，都是两根不同的手指。
; 用字母而不是方向键，是为了把 Alt+←/→ 还给浏览器的前进后退。
;
; 注：中途试过改成 hjkl，又改了回来。别再来回换了——
; 编辑器圈子里 hjkl 是主流（源自 ADM-3A 终端键面上印的箭头），
; 但系统级光标/鼠标导航这一侧普遍推荐倒 T 形，本模块属于后者。
*i::MK_Press("U")
*i Up::MK_Release("U")
*j::MK_Press("L")
*j Up::MK_Release("L")
*k::MK_Press("D")
*k Up::MK_Release("D")
*l::MK_Press("R")
*l Up::MK_Release("R")

; 这两个也必须带 * ：不带通配符的热键要求"没有任何修饰键按着"，
; 而这里恰恰是在按着 Alt 的前提下触发，漏了 * 就永远不匹配。
*Enter::MK_Click("Left")
*SC035::MK_Click("Right")   ; SC035 = 斜杠键，写扫描码避免 */ 被解析歧义

; p=上滚 n=下滚，沿用 emacs 里 p/n 表示上/下的习惯（原先这两个键是整页跳）。
; 用字母而不是 PgUp/PgDn，一是手不用离开 ijkl 区域，二是把 Alt+PgUp/PgDn 还给应用。
*p::MK_WheelStart(1)
*p Up::MK_WheelStop()
*n::MK_WheelStart(-1)
*n Up::MK_WheelStop()

#HotIf
