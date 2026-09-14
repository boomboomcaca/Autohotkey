;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;; 键盘控制鼠标：按住 Alt + 方向键
;;
;;   Alt + ↑↓←→           移动光标，按住逐渐加速，松开 Alt 即停
;;                        前 0.2 秒温和（便于短距离微调），之后切入高速档奔袭
;;   Alt + Shift + ↑↓←→   拖动：自动压住左键并移动，松开 Alt 时放开
;;
;;   Alt + Enter          左键单击
;;   Alt + /              右键单击
;;   Alt + PgUp / PgDn    滚轮上 / 下
;;
;; 无模式切换，按住 Alt 即用。
;;
;; 已知冲突（都是全局接管按键的必然结果，不是 bug）：
;;   1) Alt+← 和 Alt+→ 原本是浏览器后退/前进，本模块接管它们。
;;      想保留的话把 MK_ModAlt 改成 "RAlt"，改用右 Alt 触发。
;;   2) Alt+Enter 被本模块吃掉，Chrome 地址栏"新标签打开"、Excel 单元格内换行、
;;      资源管理器"属性"等原生功能在全局失效。
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

; 菜单屏蔽键。MK_Click 里要临时松开 Alt，若这次 Alt 按下期间应用没见过任何按键，
; 松开那一刻 Windows 会把它当成"激活菜单栏"（Win32 菜单高亮、Chrome 聚焦 ⋮）。
; 必须用 AHK 自己的 A_MenuMaskKey（本机为 Ctrl vk11sc01D）；
; 早先手写的 vkE8 不起屏蔽作用，实测 13 次全部误激活了菜单。
global MK_MaskKey      := "{" . A_MenuMaskKey . "}"

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

; 用逻辑状态而非 "P" 物理状态：物理状态读不到程序注入的按键。
MK_AltDown() {
    global MK_ModAlt
    if (MK_ModAlt = "LAlt")
        return GetKeyState("LAlt")
    if (MK_ModAlt = "RAlt")
        return GetKeyState("RAlt")
    return GetKeyState("LAlt") || GetKeyState("RAlt")
}

MK_Press(k) {
    global MK_Dir, MK_DirSince, MK_Tick, MK_Dragging
    ; 按下瞬间若同时按着 Shift，进入拖动：先压住左键再开始移动
    if (GetKeyState("Shift") && !MK_Dragging) {
        Click("Left Down")
        MK_Dragging := true
    }
    MK_Dir[k] := true
    MK_DirSince[k] := A_TickCount
    SetTimer(MK_Step, MK_Tick)
}

MK_Release(k) {
    global MK_Dir
    MK_Dir[k] := false
}

MK_Stop() {
    global MK_Vel, MK_AccX, MK_AccY, MK_Dir, MK_Dragging, MK_BurstSince
    SetTimer(MK_Step, 0)
    for k in ["U", "D", "L", "R"]
        MK_Dir[k] := false
    MK_BurstSince := 0
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

    if (!MK_AltDown()) {
        MK_Stop()
        return
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

; 单击。直接 Click 不行：用户物理按着 Alt，应用会把点击当成 Alt+点击——
; Chrome 里 Alt+点链接 = 下载，VS Code 里 Alt+点击 = 加一个光标
; （真人实测：改之前 12 次点击到达时全部 Alt=1，改之后为 Alt=0）。
; 所以先逻辑松开 Alt 再点、点完按回。两处都要发屏蔽键：
; 松开前发，避免"这次 Alt 按下期间没按过键"被判成激活菜单；
; 按回后再发一次，否则用户之后物理松开 Alt 时同样会激活菜单。
; NoTimers：这几毫秒里不让 MK_Step / MK_WheelTick 插队，它们的 Alt 自查会把这次松开当成用户松手。
MK_Click(btn) {
    global MK_MaskKey
    Thread("NoTimers", true)
    Send("{Blind}" . MK_MaskKey . "{Alt up}")
    Click(btn)
    Send("{Blind}{Alt down}" . MK_MaskKey)
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
;  1) 先松 Alt 再松 PgUp/PgDn 时，带 Alt 的松开热键不会触发，光靠它停不下定时器，
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
    if (!MK_WheelDir || !MK_AltDown() || (A_TickCount - MK_WheelSince) > MK_MaxHoldMs) {
        MK_WheelStop()
        return
    }
    Send(MK_WheelDir > 0 ? "{WheelUp}" : "{WheelDown}")
}

; ---------------- 热键 ----------------
; * 前缀：允许同时按着 Shift / Ctrl 等其他修饰键，不影响触发。
; 保留它是为了别让 Alt+Shift+方向 之类的组合突然变成"什么都不做"

*!Up::MK_Press("U")
*!Up Up::MK_Release("U")
*!Down::MK_Press("D")
*!Down Up::MK_Release("D")
*!Left::MK_Press("L")
*!Left Up::MK_Release("L")
*!Right::MK_Press("R")
*!Right Up::MK_Release("R")

!Enter::MK_Click("Left")
!/::MK_Click("Right")

*!PgUp::MK_WheelStart(1)
*!PgUp Up::MK_WheelStop()
*!PgDn::MK_WheelStart(-1)
*!PgDn Up::MK_WheelStop()
