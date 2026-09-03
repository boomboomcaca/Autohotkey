;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; TTS 朗读相关函数 (Edge TTS 版)
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

global g_TtsProcPid := 0  ; 点击朗读(PlayTtsText/PollTtsPlay)的 edge-tts 进程
global g_TtsProcStartTick := 0
global g_TtsPlayText := ""
global g_TtsTempFile := ""
global g_HoverTtsStartTick := 0
global g_HoverTtsProcPid := 0  ; 悬停朗读(PlayTtsLoop/PollHoverTtsPlay)独立的 edge-tts 进程，避免与点击朗读互相覆盖
global g_HoverTtsTempFile := A_Temp . "\ahk_tts_hover.mp3"  ; 悬停朗读音频文件（PlayTtsLoop/PollHoverTtsPlay/StopTts 共用）

Gui_PlayOriginal(*)
{
  global g_OrigEditCtrl, g_IsChineseMode
  text := Trim(g_OrigEditCtrl.Value)
  if (text = "")
    return
  
  ; 英文模式下通常朗读原文，中文模式下也可以朗读（Edge TTS 支持良好）
  PlayTtsText(text)
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

Gui_PlayQuestion(*)
{
  global g_QuestionEditCtrl
  PlayTtsText(g_QuestionEditCtrl.Value)
}

RestorePrevForeground()
{
  global g_PrevForegroundHwnd, g_MainGui
  try {
    if (g_PrevForegroundHwnd && WinExist("ahk_id " . g_PrevForegroundHwnd))
      WinActivate("ahk_id " . g_PrevForegroundHwnd)
  }
}

; 判断文本是否为界面占位提示（不应朗读）。
; 必须精确/锚定匹配：用 InStr(text, "正在") 会误伤含"正在"的正常句子（如"我正在学习"）
IsTtsPlaceholder(text)
{
  static placeholders := Map(
    "正在处理...", 1, "正在思考...", 1, "正在生成输出...", 1,
    "正在查询...", 1, "⏳ 正在查询...", 1,
    "正在切换语言并重新查询...", 1,
    "Querying...", 1, "⏳ Querying...", 1,
    "Switching language and re-querying...", 1)
  ; 以状态图标或"请求失败/错误:"开头的都是界面提示文字，不朗读
  return placeholders.Has(text) || RegExMatch(text, "^(⏳|⏱|⚠|请求失败|错误[:：])")
}

; 将文本安全地作为 edge-tts 的命令行参数：
; 反斜杠换为空格（末尾的 \ 会与闭合引号结合破坏参数边界，且朗读反斜杠无意义）、
; 换行压成空格、引号转义
EscapeTtsArg(text)
{
  text := StrReplace(text, "\", " ")
  text := StrReplace(text, "`r", "")
  text := StrReplace(text, "`n", " ")
  return StrReplace(text, '"', '\"')
}

; 核心朗读函数：支持中英自动识别
PlayTtsText(text, isRetry := false)
{
  ; 这三个必须声明 global：PollTtsPlay 读取它们做超时/播放判断；
  ; 漏声明会写入局部，全局停留在 0/""，导致点击朗读 100ms 即“超时”、永不播放。
  global g_TtsProcPid, g_TtsRetryCount, g_TtsProcStartTick, g_TtsPlayText, g_TtsTempFile
  static tempFile := A_Temp . "\ahk_tts_edge.mp3"
  
  if (!isRetry)
    g_TtsRetryCount := 0
  
  text := Trim(text)
  if (text = "" || IsTtsPlaceholder(text))
    return

  ; 1. 停止之前的播放和生成任务
  try {
    if (g_TtsProcPid > 0)
      ProcessClose(g_TtsProcPid)
    SoundPlay("NonExistent.zzz")
  }
  g_TtsProcPid := 0
  Sleep(50)

  ; 2. 自动检测语言并选择语音
  isChinese := RegExMatch(text, "[\x{4e00}-\x{9fff}]")
  voice := isChinese ? "zh-CN-XiaoxiaoNeural" : "en-US-AriaNeural"

  ; 3. 调用 edge-tts 生成音频
  RestorePrevForeground()
  escapedText := EscapeTtsArg(text)

  try {
    ; 使用非阻塞启动，并通过定时器轮询检测结束
    Run('edge-tts --voice ' . voice . ' --text "' . escapedText . '" --write-media "' . tempFile . '"', , "Hide", &outPid)
    g_TtsProcPid := outPid
    
    g_TtsProcStartTick := A_TickCount
    g_TtsPlayText := text
    g_TtsTempFile := tempFile
    
    SetTimer(PollTtsPlay, 100)
  } catch Error as e {
    ; 静默失败
  }
}

PollTtsPlay()
{
  global g_TtsProcPid, g_TtsProcStartTick, g_TtsPlayText, g_TtsTempFile
  
  if (g_TtsProcPid <= 0) {
    SetTimer(PollTtsPlay, 0)
    return
  }
  
  ; 超时保护 (10秒)，超时自动重试一次
  if (A_TickCount - g_TtsProcStartTick > 10000) {
    try ProcessClose(g_TtsProcPid)
    g_TtsProcPid := 0
    SetTimer(PollTtsPlay, 0)
    
    global g_TtsRetryCount, g_TtsPlayText
    if (g_TtsRetryCount < 1) {
      g_TtsRetryCount++
      PlayTtsText(g_TtsPlayText, true)
    }
    return
  }
  
  if (!ProcessExist(g_TtsProcPid)) {
    SetTimer(PollTtsPlay, 0)
    g_TtsProcPid := 0
    if (FileExist(g_TtsTempFile) && !IsTtsPlaceholder(g_TtsPlayText)) {
      ; edge-tts 网络失败时会留下空/残缺文件，SoundPlay 对其抛异常；定时器线程里必须捕获，否则弹错误框
      try SoundPlay(g_TtsTempFile) ; 异步非阻塞播放音频
    }
  }
}

CheckTtsHover()
{
  global g_TtsOrigCtrl, g_TtsCorrectCtrl, g_TtsTranslateCtrl, g_TtsQuestionCtrl
  global g_MainGui, g_TtsPlaying, g_IsChineseMode, g_HoverTarget, g_QuestionEditCtrl
  global g_PrevForegroundHwnd
  static lastHoverCtrl := ""

  ; 安全检查：确保主窗口对象存在且有效
  if (!g_MainGui || !IsObject(g_MainGui)) {
    SetTimer(CheckTtsHover, 0)
    return
  }

  try {
    ; 记录进入弹窗前的窗口句柄，用于朗读后恢复焦点（如果需要）
    fgHwnd := WinActive("A")
    if (fgHwnd && fgHwnd != g_MainGui.Hwnd)
      g_PrevForegroundHwnd := fgHwnd
  } catch {
  }

  currentHover := ""
  try {
    ; 获取鼠标下的控件 HWND
    MouseGetPos(&mx, &my, &winUnder, &ctrlUnder, 2)
    
    ; 只有当控件变量是有效的 GUI 控件对象时，才允许访问 .Hwnd
    if (ctrlUnder) {
        if (g_TtsOrigCtrl && IsObject(g_TtsOrigCtrl) && ctrlUnder = g_TtsOrigCtrl.Hwnd)
          currentHover := "orig"
        else if (g_TtsCorrectCtrl && IsObject(g_TtsCorrectCtrl) && ctrlUnder = g_TtsCorrectCtrl.Hwnd)
          currentHover := "correct"
        else if (g_TtsTranslateCtrl && IsObject(g_TtsTranslateCtrl) && ctrlUnder = g_TtsTranslateCtrl.Hwnd)
          currentHover := "translate"
        else if (g_TtsQuestionCtrl && IsObject(g_TtsQuestionCtrl) && ctrlUnder = g_TtsQuestionCtrl.Hwnd)
          currentHover := "question"
    }
  } catch {
    ; 忽略鼠标位置探测中的偶发错误
  }

  if (currentHover != "" && currentHover != lastHoverCtrl) {
    g_TtsPlaying := true
    g_HoverTarget := currentHover
    ; 异步启动播放循环，避免阻塞检测
    SetTimer(PlayTtsLoop, -10)
  } else if (currentHover = "" && lastHoverCtrl != "") {
    StopTts()
  }

  lastHoverCtrl := currentHover
}

StopTts()
{
  global g_TtsPlaying, g_HoverTarget, g_HoverTtsProcPid, g_HoverTtsTempFile
  g_TtsPlaying := false
  g_HoverTarget := ""
  SetTimer(PollHoverTtsPlay, 0)
  try {
    if (g_HoverTtsProcPid > 0 && ProcessExist(g_HoverTtsProcPid)) {
      ; 生成被中途终止，落盘的文件是残缺的，必须删掉：
      ; 否则下次悬停同一文本时 PlayTtsLoop 会因"文本未变且文件存在"直接播放残缺文件而不重新生成
      ProcessClose(g_HoverTtsProcPid)
      ProcessWaitClose(g_HoverTtsProcPid, 1)
      try FileDelete(g_HoverTtsTempFile)
    }
    SoundPlay("NonExistent.zzz")
  }
  g_HoverTtsProcPid := 0
}

PlayTtsLoop(isRetry := false)
{
  global g_TtsPlaying, g_HoverTarget, g_HoverTtsProcPid, g_HoverTtsRetryCount, g_HoverTtsTempFile
  global g_OrigEditCtrl, g_CorrectEditCtrl, g_TranslateEditCtrl, g_QuestionEditCtrl
  static lastText := ""  ; 用于缓存上一次处理的文字
  tempFile := g_HoverTtsTempFile

  if (!isRetry)
    g_HoverTtsRetryCount := 0

  if (!g_TtsPlaying || g_HoverTarget = "") {
    lastText := "" ; 清空缓存，下次进入重新生成
    return
  }

  ; 读取控件文本必须放进 try：窗口可能已销毁，控件引用失效时 .Value 会抛出未捕获异常
  text := ""
  try {
    if (g_HoverTarget = "orig")
      text := Trim(g_OrigEditCtrl.Value)
    else if (g_HoverTarget = "correct")
      text := Trim(g_CorrectEditCtrl.Value)
    else if (g_HoverTarget = "translate")
      text := Trim(g_TranslateEditCtrl.Value)
    else if (g_HoverTarget = "question")
      text := Trim(g_QuestionEditCtrl.Value)
    else
      return
  } catch {
    return
  }

  if (text = "" || IsTtsPlaceholder(text))
    return

  isChinese := RegExMatch(text, "[\x{4e00}-\x{9fff}]")
  voice := isChinese ? "zh-CN-XiaoxiaoNeural" : "en-US-AriaNeural"

  try {
    ; 核心优化：如果文字没变且文件存在，则不重新生成 (如果是重试则强制重新生成)
    if (text != lastText || !FileExist(tempFile) || isRetry) {
        ; 必须先终止上一个生成进程并停止播放：
        ; SoundPlay 播放中会锁住 tempFile，不停止则新的 edge-tts 写入失败、之后播的还是旧内容；
        ; 上一个 edge-tts 未结束则两个进程并发写同一文件
        if (g_HoverTtsProcPid > 0) {
            try ProcessClose(g_HoverTtsProcPid)
            ; 等旧进程真正退出并释放文件句柄，否则新的 edge-tts 可能打不开同名文件而生成失败
            try ProcessWaitClose(g_HoverTtsProcPid, 1)
            g_HoverTtsProcPid := 0
        }
        try SoundPlay("NonExistent.zzz")

        escapedText := EscapeTtsArg(text)

        ; 异步非阻塞生成音频
        Run('edge-tts --voice ' . voice . ' --text "' . escapedText . '" --write-media "' . tempFile . '"', , "Hide", &outPid)
        g_HoverTtsProcPid := outPid
        lastText := text

        global g_HoverTtsStartTick
        g_HoverTtsStartTick := A_TickCount
        SetTimer(PollHoverTtsPlay, 100)
    } else {
        ; 文件已存在且还是原文本，直接采用非阻塞方式播放一次
        if (g_TtsPlaying && FileExist(tempFile)) {
            SoundPlay(tempFile)
        }
    }
  } catch {
  }
}

PollHoverTtsPlay()
{
  global g_HoverTtsProcPid, g_HoverTtsStartTick, g_TtsPlaying, g_HoverTtsTempFile
  tempFile := g_HoverTtsTempFile

  if (g_HoverTtsProcPid <= 0 || !g_TtsPlaying) {
    SetTimer(PollHoverTtsPlay, 0)
    return
  }

  if (A_TickCount - g_HoverTtsStartTick > 10000) {
    try ProcessClose(g_HoverTtsProcPid)
    g_HoverTtsProcPid := 0
    SetTimer(PollHoverTtsPlay, 0)

    global g_HoverTtsRetryCount
    if (g_HoverTtsRetryCount < 1) {
      g_HoverTtsRetryCount++
      PlayTtsLoop(true)
    }
    return
  }

  if (!ProcessExist(g_HoverTtsProcPid)) {
    SetTimer(PollHoverTtsPlay, 0)
    g_HoverTtsProcPid := 0
    if (g_TtsPlaying && FileExist(tempFile)) {
      try SoundPlay(tempFile)  ; 同 PollTtsPlay：文件为空/损坏时不能让异常冒泡成错误弹窗
    }
  }
}

