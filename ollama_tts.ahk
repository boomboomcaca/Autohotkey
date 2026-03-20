;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; TTS 朗读相关函数 (Edge TTS 版)
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

global g_TtsProcPid := 0  ; 用于跟踪 edge-tts 进程

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

; 核心朗读函数：支持中英自动识别
PlayTtsText(text)
{
  global g_TtsProcPid
  static tempFile := A_Temp . "\ahk_tts_edge.mp3"
  
  text := Trim(text)
  if (text = "" || InStr(text, "正在") || InStr(text, "切换后"))
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
  escapedText := StrReplace(text, '"', '\"')
  escapedText := StrReplace(escapedText, '`n', ' ')
  escapedText := StrReplace(escapedText, '`r', '')

  try {
    ; 使用 Run 代替 RunWait 以捕获 PID 并允许在生成过程中被 StopTts 强行终止
    Run('edge-tts --voice ' . voice . ' --text "' . escapedText . '" --write-media "' . tempFile . '"', , "Hide", &outPid)
    g_TtsProcPid := outPid
    
    ; 等待音频生成完成
    ProcessWaitClose(outPid)
    
    if (FileExist(tempFile) && !InStr(text, "正在") && !InStr(text, "切换后"))
      SoundPlay(tempFile)
  } catch Error as e {
    ; 静默失败
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
  global g_TtsPlaying, g_HoverTarget, g_TtsProcPid
  g_TtsPlaying := false
  g_HoverTarget := ""
  try {
    if (g_TtsProcPid > 0)
      ProcessClose(g_TtsProcPid)
    SoundPlay("NonExistent.zzz")
  }
  g_TtsProcPid := 0
}

PlayTtsLoop()
{
  global g_TtsPlaying, g_HoverTarget, g_TtsProcPid
  global g_OrigEditCtrl, g_CorrectEditCtrl, g_TranslateEditCtrl, g_QuestionEditCtrl
  static lastText := ""  ; 用于缓存上一次处理的文字
  static tempFile := A_Temp . "\ahk_tts_hover.mp3"

  if (!g_TtsPlaying || g_HoverTarget = "") {
    lastText := "" ; 清空缓存，下次进入重新生成
    return
  }

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

  if (text = "" || InStr(text, "正在") || InStr(text, "切换后"))
    return

  isChinese := RegExMatch(text, "[\x{4e00}-\x{9fff}]")
  voice := isChinese ? "zh-CN-XiaoxiaoNeural" : "en-US-AriaNeural"
  
  try {
    ; 核心优化：如果文字没变且文件存在，则不重新生成
    if (text != lastText || !FileExist(tempFile)) {
        escapedText := StrReplace(text, '"', '\"')
        escapedText := StrReplace(escapedText, '`n', ' ')
        
        ; 生成音频
        Run('edge-tts --voice ' . voice . ' --text "' . escapedText . '" --write-media "' . tempFile . '"', , "Hide", &outPid)
        g_TtsProcPid := outPid
        ProcessWaitClose(outPid)
        lastText := text
    }
    
    if (g_TtsPlaying && FileExist(tempFile)) {
      SoundPlay(tempFile, "Wait")
      
      ; 播放完毕后如果还在悬停，循环播放
      if (g_TtsPlaying)
        SetTimer(PlayTtsLoop, -300) ; 将间隔从 300ms 缩短到 100ms
    }
  } catch {
  }
}

