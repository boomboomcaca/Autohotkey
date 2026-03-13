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
    ; 同步生成音频
    RunWait('edge-tts --voice ' . voice . ' --text "' . escapedText . '" --write-media "' . tempFile . '"', , "Hide")
    
    if FileExist(tempFile)
      SoundPlay(tempFile)
  } catch Error as e {
    ; 静默失败或提示
  }
}

CheckTtsHover()
{
  global g_TtsOrigCtrl, g_TtsCorrectCtrl, g_TtsTranslateCtrl, g_TtsQuestionCtrl
  global g_MainGui, g_TtsPlaying, g_IsChineseMode, g_HoverTarget, g_QuestionEditCtrl
  global g_PrevForegroundHwnd
  static lastHoverCtrl := ""

  if (g_MainGui = "") {
    SetTimer(CheckTtsHover, 0)
    return
  }

  try {
    fgHwnd := WinGetID("A")
    if (fgHwnd != g_MainGui.Hwnd)
      g_PrevForegroundHwnd := fgHwnd
  }

  currentHover := ""
  try {
    MouseGetPos(&mx, &my, &winUnder, &ctrlUnder, 2)
    if (ctrlUnder = g_TtsOrigCtrl.Hwnd)
      currentHover := "orig"
    else if (ctrlUnder = g_TtsCorrectCtrl.Hwnd)
      currentHover := "correct"
    else if (ctrlUnder = g_TtsTranslateCtrl.Hwnd)
      currentHover := "translate"
    else if (g_TtsQuestionCtrl != "" && ctrlUnder = g_TtsQuestionCtrl.Hwnd)
      currentHover := "question"
  } 

  if (currentHover != "" && currentHover != lastHoverCtrl) {
    g_TtsPlaying := true
    g_HoverTarget := currentHover
    PlayTtsLoop()
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
  static tempFile := A_Temp . "\ahk_tts_hover.mp3"

  if (!g_TtsPlaying || g_HoverTarget = "")
    return

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
    escapedText := StrReplace(text, '"', '\"')
    escapedText := StrReplace(escapedText, '`n', ' ')
    
    ; 生成音频
    RunWait('edge-tts --voice ' . voice . ' --text "' . escapedText . '" --write-media "' . tempFile . '"', , "Hide")
    
    if (g_TtsPlaying && FileExist(tempFile)) {
      SoundPlay(tempFile, "Wait")
      
      ; 播放完毕后如果还在悬停，循环播放
      if (g_TtsPlaying)
        SetTimer(PlayTtsLoop, -300)
    }
  } catch {
  }
}

