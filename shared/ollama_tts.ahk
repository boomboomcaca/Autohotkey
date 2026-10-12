;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; TTS 朗读相关函数 (有道真人原声 + 本地 SAPI 版)
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

global g_TtsProcPid := 0        ; 有道发音 curl 下载进程 PID
global g_TtsProcStartTick := 0  ; 下载开始时间戳
global g_TtsPlayText := ""       ; 当前正在请求朗读的文本
global g_TtsTempFile := A_Temp . "\ahk_tts_youdao.mp3"  ; 有道临时音频文件
global g_LastHoverText := ""    ; 悬停朗读防重复文本

; ===== 本地 SAPI 离线朗读（支持中英文自动适配，零延迟、免联网、永不 403） =====
PlayLocalSapi(text)
{
  text := Trim(text)
  if (text = "" || IsTtsPlaceholder(text))
    return

  static spVoice := ""
  try {
    if (!spVoice)
      spVoice := ComObject("SAPI.SpVoice")
    
    ; 根据内容是否含中文字符智能切换英文 (Zira) 或中文 (Huihui) 语音包
    targetName := RegExMatch(text, "[\x{4e00}-\x{9fff}]") ? "Huihui" : "Zira"
    try {
      for voice in spVoice.GetVoices() {
        if InStr(voice.GetDescription(), targetName) {
          spVoice.Voice := voice
          break
        }
      }
    }
    ; 1 (SVSFlagsAsync) | 2 (SVSFPurgeBeforeSpeak) = 3，打断前次朗读并异步播放
    spVoice.Speak(text, 3)
  } catch {
  }
}

StopLocalSapi()
{
  try {
    static spVoice := ""
    if (!spVoice)
      spVoice := ComObject("SAPI.SpVoice")
    spVoice.Speak("", 2) ; 2 = SVSFPurgeBeforeSpeak 清空未说出的缓冲区
  } catch {
  }
}

Gui_PlayOriginal(*)
{
  global g_OrigEditCtrl
  try {
    text := Trim(g_OrigEditCtrl.Value)
    if (text != "")
      PlayTtsText(text)
  }
}

Gui_PlayCorrect(*)
{
  global g_CorrectEditCtrl
  try {
    text := Trim(g_CorrectEditCtrl.Value)
    if (text != "")
      PlayTtsText(text)
  }
}

Gui_PlayTranslate(*)
{
  global g_TranslateEditCtrl
  try {
    text := Trim(g_TranslateEditCtrl.Value)
    if (text != "")
      PlayTtsText(text)
  }
}

Gui_PlayQuestion(*)
{
  global g_QuestionEditCtrl
  try {
    text := Trim(g_QuestionEditCtrl.Value)
    if (text != "")
      PlayTtsText(text)
  }
}

RestorePrevForeground()
{
  global g_PrevForegroundHwnd
  try {
    if (g_PrevForegroundHwnd && WinExist("ahk_id " . g_PrevForegroundHwnd))
      WinActivate("ahk_id " . g_PrevForegroundHwnd)
  }
}

; 判断文本是否为界面占位提示（不应朗读）
IsTtsPlaceholder(text)
{
  static placeholders := Map(
    "正在处理...", 1, "正在思考...", 1, "正在生成输出...", 1,
    "正在查询...", 1, "⏳ 正在查询...", 1,
    "正在切换语言并重新查询...", 1,
    "Querying...", 1, "⏳ Querying...", 1,
    "Switching language and re-querying...", 1)
  return placeholders.Has(text) || RegExMatch(text, "^(⏳|⏱|⚠|请求失败|错误[:：])")
}

; ===== 判断是否适合有道词典真人发音（纯英文字符、短词/短语，不超过 4 个词） =====
IsYoudaoEligible(text)
{
  t := Trim(text)
  if (t = "" || InStr(t, "`n") || InStr(t, "`r"))
    return false
  ; 包含汉字则走中文 SAPI，不走有道英文原声
  if RegExMatch(t, "[\x{4e00}-\x{9fff}]")
    return false
  ; 词数检查（按空格分割，不超过 4 个词，长度 <= 40）
  words := StrSplit(RegExReplace(t, "\s+", " "), " ")
  return (words.Length <= 4 && StrLen(t) <= 40)
}

; ===== 启动有道词典真人发音下载（后台异步），返回进程 PID；失败返回 0 =====
; voiceType: 0 = 美音（默认），1 = 英音
RunYoudaoWordAudio(word, outFile, voiceType := 0)
{
  text := Trim(word)
  if (text = "")
    return 0

  ; 过滤非字母/连字符字符
  cleanWord := RegExReplace(text, "[^a-zA-Z\s'-]", "")
  if (cleanWord = "")
    cleanWord := text

  try {
    ; 使用系统自带 curl 快速下载真人原声 MP3（最大超时 3 秒）
    cmd := 'curl -s -L --max-time 3 "https://dict.youdao.com/dictvoice?audio=' . cleanWord . '&type=' . voiceType . '" -o "' . outFile . '"'
    Run(cmd, , "Hide", &outPid)
    return outPid
  }
  return 0
}

; ===== 核心朗读入口：单词/短语走有道真人原声，长句/中文走本地 SAPI =====
PlayTtsText(text)
{
  global g_TtsProcPid, g_TtsProcStartTick, g_TtsPlayText, g_TtsTempFile
  
  text := Trim(text)
  if (text = "" || IsTtsPlaceholder(text))
    return

  ; 停止之前的播放与下载
  StopTts()
  Sleep(30)
  RestorePrevForeground()

  ; 单词与短语优先使用有道真人原声
  if (IsYoudaoEligible(text)) {
    voiceType := 0
    try voiceType := Integer(IniRead(A_ScriptDir . "\ollama_config.ini", "Settings", "WordVoiceType", "0"))
    
    g_TtsPlayText := text
    g_TtsProcStartTick := A_TickCount
    outPid := RunYoudaoWordAudio(text, g_TtsTempFile, voiceType)
    if (outPid) {
      g_TtsProcPid := outPid
      SetTimer(PollYoudaoPlay, 50)
      return
    }
  }

  ; 长句、中文或有道未启动，直接使用本地 SAPI 极速朗读
  PlayLocalSapi(text)
}

; ===== 轮询等待有道音频下载完成并播放 =====
PollYoudaoPlay()
{
  global g_TtsProcPid, g_TtsProcStartTick, g_TtsPlayText, g_TtsTempFile
  
  if (g_TtsProcPid <= 0) {
    SetTimer(PollYoudaoPlay, 0)
    return
  }

  ; 超时保护 (2.5 秒)，超时自动降级 SAPI
  if (A_TickCount - g_TtsProcStartTick > 2500) {
    try ProcessClose(g_TtsProcPid)
    g_TtsProcPid := 0
    SetTimer(PollYoudaoPlay, 0)
    PlayLocalSapi(g_TtsPlayText)
    return
  }

  if (!ProcessExist(g_TtsProcPid)) {
    SetTimer(PollYoudaoPlay, 0)
    g_TtsProcPid := 0
    if (!IsTtsPlaceholder(g_TtsPlayText)) {
      if (FileExist(g_TtsTempFile) && FileGetSize(g_TtsTempFile) > 500) {
        try SoundPlay(g_TtsTempFile)
      } else {
        ; 下载失败或文件异常，降级 SAPI
        PlayLocalSapi(g_TtsPlayText)
      }
    }
  }
}

; ===== 停止所有朗读和下载 =====
StopTts()
{
  global g_TtsPlaying, g_HoverTarget, g_TtsProcPid, g_LastHoverText
  g_TtsPlaying := false
  g_HoverTarget := ""
  g_LastHoverText := ""
  SetTimer(PollYoudaoPlay, 0)
  try {
    if (g_TtsProcPid > 0 && ProcessExist(g_TtsProcPid)) {
      ProcessClose(g_TtsProcPid)
    }
    SoundPlay("NonExistent.zzz")
    StopLocalSapi()
  }
  g_TtsProcPid := 0
}

; ===== 鼠标悬停检测 =====
CheckTtsHover()
{
  global g_TtsOrigCtrl, g_TtsCorrectCtrl, g_TtsTranslateCtrl, g_TtsQuestionCtrl
  global g_MainGui, g_TtsPlaying, g_HoverTarget, g_PrevForegroundHwnd
  static lastHoverCtrl := ""

  if (!g_MainGui || !IsObject(g_MainGui)) {
    SetTimer(CheckTtsHover, 0)
    return
  }

  try {
    fgHwnd := WinActive("A")
    if (fgHwnd && fgHwnd != g_MainGui.Hwnd)
      g_PrevForegroundHwnd := fgHwnd
  } catch {
  }

  currentHover := ""
  try {
    MouseGetPos(&mx, &my, &winUnder, &ctrlUnder, 2)
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
  }

  if (currentHover != "" && currentHover != lastHoverCtrl) {
    g_TtsPlaying := true
    g_HoverTarget := currentHover
    SetTimer(PlayTtsLoop, -10)
  } else if (currentHover = "" && lastHoverCtrl != "") {
    StopTts()
  }

  lastHoverCtrl := currentHover
}

; ===== 悬停朗读调度 =====
PlayTtsLoop(*)
{
  global g_TtsPlaying, g_HoverTarget, g_LastHoverText
  global g_OrigEditCtrl, g_CorrectEditCtrl, g_TranslateEditCtrl, g_QuestionEditCtrl

  if (!g_TtsPlaying || g_HoverTarget = "") {
    g_LastHoverText := ""
    return
  }

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

  if (text != g_LastHoverText) {
    g_LastHoverText := text
    PlayTtsText(text)
  }
}
