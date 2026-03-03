;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; TTS 朗读相关函数
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

Gui_PlayOriginal(*)
{
  global g_OrigEditCtrl, g_IsChineseMode
  static tempFile := ""
  
  ; 只在英文模式下朗读
  if (g_IsChineseMode)
    return
  
  text := Trim(g_OrigEditCtrl.Value)
  if (text = "")
    return
  
  ; 恢复前台窗口，避免全屏时任务栏弹出
  RestorePrevForeground()
  
  ; 停止之前的播放
  try {
    SoundPlay("NonExistent.zzz")
  }
  Sleep(50)
  
  ; 删除旧文件
  if (tempFile != "" && FileExist(tempFile)) {
    try {
      FileDelete(tempFile)
    }
  }
  
  ; 使用 Google TTS API
  try {
    ; 构建 Google TTS URL
    encodedText := ""
    Loop Parse, text
    {
      char := A_LoopField
      if RegExMatch(char, "[a-zA-Z0-9\-_.~]")
        encodedText .= char
      else
        encodedText .= "%" . Format("{:02X}", Ord(char))
    }
    
    ttsUrl := "https://translate.google.com/translate_tts?ie=UTF-8&tl=en-US&client=tw-ob&q=" . encodedText
    
    ; 使用固定文件名
    tempFile := A_Temp . "\ahk_tts_audio.mp3"
    
    http := ComObject("WinHttp.WinHttpRequest.5.1")
    http.Open("GET", ttsUrl, false)
    http.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
    http.SetRequestHeader("Referer", "https://translate.google.com/")
    http.Send()
    http.WaitForResponse()
    
    if (http.Status = 200) {
      ; 保存音频文件
      adoStream := ComObject("ADODB.Stream")
      adoStream.Type := 1  ; Binary
      adoStream.Open()
      adoStream.Write(http.ResponseBody)
      adoStream.SaveToFile(tempFile, 2)  ; 2 = overwrite
      adoStream.Close()
      
      ; 播放音频
      SoundPlay(tempFile)
    }
  } catch Error as e {
    ; 静默失败
  }
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
  ; 恢复之前的前台窗口，防止全屏时任务栏弹出
  try {
    if (g_PrevForegroundHwnd && WinExist("ahk_id " . g_PrevForegroundHwnd))
      WinActivate("ahk_id " . g_PrevForegroundHwnd)
  }
}

PlayTtsText(text)
{
  static tempFile := ""
  
  text := Trim(text)
  if (text = "" || InStr(text, "正在") || InStr(text, "切换后"))
    return
  
  ; 恢复前台窗口，避免全屏时任务栏弹出
  RestorePrevForeground()
  
  ; 停止之前的播放
  try {
    SoundPlay("NonExistent.zzz")
  }
  Sleep(50)
  
  ; 删除旧文件
  if (tempFile != "" && FileExist(tempFile)) {
    try {
      FileDelete(tempFile)
    }
  }
  
  ; 使用 Google TTS API
  try {
    encodedText := ""
    Loop Parse, text
    {
      char := A_LoopField
      if RegExMatch(char, "[a-zA-Z0-9\-_.~]")
        encodedText .= char
      else
        encodedText .= "%" . Format("{:02X}", Ord(char))
    }
    
    ttsUrl := "https://translate.google.com/translate_tts?ie=UTF-8&tl=en-US&client=tw-ob&q=" . encodedText
    tempFile := A_Temp . "\ahk_tts_audio.mp3"
    
    http := ComObject("WinHttp.WinHttpRequest.5.1")
    http.Open("GET", ttsUrl, false)
    http.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
    http.SetRequestHeader("Referer", "https://translate.google.com/")
    http.Send()
    http.WaitForResponse()
    
    if (http.Status = 200) {
      adoStream := ComObject("ADODB.Stream")
      adoStream.Type := 1
      adoStream.Open()
      adoStream.Write(http.ResponseBody)
      adoStream.SaveToFile(tempFile, 2)
      adoStream.Close()
      
      SoundPlay(tempFile)
    }
  } catch {
  }
}

CheckTtsHover()
{
  global g_TtsOrigCtrl, g_TtsCorrectCtrl, g_TtsTranslateCtrl, g_TtsQuestionCtrl
  global g_MainGui, g_TtsPlaying, g_IsChineseMode, g_HoverTarget, g_QuestionEditCtrl
  global g_PrevForegroundHwnd
  static lastHoverCtrl := ""

  ; 如果窗口已关闭，停止定时器
  if (g_MainGui = "") {
    SetTimer(CheckTtsHover, 0)
    return
  }

  ; 记录上一个非本窗口的前台窗口，用于朗读后恢复焦点
  try {
    fgHwnd := WinGetID("A")
    if (fgHwnd != g_MainGui.Hwnd)
      g_PrevForegroundHwnd := fgHwnd
  }

  ; 检测鼠标在哪个朗读图标上
  currentHover := ""
  try {
    MouseGetPos(&mx, &my, &winUnder, &ctrlUnder, 2)
    ; 中文模式只有翻译图标，英文模式有三个图标
    if (!g_IsChineseMode && ctrlUnder = g_TtsOrigCtrl.Hwnd)
      currentHover := "orig"
    else if (!g_IsChineseMode && ctrlUnder = g_TtsCorrectCtrl.Hwnd)
      currentHover := "correct"
    else if (ctrlUnder = g_TtsTranslateCtrl.Hwnd)
      currentHover := "translate"
    else if (g_TtsQuestionCtrl != "" && ctrlUnder = g_TtsQuestionCtrl.Hwnd)
      currentHover := "question"
  } catch {
  }

  if (currentHover != "" && currentHover != lastHoverCtrl) {
    ; 进入新图标，开始播放
    g_TtsPlaying := true
    g_HoverTarget := currentHover
    PlayTtsLoop()
  } else if (currentHover = "" && lastHoverCtrl != "") {
    ; 离开图标，停止播放
    g_TtsPlaying := false
    g_HoverTarget := ""
    try {
      SoundPlay("NonExistent.zzz")
    }
  }

  lastHoverCtrl := currentHover
}

PlayTtsLoop()
{
  global g_TtsPlaying, g_IsChineseMode, g_HoverTarget
  global g_OrigEditCtrl, g_CorrectEditCtrl, g_TranslateEditCtrl, g_QuestionEditCtrl
  static tempFile := ""

  if (!g_TtsPlaying || g_HoverTarget = "")
    return

  ; 根据悬停目标获取文本
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

  ; 使用 Google TTS API
  try {
    encodedText := ""
    Loop Parse, text
    {
      char := A_LoopField
      if RegExMatch(char, "[a-zA-Z0-9\-_.~]")
        encodedText .= char
      else
        encodedText .= "%" . Format("{:02X}", Ord(char))
    }

    ttsUrl := "https://translate.google.com/translate_tts?ie=UTF-8&tl=en-US&client=tw-ob&q=" . encodedText
    tempFile := A_Temp . "\ahk_tts_audio.mp3"

    http := ComObject("WinHttp.WinHttpRequest.5.1")
    http.Open("GET", ttsUrl, false)
    http.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
    http.SetRequestHeader("Referer", "https://translate.google.com/")
    http.Send()
    http.WaitForResponse()

    if (http.Status = 200) {
      adoStream := ComObject("ADODB.Stream")
      adoStream.Type := 1
      adoStream.Open()
      adoStream.Write(http.ResponseBody)
      adoStream.SaveToFile(tempFile, 2)
      adoStream.Close()

      ; 播放并等待完成
      SoundPlay(tempFile, "Wait")

      ; 播放完毕后如果还在悬停，继续播放
      if (g_TtsPlaying)
        SetTimer(PlayTtsLoop, -100)
    }
  } catch {
  }
}
