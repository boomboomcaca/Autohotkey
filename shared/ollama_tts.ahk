;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; TTS 朗读相关函数 (有道真人原声 + Google TTS 神经网络兜底 + 本地 SAPI)
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

global g_TtsMinGoodBitrate := 48 ; 低于该码率 (kbps) 视作低清晰度音频，触发 Google TTS 升级
global g_TtsCacheWriteCount := 0 ; 缓存写入次数，定期触发轻量淘汰
global g_TtsProcPid := 0         ; TTS 后台 curl 下载进程 PID
global g_TtsProcStartTick := 0   ; 下载开始时间戳
global g_TtsPlayText := ""       ; 当前正在请求朗读的文本
global g_TtsStage := ""          ; 翻译窗口当前下载阶段: "" / "youdao" / "google"
global g_TtsTmpYoudao := ""      ; 有道阶段临时文件路径
global g_TtsFallbackFile := ""   ; 有道有效但低码率时的兜底文件
global g_TtsTargetCache := ""    ; 目标缓存文件路径
global g_TtsLang := "en-US"      ; 目标语种标签
global g_TtsGoogleJob := ""      ; Google 分片下载状态对象
global g_LastHoverText := ""     ; 悬停朗读防重复文本

; ===== 本地 SAPI 离线朗读（支持中英文自动适配，零延迟、免联网、最后兜底） =====
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

; ===== 将字符串按 UTF-8 做 RFC 3986 URL 编码（处理空格与撇号等特殊字符） =====
UriEncodeForUrl(str)
{
  if (str = "")
    return ""
  buf := Buffer(StrPut(str, "UTF-8"))
  cb := StrPut(str, buf, "UTF-8") - 1
  if (cb <= 0)
    return ""
  out := ""
  Loop cb {
    b := NumGet(buf, A_Index - 1, "UChar")
    if ((b >= 65 && b <= 90) || (b >= 97 && b <= 122) || (b >= 48 && b <= 57)
        || b == 45 || b == 46 || b == 95 || b == 126) {
      out .= Chr(b)
    } else {
      out .= Format("%{:02X}", b)
    }
  }
  return out
}

; ===== 读取与归一化发音口音配置 =====
TtsReadVoiceType()
{
  voiceType := 0
  try voiceType := Integer(IniRead(A_ScriptDir . "\ollama_config.ini", "Settings", "WordVoiceType", "0"))
  return voiceType
}

; 将历史 WordVoiceType（0/1/2）映射为标准有道 type（2=美音，1=英音）与 Google 语种
TtsNormalizeVoiceType(voiceType)
{
  vt := 0
  try vt := Integer(voiceType)
  if (vt == 1)
    return {youdaoType: 1, googleLang: "en-GB"}
  return {youdaoType: 2, googleLang: "en-US"}
}

; 根据文本内容与用户口音偏好判定 Google TTS 语种标签
TtsDetectGoogleLang(text, voiceType := -1)
{
  if RegExMatch(text, "[\x{4e00}-\x{9fff}]")
    return "zh-CN"
  vt := (voiceType < 0) ? TtsReadVoiceType() : voiceType
  return TtsNormalizeVoiceType(vt).googleLang
}

; ===== 归一化有道查词参数（清洗标点并统一转小写，避免有道返回 500 JSON） =====
TtsNormalizeYoudaoWord(word)
{
  text := Trim(word)
  if (text = "")
    return ""
  text := StrReplace(StrReplace(text, Chr(0x2019), "'"), Chr(0x2018), "'")
  cleanWord := Trim(RegExReplace(RegExReplace(text, "[^a-zA-Z0-9\s'-]", ""), "\s+", " "))
  if (cleanWord = "")
    cleanWord := text
  return StrLower(cleanWord)
}

; ===== 32-bit FNV-1a 哈希（用于本地音频缓存文件名去碰撞） =====
TtsFnv1aHex(str)
{
  buf := Buffer(StrPut(str, "UTF-8"))
  cb := StrPut(str, buf, "UTF-8") - 1
  h := 0x811C9DC5
  Loop cb {
    b := NumGet(buf, A_Index - 1, "UChar")
    h := ((h ^ b) * 0x01000193) & 0xFFFFFFFF
  }
  return Format("{:08x}", h)
}

; ===== 本地音频缓存目录与路径管理 =====
TtsGetCacheDir()
{
  dir := A_Temp . "\ahk_tts_cache"
  if !DirExist(dir)
    try DirCreate(dir)
  return dir
}

TtsIsInCacheDir(filePath)
{
  if (filePath = "")
    return false
  dir := RTrim(A_Temp . "\ahk_tts_cache", "\") . "\"
  normPath := StrReplace(filePath, "/", "\")
  return (InStr(normPath, dir, false) = 1)
}

TtsDeleteTempFile(filePath)
{
  if (filePath != "" && !TtsIsInCacheDir(filePath))
    try FileDelete(filePath)
}

TtsCleanupFileList(fileList)
{
  if (!IsObject(fileList))
    return
  for fp in fileList
    TtsDeleteTempFile(fp)
}

TtsGetCachePath(cacheKey, langKey)
{
  norm := StrLower(Trim(RegExReplace(cacheKey, "[\r\n\s]+", " ")))
  safeLang := RegExReplace(langKey, "[^a-zA-Z0-9_-]", "")
  if (safeLang = "")
    safeLang := "en-US"
  prefix := RegExReplace(norm, "[^a-z0-9_-]", "_")
  prefix := Trim(SubStr(prefix, 1, 24), "_")
  if (prefix = "")
    prefix := "tts"
  hash := TtsFnv1aHex(safeLang . "|" . norm)
  return TtsGetCacheDir() . "\" . prefix . "_" . safeLang . "_" . hash . ".mp3"
}

TtsResolveCacheInfo(text, voiceType := -1)
{
  vt := (voiceType < 0) ? TtsReadVoiceType() : voiceType
  lang := TtsDetectGoogleLang(text, vt)
  isYd := IsYoudaoEligible(text)
  cacheKey := isYd ? ("w:" . TtsNormalizeYoudaoWord(text)) : ("s:" . text)
  cacheFile := TtsGetCachePath(cacheKey, lang)
  return {voiceType: vt, lang: lang, isYoudao: isYd, cacheKey: cacheKey, cacheFile: cacheFile}
}

; ===== 纯 AHK MP3 帧头解析与码率校验 =====
TtsParseMp3FrameHeader(buf, offset, bufLen)
{
  static brV1 := [32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320]
  static brV2 := [8, 16, 24, 32, 40, 48, 56, 64, 80, 96, 112, 128, 144, 160]
  static srV1 := [44100, 48000, 32000]
  static srV2 := [22050, 24000, 16000]
  static srV25 := [11025, 12000, 8000]
  static invalid := {valid: false, bitrateKbps: 0, sampleRate: 0, frameLen: 0}

  if (offset < 0 || offset + 4 > bufLen)
    return invalid

  b0 := NumGet(buf, offset, "UChar")
  b1 := NumGet(buf, offset + 1, "UChar")
  if (b0 != 0xFF || (b1 & 0xE0) != 0xE0)
    return invalid

  verBits := (b1 >> 3) & 3
  layerBits := (b1 >> 1) & 3
  if (verBits == 1 || layerBits != 1)
    return invalid

  b2 := NumGet(buf, offset + 2, "UChar")
  brIdx := (b2 >> 4) & 0xF
  srIdx := (b2 >> 2) & 3
  if (brIdx < 1 || brIdx > 14 || srIdx > 2)
    return invalid
  pad := (b2 >> 1) & 1

  if (verBits == 3) {
    br := brV1[brIdx]
    sr := srV1[srIdx + 1]
    frameLen := Floor(144 * br * 1000 / sr) + pad
  } else {
    br := brV2[brIdx]
    sr := (verBits == 2) ? srV2[srIdx + 1] : srV25[srIdx + 1]
    frameLen := Floor(72 * br * 1000 / sr) + pad
  }
  if (frameLen < 16)
    return invalid
  return {valid: true, bitrateKbps: br, sampleRate: sr, frameLen: frameLen}
}

TtsIsXingInfoFrame(buf, frameOffset, frameLen, bufLen)
{
  ; MPEG1/2/2.5 各声道模式下 Xing/Info 元数据头固定位于帧偏移 +13 / +21 / +36
  for off in [13, 21, 36] {
    if (off + 4 <= frameLen && frameOffset + off + 4 <= bufLen) {
      b0 := NumGet(buf, frameOffset + off, "UChar")
      b1 := NumGet(buf, frameOffset + off + 1, "UChar")
      b2 := NumGet(buf, frameOffset + off + 2, "UChar")
      b3 := NumGet(buf, frameOffset + off + 3, "UChar")
      if ((b0 == 0x58 && b1 == 0x69 && b2 == 0x6E && b3 == 0x67)     ; "Xing"
          || (b0 == 0x49 && b1 == 0x6E && b2 == 0x66 && b3 == 0x6F)) ; "Info"
        return true
    }
  }
  return false
}

; 校验文件是否为合法 MP3，并采样多帧获取真实音频码率（拒绝 JSON/HTML 错误体）
TtsGetMp3Info(filePath)
{
  invalid := {valid: false, bitrateKbps: 0, sampleRate: 0, size: 0}
  if (filePath = "" || !FileExist(filePath))
    return invalid

  fileSize := 0
  try fileSize := FileGetSize(filePath)
  if (fileSize < 256)
    return {valid: false, bitrateKbps: 0, sampleRate: 0, size: fileSize}

  f := ""
  buf := ""
  bufLen := 0
  scanStart := 0
  try {
    f := FileOpen(filePath, "r")
    if (!f)
      return invalid
    readLen := Min(fileSize, 16384)
    buf := Buffer(readLen)
    bufLen := f.RawRead(buf, readLen)
    if (bufLen < 10) {
      f.Close()
      return invalid
    }

    firstByte := NumGet(buf, 0, "UChar")
    ; 拒绝 JSON / HTML 错误响应
    if (firstByte == 0x7B || firstByte == 0x3C || firstByte == 0x5B) {
      f.Close()
      return invalid
    }

    ; 跳过 ID3v2 标签头
    if (NumGet(buf, 0, "UChar") == 0x49 && NumGet(buf, 1, "UChar") == 0x44 && NumGet(buf, 2, "UChar") == 0x33) {
      flags := NumGet(buf, 5, "UChar")
      tagSize := ((NumGet(buf, 6, "UChar") & 0x7F) << 21)
               | ((NumGet(buf, 7, "UChar") & 0x7F) << 14)
               | ((NumGet(buf, 8, "UChar") & 0x7F) << 7)
               | (NumGet(buf, 9, "UChar") & 0x7F)
      scanStart := 10 + tagSize + ((flags & 0x10) ? 10 : 0)
      if (scanStart >= fileSize) {
        f.Close()
        return invalid
      }
      if (scanStart >= bufLen || (bufLen - scanStart < 4096 && fileSize - scanStart > bufLen - scanStart)) {
        readLen2 := Min(fileSize - scanStart, 8192)
        if (readLen2 < 4) {
          f.Close()
          return invalid
        }
        buf := Buffer(readLen2)
        f.Seek(scanStart, 0)
        bufLen := f.RawRead(buf, readLen2)
        scanStart := 0
      }
    }
    f.Close()
    f := ""
  } catch {
    if (f)
      try f.Close()
    return invalid
  }

  scanLimit := Min(bufLen - 4, scanStart + 4096)
  pos := scanStart
  while (pos <= scanLimit) {
    h1 := TtsParseMp3FrameHeader(buf, pos, bufLen)
    if (!h1.valid) {
      pos++
      continue
    }
    nextPos := pos + h1.frameLen
    if (nextPos + 4 > bufLen) {
      pos++
      continue
    }
    h2 := TtsParseMp3FrameHeader(buf, nextPos, bufLen)
    if (!h2.valid) {
      pos++
      continue
    }

    ; 跳过首帧 LAME Xing/Info 无声元数据头，避免其高码率掩盖后续真实的 32 kbps 低码率语音帧
    isMeta0 := TtsIsXingInfoFrame(buf, pos, h1.frameLen, bufLen)
    maxBr := isMeta0 ? h2.bitrateKbps : Max(h1.bitrateKbps, h2.bitrateKbps)
    sr := h1.sampleRate
    stepPos := nextPos + h2.frameLen
    frames := isMeta0 ? 1 : 2
    while (frames < 6 && stepPos + 4 <= bufLen) {
      hNext := TtsParseMp3FrameHeader(buf, stepPos, bufLen)
      if (!hNext.valid)
        break
      if (hNext.bitrateKbps > maxBr)
        maxBr := hNext.bitrateKbps
      stepPos += hNext.frameLen
      frames++
    }
    return {valid: true, bitrateKbps: maxBr, sampleRate: sr, size: fileSize}
  }
  return invalid
}

TtsCacheHit(filePath, requireGoodBitrate := true)
{
  global g_TtsMinGoodBitrate
  info := TtsGetMp3Info(filePath)
  return info.valid && (!requireGoodBitrate || info.bitrateKbps >= g_TtsMinGoodBitrate)
}

; 按修改时间清理超出上限的最旧缓存文件
TtsPruneCache(maxFiles := 200, keepFiles := 150)
{
  dir := TtsGetCacheDir()
  list := ""
  count := 0
  Loop Files, dir . "\*.mp3" {
    list .= A_LoopFileTimeModified . "`t" . A_LoopFileFullPath . "`n"
    count++
  }
  if (count <= maxFiles)
    return
  list := RTrim(Sort(list), "`n")
  deleteCount := count - keepFiles
  idx := 0
  for line in StrSplit(list, "`n") {
    idx++
    if (idx > deleteCount)
      break
    parts := StrSplit(line, "`t", , 2)
    if (parts.Length = 2)
      try FileDelete(parts[2])
  }
}

TtsSaveToCache(srcFile, cacheFile)
{
  global g_TtsCacheWriteCount
  if (srcFile = "" || !FileExist(srcFile))
    return srcFile
  if (StrLower(srcFile) = StrLower(cacheFile))
    return cacheFile
  TtsGetCacheDir()
  try {
    FileMove(srcFile, cacheFile, true)
    g_TtsCacheWriteCount++
    if (Mod(g_TtsCacheWriteCount, 20) = 0)
      TtsPruneCache(200, 150)
    return cacheFile
  } catch {
    if (TtsGetMp3Info(cacheFile).valid) {
      TtsDeleteTempFile(srcFile)
      return cacheFile
    }
    return srcFile
  }
}

; ===== Google TTS 长文本智能分片（每片 <= 180 字符，最多 3 片） =====
TtsSplitTextForGoogle(text, maxChars := 180, maxChunks := 3)
{
  clean := Trim(RegExReplace(text, "[\r\n\s]+", " "))
  clean := StrReplace(StrReplace(clean, Chr(0x2019), "'"), Chr(0x2018), "'")
  if (clean = "")
    return []
  if (StrLen(clean) <= maxChars)
    return [clean]

  chunks := []
  rest := clean
  minCut := Floor(maxChars * 0.45)

  Loop maxChunks {
    if (rest = "")
      break
    if (StrLen(rest) <= maxChars) {
      chunks.Push(rest)
      rest := ""
      break
    }

    sub := SubStr(rest, 1, maxChars)
    cutPos := 0
    if RegExMatch(sub, "^(.*(?:[。？！；，：|]|[.?!,;:](?!\d)))", &m) && StrLen(m[1]) >= minCut {
      cutPos := StrLen(m[1])
    } else {
      spPos := InStr(sub, " ", , -1)
      cutPos := (spPos >= minCut) ? spPos : maxChars
    }

    piece := Trim(SubStr(rest, 1, cutPos))
    if (piece != "")
      chunks.Push(piece)
    rest := Trim(SubStr(rest, cutPos + 1))
  }

  ; 超出 maxChunks 上限（>540 字），交由本地 SAPI 朗读，避免多段等待
  if (rest != "")
    return []
  return chunks
}

; ===== 合并多个无 ID3 头的 Google TTS MP3 分片为单个连续文件 =====
TtsConcatMp3Files(fileList, outFile)
{
  if (!IsObject(fileList) || fileList.Length = 0)
    return false
  for fp in fileList {
    if (!TtsGetMp3Info(fp).valid)
      return false
  }

  TtsDeleteTempFile(outFile)
  outF := ""
  inF := ""
  try {
    outF := FileOpen(outFile, "w")
    if (!outF)
      return false
    buf := Buffer(65536)
    for fp in fileList {
      inF := FileOpen(fp, "r")
      if (!inF) {
        outF.Close()
        TtsDeleteTempFile(outFile)
        return false
      }
      while ((bytesRead := inF.RawRead(buf, 65536)) > 0)
        outF.RawWrite(buf, bytesRead)
      inF.Close()
      inF := ""
    }
    outF.Close()
    outF := ""
    return TtsGetMp3Info(outFile).valid
  } catch {
    if (inF)
      try inF.Close()
    if (outF)
      try outF.Close()
    TtsDeleteTempFile(outFile)
    return false
  }
}

; 收尾 Google TTS 分片：单片直接入缓存，多片合并校验后入缓存
TtsFinalizeGoogleAudio(chunkFiles, targetCache, tempMergedPath)
{
  if (!IsObject(chunkFiles) || chunkFiles.Length = 0)
    return {ok: false, file: ""}

  srcFile := ""
  if (chunkFiles.Length = 1) {
    srcFile := chunkFiles[1]
  } else {
    if !TtsConcatMp3Files(chunkFiles, tempMergedPath) {
      TtsCleanupFileList(chunkFiles)
      TtsDeleteTempFile(tempMergedPath)
      return {ok: false, file: ""}
    }
    TtsCleanupFileList(chunkFiles)
    srcFile := tempMergedPath
  }

  if (!TtsGetMp3Info(srcFile).valid) {
    TtsDeleteTempFile(srcFile)
    return {ok: false, file: ""}
  }
  return {ok: true, file: TtsSaveToCache(srcFile, targetCache)}
}

TtsPlayAudioFileNow(filePath)
{
  if (filePath = "")
    return
  try SoundPlay("NonExistent.zzz")
  StopLocalSapi()
  try SoundPlay(filePath)
}

; ===== 启动有道词典真人发音下载（后台异步，自动转小写并映射标准 type），返回进程 PID；失败返回 0 =====
; voiceType: 0/2 = 美音(type=2)，1 = 英音(type=1)
RunYoudaoWordAudio(word, outFile, voiceType := 0)
{
  lowerWord := TtsNormalizeYoudaoWord(word)
  if (lowerWord = "")
    return 0

  ydType := TtsNormalizeVoiceType(voiceType).youdaoType
  encodedWord := UriEncodeForUrl(lowerWord)
  TtsDeleteTempFile(outFile)

  try {
    ; 使用 -f 确保 500 JSON 等错误体不落盘
    cmd := 'curl.exe -s -f -L --max-time 3 "https://dict.youdao.com/dictvoice?audio=' . encodedWord . '&type=' . ydType . '" -o "' . outFile . '"'
    Run(cmd, , "Hide", &outPid)
    return outPid
  }
  return 0
}

; ===== 启动 Google 神经网络 TTS 下载（后台异步），返回进程 PID；失败返回 0 =====
RunGoogleTtsAudio(text, outFile, lang := "en-US")
{
  clean := Trim(RegExReplace(text, "[\r\n\s]+", " "))
  clean := StrReplace(StrReplace(clean, Chr(0x2019), "'"), Chr(0x2018), "'")
  if (clean = "" || StrLen(clean) > 200)
    return 0

  encoded := UriEncodeForUrl(clean)
  TtsDeleteTempFile(outFile)

  try {
    cmd := 'curl.exe -s -f -L --max-time 4 -A "Mozilla/5.0" "https://translate.google.com/translate_tts?ie=UTF-8&client=tw-ob&tl=' . lang . '&q=' . encoded . '" -o "' . outFile . '"'
    Run(cmd, , "Hide", &outPid)
    return outPid
  }
  return 0
}

; ===== 共享 Google 分片下载状态机（供翻译窗口与查词整句朗读复用） =====
TtsGoogleJobStart(chunks, lang, targetCache, prefix)
{
  if (!IsObject(chunks) || chunks.Length = 0)
    return ""
  tmpFile := A_Temp . "\" . prefix . "_c_" . A_TickCount . "_1.mp3"
  outPid := RunGoogleTtsAudio(chunks[1], tmpFile, lang)
  if (!outPid) {
    TtsDeleteTempFile(tmpFile)
    return ""
  }
  return {
    pid: outPid,
    startTick: A_TickCount,
    lang: lang,
    targetCache: targetCache,
    prefix: prefix,
    chunks: chunks,
    idx: 1,
    files: [],
    curTmp: tmpFile,
    mergedTmp: ""
  }
}

TtsGoogleJobStop(job)
{
  if (!IsObject(job))
    return
  if (job.pid > 0 && ProcessExist(job.pid))
    try ProcessClose(job.pid)
  job.pid := 0
  TtsDeleteTempFile(job.curTmp)
  TtsDeleteTempFile(job.mergedTmp)
  TtsCleanupFileList(job.files)
  job.curTmp := ""
  job.mergedTmp := ""
  job.files := []
}

; 轮询一步 Google 分片任务；返回 {done: bool, ok: bool, file: string}
TtsGoogleJobStep(job, timeoutMs := 4500)
{
  if (!IsObject(job) || job.pid <= 0)
    return {done: true, ok: false, file: ""}

  if (ProcessExist(job.pid)) {
    if (A_TickCount - job.startTick <= timeoutMs)
      return {done: false, ok: false, file: ""}
    try ProcessClose(job.pid)
  }
  job.pid := 0

  curFile := job.curTmp
  job.curTmp := ""
  if (!TtsGetMp3Info(curFile).valid) {
    TtsDeleteTempFile(curFile)
    TtsGoogleJobStop(job)
    return {done: true, ok: false, file: ""}
  }

  job.files.Push(curFile)
  if (job.idx < job.chunks.Length) {
    job.idx++
    nextTmp := A_Temp . "\" . job.prefix . "_c_" . A_TickCount . "_" . job.idx . ".mp3"
    outPid := RunGoogleTtsAudio(job.chunks[job.idx], nextTmp, job.lang)
    if (!outPid) {
      TtsDeleteTempFile(nextTmp)
      TtsGoogleJobStop(job)
      return {done: true, ok: false, file: ""}
    }
    job.curTmp := nextTmp
    job.pid := outPid
    job.startTick := A_TickCount
    return {done: false, ok: false, file: ""}
  }

  mergedTmp := A_Temp . "\" . job.prefix . "_m_" . A_TickCount . ".mp3"
  res := TtsFinalizeGoogleAudio(job.files, job.targetCache, mergedTmp)
  job.files := []
  if (!res.ok)
    return {done: true, ok: false, file: ""}
  if (!TtsIsInCacheDir(res.file))
    job.mergedTmp := res.file
  return {done: true, ok: true, file: res.file}
}

TtsResetTranslatePipelineFiles()
{
  global g_TtsTmpYoudao, g_TtsFallbackFile, g_TtsGoogleJob
  TtsDeleteTempFile(g_TtsTmpYoudao)
  TtsDeleteTempFile(g_TtsFallbackFile)
  g_TtsTmpYoudao := ""
  g_TtsFallbackFile := ""
  if IsObject(g_TtsGoogleJob)
    TtsGoogleJobStop(g_TtsGoogleJob)
  g_TtsGoogleJob := ""
}

TtsStartGoogleStageForTranslate(chunks)
{
  global g_TtsProcPid, g_TtsProcStartTick, g_TtsStage, g_TtsLang, g_TtsTargetCache, g_TtsGoogleJob
  job := TtsGoogleJobStart(chunks, g_TtsLang, g_TtsTargetCache, "ahk_tts_tmp")
  if (!IsObject(job))
    return false
  g_TtsGoogleJob := job
  g_TtsStage := "google"
  g_TtsProcPid := job.pid
  g_TtsProcStartTick := job.startTick
  SetTimer(PollTtsPipeline, 50)
  return true
}

; ===== 核心朗读入口：缓存 → 有道高质量真人 → Google 神经网络分片 → 本地 SAPI 兜底 =====
PlayTtsText(text)
{
  global g_TtsProcPid, g_TtsProcStartTick, g_TtsPlayText, g_TtsStage
  global g_TtsTmpYoudao, g_TtsTargetCache, g_TtsLang, g_TtsFallbackFile

  text := Trim(text)
  if (text = "" || IsTtsPlaceholder(text))
    return

  ; 停止之前的播放与下载
  StopTts()
  Sleep(30)
  RestorePrevForeground()

  info := TtsResolveCacheInfo(text)
  if TtsCacheHit(info.cacheFile, info.isYoudao) {
    TtsPlayAudioFileNow(info.cacheFile)
    return
  }

  g_TtsPlayText := text
  g_TtsLang := info.lang
  g_TtsTargetCache := info.cacheFile
  g_TtsFallbackFile := ""

  if (info.isYoudao) {
    tmpYd := A_Temp . "\ahk_tts_tmp_yd_" . A_TickCount . ".mp3"
    outPid := RunYoudaoWordAudio(text, tmpYd, info.voiceType)
    if (outPid) {
      g_TtsTmpYoudao := tmpYd
      g_TtsStage := "youdao"
      g_TtsProcPid := outPid
      g_TtsProcStartTick := A_TickCount
      SetTimer(PollTtsPipeline, 50)
      return
    }
  }

  chunks := TtsSplitTextForGoogle(text, 180, 3)
  if TtsStartGoogleStageForTranslate(chunks)
    return

  PlayLocalSapi(text)
}

TtsFallbackOrSapiForTranslate()
{
  global g_TtsFallbackFile, g_TtsPlayText, g_TtsStage
  fb := g_TtsFallbackFile
  txt := g_TtsPlayText
  g_TtsStage := ""
  if (fb != "" && TtsGetMp3Info(fb).valid) {
    if (!IsTtsPlaceholder(txt))
      TtsPlayAudioFileNow(fb)
    return
  }
  TtsDeleteTempFile(fb)
  g_TtsFallbackFile := ""
  if (txt != "" && !IsTtsPlaceholder(txt))
    PlayLocalSapi(txt)
}

; ===== 轮询翻译窗口 TTS 下载管线并自动升级/兜底 =====
PollTtsPipeline()
{
  global g_TtsProcPid, g_TtsProcStartTick, g_TtsPlayText, g_TtsStage, g_TtsMinGoodBitrate
  global g_TtsTmpYoudao, g_TtsFallbackFile, g_TtsTargetCache, g_TtsGoogleJob

  if (g_TtsStage = "youdao") {
    if (g_TtsProcPid <= 0) {
      SetTimer(PollTtsPipeline, 0)
      return
    }
    if (ProcessExist(g_TtsProcPid)) {
      if (A_TickCount - g_TtsProcStartTick <= 2500)
        return
      try ProcessClose(g_TtsProcPid)
    }
    g_TtsProcPid := 0

    ydFile := g_TtsTmpYoudao
    g_TtsTmpYoudao := ""
    mp3 := TtsGetMp3Info(ydFile)
    if (mp3.valid && mp3.bitrateKbps >= g_TtsMinGoodBitrate) {
      SetTimer(PollTtsPipeline, 0)
      g_TtsStage := ""
      playPath := TtsSaveToCache(ydFile, g_TtsTargetCache)
      if (!TtsIsInCacheDir(playPath))
        g_TtsTmpYoudao := playPath
      if (!IsTtsPlaceholder(g_TtsPlayText))
        TtsPlayAudioFileNow(playPath)
      return
    }
    if (mp3.valid)
      g_TtsFallbackFile := ydFile
    else
      TtsDeleteTempFile(ydFile)

    chunks := TtsSplitTextForGoogle(g_TtsPlayText, 180, 3)
    if TtsStartGoogleStageForTranslate(chunks)
      return
    SetTimer(PollTtsPipeline, 0)
    TtsFallbackOrSapiForTranslate()
    return
  }

  if (g_TtsStage = "google") {
    step := TtsGoogleJobStep(g_TtsGoogleJob, 4500)
    g_TtsProcPid := IsObject(g_TtsGoogleJob) ? g_TtsGoogleJob.pid : 0
    if (!step.done)
      return
    SetTimer(PollTtsPipeline, 0)
    g_TtsStage := ""
    if (step.ok) {
      TtsDeleteTempFile(g_TtsFallbackFile)
      g_TtsFallbackFile := ""
      if (!IsTtsPlaceholder(g_TtsPlayText))
        TtsPlayAudioFileNow(step.file)
    } else {
      TtsFallbackOrSapiForTranslate()
    }
    return
  }

  SetTimer(PollTtsPipeline, 0)
}

; ===== 停止所有朗读和下载 =====
StopTts()
{
  global g_TtsPlaying, g_HoverTarget, g_TtsProcPid, g_LastHoverText
  global g_TtsPlayText, g_TtsStage
  g_TtsPlaying := false
  g_HoverTarget := ""
  g_LastHoverText := ""
  g_TtsPlayText := ""
  g_TtsStage := ""
  SetTimer(PollTtsPipeline, 0)
  try {
    if (g_TtsProcPid > 0 && ProcessExist(g_TtsProcPid)) {
      ProcessClose(g_TtsProcPid)
    }
    SoundPlay("NonExistent.zzz")
    StopLocalSapi()
  }
  g_TtsProcPid := 0
  TtsResetTranslatePipelineFiles()
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
