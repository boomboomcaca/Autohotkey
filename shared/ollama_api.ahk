; ===== Ollama/Mistral API 共享核心模块 =====
; 统一处理 JSON 转义、文本清理和 Prompt 模板
; （API 配置 g_Mistral* 由 ollama_translate.ahk 顶层从 ini 读取，各调用方直接使用全局变量）

; ===== JSON 转义（核心逻辑）=====
EscapeJsonForApi(text)
{
    text := StrReplace(text, "\", "\\")
    text := StrReplace(text, "`"", "\`"")
    text := StrReplace(text, "`n", "\n")
    text := StrReplace(text, "`r", "\r")
    text := StrReplace(text, "`t", "\t")
    ; 其余 C0 控制字符（如 \f、\b、ESC）在 JSON 字符串里必须转义，原样发出接口会直接 400；
    ; 它们对文本没有意义，直接删除
    text := RegExReplace(text, "[\x01-\x08\x0B\x0C\x0E-\x1F]", "")
    return text
}

; ===== JSON 反转义（核心逻辑）=====
; 关键：必须先把转义反斜杠 (\\) 抽到占位符，最后再还原，
; 否则像 "\\n"（字面反斜杠+n）会被先解析成换行而损坏内容。
UnescapeApiJson(text)
{
    ph := Chr(0xE000)  ; 私有区占位符，正常文本不会出现
    text := StrReplace(text, "\\", ph)
    ; \uXXXX → 实际字符（含代理对）。必须放在 \\ 抽走之后，否则字面 "\\u0041" 会被误解码。
    ; Ollama 等 Go 实现的服务端会把 < > & 输出成 \u003c \u003e \u0026，不解码就会原样显示
    if InStr(text, "\u") {
        pos := 1
        while (pos := RegExMatch(text, "i)\\u([0-9A-F]{4})", &m, pos)) {
            cp := Integer("0x" . m[1])
            len := 6
            if (cp >= 0xD800 && cp <= 0xDBFF && RegExMatch(SubStr(text, pos + 6, 6), "i)^\\u(D[C-F][0-9A-F]{2})$", &m2)) {
                cp := 0x10000 + ((cp - 0xD800) << 10) + (Integer("0x" . m2[1]) - 0xDC00)
                len := 12
            }
            ch := (cp = 0 || (cp >= 0xD800 && cp <= 0xDFFF)) ? "" : Chr(cp)  ; NUL 与孤立代理项丢弃
            text := SubStr(text, 1, pos - 1) . ch . SubStr(text, pos + len)
            pos += StrLen(ch)
        }
    }
    text := StrReplace(text, "\n", "`n")
    text := StrReplace(text, "\r", "`r")
    text := StrReplace(text, "\t", "`t")
    text := StrReplace(text, "\`"", "`"")
    text := StrReplace(text, "\/", "/")
    text := StrReplace(text, ph, "\")
    return text
}

; ===== 过滤 Emoji 和不可渲染的 Unicode 字符 =====
; 注意：此函数是全局唯一的，所有模块通过 Include 共享
StripEmoji_FromApiModule(text)
{
    ; 移除零宽字符、变体选择符、对象替换字符、装饰符号
    text := RegExReplace(text, "[\x{200B}-\x{200F}\x{200D}\x{2060}-\x{206F}\x{FEFF}\x{FFFC}\x{FFFD}\x{FE00}-\x{FE0F}\x{2600}-\x{27BF}\x{2B50}-\x{2B55}]", "")
    ; 移除补充平面字符（Emoji 等），但保留 CJK 扩展区（0x20000-0x3FFFF）的生僻汉字，
    ; 否则像 "𠀀" 这类扩展 B 区汉字会被连同 Emoji 一起静默删除
    result := ""
    i := 1
    len := StrLen(text)
    while (i <= len) {
        cp := Ord(SubStr(text, i, 2))  ; 完整代理对时返回补充平面码点（>= 0x10000）
        if (cp >= 0x10000) {
            if (cp >= 0x20000 && cp <= 0x3FFFF)
                result .= SubStr(text, i, 2)
            i += 2
        } else {
            if !(cp >= 0xD800 && cp <= 0xDFFF)  ; 孤立代理项直接丢弃
                result .= SubStr(text, i, 1)
            i++
        }
    }
    return result
}

; 保持向后兼容的别名
StripEmoji(text) => StripEmoji_FromApiModule(text)

; ===== 读取 UTF-8 ini 中的键值 =====
; IniRead 底层是 GetPrivateProfileString：对没有 UTF-16 BOM 的文件按系统 ANSI 代码页解析，
; 本项目的 ini 是无 BOM 的 UTF-8，中文值（模板名、Anki 牌组/字段名）会被读成乱码。
; 含中文的值必须用本函数按 UTF-8 自行解析；纯 ASCII 的键（API key、URL）用 IniRead 即可。
IniReadUtf8(file, section, key, default := "")
{
    content := ""
    try content := FileRead(file, "UTF-8")
    catch
        return default
    inSection := false
    Loop Parse, content, "`n", "`r"
    {
        line := Trim(A_LoopField)
        if (line = "" || SubStr(line, 1, 1) = ";")
            continue
        if (RegExMatch(line, "^\[(.*)\]$", &m)) {
            inSection := (m[1] = section)
            continue
        }
        if (inSection && RegExMatch(line, "^([^=]*?)\s*=\s*(.*)$", &kv) && kv[1] = key)
            return kv[2]
    }
    return default
}

; ===== 预设 Prompt 模板 =====
; 注意：模板里的换行必须写成 `n。EscapeJsonForApi 会把反斜杠转义，
; 写成字面 \n 时模型收到的是 "\" "n" 两个字符而不是换行

; 单词英英释义
GetWordLookupPromptEn(word, context)
{
    prompt := "You are an English-English dictionary. Explain the word '" . word . "' entirely in simple English."
    if (context != "" && context != word)
        prompt .= " Please explain its meaning in the following context:`nContext: " . context

    prompt .= "`n`nPlease output using the following format (plain text only):`n● Part of Speech: xxx /American English IPA/ (phonetics is REQUIRED, always provide American English IPA)`n● Word Roots: [One-line brief breakdown, e.g. pre-(before) + dict(speak) + -ion(noun suffix)]`n● Definition: [Simple English definition]`n● Context Meaning: [Explanation based on the given context]`n● Collocations: [Common collocations or examples]"

    return prompt
}

; 单词英汉释义
GetWordLookupPromptZh(word, context)
{
    prompt := "你是一个英语词典。解释单词 '" . word . "'"
    if (context != "" && context != word)
        prompt .= " 在以下语境中的含义。`n语境：" . context
    else
        prompt .= " 的含义。"

    prompt .= "`n`n请用以下格式输出（纯文本）：`n● 词性：xxx /美式音标/（音标为必填项，必须给出美式英语 IPA 音标）`n● 词根拆解：用一行简洁列出，格式如 pre-(前缀,'之前') + dict(词根,'说') + -ion(后缀,名词)`n● 释义：xxx`n● 语境释义：在这个句子中表示...`n● 常见搭配：xxx"

    return prompt
}

; 组合翻译+纠错（中文模式）
GetCombinedPromptChinese(text)
{
    return "请对以下中文进行润色和翻译。不要使用Markdown格式。`n`n输出格式(严格遵守):`n===CORRECT===`n润色后的中文`n===TRANSLATE===`n英文翻译`n`n原文: " . text
}

; 组合翻译+纠错（英文模式）
GetCombinedPromptEnglish(text)
{
    return "纠正并翻译以下英文。纯文本输出，不要用任何符号包裹单词。`n`n格式：`n===CORRECT===`n纠正后的英文`n---`n错误: 原文 → 修正 (解释)`n===TRANSLATE===`n中文翻译`n`n英文: " . text
}