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

; ===== 按 UTF-8 整体读取文本文件（读不到时返回空串，不抛异常）=====
ReadTextFileUtf8(filePath)
{
    if (filePath = "" || !FileExist(filePath))
        return ""
    content := ""
    try {
        f := FileOpen(filePath, "r", "UTF-8")
        if (!f)
            return ""
        content := f.Read()
        f.Close()
    }
    return content
}

; ===== 解析 OpenAI 兼容的 SSE 流，返回拼接后的正文 =====
; 三个调用方（取词浮窗、Prompt 对话、翻译窗）此前各有一份逐行相同的拷贝，
; 改一个字段名就要同步改三处、漏一处就是某条链路静默返回空串，故收归此处。
;
; 只负责"从流里把正文抠出来"，不做 Trim、不压缩空行、不过滤 Emoji——
; 这些各调用方的要求不一样（流式热路径上还要避免重复执行），留给调用方自己做。
;
; 遇到接口错误体时立即停止解析，把消息写进 errMsg 并返回已拼接的内容。
; 调用方只要 errMsg 非空就应当优先显示它：流到一半才报错时，
; 继续显示那半截内容会让用户以为回答已经完整。
;
; errMsg 是必填的输出参数，不用写成可选：AHK v2 里 "&errMsg := """ 并不能让 ByRef 参数变成可选，
; 省略实参会直接报加载期错误（v2 的可选 ByRef 写法是 &errMsg?）。
; 不关心错误的调用方传一个丢弃用的变量即可。
ParseSseContent(rawText, &errMsg)
{
    errMsg := ""
    result := ""
    if (rawText = "")
        return ""

    Loop Parse, rawText, "`n", "`r"
    {
        line := Trim(A_LoopField)
        if (line = "")
            continue

        ; SSE 格式: 每行以 "data: " 开头
        if (SubStr(line, 1, 6) = "data: ")
            line := SubStr(line, 7)

        ; 跳过结束标记
        if (line = "[DONE]")
            continue

        if (!InStr(line, "{"))
            continue

        ; 错误体的两种形态：{"error":{"message":"..."}} 与 {"error":"..."}。
        ; 正文里出现的 "error" 会被转义成 \"error\"，不会匹配到这里
        if RegExMatch(line, '"error"\s*:\s*\{[^}]*"message"\s*:\s*"((?:[^"\\]|\\.)*)"', &m)
        {
            errMsg := StrReplace(StrReplace(m[1], "\n", "`n"), '\"', '"')
            return result
        }
        if RegExMatch(line, '"error"\s*:\s*"((?:[^"\\]|\\.)*)"', &m)
        {
            errMsg := StrReplace(StrReplace(m[1], "\n", "`n"), '\"', '"')
            return result
        }

        ; OpenAI 格式: choices[0].delta.content / choices[0].message.content
        if RegExMatch(line, '"content"\s*:\s*"((?:[^"\\]|\\.)*)"', &m)
            result .= UnescapeApiJson(m[1])
    }
    return result
}

; ===== 从响应正文中提取接口错误信息 =====
; curl 用 -s 且不带 -f 时，HTTP 4xx/5xx 的退出码仍然是 0，错误 JSON 被原样写进 -o 指定的文件；
; 而 SSE 解析器只认 "content" 字段，错误体里没有这个键，解析结果就是空串。
; 各调用方若不先判断接口错误，会把 403（模型不在订阅套餐内）、401（key 失效）、
; 429（限流）一律显示成"请检查网络连接"，把排查方向完全带偏。
; 仅应在正常内容解析为空时调用。
ExtractApiError(rawText)
{
    if (rawText = "")
        return ""
    ; Mistral/OpenAI 两种错误体格式：
    ;   {"object":"error","message":"...","type":"tier_not_allowed","code":"1910","raw_status_code":403}
    ;   {"error":{"message":"...","type":"invalid_request_error","code":null}}
    if !RegExMatch(rawText, '"message"\s*:\s*"((?:[^"\\]|\\.)*)"', &m)
        return ""
    msg := Trim(UnescapeApiJson(m[1]))
    if (msg = "")
        return ""
    ; 附带状态码/错误类型，便于区分是套餐问题、key 问题还是限流
    detail := ""
    if RegExMatch(rawText, '"raw_status_code"\s*:\s*(\d+)', &s)
        detail := "HTTP " . s[1]
    if RegExMatch(rawText, '"type"\s*:\s*"([^"]+)"', &t)
        detail := (detail = "") ? t[1] : detail . " " . t[1]
    return (detail = "") ? msg : msg . " (" . detail . ")"
}

; ===== curl 流式请求 =====
; 一次请求要用三个临时文件，路径统一由 prefix 派生。
; 此前这些文件名以字面量散落在各模块里（ahk_wl_curl.cfg 一个名字就出现在 5 处），
; 改名要同步改 6 个地方，漏一处轻则 TEMP 里留垃圾，
; 重则删不掉那个写着 API key 的 cfg 文件。
CurlStreamFile(prefix)  => A_Temp . "\ahk_" . prefix . "_stream.txt"
CurlRequestFile(prefix) => A_Temp . "\ahk_" . prefix . "_request.json"
CurlConfigFile(prefix)  => A_Temp . "\ahk_" . prefix . "_curl.cfg"

; 删除一次请求留下的全部临时文件。cfg 里写着 API key，任何收尾路径都必须删到它
CurlCleanupTempFiles(prefix)
{
    try FileDelete(CurlStreamFile(prefix))
    try FileDelete(CurlRequestFile(prefix))
    try FileDelete(CurlConfigFile(prefix))
}

; 用 curl.exe 发起一次 OpenAI 兼容的流式请求，响应由 curl 直接写进 CurlStreamFile(prefix)，
; 调用方轮询该文件即可。
;
; 用 curl 而不是 WinHttp：兼容 TUN 模式代理。
; Authorization 写进 -K 配置文件而不是命令行：命令行参数本机任意进程都能从进程列表读到，
; 直接写在命令行上等于公开 API key。
;
; 返回 {pid, stage, detail}：
;   pid    成功时为 curl 进程 PID，失败为 0
;   stage  失败环节——"tempfile"（临时文件写不进去）或 "curl"（curl.exe 起不来）
;   detail curl 环节的异常消息，供调用方拼进提示
; 只报告失败、不显示提示、也不复位调用方的 pending 状态：
; 各调用方的报错界面和状态变量都不一样，硬塞进来只会把耦合搬个地方。
StartCurlStream(prefix, json, timeoutSec)
{
    global g_MistralApiKey, g_MistralEndpoint

    jsonFile := CurlRequestFile(prefix)
    curlCfg := CurlConfigFile(prefix)
    CurlCleanupTempFiles(prefix)

    try {
        FileAppend(json, jsonFile, "UTF-8-RAW")
        FileAppend('header = "Authorization: Bearer ' . g_MistralApiKey . '"`n', curlCfg, "UTF-8-RAW")
    } catch {
        return {pid: 0, stage: "tempfile", detail: ""}
    }

    try {
        curlCmd := 'curl.exe -s -N --connect-timeout 10 -m ' . timeoutSec
                 . ' -X POST "' . g_MistralEndpoint . '"'
                 . ' -H "Content-Type: application/json"'
                 . ' -K "' . curlCfg . '"'
                 . ' -d "@' . jsonFile . '"'
                 . ' -o "' . CurlStreamFile(prefix) . '"'
        Run(curlCmd, , "Hide", &outPid)
        return {pid: outPid, stage: "", detail: ""}
    } catch Error as e {
        return {pid: 0, stage: "curl", detail: e.Message}
    }
}

; ===== Prompt 正文在 ini 中的转义 =====
; ini 是行式格式，一个键值就是一行。把含换行的正文原样写进 prompt=，
; 读回来时只能靠"后续行都算续行"来还原，由此带来三处内容丢失：
;   1) 正文里的空行会被当成分隔空行吃掉；
;   2) 以 [ 开头的正文行会被误判成 section 头，该行及之后的内容全部丢失；
;   3) 首尾空格被解析时的 Trim 吃掉。
; 因此正文一旦"不安全"就转成单行转义形式写进 prompt_esc=。
; 普通正文（绝大多数情况）仍写成可读的 prompt=，手工编辑 ini 的体验保持不变。
EscapePromptValue(text)
{
    text := StrReplace(text, "\", "\\")
    text := StrReplace(text, "`r", "\r")
    text := StrReplace(text, "`n", "\n")
    text := StrReplace(text, "`t", "\t")
    return text
}

; 反转义。同 UnescapeApiJson：必须先把 \\ 抽到占位符，
; 否则字面反斜杠加 n（"\\n"）会被先解析成换行而损坏内容。
UnescapePromptValue(text)
{
    ph := Chr(0xE000)  ; 私有区占位符，正常文本不会出现
    text := StrReplace(text, "\\", ph)
    text := StrReplace(text, "\n", "`n")
    text := StrReplace(text, "\r", "`r")
    text := StrReplace(text, "\t", "`t")
    text := StrReplace(text, ph, "\")
    return text
}

; 判断正文原样写进 ini 是否会被破坏
PromptValueNeedsEscaping(text)
{
    if (text = "")
        return false
    ; 换行/制表符：ini 一行只能放一个键值
    ; 反斜杠：不转义的话读回来会被当成转义序列
    if (InStr(text, "`n") || InStr(text, "`r") || InStr(text, "`t") || InStr(text, "\"))
        return true
    ; 以 [ 开头会被误判成 section 头
    if (SubStr(text, 1, 1) = "[")
        return true
    ; 首尾空格会被解析时的 Trim 吃掉
    if (text != Trim(text))
        return true
    return false
}

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