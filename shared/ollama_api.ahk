; ===== Ollama/Mistral API 共享核心模块 =====
; 统一处理所有 AI API 调用，消除代码重复

; ===== API 配置 =====
GetApiConfig()
{
    global g_MistralApiKey, g_MistralModel, g_MistralEndpoint

    ; 从配置文件读取（如果还未读取）
    if (!IsSet(g_MistralApiKey) || g_MistralApiKey = "") {
        g_MistralApiKey := IniRead(A_ScriptDir . "\ollama_config.ini", "Settings", "MistralApiKey", "")
    }
    if (!IsSet(g_MistralModel) || g_MistralModel = "") {
        g_MistralModel := IniRead(A_ScriptDir . "\ollama_config.ini", "Settings", "MistralModel", "mistral-large-latest")
    }
    if (!IsSet(g_MistralEndpoint) || g_MistralEndpoint = "") {
        g_MistralEndpoint := IniRead(A_ScriptDir . "\ollama_config.ini", "Settings", "MistralEndpoint", "https://api.mistral.ai/v1/chat/completions")
    }

    return {apiKey: g_MistralApiKey, model: g_MistralModel, endpoint: g_MistralEndpoint}
}

; ===== JSON 转义（核心逻辑）=====
EscapeJsonForApi(text)
{
    text := StrReplace(text, "\", "\\")
    text := StrReplace(text, "`"", "\`"")
    text := StrReplace(text, "`n", "\n")
    text := StrReplace(text, "`r", "\r")
    text := StrReplace(text, "`t", "\t")
    return text
}

; ===== 构建标准 API JSON =====
BuildApiJson(prompt, model?, endpoint?, temperature := 0, maxTokens := 1024, stream := false)
{
    config := GetApiConfig()
    model := model ?? config.model
    endpoint := endpoint ?? config.endpoint

    ; 系统提示：强制禁用 Markdown 和符号
    sysPrompt := "纯文本输出，不要用任何符号（如反斜杠、星号、井号）包裹或强调单词。"

    escapedPrompt := EscapeJsonForApi(prompt)

    json := '{"model":"' . model . '","messages":[{"role":"system","content":"' . sysPrompt . '"},{"role":"user","content":"' . escapedPrompt . '"}],"temperature":' . temperature . ',"max_tokens":' . maxTokens . ',"stream":' . (stream ? "true" : "false") . '}'

    return json
}

; ===== 同步调用 API（使用 WinHttp）=====
CallApiSync(prompt, model?, timeout := 30)
{
    config := GetApiConfig()
    model := model ?? config.model
    endpoint := endpoint ?? config.endpoint

    escapedPrompt := EscapeJsonForApi(prompt)
    sysPrompt := "纯文本输出，不要用任何符号（如反斜杠、星号、井号）包裹或强调单词。"

    json := '{"model":"' . model . '","messages":[{"role":"system","content":"' . sysPrompt . '"},{"role":"user","content":"' . escapedPrompt . '"}],"temperature":0,"max_tokens":1024,"stream":false}'

    try {
        http := ComObject("WinHttp.WinHttpRequest.5.1")
        http.Open("POST", endpoint, false)
        http.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
        http.SetRequestHeader("Authorization", "Bearer " . config.apiKey)
        http.Send(json)
        http.WaitForResponse(timeout)

        response := http.ResponseText

        ; 解析 OpenAI/Mistral 格式响应
        if RegExMatch(response, '"content"\s*:\s*"((?:[^"\\]|\\.)*)"', &m)
            result := m[1]
        else
            return "解析失败: " . SubStr(response, 1, 200)

        ; 转义还原
        result := StrReplace(result, "\n", "`n")
        result := StrReplace(result, "\r", "`r")
        result := StrReplace(result, "\t", "`t")
        result := StrReplace(result, "\`"", "`"")
        result := StrReplace(result, "\\", "\")

        result := Trim(result)
        return StripEmoji(result)
    } catch Error as e {
        return "请求失败: " . e.Message
    }
}

; ===== 异步调用 API（使用 curl，流式响应）=====
CallApiAsync(prompt, streamFile, &outPid, model?, temperature := 0, maxTokens := 1024)
{
    config := GetApiConfig()
    model := model ?? config.model
    endpoint := endpoint ?? config.endpoint

    ; 构建 JSON
    json := BuildApiJson(prompt, model, endpoint, temperature, maxTokens, true)

    ; 写入临时文件
    jsonFile := A_Temp . "\ahk_api_request.json"
    try FileDelete(streamFile)
    try FileDelete(jsonFile)

    try {
        FileAppend(json, jsonFile, "UTF-8-RAW")
    } catch {
        return false
    }

    ; 使用 curl.exe 调用 API（兼容 TUN 代理）
    try {
        curlCmd := 'curl.exe -s -N --connect-timeout 10 -m 120 -X POST "' . endpoint . '" -H "Content-Type: application/json" -H "Authorization: Bearer ' . config.apiKey . '" -d "@' . jsonFile . '" -o "' . streamFile . '"'
        Run(curlCmd, , "Hide", &outPid)
        return true
    } catch {
        return false
    }
}

; ===== 读取流式响应内容 =====
ReadStreamContent(filePath)
{
    if (!FileExist(filePath))
        return ""

    try {
        f := FileOpen(filePath, "r", "UTF-8")
        if (!f)
            return ""
        content := f.Read()
        f.Close()
    } catch {
        return ""
    }

    ; 解析 OpenAI SSE 流式 JSON
    result := ""
    Loop Parse, content, "`n", "`r"
    {
        line := Trim(A_LoopField)
        if (line = "")
            continue

        ; OpenAI SSE 格式: 每行以 "data: " 开头
        if (SubStr(line, 1, 6) = "data: ")
            line := SubStr(line, 7)

        ; 跳过 [DONE] 标记
        if (line = "[DONE]")
            continue

        if (!InStr(line, "{"))
            continue

        ; 检测错误
        if RegExMatch(line, '"error"\s*:\s*\{[^}]*"message"\s*:\s*"((?:[^"\\]|\\.)*)"', &m) {
            errorMsg := m[1]
            errorMsg := StrReplace(errorMsg, "\n", "`n")
            errorMsg := StrReplace(errorMsg, '\"', '"')
            return "错误: " . errorMsg
        }

        ; 提取 content
        if RegExMatch(line, '"content"\s*:\s*"((?:[^"\\]|\\.)*)"', &m) {
            token := m[1]
            token := StrReplace(token, "\n", "`n")
            token := StrReplace(token, "\r", "`r")
            token := StrReplace(token, "\t", "`t")
            token := StrReplace(token, '\"', '"')
            token := StrReplace(token, "\\", "\")
            result .= token
        }
    }

    result := Trim(result)
    result := RegExReplace(result, "(\r?\n\s*){2,}", "`n")

    return result
}

; ===== 清理临时文件 =====
CleanupApiTempFiles(streamFile)
{
    try FileDelete(streamFile)
    try FileDelete(A_Temp . "\ahk_api_request.json")
}

; ===== 过滤 Emoji 和不可渲染的 Unicode 字符 =====
; 注意：此函数是全局唯一的，所有模块通过 Include 共享
StripEmoji_FromApiModule(text)
{
    ; 移除零宽字符、变体选择符、对象替换字符、装饰符号
    text := RegExReplace(text, "[\x{200B}-\x{200F}\x{200D}\x{2060}-\x{206F}\x{FEFF}\x{FFFC}\x{FFFD}\x{FE00}-\x{FE0F}\x{2600}-\x{27BF}\x{2B50}-\x{2B55}]", "")
    ; 移除补充平面字符（Emoji 等）：过滤 UTF-16 代理对
    result := ""
    Loop Parse, text {
        cp := Ord(A_LoopField)
        if (cp >= 0xD800 && cp <= 0xDFFF)
            continue
        result .= A_LoopField
    }
    return result
}

; 保持向后兼容的别名
StripEmoji(text) => StripEmoji_FromApiModule(text)

; ===== 预设 Prompt 模板 =====

; 翻译（中→英）
GetTranslatePromptZhToEn(text)
{
    return "Translate to English. Keep the exact same formatting, including punctuation marks, line breaks, and spacing. Output only the translation:`n" . text
}

; 翻译（英→中）
GetTranslatePromptEnToZh(text)
{
    return "Translate to Chinese. Keep the exact same formatting, including punctuation marks, line breaks, and spacing. Output only the translation:`n" . text
}

; 英文纠错
GetCorrectPromptEnglish(text)
{
    return "Correct this English text for a Chinese learner.`n`nRules:`n1. First line: ONLY the corrected sentence, nothing else`n2. Second line: exactly three dashes: ---`n3. Then list errors in Chinese: 错误1: 原文 → 修正 (解释)`n`nExample output:`nI am a real team member.`n---`n错误1: i → I (句首字母需要大写)`n错误2: real team → a real team (需要冠词 a)`n`nNow correct: " . text
}

; 中文润色
GetCorrectPromptChinese(text)
{
    return "You are a Chinese language tutor. Correct and improve the following Chinese text. Fix grammar, punctuation, and improve expression while keeping the original meaning. Output only the corrected text without any explanation:`n" . text
}

; 单词英英释义
GetWordLookupPromptEn(word, context)
{
    prompt := "You are an English-English dictionary. Explain the word '" . word . "' entirely in simple English."
    if (context != "" && context != word)
        prompt .= " Please explain its meaning in the following context:\nContext: " . context

    prompt .= "\n\nPlease output using the following format (plain text only):\n● Part of Speech: xxx /American English IPA/ (phonetics is REQUIRED, always provide American English IPA)\n● Word Roots: [One-line brief breakdown, e.g. pre-(before) + dict(speak) + -ion(noun suffix)]\n● Definition: [Simple English definition]\n● Context Meaning: [Explanation based on the given context]\n● Collocations: [Common collocations or examples]"

    return prompt
}

; 单词英汉释义
GetWordLookupPromptZh(word, context)
{
    prompt := "你是一个英语词典。解释单词 '" . word . "'"
    if (context != "" && context != word)
        prompt .= " 在以下语境中的含义。\n语境：" . context
    else
        prompt .= " 的含义。"

    prompt .= "\n\n请用以下格式输出（纯文本）：\n● 词性：xxx /美式音标/（音标为必填项，必须给出美式英语 IPA 音标）\n● 词根拆解：用一行简洁列出，格式如 pre-(前缀,'之前') + dict(词根,'说') + -ion(后缀,名词)\n● 释义：xxx\n● 语境释义：在这个句子中表示...\n● 常见搭配：xxx"

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