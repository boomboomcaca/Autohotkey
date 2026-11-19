; ===== OCR 字符混淆修正 =====
; 等宽/代码/艺术字体下 OCR 容易把 l/I/|/!↔1、o/O↔0 混淆
; 策略：根据相邻字母的大小写语境决定修正目标
FixOcrConfusion(word)
{
    ; 纯数字不修正
    if RegExMatch(word, "^\d+$")
        return word

    result := ""
    len := StrLen(word)
    Loop len {
        ch := SubStr(word, A_Index, 1)
        prev := A_Index > 1 ? SubStr(word, A_Index - 1, 1) : ""
        next := A_Index < len ? SubStr(word, A_Index + 1, 1) : ""
        prevIsLetter := RegExMatch(prev, "[a-zA-Z]")
        nextIsLetter := RegExMatch(next, "[a-zA-Z]")

        if (prevIsLetter && nextIsLetter) {
            prevIsUpper := RegExMatch(prev, "[A-Z]")
            nextIsUpper := RegExMatch(next, "[A-Z]")
            if (ch = "1" || ch = "|") {
                ; 1 / | 夹在字母间 → I 或 l
                result .= (prevIsUpper && nextIsUpper) ? "I" : "l"
            } else if (ch = "!" && prevIsUpper && nextIsUpper) {
                ; 大写单词中的 !（如 LAV!YRJM）→ I
                result .= "I"
            } else if (ch = "0") {
                ; 0 夹在字母间 → o 或 O
                result .= (prevIsUpper && nextIsUpper) ? "O" : "o"
            } else {
                result .= ch
            }
        } else {
            result .= ch
        }
    }

    ; 全大写长词里夹杂的单个小写 l（如 MlLITARY）归正为 I，避开 WiFi/URLs 等混写词
    if (StrLen(result) >= 4 && RegExMatch(result, "^[A-Zl]+$") && RegExMatch(result, "[A-Z].*[A-Z]"))
        result := StrReplace(result, "l", "I", true)

    return result
}

; ===== 清洗单词：统一弯撇号、去掉零宽/占位符及首尾标点（保留词内连字符与撇号） =====
WL_CleanWord(raw)
{
    if (raw = "")
        return ""
    s := StrReplace(raw, Chr(0x2019), "'")
    s := StrReplace(s, Chr(0x2018), "'")
    s := RegExReplace(s, "[\x{200B}-\x{200D}\x{FEFF}\x{FFFC}]", "")
    s := Trim(s)
    return RegExReplace(s, "^[^\w\x{4e00}-\x{9fff}]+|[^\w\x{4e00}-\x{9fff}]+$", "")
}

; ===== 清洗划词选区短语（允许短语/词组，拒绝多行整段误选） =====
WL_CleanPhrase(raw)
{
    if (raw = "")
        return ""
    if RegExMatch(Trim(raw), "(\r?\n.*){2,}")
        return ""
    s := StrReplace(raw, Chr(0x2019), "'")
    s := StrReplace(s, Chr(0x2018), "'")
    s := RegExReplace(s, "[\x{200B}-\x{200D}\x{FEFF}\x{FFFC}]", " ")
    s := Trim(RegExReplace(s, "[\r\n\s]+", " "))
    s := RegExReplace(s, "^[^\w\x{4e00}-\x{9fff}]+|[^\w\x{4e00}-\x{9fff}]+$", "")
    if (s = "" || StrLen(s) > 70)
        return ""
    if !RegExMatch(s, "[A-Za-z\x{4e00}-\x{9fff}]")
        return ""
    if (StrSplit(s, " ").Length > 8)
        return ""
    if (!InStr(s, " ") && RegExMatch(s, "[\x{4e00}-\x{9fff}]") && StrLen(s) > 20)
        return ""
    return s
}

; ===== 判断文本是否包含目标词或短语（英文/数字按词边界匹配，避免短词误命中子串） =====
WL_ContainsWord(haystack, word)
{
    if (haystack = "" || word = "")
        return false
    normWord := StrReplace(StrReplace(word, Chr(0x2019), "'"), Chr(0x2018), "'")
    if RegExMatch(normWord, "^[A-Za-z0-9]+(?:[\s'\-][A-Za-z0-9]+)*$") {
        normHay := (InStr(haystack, Chr(0x2019)) || InStr(haystack, Chr(0x2018)))
            ? StrReplace(StrReplace(haystack, Chr(0x2019), "'"), Chr(0x2018), "'")
            : haystack
        escaped := RegExReplace(normWord, "[\\^$.|?*+()\[\]{}]", "\$0")
        escaped := RegExReplace(escaped, "\s+", "\s+")
        return RegExMatch(normHay, "i)(?<![A-Za-z0-9])" . escaped . "(?![A-Za-z0-9])") > 0
    }
    if (InStr(haystack, word) > 0)
        return true
    if (InStr(normWord, "'") || normWord != word) {
        norm := StrReplace(StrReplace(haystack, Chr(0x2019), "'"), Chr(0x2018), "'")
        return InStr(norm, normWord) > 0
    }
    return false
}

; ===== 清洗 UIA 文本块：按 ￼（U+FFFC）控件边界切分并保留包含目标词的真实段落 =====
WL_CleanUiaBlock(rawText, targetWord := "")
{
    if (rawText = "")
        return ""
    s := rawText
    if (InStr(s, Chr(0xFFFC))) {
        sep := Chr(0xE001)
        s := RegExReplace(s, "(?:\r?\n[\s\x{FFFC}]*\x{FFFC}[\s\x{FFFC}]*|[\s\x{FFFC}]*\x{FFFC}[\s\x{FFFC}]*\r?\n|\x{FFFC}{2,})", sep)
        segs := StrSplit(s, sep)
        picked := ""
        longest := ""
        nonEmpty := []
        for seg in segs {
            cleanSeg := RegExReplace(seg, "[\x{200B}-\x{200D}\x{FEFF}\x{FFFC}]", " ")
            cleanSeg := Trim(RegExReplace(cleanSeg, "[\r\n\s]+", " "))
            if (cleanSeg = "")
                continue
            nonEmpty.Push(cleanSeg)
            if (StrLen(cleanSeg) > StrLen(longest))
                longest := cleanSeg
        }
        if (targetWord != "") {
            for seg in nonEmpty {
                if (WL_ContainsWord(seg, targetWord)) {
                    picked := seg
                    break
                }
            }
            if (picked = "") {
                for seg in nonEmpty {
                    if (InStr(seg, targetWord)) {
                        picked := seg
                        break
                    }
                }
            }
            s := (picked != "") ? picked : longest
        } else {
            joined := ""
            for seg in nonEmpty
                joined .= (joined = "" ? "" : " ") . seg
            s := joined
        }
    }
    s := RegExReplace(s, "[\x{200B}-\x{200D}\x{FEFF}\x{FFFC}]", " ")
    return Trim(RegExReplace(s, "[\r\n\s]+", " "))
}

; ===== 从长段文本中按句末标点提取包含目标词的句子；无标点长文本按词位置截取窗口 =====
WL_ExtractTargetSentence(text, word, maxChars := 260)
{
    text := Trim(RegExReplace(text, "[\r\n\s]+", " "))
    if (text = "" || word = "" || StrLen(text) <= maxChars)
        return text

    phDot := Chr(0xE002)
    sep := Chr(0xE001)

    tmp := RegExReplace(text, "\b(Mr|Mrs|Ms|Dr|Prof|Sr|Jr|St|Fig|No|Vol|Ch|vs|etc)\.", "$1" . phDot)
    tmp := RegExReplace(tmp, "i)\b(i)\.(e)\.", "$1" . phDot . "$2" . phDot)
    tmp := RegExReplace(tmp, "i)\b(e)\.(g)\.", "$1" . phDot . "$2" . phDot)
    tmp := RegExReplace(tmp, "\b([A-Z])\.", "$1" . phDot)
    tmp := RegExReplace(tmp, "(?<=\d)\.(?=\d)", phDot)

    tmp := RegExReplace(tmp, "([.?!][`"')\]»”’]*)\s+", "$1" . sep)
    tmp := RegExReplace(tmp, "([。？！；][”’）\]】]*)\s*", "$1" . sep)
    tmp := StrReplace(tmp, phDot, ".")

    sents := []
    for part in StrSplit(tmp, sep) {
        p := Trim(part)
        if (p != "")
            sents.Push(p)
    }

    hitIdx := 0
    for i, s in sents {
        if (WL_ContainsWord(s, word)) {
            hitIdx := i
            break
        }
    }
    if (hitIdx = 0) {
        for i, s in sents {
            if (InStr(s, word)) {
                hitIdx := i
                break
            }
        }
    }

    if (hitIdx > 0) {
        res := sents[hitIdx]
        if (StrLen(res) < 45) {
            if (hitIdx < sents.Length && StrLen(res) + 1 + StrLen(sents[hitIdx + 1]) <= maxChars)
                res .= " " . sents[hitIdx + 1]
            else if (hitIdx > 1 && StrLen(sents[hitIdx - 1]) + 1 + StrLen(res) <= maxChars)
                res := sents[hitIdx - 1] . " " . res
        }
    } else {
        res := text
    }

    maxHard := Round(maxChars * 1.4)
    if (StrLen(res) > maxHard) {
        pos := 0
        if RegExMatch(word, "^[A-Za-z0-9]+$") {
            escaped := RegExReplace(word, "[\\^$.|?*+()\[\]{}]", "\$0")
            pos := RegExMatch(res, "i)(?<![A-Za-z0-9])" . escaped . "(?![A-Za-z0-9])")
        }
        if (pos = 0)
            pos := InStr(res, word)
        if (pos > 0) {
            half := maxChars // 2
            startPos := Max(1, pos - half)
            endPos := Min(StrLen(res), startPos + maxChars - 1)
            if (endPos - startPos + 1 < maxChars && startPos > 1)
                startPos := Max(1, endPos - maxChars + 1)
            slice := SubStr(res, startPos, endPos - startPos + 1)
            if (startPos > 1) {
                sp := InStr(slice, " ")
                if (sp > 0 && sp < 30)
                    slice := SubStr(slice, sp + 1)
            }
            if (endPos < StrLen(res)) {
                sp := InStr(slice, " ", , -1)
                if (sp > 0 && StrLen(slice) - sp < 30)
                    slice := SubStr(slice, 1, sp - 1)
            }
            res := Trim(slice)
        } else {
            res := Trim(SubStr(res, 1, maxHard))
        }
    }
    return Trim(res)
}

; ===== 点到矩形边界的最短距离（垂直偏差按 2.5 倍加权，防止跨行吸词） =====
WL_RectEdgeDist(px, py, rx, ry, rw, rh, &dx := 0, &dy := 0)
{
    dx := Max(0, rx - px, px - (rx + rw))
    dy := Max(0, ry - py, py - (ry + rh))
    return Sqrt(dx * dx + (dy * 2.5) ** 2)
}

; ===== 校验 UIA TextRange 屏幕坐标是否贴近鼠标（防止 Document/穿透误吸远处文本） =====
WL_IsRangeNearPoint(textRange, mouseX, mouseY, maxDx, maxDy, allowEmptyRects := false, boundRect := 0)
{
    rects := ""
    try rects := textRange.GetBoundingRectangles()
    if (!IsObject(rects) || rects.Length = 0)
        return allowEmptyRects

    hasBound := IsObject(boundRect) && (boundRect.r > boundRect.l) && (boundRect.b > boundRect.t)
    for r in rects {
        if (r.w <= 0 || r.h <= 0)
            continue
        if (hasBound) {
            if (Min(r.x + r.w, boundRect.r + 8) <= Max(r.x, boundRect.l - 8)
                || Min(r.y + r.h, boundRect.b + 8) <= Max(r.y, boundRect.t - 8))
                continue
        }
        WL_RectEdgeDist(mouseX, mouseY, r.x, r.y, r.w, r.h, &dx, &dy)
        if (dx <= maxDx && dy <= maxDy)
            return true
    }
    return false
}

; ===== 判断 UIA 相邻段落是否应并入上下文（过滤按钮分隔、跨栏跳跃及已有完整句） =====
WL_ShouldMergeUiaNeighbor(currRange, neighborRange, neighborRaw, currClean, isPrev)
{
    if (neighborRaw = "" || InStr(neighborRaw, Chr(0xFFFC)))
        return false
    nClean := Trim(RegExReplace(neighborRaw, "[\r\n\s]+", " "))
    if (nClean = "" || nClean = currClean || InStr(currClean, nClean))
        return false

    if (!isPrev && StrLen(currClean) >= 35 && RegExMatch(currClean, "[.?!。？！][`"')\]»”’]*$"))
        return false
    if (isPrev && StrLen(currClean) >= 35 && RegExMatch(nClean, "[.?!。？！][`"')\]»”’]*$"))
        return false

    try {
        cRects := currRange.GetBoundingRectangles()
        nRects := neighborRange.GetBoundingRectangles()
        cMinX := 9999999, cMinY := 9999999, cMaxX := -9999999, cMaxY := -9999999, refH := 0
        for r in cRects {
            if (r.w <= 0 || r.h <= 0)
                continue
            cMinX := Min(cMinX, r.x), cMinY := Min(cMinY, r.y)
            cMaxX := Max(cMaxX, r.x + r.w), cMaxY := Max(cMaxY, r.y + r.h)
            if (refH = 0)
                refH := r.h
        }
        nMinX := 9999999, nMinY := 9999999, nMaxX := -9999999, nMaxY := -9999999
        for r in nRects {
            if (r.w <= 0 || r.h <= 0)
                continue
            nMinX := Min(nMinX, r.x), nMinY := Min(nMinY, r.y)
            nMaxX := Max(nMaxX, r.x + r.w), nMaxY := Max(nMaxY, r.y + r.h)
        }
        if (cMaxX > cMinX && nMaxX > nMinX) {
            refH := Max(8, refH)
            vGap := isPrev ? (cMinY - nMaxY) : (nMinY - cMaxY)
            if (vGap < -Round(refH * 0.4) || vGap > Round(refH * 1.4))
                return false
            hOverlap := Min(cMaxX, nMaxX) - Max(cMinX, nMinX)
            if (hOverlap < Min(cMaxX - cMinX, nMaxX - nMinX) * 0.2)
                return false
        }
    }
    return true
}

; ===== 合并 UIA 两侧紧邻连字符/撇号片段，保持复合词与缩写完整性 =====
WL_MergeCompoundNeighbors(leftText, rawWord, rightText)
{
    sLeft := StrReplace(StrReplace(leftText, Chr(0x2019), "'"), Chr(0x2018), "'")
    sWord := StrReplace(StrReplace(rawWord, Chr(0x2019), "'"), Chr(0x2018), "'")
    sRight := StrReplace(StrReplace(rightText, Chr(0x2019), "'"), Chr(0x2018), "'")

    hasLeadSpace := RegExMatch(sWord, "^\s") > 0
    hasTrailSpace := RegExMatch(sWord, "\s$") > 0
    core := Trim(sWord)
    if (core = "" || RegExMatch(core, "\s"))
        return ""

    if (core = "-" || core = "'") {
        if (!hasLeadSpace && !hasTrailSpace
            && RegExMatch(sLeft, "(?<![A-Za-z0-9'\-])([A-Za-z0-9]+(?:['\-][A-Za-z0-9]+)*)$", &mL)
            && RegExMatch(sRight, "^([A-Za-z0-9]+(?:['\-][A-Za-z0-9]+)*)(?![A-Za-z0-9'\-])", &mR))
            return mL[1] . core . mR[1]
        return ""
    }

    if (!hasLeadSpace) {
        if (RegExMatch(core, "^['\-][A-Za-z0-9]")) {
            if (RegExMatch(sLeft, "(?<![A-Za-z0-9])([A-Za-z0-9]+(?:['\-][A-Za-z0-9]+)*)$", &mL))
                core := mL[1] . core
        } else if (RegExMatch(core, "^[A-Za-z0-9]")) {
            if (RegExMatch(sLeft, "(?<![A-Za-z0-9])([A-Za-z0-9]+(?:['\-][A-Za-z0-9]+)*['\-])$", &mL))
                core := mL[1] . core
        }
    }

    if (!hasTrailSpace) {
        if (RegExMatch(core, "[A-Za-z0-9]['\-]$")) {
            if (RegExMatch(sRight, "^([A-Za-z0-9]+(?:['\-][A-Za-z0-9]+)*)(?![A-Za-z0-9])", &mR))
                core := core . mR[1]
        } else if (RegExMatch(core, "[A-Za-z0-9]$")) {
            if (RegExMatch(sRight, "^(['\-][A-Za-z0-9]+(?:['\-][A-Za-z0-9]+)*)(?![A-Za-z0-9])", &mR))
                core := core . mR[1]
        }
    }

    return core
}

; ===== 探测 UIA Word 范围左右字符以补全连字符复合词（如 Franco-Prussian） =====
WL_ExpandUiaCompoundWord(wordRange)
{
    rawWord := ""
    try rawWord := wordRange.GetText()
    if (rawWord = "")
        return ""

    leftText := ""
    rightText := ""
    if (!RegExMatch(rawWord, "^\s")) {
        try {
            leftProbe := wordRange.Clone()
            leftProbe.MoveEndpointByUnit(UIA.TextPatternRangeEndpoint.Start, UIA.TextUnit.Character, -32)
            leftProbe.MoveEndpointByRange(UIA.TextPatternRangeEndpoint.End, wordRange, UIA.TextPatternRangeEndpoint.Start)
            leftText := leftProbe.GetText()
        }
    }
    if (!RegExMatch(rawWord, "\s$")) {
        try {
            rightProbe := wordRange.Clone()
            rightProbe.MoveEndpointByUnit(UIA.TextPatternRangeEndpoint.End, UIA.TextUnit.Character, 32)
            rightProbe.MoveEndpointByRange(UIA.TextPatternRangeEndpoint.Start, wordRange, UIA.TextPatternRangeEndpoint.End)
            rightText := rightProbe.GetText()
        }
    }

    merged := WL_MergeCompoundNeighbors(leftText, rawWord, rightText)
    return (merged != "") ? merged : Trim(rawWord)
}

; ===== 正则兜底剥离扁平化 Hyperlink aria-label 末尾的视频时长/播放量无障碍后缀 =====
WL_StripLinkAriaSuffix(text)
{
    s := Trim(text)
    if (StrLen(s) < 6)
        return s

    if (RegExMatch(s, "i)^(.+?)\s+by\s+.{1,60}?\s+\d[\d,.]*[KMB]?\s+views\s+\d+\s+(?:seconds?|minutes?|hours?|days?|weeks?|months?|years?)\s+ago\b.*$", &m)
        && StrLen(Trim(m[1])) >= 4) {
        s := Trim(m[1])
    } else if (RegExMatch(s, "i)^(.+?)\s+\d[\d,.]*[KMB]?\s+views\s+\d+\s+(?:seconds?|minutes?|hours?|days?|weeks?|months?|years?)\s+ago\b.*$", &m)
        && StrLen(Trim(m[1])) >= 4) {
        s := Trim(m[1])
    }

    cand := RegExReplace(s, "i)\s+\d+\s+(?:hours?|minutes?)(?:,\s*|\s+and\s+|\s+)\d+\s+(?:minutes?|seconds?)(?:(?:,\s*|\s+and\s+|\s+)\d+\s+seconds?)?\.?$", "")
    if (cand != s && StrLen(Trim(cand)) >= 4) {
        s := Trim(cand)
    } else if (RegExMatch(s, "i)^(.+?)\s+\d+\s+(?:hours?|minutes?|seconds?)\.?$", &m)
        && StrLen(Trim(m[1])) >= 4
        && !RegExMatch(m[1], "i)\b(?:in|for|within|after|under|over|every|last|first|next|final|about|around|nearly|almost|of)$")) {
        s := Trim(m[1])
    }

    if (RegExMatch(s, "^(.+?)\s+\d+\s*(?:小时|分钟)\s*\d+\s*(?:分钟|秒钟?)(?:\s*\d+\s*秒钟?)?$", &m)
        && StrLen(Trim(m[1])) >= 2) {
        s := Trim(m[1])
    } else if (RegExMatch(s, "^(.+?)\s+\d+\s*(?:小时|分钟|秒钟?)$", &m)
        && StrLen(Trim(m[1])) >= 2
        && !RegExMatch(m[1], "[在用花过近约第前后]$")) {
        s := Trim(m[1])
    }

    return s
}

; ===== 清洗元素文本（对 Hyperlink 剥离 aria-label 中的时长等后缀并收紧边界矩形） =====
WL_GetCleanElementText(el, elType, rawName, dpiScale, &outRect)
{
    outRect := 0
    try outRect := el.BoundingRectangle
    cleanCtx := Trim(RegExReplace(rawName, "s)[\r\n\s]+", " "))
    matchedChild := false

    if (elType == 50005) {
        if (IsObject(outRect)) {
            rw := outRect.r - outRect.l
            rh := outRect.b - outRect.t
            if (rw > 0 && rh > 0 && rw <= Round(2000 * dpiScale) && rh <= Round(360 * dpiScale)) {
                try {
                    kids := el.FindAll({Type: 50020})
                    if (kids.Length >= 1 && kids.Length <= 6) {
                        bestKn := ""
                        bestKid := 0
                        for k in kids {
                            off := false
                            try off := k.IsOffscreen
                            if (off)
                                continue
                            kn := ""
                            try kn := Trim(RegExReplace(k.Name, "s)[\r\n\s]+", " "))
                            if (StrLen(kn) >= 3 && InStr(cleanCtx, kn) == 1 && StrLen(kn) > StrLen(bestKn)) {
                                bestKn := kn
                                bestKid := k
                            }
                        }
                        if (bestKn != "") {
                            cleanCtx := bestKn
                            matchedChild := true
                            try {
                                kr := bestKid.BoundingRectangle
                                if (IsObject(kr) && kr.r > kr.l && kr.b > kr.t)
                                    outRect := kr
                            }
                        }
                    }
                }
            }
        }
        if (!matchedChild)
            cleanCtx := WL_StripLinkAriaSuffix(cleanCtx)
    }
    return WL_CleanUiaBlock(cleanCtx)
}

; ===== 逐词渲染场景（如 YouTube 软字幕）向上回溯容器并聚合相邻 Text 节点拼回完整句 =====
WL_GetUiaContainerLine(el, cleanedWord, mouseX, mouseY, dpiScale)
{
    if (cleanedWord = "")
        return ""
    refH := 0
    try {
        er := el.BoundingRectangle
        if (IsObject(er))
            refH := er.b - er.t
    }

    anc := el
    Loop 8 {
        try anc := anc.Parent
        catch
            break
        if (!anc)
            break

        ancType := 0
        try ancType := anc.Type
        if (ancType == 50032) ; Window
            break

        isTp := false
        try isTp := anc.IsTextPatternAvailable
        if (isTp) {
            try {
                r := anc.TextPattern.RangeFromPoint(mouseX, mouseY)
                rPara := r.Clone()
                rPara.ExpandToEnclosingUnit(UIA.TextUnit.Paragraph)
                txt := WL_CleanUiaBlock(rPara.GetText(), cleanedWord)
                if (StrLen(txt) <= StrLen(cleanedWord) + 4) {
                    rLine := r.Clone()
                    rLine.ExpandToEnclosingUnit(UIA.TextUnit.Line)
                    txt := WL_CleanUiaBlock(rLine.GetText(), cleanedWord)
                }
                if (StrLen(txt) > StrLen(cleanedWord) + 4 && (WL_ContainsWord(txt, cleanedWord) || InStr(txt, cleanedWord)))
                    return txt
            }
        }

        if (ancType == 50030) ; Document
            break

        ancName := ""
        try ancName := Trim(anc.Name)
        if (ancName != "" && StrLen(ancName) <= 300 && StrLen(ancName) > StrLen(cleanedWord) + 4) {
            dummyRect := 0
            ancClean := WL_GetCleanElementText(anc, ancType, ancName, dpiScale, &dummyRect)
            if (StrLen(ancClean) > StrLen(cleanedWord) + 4 && (WL_ContainsWord(ancClean, cleanedWord) || InStr(ancClean, cleanedWord)))
                return ancClean
        }

        try {
            ar := anc.BoundingRectangle
            if (!IsObject(ar))
                continue
            arW := ar.r - ar.l
            arH := ar.b - ar.t
            if (arW > 0 && arH > 0 && arH <= Round(220 * dpiScale) && arW <= Round(2200 * dpiScale)) {
                kids := anc.FindAll({Type: 50020})
                if (kids.Length >= 2 && kids.Length <= 45) {
                    items := []
                    for k in kids {
                        off := false
                        try off := k.IsOffscreen
                        if (off)
                            continue
                        kn := ""
                        try kn := k.Name
                        t := WL_CleanUiaBlock(kn)
                        if (t = "")
                            continue
                        kr := 0
                        try kr := k.BoundingRectangle
                        if (!IsObject(kr))
                            continue
                        kw := kr.r - kr.l
                        kh := kr.b - kr.t
                        if (kw <= 0 || kh <= 0)
                            continue
                        cy := (kr.t + kr.b) / 2
                        if (Abs(cy - mouseY) > Round(90 * dpiScale))
                            continue
                        if (refH > 4 && (kh < refH * 0.55 || kh > refH * 1.75))
                            continue

                        dup := false
                        for it in items {
                            if (it.t == t && Abs(it.x - kr.l) <= 4 && Abs(it.cy - cy) <= 4) {
                                dup := true
                                break
                            }
                        }
                        if (!dup)
                            items.Push({t: t, x: kr.l, y: kr.t, w: kw, h: kh, cy: cy})
                    }

                    if (items.Length >= 2) {
                        Loop items.Length - 1 {
                            i := A_Index + 1
                            cur := items[i]
                            j := i - 1
                            while (j >= 1) {
                                it := items[j]
                                sameRow := Abs(cur.cy - it.cy) <= Min(cur.h, it.h) * 0.55
                                comesBefore := sameRow ? (cur.x < it.x) : (cur.cy < it.cy)
                                if (comesBefore) {
                                    items[j + 1] := items[j]
                                    j--
                                } else {
                                    break
                                }
                            }
                            items[j + 1] := cur
                        }

                        joined := ""
                        for it in items {
                            if (joined = "") {
                                joined := it.t
                            } else {
                                lastCh := SubStr(joined, -1)
                                firstCh := SubStr(it.t, 1, 1)
                                noSpace := RegExMatch(firstCh, "^[,.;:?!\)\]]$")
                                    || (RegExMatch(lastCh, "[\x{4e00}-\x{9fff}]") && RegExMatch(firstCh, "[\x{4e00}-\x{9fff}]"))
                                joined .= (noSpace ? "" : " ") . it.t
                            }
                        }
                        joined := Trim(RegExReplace(joined, "\s+", " "))
                        if (StrLen(joined) > StrLen(cleanedWord) + 4 && (WL_ContainsWord(joined, cleanedWord) || InStr(joined, cleanedWord)))
                            return joined
                    }
                }
            }
        }
    }
    return ""
}

; ===== 按水平大间距切分 OCR 单行，分离同行的独立按钮/时间戳列 =====
WL_TrimOcrLineColumns(ocrLine, anchorWordIdx := 0, refLeft := 0, refRight := 0)
{
    words := ""
    try words := ocrLine.Words
    if (!IsObject(words) || words.Length = 0) {
        if (anchorWordIdx > 0 || Min(ocrLine.x + ocrLine.w, refRight) > Max(ocrLine.x, refLeft))
            return {x: ocrLine.x, y: ocrLine.y, w: ocrLine.w, h: ocrLine.h, text: Trim(ocrLine.Text)}
        return {x: 0, y: 0, w: 0, h: 0, text: ""}
    }

    lineH := (ocrLine.h > 0) ? ocrLine.h : words[1].h
    gapThreshold := Max(24, Round(lineH * 2.5))
    joinTight := Max(5, Round(lineH * 0.35))

    segs := []
    w1 := words[1]
    segMinX := w1.x, segMinY := w1.y, segMaxX := w1.x + w1.w, segMaxY := w1.y + w1.h
    segText := w1.Text
    hasAnchor := (anchorWordIdx == 1)

    i := 2
    while (i <= words.Length) {
        prevW := words[i - 1]
        currW := words[i]
        gap := currW.x - (prevW.x + prevW.w)
        if (gap > gapThreshold) {
            segs.Push({x: segMinX, y: segMinY, w: segMaxX - segMinX, h: segMaxY - segMinY, text: Trim(segText), hasAnchor: hasAnchor})
            segMinX := currW.x, segMinY := currW.y, segMaxX := currW.x + currW.w, segMaxY := currW.y + currW.h
            segText := currW.Text
            hasAnchor := (anchorWordIdx == i)
        } else {
            segMinX := Min(segMinX, currW.x), segMinY := Min(segMinY, currW.y)
            segMaxX := Max(segMaxX, currW.x + currW.w), segMaxY := Max(segMaxY, currW.y + currW.h)
            noSpace := (gap <= joinTight) && (RegExMatch(currW.Text, "^['’]") || (RegExMatch(prevW.Text, "[\w]-$") && RegExMatch(currW.Text, "^[\w]")))
            segText .= (noSpace ? "" : " ") . currW.Text
            if (anchorWordIdx == i)
                hasAnchor := true
        }
        i++
    }
    segs.Push({x: segMinX, y: segMinY, w: segMaxX - segMinX, h: segMaxY - segMinY, text: Trim(segText), hasAnchor: hasAnchor})

    if (anchorWordIdx > 0) {
        for s in segs {
            if (s.hasAnchor)
                return s
        }
        return segs[1]
    }

    mMinX := 9999999, mMinY := 9999999, mMaxX := -9999999, mMaxY := -9999999
    mText := ""
    for s in segs {
        if (Min(s.x + s.w, refRight) > Max(s.x, refLeft)) {
            mMinX := Min(mMinX, s.x), mMinY := Min(mMinY, s.y)
            mMaxX := Max(mMaxX, s.x + s.w), mMaxY := Max(mMaxY, s.y + s.h)
            mText .= (mText = "" ? "" : " ") . s.text
        }
    }
    if (mText = "")
        return {x: 0, y: 0, w: 0, h: 0, text: ""}
    return {x: mMinX, y: mMinY, w: mMaxX - mMinX, h: mMaxY - mMinY, text: Trim(mText)}
}

; ===== 按字高、水平重叠和连续行距聚合同段 OCR 行，避免拼进标题或画面杂字 =====
WL_BuildOcrContext(ocrLines, bestLineIndex, bestWordIdx, targetWord)
{
    baseSeg := WL_TrimOcrLineColumns(ocrLines[bestLineIndex], bestWordIdx)
    if (baseSeg.text = "")
        return ""

    baseH := Max(8, baseSeg.h)
    baseCy := baseSeg.y + baseSeg.h / 2
    baseRight := baseSeg.x + baseSeg.w

    above := []
    below := []
    for i, ln in ocrLines {
        if (i == bestLineIndex)
            continue
        seg := WL_TrimOcrLineColumns(ln, 0, baseSeg.x, baseRight)
        if (seg.text = "" || seg.h <= 0 || seg.w <= 0)
            continue
        ratio := seg.h / baseH
        isSameSize := (ratio >= 0.68 && ratio <= 1.45)
        isTitlePair := false
        if (!isSameSize && ratio >= 0.42 && ratio <= 2.4 && RegExMatch(seg.text, "[A-Za-z]{2,}|[\x{4e00}-\x{9fff}]")) {
            maxH := Max(baseH, seg.h)
            maxW := Max(baseSeg.w, seg.w)
            leftAligned := Abs(seg.x - baseSeg.x) <= maxH * 1.5
            centerAligned := Abs((seg.x + seg.w / 2) - (baseSeg.x + baseSeg.w / 2)) <= maxW * 0.25
            if (leftAligned || centerAligned)
                isTitlePair := true
        }
        if (!isSameSize && !isTitlePair)
            continue
        overlapW := Min(baseRight, seg.x + seg.w) - Max(baseSeg.x, seg.x)
        if (overlapW < Min(baseSeg.w, seg.w) * 0.25 && Abs(seg.x - baseSeg.x) > Max(baseH, seg.h) * 2.5)
            continue
        segCy := seg.y + seg.h / 2
        if (Abs(segCy - baseCy) <= Min(baseH, seg.h) * 0.45)
            continue
        seg.isTitlePair := isTitlePair
        if (segCy < baseCy)
            above.Push(seg)
        else
            below.Push(seg)
    }

    if (above.Length > 1) {
        Loop above.Length - 1 {
            i := A_Index + 1
            cur := above[i], j := i - 1
            while (j >= 1 && above[j].y < cur.y) {
                above[j + 1] := above[j]
                j--
            }
            above[j + 1] := cur
        }
    }
    if (below.Length > 1) {
        Loop below.Length - 1 {
            i := A_Index + 1
            cur := below[i], j := i - 1
            while (j >= 1 && below[j].y > cur.y) {
                below[j + 1] := below[j]
                j--
            }
            below[j + 1] := cur
        }
    }

    selectedAbove := []
    currTop := baseSeg.y
    for seg in above {
        if (selectedAbove.Length >= 3)
            break
        vGap := currTop - (seg.y + seg.h)
        minGap := -Round(Min(baseH, seg.h) * 0.35)
        maxGap := seg.isTitlePair ? Round(Max(baseH, seg.h) * 0.85) : Round(baseH * 1.35)
        if (vGap < minGap || vGap > maxGap)
            break
        selectedAbove.Push(seg)
        currTop := seg.y
    }

    selectedBelow := []
    currBottom := baseSeg.y + baseSeg.h
    for seg in below {
        if (selectedBelow.Length >= 3)
            break
        vGap := seg.y - currBottom
        minGap := -Round(Min(baseH, seg.h) * 0.35)
        maxGap := seg.isTitlePair ? Round(Max(baseH, seg.h) * 0.85) : Round(baseH * 1.35)
        if (vGap < minGap || vGap > maxGap)
            break
        selectedBelow.Push(seg)
        currBottom := seg.y + seg.h
    }

    joined := ""
    Loop selectedAbove.Length {
        seg := selectedAbove[selectedAbove.Length - A_Index + 1]
        joined .= (joined = "" ? "" : " ") . seg.text
    }
    joined .= (joined = "" ? "" : " ") . baseSeg.text
    for seg in selectedBelow
        joined .= " " . seg.text

    return WL_ExtractTargetSentence(joined, targetWord)
}

; ===== 从 OCR 结果中选出鼠标下最匹配的词并生成语境句子 =====
WL_PickFromOcrResult(ocrResult, mouseX, mouseY, dpiScale, uiaRect, uiaContext, &outWord, &outLine, &outSrcTag)
{
    if (!IsObject(ocrResult))
        return false
    lines := ""
    try lines := ocrResult.Lines
    if (!IsObject(lines) || lines.Length = 0)
        return false

    maxDx := Round(28 * dpiScale)
    maxDy := Round(18 * dpiScale)
    hasUiaRect := IsObject(uiaRect) && (uiaRect.r > uiaRect.l) && (uiaRect.b > uiaRect.t)
    pad := Round(8 * dpiScale)

    bestScore := 999999999
    bestCenterDist := 999999999
    bestWord := ""
    bestIndex := 0
    bestLineIndex := 0
    bestLineObj := ""
    bestInRect := false

    passCount := hasUiaRect ? 2 : 1
    Loop passCount {
        restrictToRect := (hasUiaRect && A_Index == 1)
        for lineIndex, ocrLine in lines {
            for index, ocrWord in ocrLine.Words {
                if (WL_CleanWord(ocrWord.Text) = "")
                    continue
                cx := ocrWord.x + ocrWord.w / 2
                cy := ocrWord.y + ocrWord.h / 2
                if (restrictToRect) {
                    if (cx < uiaRect.l - pad || cx > uiaRect.r + pad || cy < uiaRect.t - pad || cy > uiaRect.b + pad)
                        continue
                }
                score := WL_RectEdgeDist(mouseX, mouseY, ocrWord.x, ocrWord.y, ocrWord.w, ocrWord.h, &dx, &dy)
                if (dx > maxDx || dy > maxDy)
                    continue
                cDist := Sqrt((mouseX - cx) ** 2 + (mouseY - cy) ** 2)
                if (score < bestScore || (score == bestScore && cDist < bestCenterDist)) {
                    bestScore := score
                    bestCenterDist := cDist
                    bestWord := ocrWord.Text
                    bestIndex := index
                    bestLineIndex := lineIndex
                    bestLineObj := ocrLine
                    bestInRect := restrictToRect
                }
            }
        }
        if (bestWord != "")
            break
    }

    if (bestWord = "" || !IsObject(bestLineObj))
        return false

    rawToken := bestWord
    hitH := Max(8, bestLineObj.Words[bestIndex].h)
    tightGap := Max(3, Round(hitH * 0.22))
    joinGap := Max(5, Round(hitH * 0.45))
    minOverlap := -Round(hitH * 0.5)

    tempIndex := bestIndex - 1
    while (tempIndex > 0) {
        prevWord := bestLineObj.Words[tempIndex]
        currWord := bestLineObj.Words[tempIndex + 1]
        if (!RegExMatch(prevWord.Text, "^[\w\x{4e00}-\x{9fff}'’\-]+$") || !RegExMatch(currWord.Text, "^[\w\x{4e00}-\x{9fff}'’\-]"))
            break
        gap := currWord.x - (prevWord.x + prevWord.w)
        hasJoinChar := RegExMatch(prevWord.Text, "['’\-]$") || RegExMatch(currWord.Text, "^['’\-]")
        limit := hasJoinChar ? joinGap : tightGap
        if (gap >= minOverlap && gap <= limit) {
            bestWord := prevWord.Text . bestWord
            tempIndex--
        } else {
            break
        }
    }

    tempIndex := bestIndex + 1
    while (tempIndex <= bestLineObj.Words.Length) {
        nextWord := bestLineObj.Words[tempIndex]
        currWord := bestLineObj.Words[tempIndex - 1]
        if (!RegExMatch(nextWord.Text, "^[\w\x{4e00}-\x{9fff}'’\-]+$") || !RegExMatch(currWord.Text, "[\w\x{4e00}-\x{9fff}'’\-]$"))
            break
        gap := nextWord.x - (currWord.x + currWord.w)
        hasJoinChar := RegExMatch(currWord.Text, "['’\-]$") || RegExMatch(nextWord.Text, "^['’\-]")
        limit := hasJoinChar ? joinGap : tightGap
        if (gap >= minOverlap && gap <= limit) {
            bestWord := bestWord . nextWord.Text
            tempIndex++
        } else {
            break
        }
    }

    cleaned := WL_CleanWord(bestWord)
    if (cleaned = "")
        return false

    if (uiaContext != "" && (WL_ContainsWord(uiaContext, cleaned) || ((StrLen(cleaned) >= 3 || bestInRect) && (InStr(uiaContext, rawToken) || InStr(uiaContext, cleaned))))) {
        outLine := WL_ExtractTargetSentence(uiaContext, cleaned)
        outSrcTag := "ocr+uiactx"
    } else {
        outLine := WL_BuildOcrContext(lines, bestLineIndex, bestIndex, cleaned)
        outSrcTag := "ocr-lines"
    }
    outWord := cleaned
    return true
}

; ===== 按比例字体宽度的近似字符宽度单位估算 =====
WL_EstimateTokenUnits(token)
{
    total := 0.0
    Loop Parse, token {
        ch := A_LoopField
        if RegExMatch(ch, "[\x{4e00}-\x{9fff}\x{3000}-\x{303F}\x{FF00}-\x{FFEF}]") {
            total += 1.9
        } else if (ch == "M" || ch == "W") {
            total += 1.45
        } else if (ch == "I") {
            total += 0.55
        } else if RegExMatch(ch, "[A-Z]") {
            total += 1.15
        } else if (ch == "m" || ch == "w") {
            total += 1.35
        } else if RegExMatch(ch, "[ijltfr1!|:;.,'`"()\[\]/\\\-]") {
            total += 0.55
        } else {
            total += 0.95
        }
    }
    return Max(0.5, total)
}

; ===== 按行宽容量贪心折行（记录每行累计宽度与每个 token 的起止单位） =====
WL_WrapTokensToLines(items, capUnits, spaceUnits := 0.55)
{
    lines := []
    currTokens := []
    currWidth := 0.0

    for it in items {
        nextStart := (currWidth > 0) ? (currWidth + spaceUnits) : 0.0
        if (currWidth > 0 && nextStart + it.w > capUnits) {
            lines.Push({tokens: currTokens, width: currWidth})
            currTokens := []
            currWidth := 0.0
            nextStart := 0.0
        }
        endU := nextStart + it.w
        currTokens.Push({clean: it.clean, startU: nextStart, endU: endU})
        currWidth := endU
    }
    if (currTokens.Length > 0)
        lines.Push({tokens: currTokens, width: currWidth})
    return lines
}

; ===== 模拟按行折行与字符比例权重从 UIA 标签矩形内定位鼠标下的单词 =====
WL_SplitTokenByMousePos(uiaContext, uiaRect, mouseX, mouseY, dpiScale)
{
    if (uiaContext = "" || StrLen(uiaContext) > 240)
        return ""

    items := []
    fallbackToken := ""
    validWordCount := 0
    for pt in StrSplit(Trim(uiaContext), " ", " `t") {
        if (pt = "")
            continue
        clean := WL_CleanWord(pt)
        if (clean != "") {
            validWordCount++
            if (fallbackToken = "")
                fallbackToken := clean
        }
        items.Push({raw: pt, clean: clean, w: WL_EstimateTokenUnits(pt)})
    }
    if (validWordCount = 0 || items.Length > 45)
        return ""
    if (validWordCount = 1 || !IsObject(uiaRect))
        return fallbackToken

    wRect := uiaRect.r - uiaRect.l
    hRect := uiaRect.b - uiaRect.t
    if (wRect <= 0)
        return fallbackToken

    estLines := (hRect <= Round(38 * dpiScale)) ? 1 : Max(1, Min(4, Round(hRect / (28 * dpiScale))))
    lineH := Max(12 * dpiScale, hRect / estLines)
    lineCapUnits := Max(6.0, wRect / (lineH * 0.38))

    wrapped := WL_WrapTokensToLines(items, lineCapUnits)
    if (wrapped.Length > estLines) {
        Loop 4 {
            lineCapUnits *= 1.12
            wrapped := WL_WrapTokensToLines(items, lineCapUnits)
            if (wrapped.Length <= estLines)
                break
        }
    }

    relX := Max(0.0, Min(0.9999, (mouseX - uiaRect.l) / wRect))
    if (wrapped.Length == 1) {
        targetLineIdx := 1
    } else {
        lineIdx := Max(0, Min(wrapped.Length - 1, Floor((mouseY - uiaRect.t) / Max(1, hRect) * wrapped.Length)))
        targetLineIdx := lineIdx + 1
    }

    targetLine := wrapped[targetLineIdx]
    if (wrapped.Length > 1) {
        lineSpanUnits := Max(targetLine.width, lineCapUnits)
    } else {
        lineSpanUnits := (targetLine.width < lineCapUnits * 0.78) ? lineCapUnits : targetLine.width
    }
    targetU := relX * lineSpanUnits

    bestClean := ""
    bestDist := 999999.0
    for tk in targetLine.tokens {
        if (tk.clean = "")
            continue
        dist := Max(0.0, tk.startU - targetU, targetU - tk.endU)
        if (dist < bestDist) {
            bestDist := dist
            bestClean := tk.clean
            if (dist == 0.0)
                break
        }
    }
    return (bestClean != "") ? bestClean : fallbackToken
}

; ===== 切换图钉状态 =====
WL_TogglePin()
{
    global g_WL_IsPinned, g_WL_PinBtn
    g_WL_IsPinned := !g_WL_IsPinned
    if (g_WL_PinBtn) {
        g_WL_PinBtn.Text := g_WL_IsPinned ? "📍" : "📌"
        g_WL_PinBtn.SetFont("s9 " . (g_WL_IsPinned ? "cCC0000 Bold" : "c333333 Norm"), "Microsoft YaHei")
    }
}

; ===== 支持拖动窗口 =====
WL_WM_LBUTTONDOWN(wParam, lParam, msg, hwnd)
{
    global g_WL_Gui
    if (g_WL_Gui != "" && hwnd == g_WL_Gui.Hwnd) {
        PostMessage(0xA1, 2, 0, , "ahk_id " . hwnd)
    }
}

; ===== 拦截上下文菜单（阻止默认的系统文本框右键菜单） =====
WL_WM_CONTEXTMENU(wParam, lParam, msg, hwnd)
{
    global g_WL_Gui
    if (g_WL_Gui != "") {
        isOurGui := false
        if (hwnd == g_WL_Gui.Hwnd || wParam == g_WL_Gui.Hwnd) {
            isOurGui := true
        } else {
            try {
                ctrl := GuiCtrlFromHwnd(hwnd)
                if (ctrl && ctrl.Gui && ctrl.Gui.Hwnd == g_WL_Gui.Hwnd)
                    isOurGui := true
                ctrlW := GuiCtrlFromHwnd(wParam)
                if (ctrlW && ctrlW.Gui && ctrlW.Gui.Hwnd == g_WL_Gui.Hwnd)
                    isOurGui := true
            }
        }
        if (!isOurGui && WinActive("ahk_id " . g_WL_Gui.Hwnd)) {
            isOurGui := true
        }

        if (isOurGui) {
            global g_WL_InitMouseX, g_WL_InitMouseY, g_WL_MouseMoved, g_WL_ShowTick
            WL_HandleRightClick()
            CoordMode("Mouse", "Screen")
            MouseGetPos(&g_WL_InitMouseX, &g_WL_InitMouseY)
            g_WL_MouseMoved := false
            g_WL_ShowTick := A_TickCount
            return 0
        }
    }
}

; ===== 拦截底层 Edit 控件的右键按下和抬起，防止其私自呼出菜单 =====
WL_WM_RBUTTON(wParam, lParam, msg, hwnd)
{
    global g_WL_Gui
    if (g_WL_Gui != "") {
        try {
            ctrl := GuiCtrlFromHwnd(hwnd)
            ; 若目标是我们查词窗口内的控件
            if (ctrl && ctrl.Gui && ctrl.Gui.Hwnd == g_WL_Gui.Hwnd) {
                if (msg == 0x0204) { ; WM_RBUTTONDOWN
                    global g_WL_InitMouseX, g_WL_InitMouseY, g_WL_MouseMoved, g_WL_ShowTick
                    WL_HandleRightClick()
                    CoordMode("Mouse", "Screen")
                    MouseGetPos(&g_WL_InitMouseX, &g_WL_InitMouseY)
                    g_WL_MouseMoved := false
                    g_WL_ShowTick := A_TickCount
                }
                return 0 ; 让控件彻底忽略右键事件
            }
        }
    }
}

; ===== 关闭浮窗 =====
CloseWordGui()
{
  global g_WL_Gui, g_WL_ResultCtrl, g_WL_TitleCtrl
  global g_WL_StreamPid, g_WL_Pending, g_WL_StreamFile
  global g_WL_WordEdit, g_WL_ContextEdit
  global g_StreamPidChat, g_ChatPending, g_QuestionEditCtrl, g_AnswerEditCtrl, g_SendBtnCtrl
  global g_TtsPlaying, g_HoverTarget, g_MainGui, g_PromptDropdown, g_WL_AnkiBtn, g_OrigEditCtrl
  global g_WL_TtsPlayPending

  ; 终止请求
  if (g_WL_StreamPid > 0) {
    try ProcessClose(g_WL_StreamPid)
    g_WL_StreamPid := 0
  }
  if (g_StreamPidChat > 0) {
    try ProcessClose(g_StreamPidChat)
    g_StreamPidChat := 0
  }
  g_WL_Pending := false
  g_ChatPending := false
  SetTimer(CheckWordResult, 0)
  SetTimer(CheckChatResult, 0)
  SetTimer(WL_CheckClickOutside, 0)
  SetTimer(CheckTtsHover, 0)
  g_WL_TtsPlayPending := false
  WL_StopSentenceTts()
  try SoundPlay("NonExistent.zzz")
  StopLocalSapi()
  g_TtsPlaying := false
  g_HoverTarget := ""

  ; 彻底清理临时文件
  CurlCleanupTempFiles("wl")

  ; 销毁 GUI
  if (g_WL_Gui != "") {
    ; 共享控件变量（g_QuestionEditCtrl/g_AnswerEditCtrl 等）同时被翻译主窗口使用：
    ; 仅当它们仍指向本浮窗（g_MainGui 即 g_WL_Gui）时才清空。
    ; 若翻译窗口在浮窗之后打开过，这些变量已指向翻译窗口的控件，
    ; 粗暴置空会使翻译窗口的 Tab/^v/发送/Enter 访问空引用而抛错
    sharedOwnedByUs := false
    try sharedOwnedByUs := (g_MainGui != "" && g_MainGui.Hwnd == g_WL_Gui.Hwnd)
    try {
      HotIfWinActive("ahk_id " g_WL_Gui.Hwnd)
      Hotkey("Escape", WL_HandleEsc, "Off")
      Hotkey("Enter", WL_HandleEnter, "Off")
      Hotkey("NumpadEnter", WL_HandleEnter, "Off")
      Hotkey("!Left", "Off")
      Hotkey("!Right", "Off")
      ; ShowWordPopup 里注册的其余窗口专属热键也要注销，否则会在已销毁的 hwnd 上累积
      for hk in ["Tab", "^Tab", "^v", "^Backspace", "^s"]
        try Hotkey(hk, "Off")
      HotIfWinActive()
    }
    ; 注销全局右键拦截
    try Hotkey("RButton", "Off")
    try g_WL_Gui.Destroy()
    g_WL_Gui := ""
    g_WL_ResultCtrl := ""
    g_WL_TitleCtrl := ""
    g_WL_WordEdit := ""
    g_WL_ContextEdit := ""
    g_WL_AnkiBtn := ""
    if (sharedOwnedByUs) {
      g_QuestionEditCtrl := ""
      g_AnswerEditCtrl := ""
      g_SendBtnCtrl := ""
      g_PromptDropdown := ""
      g_MainGui := ""
      g_OrigEditCtrl := ""  ; 它指向本浮窗的 g_WL_ContextEdit，一并清理
    }
  }
}

; ===== 查词单词预生成与整句朗读状态（配合 word_lookup.ahk 里的基础 g_WL_Tts* 变量） =====
global g_WL_TtsState := "idle"       ; "idle" / "loading" / "ready" / "failed"
global g_WL_TtsStage := ""           ; "" / "youdao" / "google"
global g_WL_TtsPlayPending := false  ; 加载完成后是否自动触发播放
global g_WL_TtsTmpYoudao := ""
global g_WL_TtsTmpGoogle := ""
global g_WL_TtsFallbackFile := ""
global g_WL_TtsTargetCache := ""
global g_WL_TtsLang := "en-US"
global g_WL_SentPid := 0             ; 整句 Google TTS 当前 curl PID
global g_WL_SentText := ""
global g_WL_SentJob := ""

WL_ResetPregenTempFiles()
{
  global g_WL_TtsTmpYoudao, g_WL_TtsTmpGoogle, g_WL_TtsFallbackFile
  TtsDeleteTempFile(g_WL_TtsTmpYoudao)
  TtsDeleteTempFile(g_WL_TtsTmpGoogle)
  TtsDeleteTempFile(g_WL_TtsFallbackFile)
  g_WL_TtsTmpYoudao := ""
  g_WL_TtsTmpGoogle := ""
  g_WL_TtsFallbackFile := ""
}

WL_ApplyPregenFallback()
{
  global g_WL_TtsFallbackFile, g_WL_TtsFile, g_WL_TtsState
  fb := g_WL_TtsFallbackFile
  g_WL_TtsFallbackFile := ""
  if (fb != "" && TtsGetMp3Info(fb).valid) {
    g_WL_TtsFile := fb
    g_WL_TtsState := "ready"
  } else {
    TtsDeleteTempFile(fb)
    g_WL_TtsState := "failed"
  }
}

; ===== 预生成 TTS 音频（缓存 → 有道高质量真人 → Google 神经网络自动升级） =====
WL_PregenTts(word)
{
  global g_WL_TtsFile, g_WL_TtsPid, g_WL_TtsWord, g_WL_TtsGenTick
  global g_WL_TtsState, g_WL_TtsStage, g_WL_TtsPlayPending
  global g_WL_TtsTmpYoudao, g_WL_TtsTmpGoogle, g_WL_TtsTargetCache, g_WL_TtsLang

  SetTimer(WL_PollPregenTts, 0)
  if (g_WL_TtsPid > 0) {
    try ProcessClose(g_WL_TtsPid)
    g_WL_TtsPid := 0
  }
  g_WL_TtsPlayPending := false

  oldFile := g_WL_TtsFile
  g_WL_TtsFile := ""
  TtsDeleteTempFile(oldFile)
  WL_ResetPregenTempFiles()

  text := Trim(word)
  if (text = "") {
    g_WL_TtsWord := ""
    g_WL_TtsState := "idle"
    g_WL_TtsStage := ""
    return
  }

  info := TtsResolveCacheInfo(text)
  g_WL_TtsWord := text
  g_WL_TtsLang := info.lang
  g_WL_TtsTargetCache := info.cacheFile

  if TtsCacheHit(info.cacheFile, info.isYoudao) {
    g_WL_TtsFile := info.cacheFile
    g_WL_TtsState := "ready"
    g_WL_TtsStage := ""
    return
  }

  g_WL_TtsState := "loading"
  if (info.isYoudao) {
    tmpYd := A_Temp . "\ahk_wl_tts_yd_" . A_TickCount . ".mp3"
    outPid := RunYoudaoWordAudio(text, tmpYd, info.voiceType)
    if (outPid) {
      g_WL_TtsTmpYoudao := tmpYd
      g_WL_TtsStage := "youdao"
      g_WL_TtsPid := outPid
      g_WL_TtsGenTick := A_TickCount
      SetTimer(WL_PollPregenTts, 50)
      return
    }
  }

  if (StrLen(text) <= 190) {
    tmpG := A_Temp . "\ahk_wl_tts_g_" . A_TickCount . ".mp3"
    outPid := RunGoogleTtsAudio(text, tmpG, info.lang)
    if (outPid) {
      g_WL_TtsTmpGoogle := tmpG
      g_WL_TtsStage := "google"
      g_WL_TtsPid := outPid
      g_WL_TtsGenTick := A_TickCount
      SetTimer(WL_PollPregenTts, 50)
      return
    }
  }

  g_WL_TtsState := "failed"
  g_WL_TtsStage := ""
}

; ===== 后台轮询单词预生成，低码率或失败时自动升级 Google TTS =====
WL_PollPregenTts()
{
  global g_WL_TtsFile, g_WL_TtsPid, g_WL_TtsWord, g_WL_TtsGenTick, g_TtsMinGoodBitrate
  global g_WL_TtsState, g_WL_TtsStage, g_WL_TtsPlayPending
  global g_WL_TtsTmpYoudao, g_WL_TtsTmpGoogle, g_WL_TtsFallbackFile, g_WL_TtsTargetCache, g_WL_TtsLang

  if (g_WL_TtsPid <= 0) {
    SetTimer(WL_PollPregenTts, 0)
    return
  }

  timeoutMs := (g_WL_TtsStage = "youdao") ? 2500 : 4500
  if (ProcessExist(g_WL_TtsPid)) {
    if (A_TickCount - g_WL_TtsGenTick <= timeoutMs)
      return
    try ProcessClose(g_WL_TtsPid)
  }
  g_WL_TtsPid := 0

  if (g_WL_TtsStage = "youdao") {
    ydFile := g_WL_TtsTmpYoudao
    g_WL_TtsTmpYoudao := ""
    mp3 := TtsGetMp3Info(ydFile)
    if (mp3.valid && mp3.bitrateKbps >= g_TtsMinGoodBitrate) {
      SetTimer(WL_PollPregenTts, 0)
      g_WL_TtsFile := TtsSaveToCache(ydFile, g_WL_TtsTargetCache)
      g_WL_TtsStage := ""
      g_WL_TtsState := "ready"
      if (g_WL_TtsPlayPending) {
        g_WL_TtsPlayPending := false
        WL_PlayCurrentWordAudio()
      }
      return
    }
    if (mp3.valid)
      g_WL_TtsFallbackFile := ydFile
    else
      TtsDeleteTempFile(ydFile)

    if (StrLen(g_WL_TtsWord) <= 190) {
      tmpG := A_Temp . "\ahk_wl_tts_g_" . A_TickCount . ".mp3"
      outPid := RunGoogleTtsAudio(g_WL_TtsWord, tmpG, g_WL_TtsLang)
      if (outPid) {
        g_WL_TtsTmpGoogle := tmpG
        g_WL_TtsStage := "google"
        g_WL_TtsPid := outPid
        g_WL_TtsGenTick := A_TickCount
        return
      }
    }

    SetTimer(WL_PollPregenTts, 0)
    g_WL_TtsStage := ""
    WL_ApplyPregenFallback()
    if (g_WL_TtsPlayPending) {
      g_WL_TtsPlayPending := false
      WL_PlayCurrentWordAudio()
    }
    return
  }

  if (g_WL_TtsStage = "google") {
    SetTimer(WL_PollPregenTts, 0)
    gFile := g_WL_TtsTmpGoogle
    g_WL_TtsTmpGoogle := ""
    if (TtsGetMp3Info(gFile).valid) {
      g_WL_TtsFile := TtsSaveToCache(gFile, g_WL_TtsTargetCache)
      TtsDeleteTempFile(g_WL_TtsFallbackFile)
      g_WL_TtsFallbackFile := ""
      g_WL_TtsState := "ready"
    } else {
      TtsDeleteTempFile(gFile)
      WL_ApplyPregenFallback()
    }
    g_WL_TtsStage := ""
    if (g_WL_TtsPlayPending) {
      g_WL_TtsPlayPending := false
      WL_PlayCurrentWordAudio()
    }
    return
  }

  SetTimer(WL_PollPregenTts, 0)
}

WL_PlayCurrentWordAudio()
{
  global WL_CurrentWord, g_WL_TtsWord, g_WL_TtsFile
  if (g_WL_TtsFile != "" && TtsGetMp3Info(g_WL_TtsFile).valid) {
    TtsPlayAudioFileNow(g_WL_TtsFile)
    return
  }
  txt := Trim(WL_CurrentWord)
  if (txt = "")
    txt := g_WL_TtsWord
  if (txt != "")
    PlayLocalSapi(txt)
}

; ===== 强制单次朗读单词（右键触发，就绪即秒播，加载中则在完成瞬间自动播） =====
WL_PlayTtsOnce()
{
  global WL_CurrentWord, g_WL_TtsWord, g_WL_TtsState, g_WL_TtsPlayPending

  WL_StopSentenceTts()
  text := Trim(WL_CurrentWord)
  if (text = "")
    return

  if (text != g_WL_TtsWord || g_WL_TtsState = "idle")
    WL_PregenTts(text)

  if (g_WL_TtsState = "loading") {
    g_WL_TtsPlayPending := true
    return
  }

  g_WL_TtsPlayPending := false
  WL_PlayCurrentWordAudio()
}

; ===== 全局右键拦截处理函数（由 Hotkey() 动态注册） =====
WL_RButtonHandler(*)
{
  global g_WL_InitMouseX, g_WL_InitMouseY, g_WL_MouseMoved, g_WL_ShowTick
  
  ; 1. 统一右键朗读长按/短按处理中心
  WL_HandleRightClick()
  
  ; 2. 重置自动关闭防抖动计时器（防止因触发而导致抖动退出）
  CoordMode("Mouse", "Screen")
  MouseGetPos(&g_WL_InitMouseX, &g_WL_InitMouseY)
  g_WL_MouseMoved := false
  g_WL_ShowTick := A_TickCount
}

; ===== 统一右键长按/短按处理中心 =====
WL_HandleRightClick()
{
  global WL_CurrentWord, WL_CurrentContext
  
  ; 探测 RButton 是否已经被释放，设置 0.3 秒判定界限
  isClick := KeyWait("RButton", "T0.3")
  
  if (isClick) {
    ; 单击 -> 朗读单词
    WL_PlayTtsOnce()
  } else {
    ; 长按 -> 朗读句子
    if (IsSet(WL_CurrentContext) && WL_CurrentContext != "") {
      WL_PlaySentenceTts(WL_CurrentContext)
    } else {
      ; 无语境时，降级朗读单字
      WL_PlayTtsOnce()
    }
    ; 等待释放，防止松开时触发二级事件或干扰
    KeyWait("RButton")
  }
}

WL_StopSentenceTts()
{
  global g_WL_SentPid, g_WL_SentText, g_WL_SentJob
  SetTimer(WL_PollSentenceTts, 0)
  if IsObject(g_WL_SentJob)
    TtsGoogleJobStop(g_WL_SentJob)
  g_WL_SentJob := ""
  if (g_WL_SentPid > 0 && ProcessExist(g_WL_SentPid))
    try ProcessClose(g_WL_SentPid)
  g_WL_SentPid := 0
  g_WL_SentText := ""
}

; ===== 朗读完整语境句子（缓存 → Google 神经网络分片 → 本地 SAPI 兜底；保留后台单词预生成不中断） =====
WL_PlaySentenceTts(sentence)
{
  global g_WL_TtsPlayPending, g_WL_SentPid, g_WL_SentText, g_WL_SentJob

  WL_StopSentenceTts()
  g_WL_TtsPlayPending := false
  try SoundPlay("NonExistent.zzz")
  StopLocalSapi()

  sentence := Trim(RegExReplace(sentence, "[\r\n\s]+", " "))
  if (sentence = "")
    return

  info := TtsResolveCacheInfo(sentence)
  sentCache := TtsGetCachePath("s:" . sentence, info.lang)
  if TtsCacheHit(sentCache, false) {
    TtsPlayAudioFileNow(sentCache)
    return
  }

  chunks := TtsSplitTextForGoogle(sentence, 180, 3)
  job := TtsGoogleJobStart(chunks, info.lang, sentCache, "ahk_wl_sent")
  if IsObject(job) {
    g_WL_SentJob := job
    g_WL_SentPid := job.pid
    g_WL_SentText := sentence
    SetTimer(WL_PollSentenceTts, 50)
    return
  }

  PlayLocalSapi(sentence)
}

WL_PollSentenceTts()
{
  global g_WL_SentPid, g_WL_SentText, g_WL_SentJob
  if !IsObject(g_WL_SentJob) {
    SetTimer(WL_PollSentenceTts, 0)
    g_WL_SentPid := 0
    return
  }
  step := TtsGoogleJobStep(g_WL_SentJob, 4500)
  g_WL_SentPid := g_WL_SentJob.pid
  if (!step.done)
    return

  SetTimer(WL_PollSentenceTts, 0)
  g_WL_SentPid := 0
  sentText := g_WL_SentText
  if (step.ok) {
    TtsPlayAudioFileNow(step.file)
  } else {
    TtsGoogleJobStop(g_WL_SentJob)
    g_WL_SentJob := ""
    if (sentText != "")
      PlayLocalSapi(sentText)
  }
}


; ===== 历史记录导航 =====
WL_NavHistory(dir)
{
  global g_WL_History, g_WL_HistoryIdx, g_WL_WordEdit, g_WL_ContextEdit, g_WL_ResultCtrl
  global WL_CurrentWord, WL_CurrentContext, g_WL_LangMode

  if (g_WL_History.Length = 0)
    return

  newIdx := g_WL_HistoryIdx + dir
  if (newIdx < 1)
    newIdx := 1
  if (newIdx > g_WL_History.Length)
    newIdx := g_WL_History.Length

  if (newIdx == g_WL_HistoryIdx)
    return

  g_WL_HistoryIdx := newIdx
  item := g_WL_History[newIdx]

  WL_CurrentWord := item.word
  WL_CurrentContext := item.context

  if (g_WL_WordEdit != "")
    g_WL_WordEdit.Value := item.word
  if (g_WL_ContextEdit != "")
    g_WL_ContextEdit.Value := item.context
  
  global g_QuestionEditCtrl
  if (g_QuestionEditCtrl != "") {
    g_QuestionEditCtrl.Value := item.word
    g_QuestionEditCtrl.Focus()
    SendMessage(0x00B1, -1, -1, g_QuestionEditCtrl.Hwnd)
  }

  if (item.result != "") {
    if (g_WL_ResultCtrl != "")
      g_WL_ResultCtrl.Value := item.result
    
    ; 终止后台可能还在进行的请求，直接显示保存的结果
    global g_WL_StreamPid, g_WL_Pending
    if (g_WL_StreamPid > 0) {
      try ProcessClose(g_WL_StreamPid)
      g_WL_StreamPid := 0
    }
    g_WL_Pending := false
    SetTimer(CheckWordResult, 0)

  } else {
    if (g_WL_ResultCtrl != "")
      g_WL_ResultCtrl.Value := (g_WL_LangMode = "EN" ? "⏳ Querying..." : "⏳ 正在查询...")
    StartWordOllamaRequest(item.word, item.context, true)
  }

  WL_PregenTts(item.word)

  ; 导航操作视同活跃操作，重置自动关闭的计时器
  global g_WL_InitMouseX, g_WL_InitMouseY, g_WL_MouseMoved, g_WL_ShowTick
  CoordMode("Mouse", "Screen")
  MouseGetPos(&g_WL_InitMouseX, &g_WL_InitMouseY)
  g_WL_MouseMoved := false
  g_WL_ShowTick := A_TickCount
}

; ===== 发送至 Anki 的核心通信模块 =====
WL_SendToAnki(*)
{
    global g_WL_WordEdit, g_WL_ContextEdit, g_WL_ResultCtrl
    global g_WL_TtsFile, g_WL_AnkiBtn

    if (!g_WL_WordEdit || !g_WL_ResultCtrl || !g_WL_AnkiBtn)
        return

    word := Trim(g_WL_WordEdit.Value)
    context := Trim(g_WL_ContextEdit.Value)
    explanation := Trim(g_WL_ResultCtrl.Value)

    if (word = "")
        return

    isAdd := InStr(g_WL_AnkiBtn.Text, "➕") || InStr(g_WL_AnkiBtn.Text, "添加")

    ; 从配置文件动态读取 Anki 卡片类型映射关系
    ; （用 IniReadUtf8 而非 IniRead：文件是无 BOM 的 UTF-8，IniRead 会把中文读成乱码）
    deckName := "英语生词"
    deckName := IniReadUtf8(A_ScriptDir . "\ollama_config.ini", "Anki", "DeckName", deckName)
    
    modelName := "问答题"
    modelName := IniReadUtf8(A_ScriptDir . "\ollama_config.ini", "Anki", "ModelName", modelName)
    
    frontField := "正面"
    frontField := IniReadUtf8(A_ScriptDir . "\ollama_config.ini", "Anki", "FrontField", frontField)
    
    backField := "背面"
    backField := IniReadUtf8(A_ScriptDir . "\ollama_config.ini", "Anki", "BackField", backField)
    
    ; 兜底防乱码
    if (InStr(modelName, "闁") || InStr(modelName, "瓟") || InStr(modelName, "ue1be") || InStr(modelName, "u95c2")) {
        deckName := "英语生词"
        modelName := "问答题"
        frontField := "正面"
        backField := "背面"
    }

    try {
        http := ComObject("WinHttp.WinHttpRequest.5.1")
        
        if (isAdd) {
            if (explanation = "" || InStr(explanation, "Querying") || InStr(explanation, "正在查询")) {
                return
            }

            ; 格式化卡片文本
            frontText := "<h2>" . word . "</h2>"
            if (context != "" && context != word)
                frontText .= "<br><br><span style='color:grey;'>" . StrReplace(context, "`n", "<br>") . "</span>"
            
            backText := StrReplace(explanation, "`n", "<br>")

            ; 文本转 JSON 安全字符串闭包
            EscapeJSON := (str) => StrReplace(StrReplace(StrReplace(StrReplace(str, "\", "\\"), "`n", "\n"), "`r", ""), "`"", "\`"")

            frontJson := EscapeJSON(frontText)
            backJson := EscapeJSON(backText)

            ; 处理音频（仅附带通过 MP3 帧头校验的合法文件）
            audioJson := ""
            if (g_WL_TtsFile != "" && TtsGetMp3Info(g_WL_TtsFile).valid) {
                absPath := StrReplace(g_WL_TtsFile, "\", "\\")
                ; 文件名安全化：去掉引号/反斜杠/空格等非法字符，避免破坏 JSON 报文与文件名
                safeName := RegExReplace(word, "[^\w\x{4e00}-\x{9fff}]", "_")
                audioJson := ',"audio": [{"path": "' . absPath . '", "filename": "ahk_tts_' . safeName . '.mp3", "fields": ["' . frontField . '"]}]'
            }

            ; 构建 JSON 报文
            payload := '{"action": "addNote", "version": 6, "params": {"note": {"deckName": "' . deckName . '", "modelName": "' . modelName . '", "fields": {"' . frontField . '": "' . frontJson . '", "' . backField . '": "' . backJson . '"}, "options": {"allowDuplicate": false}, "tags": ["AHK抓取"]' . audioJson . '}}}'

            try {
                deckPayload := '{"action": "createDeck", "version": 6, "params": {"deck": "' . deckName . '"}}'
                http.Open("POST", "http://127.0.0.1:8765", false)
                http.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
                http.Send(deckPayload)
                http.WaitForResponse()
            }

            http.Open("POST", "http://127.0.0.1:8765", false)
            http.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
            http.Send(payload)
            http.WaitForResponse()
            res := http.ResponseText
            
            if (InStr(res, '"error": null')) {
                g_WL_AnkiBtn.Text := "➖ Anki"
                g_WL_AnkiBtn.SetFont("c008800 Norm")
            } else if (InStr(res, "cannot create note because it is a duplicate")) {
                g_WL_AnkiBtn.Text := "➖ Anki"
                g_WL_AnkiBtn.SetFont("c008800 Norm")
            } else {
            }
        } else {
            ; 删除逻辑（仅搜索正面字段，避免释义误匹配）
            payload := WL_BuildAnkiFindPayload(deckName, frontField, word)
            
            http.Open("POST", "http://127.0.0.1:8765", false)
            http.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
            http.Send(payload)
            http.WaitForResponse()
            res := http.ResponseText
            
            if (RegExMatch(res, '"result":\s*\[(.*?)\]', &m)) {
                ids := Trim(m[1])
                if (ids != "") {
                    delPayload := '{"action": "deleteNotes", "version": 6, "params": {"notes": [' . ids . ']}}'
                    http.Open("POST", "http://127.0.0.1:8765", false)
                    http.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
                    http.Send(delPayload)
                    http.WaitForResponse()
                    
                    g_WL_AnkiBtn.Text := "➕ Anki"
                    g_WL_AnkiBtn.SetFont("c333333 Norm")
                    return
                }
            }
            g_WL_AnkiBtn.Text := "➕ Anki"
            g_WL_AnkiBtn.SetFont("c333333 Norm")
        }
        http := "" ; 使用完毕，释放 COM 对象
    } catch Error as e {
        http := ""
    }
}

; ===== 构造 AnkiConnect findNotes 请求体 =====
; 单词先做正则转义（re: 搜索里 c++ 这类会成为非法正则），整条查询再做一次 JSON 转义。
; 原先对单词转义一次引号、再对整条查询转义一次引号，含引号的词会被双重转义而破坏 JSON
WL_BuildAnkiFindPayload(deckName, frontField, word)
{
    reWord := RegExReplace(word, "[\\^$.|?*+()\[\]{}]", "\$0")
    query := 'deck:"' . deckName . '" ' . frontField . ':re:<h2>' . reWord . '</h2>'
    return '{"action": "findNotes", "version": 6, "params": {"query": "' . EscapeJsonForApi(query) . '"}}'
}

WL_CheckAnkiStatus(word) {
    global g_WL_AnkiBtn, g_WL_Gui
    if (!g_WL_AnkiBtn || !g_WL_Gui) {
        return
    }

    ; 默认显示为未收藏
    try g_WL_AnkiBtn.Text := "➕ Anki"
    try g_WL_AnkiBtn.SetFont("c333333 Norm")

    ; 使用单次定时器延迟执行同步请求，避免阻塞 GUI 初始渲染
    ; （本地 AnkiConnect 请求通常 <50ms，不会卡界面）
    checkWord := word
    AnkiSyncCheck() {
        SetTimer(, 0) ; 确保只执行一次
        global g_WL_AnkiBtn, g_WL_Gui
        if (!g_WL_AnkiBtn || !g_WL_Gui)
            return

        try {
            deckName := "英语生词"
            deckName := IniReadUtf8(A_ScriptDir . "\ollama_config.ini", "Anki", "DeckName", deckName)
            if (InStr(deckName, "闁") || InStr(deckName, "ue1be")) {
                deckName := "英语生词"
            }

            frontField := "正面"
            frontField := IniReadUtf8(A_ScriptDir . "\ollama_config.ini", "Anki", "FrontField", frontField)

            payload := WL_BuildAnkiFindPayload(deckName, frontField, checkWord)

            http := ComObject("WinHttp.WinHttpRequest.5.1")
            http.Open("POST", "http://127.0.0.1:8765", false) ; 同步模式，可靠获取结果
            http.SetTimeouts(2000, 2000, 2000, 2000)
            http.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
            http.Send(payload)
            http.WaitForResponse()

            res := http.ResponseText
            http := ""

            if (!g_WL_AnkiBtn || !g_WL_Gui)
                return

            if (RegExMatch(res, '"result":\s*\[(.*?)\]', &m)) {
                ids := Trim(m[1])
                if (ids != "") {
                    try g_WL_AnkiBtn.Text := "➖ Anki"
                    try g_WL_AnkiBtn.SetFont("c008800 Norm")
                    return
                }
            }
        } catch {
            ; AnkiConnect 未运行或连接失败，保持默认 ➕ 状态
        }
    }

    SetTimer(AnkiSyncCheck, -1) ; 负数 = 单次触发，延迟 1ms 后执行
}

; ===== F2 / 组合键单双击判定状态机（记录可选鼠标快照供单按延迟执行，屏蔽长按自动重复与机械抖动） =====
F2_OnKeyDown(state, nowTick, snapX := 0, snapY := 0, snapWin := 0)
{
    if (state.isDown && (nowTick - state.downTick < 1000))
        return "ignore"
    state.isDown := true
    state.downTick := nowTick

    if (state.lastDoubleTick && (nowTick - state.lastDoubleTick < 250))
        return "ignore"

    if (state.lastTapTick && (nowTick - state.lastTapTick < 45))
        return "ignore"
    state.lastTapTick := nowTick

    if (state.pendingSingle) {
        state.pendingSingle := false
        state.lastDoubleTick := nowTick
        return "double"
    }

    state.pendingSingle := true
    state.snapX := snapX
    state.snapY := snapY
    state.snapWin := snapWin
    return "wait_single"
}

; ===== 退出清理 =====
; 终止本脚本仍在运行的子进程（curl），删除临时文件（含写有 API key 的 curl 配置文件）。
WL_CleanupOnExit(*)
{
  global g_WL_TtsPid, g_WL_SentPid, g_WL_StreamPid, g_StreamPidChat, g_TtsProcPid
  try SoundPlay("NonExistent.zzz")
  StopLocalSapi()
  for pid in [g_WL_TtsPid, g_WL_SentPid, g_WL_StreamPid, g_StreamPidChat, g_TtsProcPid] {
    try {
      if (pid > 0 && ProcessExist(pid)) {
        name := ProcessGetName(pid)
        if (name = "curl.exe")
          ProcessClose(pid)
      }
    }
  }
  CurlCleanupTempFiles("wl")
  CurlCleanupTempFiles("chat")
  ; 保留 %TEMP%\ahk_tts_cache\ 中的高音质缓存，仅清理临时下载与中间分片文件
  for name in ["ahk_wl_tts_*.mp3", "ahk_tts_tmp_*.mp3", "ahk_wl_sent_*.mp3", "ahk_tts_youdao.mp3", "ahk_wl_stream_word.txt", "ahk_wl_request_word.json"]
    try FileDelete(A_Temp . "\" . name)
}
OnExit(WL_CleanupOnExit)

