# 错题录入统一接口
# 供其他 agent 快速调用：传入一份错题 JSON，自动完成：
#   1. 更新错题库\错题索引.csv
#   2. 更新/新建错题库\<科目>\<章节>.md
#   3. 新建章节文件时同步错题库\README.md
#
# 用法：
#   .\录入错题.ps1 -InputFile .\错题.json
#   .\录入错题.ps1 -Json '{"subject":"操作系统","chapter":"第2章 进程与线程",...}'
#   Get-Content .\错题.json -Raw | .\录入错题.ps1
#
# JSON 字段：
#   必填：subject, chapter, number, topic, question, answer, wrongType
#   可选：date, myAnswer, wrongReason, reviewConclusion, options, code,
#         imageSource, questionType, reviewDate, file
#
# 示例：
# {
#   "date": "2026-08-31",
#   "subject": "操作系统",
#   "chapter": "第2章 进程与线程",
#   "number": "06",
#   "topic": "进程状态变化",
#   "question": "题干",
#   "options": { "A": "...", "B": "...", "C": "...", "D": "..." },
#   "answer": "C",
#   "myAnswer": "B",
#   "wrongType": "概念不清",
#   "wrongReason": "为什么错",
#   "reviewConclusion": "下次怎么判断"
# }

[CmdletBinding()]
param(
    [Parameter(ValueFromPipeline = $true)]
    [string]$InputObject,
    [string]$InputFile,
    [string]$Json,
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
$scriptDir = $PSScriptRoot
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$indexPath = Join-Path $scriptDir "错题索引.csv"
$readmePath = Join-Path $scriptDir "README.md"
$today = Get-Date -Format "yyyy-MM-dd"

# ---------- 读取 JSON ----------
$jsonText = ""

if ($InputFile) {
    $inputPath = $InputFile
    if (-not [System.IO.Path]::IsPathRooted($inputPath)) {
        $inputPath = Join-Path (Get-Location) $InputFile
    }
    if (-not (Test-Path -LiteralPath $inputPath)) {
        throw "找不到输入文件：$inputPath"
    }
    $jsonText = [System.IO.File]::ReadAllText($inputPath, [System.Text.Encoding]::UTF8)
}
elseif ($Json) {
    $jsonText = $Json
}
elseif ($InputObject) {
    $jsonText = $InputObject
}
elseif ($MyInvocation.ExpectingInput) {
    $jsonText = @($input | ForEach-Object { [string]$_ }) -join "`n"
}
else {
    throw "请通过 -InputFile 或 -Json 传入 JSON，或用管道传入 JSON。"
}

$parsed = $null
try {
    $parsed = $jsonText | ConvertFrom-Json
}
catch {
    throw "JSON 解析失败：$($_.Exception.Message)"
}

if ($null -eq $parsed) {
    throw "JSON 内容为空。"
}

$items = @($parsed)
if ($items.Count -eq 0) {
    throw "JSON 中没有可录入的错题。"
}

# ---------- 基础映射 ----------
$subjectFolderMap = @{
    "数据结构"     = "数据结构"
    "DS"          = "数据结构"
    "操作系统"     = "操作系统"
    "OS"          = "操作系统"
    "计算机组成原理" = "计算机组成原理"
    "CO"          = "计算机组成原理"
    "计算机网络"   = "计算机网络"
    "CN"          = "计算机网络"
}

$chapterFileMap = @{
    "DS:第1章 绪论"                   = "数据结构/DS-01-绪论.md"
    "DS:第2章 线性表"                 = "数据结构/DS-02-线性表.md"
    "DS:第3章 栈队列数组"             = "数据结构/DS-03-栈队列数组.md"
    "DS:第3章 栈、队列和数组"         = "数据结构/DS-03-栈队列数组.md"
    "DS:第4章 串"                     = "数据结构/DS-04-串.md"
    "DS:第4-5章 树与二叉树"           = "数据结构/DS-04-树与二叉树.md"
    "DS:第5章 树与二叉树"             = "数据结构/DS-04-树与二叉树.md"
    "DS:第6章 图"                     = "数据结构/DS-06-图.md"
    "DS:第7章 查找"                   = "数据结构/DS-07-查找.md"
    "DS:第8章 排序"                   = "数据结构/DS-08-排序.md"
    "OS:第1章 计算机系统概述"          = "操作系统/OS-01-计算机系统概述.md"
    "OS:第2章 进程与线程"             = "操作系统/OS-02-进程与线程.md"
    "OS:第3章 内存管理"               = "操作系统/OS-03-内存管理.md"
    "OS:第4章 文件管理"               = "操作系统/OS-04-文件管理.md"
    "OS:第5章 输入、输出管理"         = "操作系统/OS-05-输入输出管理.md"
    "OS:第5章 输入输出管理"           = "操作系统/OS-05-输入输出管理.md"
    "CO:第1章 计算机系统概述"          = "计算机组成原理/CO-01-计算机系统概述.md"
    "CO:第2章 数据的表示和运算"        = "计算机组成原理/CO-02-数据的表示和运算.md"
    "CO:第3章 存储系统"               = "计算机组成原理/CO-03-存储系统.md"
    "CO:第4章 指令系统"               = "计算机组成原理/CO-04-指令系统.md"
    "CO:第5章 中央处理器"             = "计算机组成原理/CO-05-中央处理器.md"
    "CO:第6章 总线"                   = "计算机组成原理/CO-06-总线.md"
    "CO:第7章 输入、输出系统"         = "计算机组成原理/CO-07-输入输出系统.md"
    "CN:第1章 计算机网络体系结构"      = "计算机网络/CN-01-计算机网络体系结构.md"
    "CN:第2章 物理层"                 = "计算机网络/CN-02-物理层.md"
    "CN:第3章 数据链路层"             = "计算机网络/CN-03-数据链路层.md"
    "CN:第4章 网络层"                 = "计算机网络/CN-04-网络层.md"
    "CN:第5章 传输层"                 = "计算机网络/CN-05-传输层.md"
    "CN:第6章 应用层"                 = "计算机网络/CN-06-应用层.md"
}

function Resolve-SubjectFolder {
    param([string]$Subject)
    if (-not $Subject) { throw "缺少 subject" }
    $key = $Subject.Trim()
    if (-not $subjectFolderMap.ContainsKey($key)) {
        throw "不支持的 subject：$Subject（支持：数据结构/DS、操作系统/OS、计算机组成原理/CO、计算机网络/CN）"
    }
    return $subjectFolderMap[$key]
}

function Resolve-ChapterFile {
    param(
        [string]$SubjectFolder,
        [string]$Chapter,
        [string]$ProvidedFile
    )

    if ($ProvidedFile) {
        $file = $ProvidedFile.Trim().Replace("\", "/")
        if ($file -notmatch "/") {
            $file = "$SubjectFolder/$file"
        }
        return $file
    }

    if (-not $Chapter) { throw "缺少 chapter" }
    $chapterTrim = $Chapter.Trim()

    # 优先从现有错题索引里复用章节文件路径，例如 DS 第 4-5 章映射到 DS-04
    if (Test-Path -LiteralPath $indexPath) {
        $existing = @(Import-Csv -Path $indexPath -Encoding UTF8 | Where-Object { $_.科目 -eq $SubjectFolder -and $_.王道章节 -eq $chapterTrim } | Select-Object -First 1)
        if ($existing.Count -gt 0 -and $existing[0].详解文件) {
            return $existing[0].详解文件
        }
    }

    # 从内置映射取
    foreach ($entry in $chapterFileMap.Keys) {
        $parts = $entry -split ":", 2
        if ($parts[0] -eq $SubjectFolder -and $parts[1] -eq $chapterTrim) {
            return $chapterFileMap[$entry]
        }
    }

    # 兜底：按“科目缩写-章号-清理后章名.md”生成
    $abbr = switch ($SubjectFolder) {
        "数据结构" { "DS" }
        "操作系统" { "OS" }
        "计算机组成原理" { "CO" }
        "计算机网络" { "CN" }
    }
    $numMatch = [regex]::Match($chapterTrim, "第(\d+)")
    if (-not $numMatch.Success) {
        throw "无法从 chapter 推断章节文件，请传入 file 字段：$chapterTrim"
    }
    $num = $numMatch.Groups[1].Value.PadLeft(2, "0")
    $title = $chapterTrim -replace "^第[\d\-]+章\s*", ""
    $title = $title -replace "[、，,]", ""
    return "$SubjectFolder/$abbr-$num-$title.md"
}

function ConvertTo-MdCell {
    param([AllowNull()][string]$Value)
    if ($null -eq $Value) { return "" }
    return $Value.Replace("|", "\|").Replace("`r`n", "<br>").Replace("`n", "<br>")
}

function ConvertTo-CsvLine {
    param($Row)
    return ($Row | ConvertTo-Csv -NoTypeInformation)[1]
}

function Write-CsvRow {
    param(
        [string]$Date,
        [string]$Subject,
        [string]$Chapter,
        [string]$Number,
        [string]$Topic,
        [string]$WrongType,
        [string]$Answer,
        [string]$ReviewDate,
        [string]$DetailFile
    )

    $row = [pscustomobject]@{
        日期       = $Date
        科目       = $Subject
        王道章节   = $Chapter
        题号       = $Number
        考点       = $Topic
        错误类型   = $WrongType
        正确答案   = $Answer
        复查日期   = $ReviewDate
        详解文件   = $DetailFile
    }

    $csvLine = ConvertTo-CsvLine -Row $row

    if (-not (Test-Path -LiteralPath $indexPath)) {
        $header = "日期,科目,王道章节,题号,考点,错误类型,正确答案,复查日期,详解文件"
        [System.IO.File]::WriteAllText($indexPath, $header + [Environment]::NewLine, $utf8NoBom)
    }
    else {
        $existing = [System.IO.File]::ReadAllText($indexPath, $utf8NoBom)
        if ($existing.Length -gt 0 -and -not $existing.EndsWith("`n")) {
            [System.IO.File]::AppendAllText($indexPath, [Environment]::NewLine, $utf8NoBom)
        }
    }

    [System.IO.File]::AppendAllText($indexPath, $csvLine + [Environment]::NewLine, $utf8NoBom)
}

function Get-OptionsTable {
    param($Options)
    if ($null -eq $Options) { return "" }
    $lines = @()
    foreach ($prop in $Options.PSObject.Properties | Sort-Object Name) {
        $lines += "| $($prop.Name) | $(ConvertTo-MdCell $prop.Value) |"
    }
    if ($lines.Count -eq 0) { return "" }
    return "| 选项 | 内容 |`n| --- | --- |`n" + ($lines -join "`n")
}

function Build-MarkdownEntry {
    param(
        [string]$Date,
        [string]$Number,
        [string]$Topic,
        [string]$Question,
        [string]$QuestionType,
        [string]$Answer,
        [string]$MyAnswer,
        [string]$WrongType,
        [string]$WrongReason,
        [string]$ReviewConclusion,
        $Options,
        [string]$Code,
        [string]$ImageSource
    )

    $entry = @"
## $Date｜$Number｜$Topic

| 项目 | 内容 |
| --- | --- |
| 题型 | $(ConvertTo-MdCell $QuestionType) |
| 正确答案 | $(ConvertTo-MdCell $Answer) |
| 我的选择 | $(ConvertTo-MdCell $MyAnswer) |
| 错误类型 | $(ConvertTo-MdCell $WrongType) |

**原题：$Question**
"@

    if ($ImageSource) {
        $entry += "`n`n图片来源：$ImageSource"
    }

    if ($Code) {
        $fence = '```'
        $entry += "`n`n$fence`n$Code`n$fence"
    }

    $optionsTable = Get-OptionsTable -Options $Options
    if ($optionsTable) {
        $entry += "`n`n$optionsTable"
    }

    $entry += "`n`n错误原因：<span style=`"color: red;`">$WrongReason</span>"
    $entry += "`n`n复盘结论：$ReviewConclusion"

    return $entry
}

function Update-ReadmeForNewFile {
    param([string]$RelativeFile)

    if (-not (Test-Path -LiteralPath $readmePath)) { return }

    $rel = $RelativeFile -replace "/", "\"
    $leaf = Split-Path $rel -Leaf
    $dirName = Split-Path $rel -Parent
    $content = [System.IO.File]::ReadAllText($readmePath, [System.Text.Encoding]::UTF8)

    # 已存在就不动
    if ($content.Contains($leaf)) { return }

    # 在文件结构代码块里找到对应科目目录行，插入缩进文件名
    $dirLinePattern = "(?m)^[ ]{2}" + [regex]::Escape($dirName + "/") + "[ ]*\r?$"
    if ($content -match $dirLinePattern) {
        $content = [regex]::Replace($content, "(?m)^([ ]{2}$([regex]::Escape($dirName + "/"))[ ]*)\r?\n", "`$1`n    $leaf`n", 1)
    }
    else {
        # 没有这个科目目录时，在根目录行后补一个科目目录
        $content = [regex]::Replace($content, "(?m)^(错题库/[ ]*)\r?\n", "`$1`n  $dirName/`n    $leaf`n", 1)
    }

    [System.IO.File]::WriteAllText($readmePath, $content, $utf8NoBom)
}

function Add-WrongQuestion {
    param($Data)

    $subject = [string]($Data.subject)
    $chapter = [string]($Data.chapter)
    $number = [string]($Data.number)
    $topic = [string]($Data.topic)
    $question = [string]($Data.question)
    $answer = [string]($Data.answer)
    $wrongType = [string]($Data.wrongType)

    if (-not $subject) { throw "缺少 subject" }
    if (-not $chapter) { throw "缺少 chapter" }
    if (-not $number) { throw "缺少 number" }
    if (-not $topic) { throw "缺少 topic" }
    if (-not $question) { throw "缺少 question" }
    if (-not $answer) { throw "缺少 answer" }
    if (-not $wrongType) { throw "缺少 wrongType" }

    $subjectFolder = Resolve-SubjectFolder -Subject $subject
    $subjectFull = $subjectFolder
    # 若传的是缩写，索引里仍存中文全称
    switch ($subjectFolder) {
        "数据结构" { $subjectFull = "数据结构" }
        "操作系统" { $subjectFull = "操作系统" }
        "计算机组成原理" { $subjectFull = "计算机组成原理" }
        "计算机网络" { $subjectFull = "计算机网络" }
    }

    $date = if ($Data.date) { [string]$Data.date } else { $today }
    $reviewDate = if ($Data.reviewDate) { [string]$Data.reviewDate } else { ([datetime]::ParseExact($date, "yyyy-MM-dd", $null)).AddDays(1).ToString("yyyy-MM-dd") }
    $questionType = if ($Data.questionType) { [string]$Data.questionType } else { "单项选择题" }
    $myAnswer = if ($null -ne $Data.myAnswer) { [string]$Data.myAnswer } else { "" }
    $wrongReason = if ($Data.wrongReason) { [string]$Data.wrongReason } else { "待补。" }
    $reviewConclusion = if ($Data.reviewConclusion) { [string]$Data.reviewConclusion } else { "待补。" }
    $code = if ($Data.code) { [string]$Data.code } else { "" }
    $imageSource = if ($Data.imageSource) { [string]$Data.imageSource } else { "" }
    $providedFile = if ($Data.file) { [string]$Data.file } else { "" }

    $relativeFile = Resolve-ChapterFile -SubjectFolder $subjectFolder -Chapter $chapter -ProvidedFile $providedFile
    $absoluteFile = Join-Path $scriptDir $relativeFile

    # 安全校验：不允许写入错题库目录之外
    $fullBase = [System.IO.Path]::GetFullPath($scriptDir).TrimEnd("\")
    $fullTarget = [System.IO.Path]::GetFullPath($absoluteFile).TrimEnd("\")
    if (-not $fullTarget.StartsWith($fullBase + "\", [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "章节文件路径超出错题库目录：$relativeFile"
    }

    $entry = Build-MarkdownEntry `
        -Date $date `
        -Number $number `
        -Topic $topic `
        -Question $question `
        -QuestionType $questionType `
        -Answer $answer `
        -MyAnswer $myAnswer `
        -WrongType $wrongType `
        -WrongReason $wrongReason `
        -ReviewConclusion $reviewConclusion `
        -Options $Data.options `
        -Code $code `
        -ImageSource $imageSource

    if ($DryRun) {
        Write-Host "[DryRun] 将录入：$subjectFull / $chapter / $topic"
        Write-Host "[DryRun] 索引行：$relativeFile"
        Write-Host $entry
        return
    }

    # 更新章节 md
    $dir = Split-Path -Parent $absoluteFile
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    if (Test-Path -LiteralPath $absoluteFile) {
        $sep = if ((Get-Content -Raw -Encoding UTF8 $absoluteFile).EndsWith("`n")) { "`r`n`r`n" } else { "`r`n`r`n" }
        [System.IO.File]::AppendAllText($absoluteFile, $sep + $entry + [Environment]::NewLine, $utf8NoBom)
    }
    else {
        $title = [System.IO.Path]::GetFileNameWithoutExtension($absoluteFile)
        $header = "# $title`r`n"
        [System.IO.File]::WriteAllText($absoluteFile, $header + $entry + [Environment]::NewLine, $utf8NoBom)
        Update-ReadmeForNewFile -RelativeFile $relativeFile
    }

    # 更新 CSV
    Write-CsvRow -Date $date -Subject $subjectFull -Chapter $chapter -Number $number -Topic $topic -WrongType $wrongType -Answer $answer -ReviewDate $reviewDate -DetailFile $relativeFile

    Write-Host "已录入：$subjectFull / $chapter / #$number / $topic"
    Write-Host "索引：$indexPath"
    Write-Host "详解：$relativeFile"
}

foreach ($item in $items) {
    Add-WrongQuestion -Data $item
}

if ($DryRun) {
    Write-Host "[DryRun] 未写入任何文件。"
}
