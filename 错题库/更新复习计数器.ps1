# 复习计数器自动更新脚本
# 用法：
#   1) 仅自动重算：
#      .\更新复习计数器.ps1
#   2) 记录某题今天复习了（会自动累计次数、滚动最近/上次日期）：
#      .\更新复习计数器.ps1 -MarkReview -Subject 操作系统 -Chapter "第1章 计算机系统概述" -Number 03 -ReviewTopic "模块化OS的特点"
#   3) 注册 Windows 计划任务（每天 08:30 自动运行）：
#      .\更新复习计数器.ps1 -InstallTask
#
# 输出文件：错题库\复习计数器.csv
# 该 CSV 不参与 CLAUDE.md 里错题索引的固定列顺序，只作为独立计数器。

[CmdletBinding()]
param(
    [switch]$MarkReview,
    [switch]$InstallTask,
    [string]$Subject = "",
    [string]$Chapter = "",
    [string]$Number = "",
    [string]$ReviewTopic = ""
)

$ErrorActionPreference = "Stop"
$scriptDir = $PSScriptRoot
$indexPath = Join-Path $scriptDir "错题索引.csv"
$counterPath = Join-Path $scriptDir "复习计数器.csv"
$today = Get-Date -Format "yyyy-MM-dd"

if ($InstallTask) {
    $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    $trigger = New-ScheduledTaskTrigger -Daily -At "08:30"
    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopOnIdleEnd
    Register-ScheduledTask -TaskName "408-ReviewCounter" -Action $action -Trigger $trigger -Settings $settings -Force
    Write-Host "已注册计划任务：408-ReviewCounter（每天 08:30 自动更新）"
    return
}

if (-not (Test-Path $indexPath)) {
    throw "找不到错题索引：$indexPath"
}

$index = @(Import-Csv -Path $indexPath -Encoding UTF8)
$oldCounter = @()
if (Test-Path $counterPath) {
    $oldCounter = @(Import-Csv -Path $counterPath -Encoding UTF8)
}

function Get-OldRecord($subject, $chapter, $number, $topic) {
    foreach ($r in $oldCounter) {
        if ($r.科目 -eq $subject -and $r.王道章节 -eq $chapter -and $r.题号 -eq $number -and $r.考点 -eq $topic) {
            return $r
        }
    }
    return $null
}

$records = foreach ($row in $index) {
    $subject = $row.科目
    $chapter = $row.王道章节
    $number = $row.题号
    $topic = $row.考点

    $old = Get-OldRecord $subject $chapter $number $topic
    $count = 0
    $last = ""
    $prev = ""

    if ($old) {
        $count = [int]$old.复习次数
        $last = $old.最近复习日期
        $prev = $old.上次复习日期
    }

    if ($MarkReview -and $Subject -eq $subject -and $Chapter -eq $chapter -and $Number -eq $number -and $ReviewTopic -eq $topic) {
        if ($last -ne $today) {
            $prev = $last
            $last = $today
            $count++
        }
    }

    $daysSince = ""
    if ($last) {
        $lastDate = [datetime]::ParseExact($last, "yyyy-MM-dd", $null)
        $daysSince = [int]((Get-Date).Date - $lastDate.Date).TotalDays
    }

    $lastInterval = ""
    if ($last -and $prev) {
        $lastDate = [datetime]::ParseExact($last, "yyyy-MM-dd", $null)
        $prevDate = [datetime]::ParseExact($prev, "yyyy-MM-dd", $null)
        $lastInterval = [int]($lastDate.Date - $prevDate.Date).TotalDays
    }

    [pscustomobject]@{
        科目            = $subject
        王道章节        = $chapter
        题号            = $number
        考点            = $topic
        复习次数        = $count
        最近复习日期    = $last
        上次复习日期    = $prev
        距上次复习天数  = $daysSince
        上次间隔天数    = $lastInterval
        更新日期        = $today
    }
}

$records | Export-Csv -Path $counterPath -NoTypeInformation -Encoding UTF8
Write-Host "已更新：$counterPath"
