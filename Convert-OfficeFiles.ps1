<#
.SYNOPSIS
    古いOffice形式(doc, xls, ppt)を新形式(docx, xlsx, pptx)に変換する

.DESCRIPTION
    - 引数なし: フォルダ選択ダイアログを表示
    - 引数あり: フォルダまたはファイルを指定して変換

.PARAMETER Path
    変換対象のフォルダまたはファイルのパス（複数指定可能）

.EXAMPLE
    .\Convert-OfficeFiles.ps1
    # フォルダ選択ダイアログが表示される

.EXAMPLE
    .\Convert-OfficeFiles.ps1 -Path "C:\Documents"
    # 指定フォルダ内のファイルを変換

.EXAMPLE
    .\Convert-OfficeFiles.ps1 -Path "C:\test.xls", "C:\test2.doc"
    # 指定ファイルを変換
#>

param(
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$Path
)

# 保存形式の定数
$wdFormatDocumentDefault = 16        # Word: docx
$xlOpenXMLWorkbook = 51              # Excel: xlsx
$ppSaveAsOpenXMLPresentation = 24    # PowerPoint: pptx

# 結果集計用
$script:successCount = 0
$script:errorCount = 0
$script:skipCount = 0
$script:errorFiles = @()

#------------------------------------------------------------------------------
# フォルダ選択ダイアログを表示
#------------------------------------------------------------------------------
function Show-FolderDialog {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "変換対象のフォルダを選択してください"
    $dialog.ShowNewFolderButton = $false

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.SelectedPath
    }
    return $null
}

#------------------------------------------------------------------------------
# COMオブジェクトを安全に解放
#------------------------------------------------------------------------------
function Release-ComObject {
    param([object]$obj)
    if ($null -ne $obj) {
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null
        }
        catch { }
    }
}

#------------------------------------------------------------------------------
# Wordファイルを変換 (doc → docx)
#------------------------------------------------------------------------------
function Convert-WordFile {
    param([string]$FilePath)

    $word = $null
    $doc = $null

    try {
        $outputPath = [System.IO.Path]::ChangeExtension($FilePath, ".docx")

        # 既存ファイルチェック
        if (Test-Path $outputPath) {
            Write-Host "  スキップ: 出力先が既に存在 - $outputPath" -ForegroundColor Yellow
            $script:skipCount++
            return
        }

        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0  # wdAlertsNone

        $doc = $word.Documents.Open($FilePath)
        $doc.SaveAs2([ref]$outputPath, [ref]$wdFormatDocumentDefault)
        $doc.Close($false)

        Write-Host "  完了: $outputPath" -ForegroundColor Green
        $script:successCount++
    }
    catch {
        Write-Host "  エラー: $($_.Exception.Message)" -ForegroundColor Red
        $script:errorCount++
        $script:errorFiles += $FilePath
    }
    finally {
        if ($null -ne $doc) {
            Release-ComObject $doc
        }
        if ($null -ne $word) {
            $word.Quit()
            Release-ComObject $word
        }
    }
}

#------------------------------------------------------------------------------
# Excelファイルを変換 (xls → xlsx)
#------------------------------------------------------------------------------
function Convert-ExcelFile {
    param([string]$FilePath)

    $excel = $null
    $workbook = $null

    try {
        $outputPath = [System.IO.Path]::ChangeExtension($FilePath, ".xlsx")

        # 既存ファイルチェック
        if (Test-Path $outputPath) {
            Write-Host "  スキップ: 出力先が既に存在 - $outputPath" -ForegroundColor Yellow
            $script:skipCount++
            return
        }

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($FilePath)
        $workbook.SaveAs($outputPath, $xlOpenXMLWorkbook)
        $workbook.Close($false)

        Write-Host "  完了: $outputPath" -ForegroundColor Green
        $script:successCount++
    }
    catch {
        Write-Host "  エラー: $($_.Exception.Message)" -ForegroundColor Red
        $script:errorCount++
        $script:errorFiles += $FilePath
    }
    finally {
        if ($null -ne $workbook) {
            Release-ComObject $workbook
        }
        if ($null -ne $excel) {
            $excel.Quit()
            Release-ComObject $excel
        }
    }
}

#------------------------------------------------------------------------------
# PowerPointファイルを変換 (ppt → pptx)
#------------------------------------------------------------------------------
function Convert-PowerPointFile {
    param([string]$FilePath)

    $powerpoint = $null
    $presentation = $null

    try {
        $outputPath = [System.IO.Path]::ChangeExtension($FilePath, ".pptx")

        # 既存ファイルチェック
        if (Test-Path $outputPath) {
            Write-Host "  スキップ: 出力先が既に存在 - $outputPath" -ForegroundColor Yellow
            $script:skipCount++
            return
        }

        $powerpoint = New-Object -ComObject PowerPoint.Application
        # PowerPointはVisible=$falseにできない場合がある

        $presentation = $powerpoint.Presentations.Open($FilePath, $true, $false, $false)
        $presentation.SaveAs($outputPath, $ppSaveAsOpenXMLPresentation)
        $presentation.Close()

        Write-Host "  完了: $outputPath" -ForegroundColor Green
        $script:successCount++
    }
    catch {
        Write-Host "  エラー: $($_.Exception.Message)" -ForegroundColor Red
        $script:errorCount++
        $script:errorFiles += $FilePath
    }
    finally {
        if ($null -ne $presentation) {
            Release-ComObject $presentation
        }
        if ($null -ne $powerpoint) {
            $powerpoint.Quit()
            Release-ComObject $powerpoint
        }
    }
}

#------------------------------------------------------------------------------
# ファイルを変換（拡張子で振り分け）
#------------------------------------------------------------------------------
function Convert-File {
    param([string]$FilePath)

    $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()

    Write-Host "変換中: $FilePath"

    switch ($extension) {
        ".doc" { Convert-WordFile -FilePath $FilePath }
        ".xls" { Convert-ExcelFile -FilePath $FilePath }
        ".ppt" { Convert-PowerPointFile -FilePath $FilePath }
        default {
            Write-Host "  スキップ: 対象外の拡張子" -ForegroundColor Yellow
            $script:skipCount++
        }
    }
}

#------------------------------------------------------------------------------
# フォルダ内のファイルを変換
#------------------------------------------------------------------------------
function Convert-Folder {
    param([string]$FolderPath)

    Write-Host "`nフォルダを検索中: $FolderPath" -ForegroundColor Cyan

    $files = Get-ChildItem -Path $FolderPath -Recurse -File |
             Where-Object { $_.Extension -eq ".doc" -or $_.Extension -eq ".xls" -or $_.Extension -eq ".ppt" }

    if ($files.Count -eq 0) {
        Write-Host "変換対象のファイルが見つかりませんでした。" -ForegroundColor Yellow
        return
    }

    Write-Host "対象ファイル数: $($files.Count)`n"

    foreach ($file in $files) {
        Convert-File -FilePath $file.FullName
    }
}

#------------------------------------------------------------------------------
# メイン処理
#------------------------------------------------------------------------------

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " Office形式変換ツール" -ForegroundColor Cyan
Write-Host " doc→docx, xls→xlsx, ppt→pptx" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# パスの取得
$targetPaths = @()
$cancelled = $false

if ($null -eq $Path -or $Path.Count -eq 0) {
    # 引数なし: ダイアログ表示
    $selectedPath = Show-FolderDialog
    if ($null -eq $selectedPath) {
        Write-Host "キャンセルされました。" -ForegroundColor Yellow
        $cancelled = $true
    }
    else {
        $targetPaths += $selectedPath
    }
}
else {
    $targetPaths = $Path
}

# 各パスを処理
if (-not $cancelled) {
    foreach ($targetPath in $targetPaths) {
        if (-not (Test-Path $targetPath)) {
            Write-Host "パスが存在しません: $targetPath" -ForegroundColor Red
            continue
        }

        if (Test-Path $targetPath -PathType Container) {
            # フォルダの場合
            Convert-Folder -FolderPath $targetPath
        }
        else {
            # ファイルの場合
            Convert-File -FilePath $targetPath
        }
    }

    # ガベージコレクション
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    # 結果表示
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host " 処理結果" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "成功: $script:successCount 件" -ForegroundColor Green
    Write-Host "スキップ: $script:skipCount 件" -ForegroundColor Yellow
    Write-Host "エラー: $script:errorCount 件" -ForegroundColor Red

    if ($script:errorFiles.Count -gt 0) {
        Write-Host "`nエラーが発生したファイル:" -ForegroundColor Red
        foreach ($f in $script:errorFiles) {
            Write-Host "  - $f" -ForegroundColor Red
        }
    }

    Write-Host "`n処理が完了しました。"
}

# 終了時は必ずここを通る
Write-Host ""
pause
