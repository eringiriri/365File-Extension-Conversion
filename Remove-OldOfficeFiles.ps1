<#
.SYNOPSIS
    変換済みの古いOffice形式(doc, xls, ppt)を削除する

.DESCRIPTION
    対応する新形式(docx, xlsx, pptx)が存在する場合のみ、古い形式を削除する
    - 引数なし: フォルダ選択ダイアログを表示
    - 引数あり: フォルダまたはファイルを指定

.PARAMETER Path
    対象のフォルダまたはファイルのパス（複数指定可能）

.PARAMETER Force
    確認なしで削除を実行

.EXAMPLE
    .\Remove-OldOfficeFiles.ps1
    # フォルダ選択ダイアログが表示される

.EXAMPLE
    .\Remove-OldOfficeFiles.ps1 -Path "C:\Documents" -Force
    # 確認なしで削除
#>

param(
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$Path,

    [switch]$Force
)

# 結果集計用
$script:deleteCount = 0
$script:skipCount = 0
$script:errorCount = 0
$script:filesToDelete = @()

#------------------------------------------------------------------------------
# フォルダ選択ダイアログを表示
#------------------------------------------------------------------------------
function Show-FolderDialog {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "対象のフォルダを選択してください"
    $dialog.ShowNewFolderButton = $false

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.SelectedPath
    }
    return $null
}

#------------------------------------------------------------------------------
# 新形式が存在するかチェック
#------------------------------------------------------------------------------
function Test-NewFormatExists {
    param([string]$FilePath)

    $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
    $newExtension = switch ($extension) {
        ".doc" { ".docx" }
        ".xls" { ".xlsx" }
        ".ppt" { ".pptx" }
        default { $null }
    }

    if ($null -eq $newExtension) {
        return $false
    }

    $newPath = [System.IO.Path]::ChangeExtension($FilePath, $newExtension)
    return (Test-Path $newPath)
}

#------------------------------------------------------------------------------
# ファイルをチェックして削除候補に追加
#------------------------------------------------------------------------------
function Check-File {
    param([string]$FilePath)

    $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()

    if ($extension -notin @(".doc", ".xls", ".ppt")) {
        return
    }

    if (Test-NewFormatExists -FilePath $FilePath) {
        $script:filesToDelete += $FilePath
        Write-Host "  削除対象: $FilePath" -ForegroundColor Yellow
    }
    else {
        $script:skipCount++
        Write-Host "  スキップ（新形式なし）: $FilePath" -ForegroundColor Gray
    }
}

#------------------------------------------------------------------------------
# フォルダ内のファイルをチェック
#------------------------------------------------------------------------------
function Check-Folder {
    param([string]$FolderPath)

    Write-Host "`nフォルダを検索中: $FolderPath" -ForegroundColor Cyan

    $files = Get-ChildItem -Path $FolderPath -Recurse -File |
             Where-Object { $_.Extension -eq ".doc" -or $_.Extension -eq ".xls" -or $_.Extension -eq ".ppt" }

    if ($files.Count -eq 0) {
        Write-Host "対象のファイルが見つかりませんでした。" -ForegroundColor Yellow
        return
    }

    Write-Host "チェック対象: $($files.Count) 件`n"

    foreach ($file in $files) {
        Check-File -FilePath $file.FullName
    }
}

#------------------------------------------------------------------------------
# 削除を実行
#------------------------------------------------------------------------------
function Remove-Files {
    foreach ($filePath in $script:filesToDelete) {
        try {
            Remove-Item -Path $filePath -Force
            Write-Host "  削除完了: $filePath" -ForegroundColor Green
            $script:deleteCount++
        }
        catch {
            Write-Host "  削除失敗: $filePath - $($_.Exception.Message)" -ForegroundColor Red
            $script:errorCount++
        }
    }
}

#------------------------------------------------------------------------------
# メイン処理
#------------------------------------------------------------------------------

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " 古いOffice形式 削除ツール" -ForegroundColor Cyan
Write-Host " (docx/xlsx/pptxが存在する場合のみ削除)" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# パスの取得
$targetPaths = @()
$cancelled = $false

if ($null -eq $Path -or $Path.Count -eq 0) {
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

if (-not $cancelled) {
    # 各パスをチェック
    foreach ($targetPath in $targetPaths) {
        if (-not (Test-Path $targetPath)) {
            Write-Host "パスが存在しません: $targetPath" -ForegroundColor Red
            continue
        }

        if (Test-Path $targetPath -PathType Container) {
            Check-Folder -FolderPath $targetPath
        }
        else {
            Check-File -FilePath $targetPath
        }
    }

    # 削除候補がない場合
    if ($script:filesToDelete.Count -eq 0) {
        Write-Host "`n削除対象のファイルはありません。" -ForegroundColor Yellow
    }
    else {
        # 削除確認
        Write-Host "`n----------------------------------------" -ForegroundColor Cyan
        Write-Host "削除対象: $($script:filesToDelete.Count) 件" -ForegroundColor Yellow

        $doDelete = $true
        if (-not $Force) {
            $response = Read-Host "`nこれらのファイルを削除しますか？ (y/n)"
            if ($response -ne "y") {
                Write-Host "キャンセルされました。" -ForegroundColor Yellow
                $doDelete = $false
            }
        }

        if ($doDelete) {
            Write-Host "`n削除を実行中...`n" -ForegroundColor Cyan
            Remove-Files

            # 結果表示
            Write-Host "`n========================================" -ForegroundColor Cyan
            Write-Host " 処理結果" -ForegroundColor Cyan
            Write-Host "========================================" -ForegroundColor Cyan
            Write-Host "削除: $script:deleteCount 件" -ForegroundColor Green
            Write-Host "スキップ: $script:skipCount 件" -ForegroundColor Yellow
            Write-Host "エラー: $script:errorCount 件" -ForegroundColor Red

            Write-Host "`n処理が完了しました。"
        }
    }
}

# 終了時は必ずここを通る
Write-Host ""
pause
