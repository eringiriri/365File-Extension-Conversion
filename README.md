# 365File-Extension-Conversion

古いOffice形式（doc, xls, ppt）を新形式（docx, xlsx, pptx）に変換するPowerShellスクリプト。

## 必要環境

- Windows
- PowerShell
- Microsoft Office（Word, Excel, PowerPoint）がインストールされていること

## スクリプト

### Convert-OfficeFiles.ps1

古いOffice形式を新形式に変換する。

```powershell
# フォルダ選択ダイアログから選択
.\Convert-OfficeFiles.ps1

# フォルダを指定
.\Convert-OfficeFiles.ps1 -Path "C:\Documents"

# ファイルを直接指定（複数可）
.\Convert-OfficeFiles.ps1 -Path "C:\test.xls", "C:\test2.doc"
```

- 変換先ファイルが既に存在する場合はスキップ
- サブフォルダも再帰的に処理

### Remove-OldOfficeFiles.ps1

変換済みの古いOffice形式ファイルを削除する。対応する新形式（docx, xlsx, pptx）が存在する場合のみ削除される。

```powershell
# フォルダ選択ダイアログから選択
.\Remove-OldOfficeFiles.ps1

# 確認なしで削除
.\Remove-OldOfficeFiles.ps1 -Path "C:\Documents" -Force
```

## 対応形式

| 変換元 | 変換先 |
|--------|--------|
| .doc   | .docx  |
| .xls   | .xlsx  |
| .ppt   | .pptx  |