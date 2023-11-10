# import re

# def extract_number(s):
#     pattern = "202\d{9}"
#     match = re.search(pattern, s)
#     if match:
#         return match.group()
#     else:
#         return "没有找到匹配的数字"

# s = "D:Destop办公办公data1word机器人（SI）21-1202101230001宋健202101230001实验4宋健实验四.doc"
# print(extract_number(s))

# PowerShell 脚本

Add-Type -AssemblyName System.Windows.Forms

$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = 'Select a Folder'
$folderBrowser.RootFolder = [Environment+SpecialFolder]::Desktop
$folderBrowser.ShowNewFolderButton = $false

$result = $folderBrowser.ShowDialog()

if ($result -eq [Windows.Forms.DialogResult]::OK) {
    $folderPath = $folderBrowser.SelectedPath
    $word_app = New-Object -ComObject Word.Application
    $word_app.Visible = $false
    
    $files = Get-ChildItem -Path $folderPath -Recurse -Filter *.doc*
    $fileCount = $files.Count
    $currentFileNumber = 0
    $files |Rename-Item -NewName { $_.Name -replace '\.docx','.doc' }
    $files | ForEach-Object {
        $currentFileNumber++
        $docPath = $_.FullName
        
        $pdfPath = [System.IO.Path]::ChangeExtension($docPath, 'pdf')
        
        if (-Not (Test-Path $pdfPath)) {
            Write-Progress -PercentComplete (($currentFileNumber / $fileCount) * 100) -Status "Processing $docPath" -Activity "$currentFileNumber of $fileCount files processed"
            $doc = $word_app.Documents.Open($docPath)
            $doc.SaveAs([ref]$pdfPath, [ref]17)  # 17 是 wdFormatPDF 的值
            $doc.Close($false)
        }
        else {
            Write-Progress -PercentComplete (($currentFileNumber / $fileCount) * 100) -Status "Skipping $docPath (PDF already exists)" -Activity "$currentFileNumber of $fileCount files processed"
        }
    }

    $word_app.Quit()
    Write-Output "Conversion completed!"
}