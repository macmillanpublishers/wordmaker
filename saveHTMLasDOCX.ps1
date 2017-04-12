Param(
  [string]$html,
  [string]$docx
)

# make the powershell process switch the current directory.
$oldwd = [Environment]::CurrentDirectory
[Environment]::CurrentDirectory = $pwd

$html = resolve-path $html;
$docx = [IO.Path]::GetFullPath( $docx )
[Environment]::CurrentDirectory = $oldwd

[ref]$SaveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type] 
$word = New-Object -ComObject word.application 
$word.visible = $false 

"Converting $html to $docx..." 

$doc = $word.documents.open($html) 

#$doc.saveas($docx, $SaveFormat::wdFormatDocumentDefault)
$doc.saveas([ref] $docx, [ref]$SaveFormat::wdFormatDocumentDefault) 

$doc.close() 
$word.Quit() 
$word = $null 
[gc]::collect() 
[gc]::WaitForPendingFinalizers()