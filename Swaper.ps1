
function SmartNotebook2PDF
{
    param(
        [object]$notebookfile,
        [string]$newLoc
    )
    if(((Get-Item $notebookfile).basename)[-1] -eq '.'){
    
        $rename = ($notebookfile.Substring(0, $notebookfile.Length-10)) + '.notebook'
        Rename-Item -Path $notebookfile -NewName $rename
        $notebookfile = $rename
    }
    $notebasename = (Get-Item $notebookfile).basename
   
    
    
    $dir = $notebookfile.replace($notebasename + '.notebook','')
   
    $zipfile = $notebasename + '.zip'
    $storage = $dir + $zipfile
    $storage2 = $dir + ($notebasename)

    Copy-Item $notebookfile  -Destination $storage
    try {
        Expand-Archive $storage -DestinationPath $storage2 -Force
    }
    catch {
        Remove-Item -path $storage
        continue 
    }
        if (!(Test-Path $storage2\imsmanifest.xml)) {
            
            Write-Host "sorry, this notebook seems to have no manifest."
            Out-File $notebookfile -FilePath ($dir + '.txt')
            Remove-Item -path $storage
            continue
        }
        $ovf = New-Object System.XML.XMLDocument
        $ovf.Load( $storage2+"\imsmanifest.xml")
        $pages = $ovf.manifest.resources.resource[1].file.href
    $n = 0
    ForEach ($page in $pages) {
        $n++
        Write-Host "Found $page, converting"
        convert.exe  $storage2\$page ($dir + $page.replace('.svg','.pdf'))
    }

    if (!($pages.GetType().Name -eq 'String')) {
        for ($i = 0; $i -lt $pages.Count; $i++) {
        
            $pages[$i] = $dir + $pages[$i]
        }
    }else{
        $pages = $dir + $pages
    }
    
        
        
    
    
    $pages = $pages.replace('.svg','.pdf')
    $end = $pages | Out-String 
    Merge-Pdf -InputFile $pages -OutputFile ($storage2 + '.pdf')
    
    
    Write-Host "PDF created $dir$notebasename.pdf. Cleaning up."
    
    
    Remove-Item -path $pages
    Remove-Item -path $storage2 -recurse
    Remove-Item -path $storage
    Write-Host "Done."
    directoryMoveBase -OldBaseDir ($storage2 + '.pdf') -baseDir $newLoc
}

function directoryMoveBase{

    Param(
        $OldBaseDir,
        $baseDir
    )
    $newPath = $baseDir + "\NewFolder\"

    $oldpath = $OldBaseDir.replace($baseDir,$null).split('\')
    for ($i = 0; $i -lt $oldpath.Count - 1; $i++) {
        $newPath += $oldpath[$i] + '\' 
    }
   New-Item -ItemType Directory -Force -Path $newPath
   Move-Item -Path ($OldBaseDir) -Destination $newPath -Force
   

}
Function dialog-Folder(){
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
        Description = "Please select the folder"
        RootFolder  = "Desktop"
    }
    $result = $FolderBrowser.ShowDialog()
    if ($result -ne "OK") {
         $null = [System.Windows.Forms.MessageBox]::Show("No folder chosen!", "Error", 0, [System.Windows.Forms.MessageBoxIcon]::Error)
         return ''
    }
    else {
         return $FolderBrowser.SelectedPath
    }
}

$path = dialog-Folder
if ($path -ne '') {
    $files = dir -Path $path -recurse -ea 0 *.notebook | % FullName

    if ($files -eq $null) {
        break
    }
   Write-Host 'Es sind insgesamt' $files.Length -ForegroundColor Green
   $n = 1
   Write-Host '____________________________________'
ForEach ($file in $files) {
    
    Write-Host '____________________________________'
    Write-Host $n -ForegroundColor Yellow
    SmartNotebook2PDF -notebookfile $file -newLoc $path
    #SmartNotebook2PDF -notebookfile $files[575] -newLoc $path
    $n += 1
   
    Write-Host '____________________________________'
    }
}

