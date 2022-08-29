
# .Notebook Converter

Dies ist ein Script, welches per powershell ausgeführt werden muss.




## Installation

Um das Script nutzen zu können benötigen Sie:
- [ImageMagick](https://imagemagick.org/index.php)
- MergePDF

### Installation von MergePDF

um das Programm zu installieren führen Sie einmal mit Administrationsrechte die Konsole aus. Anschließend geben sie 
folgenes ein:
```powershell
Install-Module PSWritePDF -Force
```

### Installation von ImageMagick

Folgen Sie den Anweisungen auf der Website [Der Windows Papst – IT Blog Walter](https://www.der-windows-papst.de/2021/08/25/convert-heic-to-jpg-png-powershell/)




    
## Fehlerbehandlung

Beim Script handelt es sich um ein Converter mit Basic Functions. 

Folgenes wird beim Fehler passieren:

- Bei einem Fehler innerhalb der Datei, wird die Datei nicht mehr berücksichtigt und sie wird übersprungen. Die Datei wird aber anhand einer *.TXT Datei gekennzeichnet.
- Fehler mit Sonderzeichen können in einen Fehler fallen, der noch nicht berücksichtigt wurde.
