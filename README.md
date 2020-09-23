<div align="center">

## RecurseFolderList


</div>

### Description

This code is a modified version of ShowFolderList by Bruce Lindsay. (Thanx !!) This code will recursively parse a directory defined by an path parameter. My aim was to work around

the non-recursive nature of the dir function. Bruce's original code does that to one folder/child level. Mine now returns everything below a given path. You can still use getattr to define Folder or File attributes.
 
### More Info
 
foldername - "c:\temp"

No error trapping, untested on VB3/4


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Scott Brown](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/scott-brown.md)
**Level**          |Unknown
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/scott-brown-recursefolderlist__1-2876/archive/master.zip)





### Source Code

```
Function RecurseFolderList(foldername)
 Dim fso, f, fc, fj, f1
 Set fso = CreateObject("Scripting.FileSystemObject")
 Set f = fso.GetFolder(foldername)
 Set fc = f.Subfolders
 Set fj = f.Files
 'For each subfolder in the Folder
 For Each f1 In fc
  'Do something with the Folder Name
  debug.print f1
  'Then recurse this function with the sub-folder to get any sub-folders
  RecurseFolderList(f1)
 Next
 'For each folder check for any files
 For Each f1 In fj
  debug.print f1
 Next
End Function
```

