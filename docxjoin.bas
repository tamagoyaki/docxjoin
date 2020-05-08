Sub docxjoin()
  ' select a filelist
  Set fd = Application.FileDialog(msoFileDialogOpen)
  With fd
    .AllowMultiSelect = False
    .InitialFileName = "C:"
    .Filters.Add "Word doc list", "*.docxlist", 1
    If .Show <> -1 Then Exit Sub
  End With

  ' read filelist
  Set filelist = CreateObject("Scripting.FileSystemObject").OpenTextFile(fd.SelectedItems.Item(1), 1)

  ' regex for filelist
  Set regx_comment = CreateObject("vbscript.regexp")
  With regx_comment
     .Global = True
     .Pattern = "^#.*$"
  End With
  
  Set regx_empty = CreateObject("vbscript.regexp")
  With regx_empty
     .Global = True
     .Pattern = "^ *$"
  End With

  ' join docx
  Do While filelist.AtEndOfStream <> True
    file = filelist.ReadLine

    ' ignore the line if it's a comment or emply
    comment_found = regx_comment.test(file)
    empty_found = regx_empty.test(file)

    If comment_found Or empty_found = True Then
       GoTo CONTINUE
    End If

    ' insert
　　With Selection
　　　.InsertFile file
　　　.InsertBreak wdSectionBreakNextPage
　　　.Collapse wdCollapseEnd
　　End With
CONTINUE:
  Loop
filelist.Close

End Sub