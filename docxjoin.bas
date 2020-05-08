Sub docxjoin()
　Dim doc_name As String

　Documents.Add

　ChDir "C:\Users\kiyotoito\Desktop"

　doc_name = Dir("*.doc*")

　Do While doc_name <> ""
　　With Selection
　　　.TypeText "ファイル名 = " & doc_name & vbCr
　　　.InsertBreak wdPageBreak
　　　.InsertFile doc_name
　　　.InsertBreak wdSectionBreakNextPage
　　　.Collapse wdCollapseEnd
　　End With
　　doc_name = Dir
　Loop

End Sub