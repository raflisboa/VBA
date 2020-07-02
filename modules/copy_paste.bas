Attribute VB_Name = "copy_paste"
Sub copyPaste()

    Sheet3.Visible = xlSheetVisible
    Sheets("Aux_1").Select
    Range("E2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.End(xlUp).Select
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("Table8[[#Headers],[text_Checksum]]").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Save
    Sheet3.Visible = xlSheetVeryHidden
    
End Sub

