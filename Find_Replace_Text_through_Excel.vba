'How to run:
'Copy this code into a excel file vba code
'In sheet1: place find text words into col-A and replace text words in column-B
'Run WordFindAndReplaceIAST0()
'============================================================
Option Explicit
Dim mydocfile As String

Public Sub WordFindAndReplaceIAST0()
    Dim ws As Worksheet, msWord As Object, itm As Range
    Dim outfile, outfileL, outfileR As String
    Dim strlen, pos As Integer
    Set ws = ActiveSheet
    Set msWord = CreateObject("Word.Application")

    getfilename (1)  'get file name; just pass a dummy number
        
    With msWord
        .Visible = True
        .Documents.Open mydocfile
        .Activate
        
	'accept all changes and enable trackmode
        With .ActiveDocument  
            .Revisions.AcceptAll
            If (Not .TrackRevisions) Then
                .TrackRevisions = True
            End If
        End With
        
        
        With .ActiveDocument.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting

            For Each itm In ws.UsedRange.Columns("A").Cells
                .Text = itm.Value2       'Find all strings in col A
                .Replacement.Text = itm.Offset(, 1).Value2  'Replacements from col B
                .MatchCase = False
                .MatchWholeWord = False
                .Execute Replace:=2     'wdReplaceAll (WdReplace Enumeration)
            Next
        End With
        
        strlen = Len(mydocfile)
        pos = InStrRev(mydocfile, "\")
        outfileL = Left(mydocfile, pos)
        outfileR = Right(mydocfile, strlen - pos)
             
        outfile = outfileL + "IAST_" + outfileR

        .ActiveDocument.SaveAs2 Filename:=outfile        
        .Quit SaveChanges:=False
    End With
End Sub

Sub getfilename(num)
    Dim my_FileName As Variant
    
    my_FileName = Application.GetOpenFilename(FileFilter:="Excel Files,*.doc*;*.docx*")
     If my_FileName <> False Then
      mydocfile = my_FileName
    End If
End Sub

