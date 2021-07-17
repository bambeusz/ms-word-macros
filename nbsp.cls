VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub nbsp_replace()
  Dim objRegex As RegExp
  Dim originalPosition As Range
  Dim matches As MatchCollection
  Dim m As Match
  Dim line_start_pos As Long
  Dim line_end_pos As Long
  Dim match_start_pos As Long
  Dim match_end_pos As Long

  Set objRegex = New RegExp
  Set originalPosition = Selection.Range

  With objRegex
    .Pattern = " ([a-z])([^a-zĄąĆćĘęŁłŃńÓóŚśŹźŻż ]+)? "
    .Global = True
    .IgnoreCase = True
  End With

  Selection.HomeKey Unit:=wdStory

  Do
    Selection.HomeKey Unit:=wdLine
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    line_start_pos = Selection.Start
    line_end_pos = Selection.End

    Set matches = objRegex.Execute(Selection.Text)
    For Each m In matches
      match_start_pos = line_start_pos + m.FirstIndex
      match_end_pos = match_start_pos + m.Length

      If match_end_pos = line_end_pos Then
        Selection.Collapse
        ActiveDocument.Range(Start:=match_end_pos - 1, End:=match_end_pos).Select
        Selection.TypeText Text:=Chr$(160)
        Selection.MoveUp Unit:=wdLine, Count:=1
      End If
    Next m

    Selection.MoveDown Unit:=wdLine, Count:=1
  Loop Until Selection.End = ActiveDocument.Bookmarks("\EndOfDoc").Range.End

  originalPosition.Select
End Sub
