Attribute VB_Name = "NewMacros"
Sub searchBlue()
'
' searchBlue Macro
'
'
    ActiveWindow.DocumentMap = True
    Selection.Find.ClearFormatting
    Selection.Find.Font.Color = 12611584
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
End Sub
