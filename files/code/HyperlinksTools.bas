Attribute VB_Name = "HyperlinksTools"
Sub removeAllHyperlinksInSelection()
    Selection.Hyperlinks.Delete
'    ActiveSheet.UsedRange.Hyperlinks.Delete
End Sub

Sub removeAllHyperlinksInUsedRange()
    ActiveSheet.UsedRange.Hyperlinks.Delete
End Sub

