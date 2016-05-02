Sub UpdateOldPatentClaims()

' Author: Stephen G. Nagy (stnagy@gmail.com)
' Licensed under the GPL License v 3.0: http://www.gnu.org/licenses/gpl-3.0.en.html
' If this macro helped you, consider donating at paypal.me/stnagy.

' Macro does the following in sequence:
'    (1) turns on track changes
'    (2) removes all text that is formatted with strikethrough
'    (3) removes underline formatting but leave text
'    (4) changes track changes setting to whatever it was before the macro


    ' set variable [ currentTrackChanges ] to true / false depending on whether
    ' track changes is currently enabled.
    Dim currentTrackChanges As Boolean
    currentTrackChanges = ActiveDocument.TrackRevisions

    ' (1) turns on track changes
    ActiveDocument.TrackRevisions = True
    With Options
        .InsertedTextMark = wdInsertedTextMarkUnderline
        .InsertedTextColor = wdBlue
        .DeletedTextMark = wdDeletedTextMarkStrikeThrough
        .DeletedTextColor = wdByAuthor
        .RevisedPropertiesMark = wdRevisedPropertiesMarkColorOnly
        .RevisedPropertiesColor = wdBrightGreen
        .RevisedLinesMark = wdRevisedLinesMarkOutsideBorder
        .RevisedLinesColor = wdAuto
        .CommentsColor = wdByAuthor
    End With

    ' (2) removes all text that is formatted with strikethrough
    With Selection.Find
        .ClearFormatting
        .Format = True
        .Font.StrikeThrough = True
        .Forward = True
        .Wrap = wdFindContinue
        .Replacement.ClearFormatting
        .Replacement.Text = ""
        .Execute Forward:=True, Replace:=wdReplaceAll, _
        FindText:="", ReplaceWith:=""
    End With

    ' (3) removes underline formatting but leave text
    With Selection.Find
        .ClearFormatting
        .Format = True
        .Font.Underline = True
        .Font.Bold = False
        .Font.StrikeThrough = False
        .Font.Italic = False
        .Forward = True
        .Wrap = wdFindContinue
        .Replacement.ClearFormatting
        .Replacement.Font.Underline = False
        .Execute Forward:=True, Replace:=wdReplaceAll, _
        FindText:="", ReplaceWith:=""
    End With

    ' (4) changes track changes setting to whatever it was before the macro
    ActiveDocument.TrackRevisions = currentTrackChanges

End Sub
