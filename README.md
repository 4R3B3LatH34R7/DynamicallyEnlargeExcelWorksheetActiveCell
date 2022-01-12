# Dynamically Enlarging an Excel Worksheet's ActiveCell

A few days ago, I came across a question in a facebook group where Myanmar people share, learn and discuss MS Excel related information.\
The question is about making the ActiveCell in the worksheet bigger/larger/zoomed-in.\
The OP didn't mention why he needed that feature.

Normally, I don't like to mess up the Excel UI for the user.\
And I warned the OP that changing the cell size like that would be dizzying and made sure that he wanted this knowing possible side effects.

The following .gif is provided just to clarify the OP's requirements and the answer I created.

![EnlargingActiveCell](images/EnlargingActiveCell.gif)

Also available on [StackOverflow as a self-answered Q&A post](https://stackoverflow.com/questions/69795647/excel-vba-how-to-dynamically-enlarge-expand-activecell/69795648#69795648).

## Possible solutions
I could immediately think of 4 methods that I may be able to use to achieve that goal.

1. Manipulate the ActiveCell's rowHeight and columnWidth properties. (most likely)
2. Play with ActiveCell's column's Autofit. (I don't think it'd give the desired effect)
3. Play with the Worksheet's Zoom %. (not really)
4. Place a textbox on a modeless userform and put the activecell's contents into it. (I love it)

Among those 4, method1 is most likely what the OP wanted. But I believe that it would be quite disturbing to the user's eyes.\
And the only other worthy contender, IMHO, is the method4 as it seems more elegant and less distracting to the user.\
The remaining 2, I did't even bother to try.

## The VBA Code
It seems like a trivial and minor coding fun for me but actually, I just spent 2 days working on it and I still feel like some parts need to function a bit better, especially the userform method.\
Therefore, I will just share the method1 code as the userform method still feels like it needs further polish.

```VBA
'copy paste into ThisWorkbook module
Option Explicit
Const defaultColumnWidth = 8.11
Const defaultRowHeight = 14.4
Const increasedColumnWidth = 50
Const increasedRowHeight = 50
Private saved_ActiveCell_ColumnWidth As Integer
Private saved_ActiveCell_RowHeight As Integer
Private saved_ActiveCell As Range
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
'explicit error checking was not done - use at users' own risk
'very important that sh must be a worksheet
'    If Sh.Name = "Sheet1" Then 'set sheet name here to limit to Sheet1 only
    If Target.Cells.CountLarge = 1 Then
        Application.ScreenUpdating = False
        If Target.Value <> "" Then
            If Not saved_ActiveCell Is Nothing Then
                saved_ActiveCell.EntireColumn.ColumnWidth = saved_ActiveCell_ColumnWidth
                saved_ActiveCell.EntireRow.RowHeight = saved_ActiveCell_RowHeight
            End If
            Set saved_ActiveCell = Target 'Application.ActiveCell.Address
            saved_ActiveCell_ColumnWidth = Target.ColumnWidth
            saved_ActiveCell_RowHeight = Target.RowHeight

            Target.EntireRow.RowHeight = increasedRowHeight
            Target.EntireColumn.ColumnWidth = increasedColumnWidth
        Else
            If Not saved_ActiveCell Is Nothing Then
                saved_ActiveCell.EntireColumn.ColumnWidth = saved_ActiveCell_ColumnWidth
                saved_ActiveCell.EntireRow.RowHeight = saved_ActiveCell_RowHeight
            End If
        End If
        Application.ScreenUpdating = True
    End If
'    End If
End Sub
```
In both methods, I decided to use the Workbook_SheetSelectionChange Event because I don't want to be worried about the code not working in the newly inserted Worksheets.
As mentioned in the code, it is possible to limit this feature to be available in certain Worksheets by modifying a little bit where the Worksheet name can be checked.

There is no guarantee that the userform based method will be further developed as I am currently occupied with other projects.
If it becomes satisfactorily functional and clean enough, I will release it.

***
## License
I don't actually like/want/wish to apply CC BY-SA license to what I share, really!\
However, there exists some jerks in this world who thought it's ok to derive my work without proper accreditation.\
I don't care much for fame nor finance but a little credit for the many hours of my limited life I spent on a project is appreciated.\
Shield: [![CC BY-SA 4.0][cc-by-sa-shield]][cc-by-sa]

This work is licensed under a
[Creative Commons Attribution-ShareAlike 4.0 International License][cc-by-sa].

[![CC BY-SA 4.0][cc-by-sa-image]][cc-by-sa]

[cc-by-sa]: http://creativecommons.org/licenses/by-sa/4.0/
[cc-by-sa-image]: https://licensebuttons.net/l/by-sa/4.0/88x31.png
[cc-by-sa-shield]: https://img.shields.io/badge/License-CC%20BY--SA%204.0-lightgrey.svg
***
