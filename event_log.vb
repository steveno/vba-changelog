'Copyright 2011-2013 Steven Oliver <oliver.steven@gmail.com>
'
'Licensed under the Apache License, Version 2.0 (the "License");
'you may not use this file except in compliance with the License.
'You may obtain a copy of the License at
'
'       http://www.apache.org/licenses/LICENSE-2.0
'
'Unless required by applicable law or agreed to in writing, software
'distributed under the License is distributed on an "AS IS" BASIS,
'WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'See the License for the specific language governing permissions and
'limitations under the License.

Private Function row_exist(Sh As String, Ro As String) As Boolean
    Dim sos_num As String
    Dim last_row As Integer
    
    ' Pick up the sos num from the sheet
    ' listing the sos'es
    sos_num = Sheets(Sh).Cells(Ro, 1).Text
    
    ' Go back to the This Week sheet
    ' and see if we have it already
    With Sheets("This Week")
        last_row = .UsedRange.Rows.Count
        For Each c In .Range("A1", "A" & last_row).Cells
            If c.Text = sos_num Then
                row_exist = True
                Exit Function
            End If
        Next
    End With
    
    ' Couldn't find it
    row_exist = False
End Function

Private Sub update_dates()
    Dim last_row As Integer
    
    With Sheets("Change Log")
        .Unprotect Password:="newpass"
                
        ' We're going to loop through the whole sheet
        ' and find any value in columns E or F that
        ' are dates and reformat them to look like dates
        last_row = .UsedRange.Rows.Count
        For Each c In .Range("B2", "B" & last_row).Cells
            If Left(c.Text, 2) = "$F" Or Left(c.Text, 2) = "$E" Then
                .Cells(c.Row, 4).NumberFormat = "m/d/yyyy"
            End If
        Next
        .Protect Password:="newpass"
    End With
End Sub

Private Sub add_row(Sh As String, Ro As String)
    Dim last_row As Integer
    
    ' Grab the SOS from the year sheet
    Sheets(Sh).Rows(Ro).Copy
       
    ' Move the copied row to the This Week sheet
    ' to add
    With Sheets("This Week")
        ' Find the last row and make sure
        ' we always use it
        last_row = .UsedRange.Rows.Count + 1
        
        ' Move the button so it doesn't get clobbered
        .Shapes("Button 1").Select
        Selection.ShapeRange.IncrementTop 15
        
        ' Paste in the work row and unhide it
        ' I don't know why it pastes hidden
        .Columns("A").Rows(last_row).Select
        ActiveSheet.Paste
        Selection.EntireRow.Hidden = False
    End With
End Sub

Private Sub update_row(Sh As String, Ro As String)
    Dim sos_num As String
    Dim last_row As Integer
    
    ' Grab the SOS from the year sheet
    Sheets(Sh).Rows(Ro).Copy
    
    ' Pick up the sos num from the sheet
    ' listing the sos'es
    sos_num = Sheets(Sh).Cells(Ro, 1).Text
       
    ' Move the copied row to the This Week sheet
    ' to complete the update
    With Sheets("This Week")
        ' Find this sos in the This Week sheet
        last_row = .UsedRange.Rows.Count
        For Each c In .Range("A1", "A" & last_row).Cells
            If c.Text = sos_num Then
                ' Paste in the work row and unhide it
                .Columns("A").Rows(c.Row).Select
                ActiveSheet.Paste
                Selection.EntireRow.Hidden = False
                Exit Sub
            End If
        Next
    End With
End Sub

Private Sub remove_row(Sh As String, Ro As String)
    ' Delete the requested row shifting the left over
    ' cells upward
    Sheets(Sh).Rows(Ro).Delete Shift:=xlUp
End Sub

Private Sub clean_sheet(Sh As String)
    Dim c As Integer
    
    ' Loop through the whole sheet and
    ' delete every row
    With Sheets(Sh)
        last_row = .UsedRange.Rows.Count
        For c = 2 To last_row Step 1
            Call remove_row(Sh, 2)
        Next
    End With
End Sub

Private Sub update_from_changelog()
    Sheets("Change Log").Visible = True
    
    ' Update the dates on the changelog to make sure we
    ' don't run into anything unexpected
    Call update_dates
        
    Dim sos_sheet As String, sos_row As String
    Dim last_row As Integer
    
    ' Scroll through changelog and look for entries within
    ' the past 7 days
    With Sheets("Change Log")
        last_row = .UsedRange.Rows.Count
        For Each c In .Range("F2", "F" & last_row).Cells
            ' If adding a column gets picked up in the log
            ' this should catch it
            Dim test As Variant
            test = Split(.Cells(c.Row, "B").Text, "$")
            If Not IsNumeric(test(2)) Then
                GoTo NextStatement
            End If
                
            If CDate(c.Text) >= DateAdd("d", -7, Date) Then
                sos_row = .Cells(c.Row, "B").Text
                sos_row = Right(sos_row, Len(sos_row) - InStrRev(sos_row, "$"))
                sos_sheet = .Cells(c.Row, "A").Text
                
                If sos_sheet = "Data" Then
                    ' do nothing
                Else
                    If row_exist(sos_sheet, sos_row) Then
                        Call update_row(sos_sheet, sos_row)
                    Else
                        Call add_row(sos_sheet, sos_row)
                    End If
                End If
            End If
NextStatement:
        Next
    End With
    
    Sheets("Change Log").Visible = False
End Sub

Private Sub update_from_years()
    Dim curr_year As Integer
    Dim last_row As Integer, cnt As Integer
    
    ' We're going to loop through the current year
    ' and the year before it. Anything older than
    ' that should be cancelled for not being touched
    ' for such a signifigant amount of time
    curr_year = Year(Now())
    For cnt = curr_year - 1 To curr_year Step 1
        With Sheets(CStr(cnt))
            last_row = .UsedRange.Rows.Count
            For Each c In .Range("D2", "D" & last_row).Cells
                If c.Text = "40 In Process" Or c.Text = "30 Approved and Assigned" Or c.Text = "25 Approved by BPM" Then
                    If row_exist(CStr(cnt), c.Row) Then
                        Call update_row(CStr(cnt), c.Row)
                    Else
                        Call add_row(CStr(cnt), c.Row)
                    End If
                End If
            Next
        End With
    Next
End Sub

Private Sub cleanup()
    last_row = Sheets("This Week").UsedRange.Rows.Count
    
    With Sheets("This Week")
        ' Remove the coloring from the priorities
        ' column since this page will not be
        ' autoformatted
        Range("C2").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.ClearFormats
        
        ' Delete columns I don't care to see
        Range("G:G,H:H,J:J").Select
        Selection.Delete Shift:=xlToLeft
        
        ' The deletion above removes my column
        ' header so readd it
        Range("F1").Select
        Selection.Copy
        Range("G1").Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "Transport(s)"
        
        ' Sort the SOS list
        Range("A2").Select
        ActiveWorkbook.Worksheets("This Week").Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("This Week").Sort.SortFields.Add Key:=Range( _
            "A2", "A" & last_row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        With ActiveWorkbook.Worksheets("This Week").Sort
            .SetRange Range("A1", "A" & last_row)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        ' Filter the table and make the header row pretty
        Range("A1:H1").Select
        Columns("A:A").EntireColumn.AutoFit
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorLight2
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        
        ' Auto fit all of the cells
        Cells.Select
        Cells.EntireColumn.AutoFit
    End With
End Sub

Public Sub update()
    With Application
        ' Do away with screen updating and any automatic
        ' calculations to improve performance
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        
        Call clean_sheet("This Week")
        Call update_from_changelog
        Call update_from_years
        Call cleanup
        
        ' Restore Excel to it's default state
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub
