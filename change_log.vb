'Copyright 2011, 2012 Steven Oliver <oliver.steven@gmail.com>
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

Public multi As Boolean

Dim vOldValR() As Range
Dim vOldVal() As String

Private Function check_sheet(sheet As String) As Boolean
    If sheet = "This Week" Or sheet = "Statistics" Or sheet = "Change Log" Then
        check_sheet = True
    Else
        check_sheet = False
    End If
End Function

Private Sub Assign_Range_Values(ByVal Target As Range)
    multi = True
    Dim i As Integer
    i = 0
    For Each lRange In Target.Cells
        Set vOldValR(i) = lRange
        vOldVal(i) = lRange.Value
        i = i + 1
    Next
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If check_sheet(Target.Cells.Worksheet.Name) or  Target.Cells.Count >= 500 Then
        Exit Sub
    ' Excel forbids you from assigning a whole column
    ' or row to an array for performance reasons so
    ' keep it simple and set a hard limit
    '
    ' http://support.microsoft.com/kb/166342
    ElseIf Target.Cells.Count < 500 Then
        ReDim vOldValR(0 To Target.Cells.Count - 1) As Range
        ReDim vOldVal(0 To Target.Cells.Count - 1) As String
    End If
    
    If Target.Cells.Count > 1 Then
        Call Assign_Range_Values(Target)
    ElseIf IsNull(Target.Text) Then
        multi = False
        Set vOldValR(0) = Target
        vOldVal(0) = vbNullString
    Else
        multi = False
        Set vOldValR(0) = Target
        vOldVal(0) = Target.Text
    End If
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If check_sheet(Target.Cells.Worksheet.Name) Then
        Exit Sub
    End If
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
        
    If multi Then
        Dim i As Integer
        i = 0

        For Each lArr In vOldValR()
            Call Write_Change(lArr, i)
            i = i + 1
        Next lArr
    Else
        Call Write_Change(Target, 0)
    End If
    
    With Application
            .ScreenUpdating = True
            .EnableEvents = True
    End With
End Sub

Private Sub Write_Change(ByVal Target As Range, i As Integer)    
    Dim bBold As Boolean
            
    If vOldVal(i) = "" Or IsNull(vOldVal(i)) Then
        If Target = "" Or IsNull(Target) Then
            Exit Sub
        Else
            vOldVal(i) = "(null)"
        End If
    End If
              
    bBold = Target.HasFormula
    With Sheet8
        .Unprotect Password:="newpass"
        If .Range("A1") = vbNullString Then
            .Range("A1:G1") = Array("SHEET", "CELL", "OLD VALUE", _
                "NEW VALUE", "TIME", "DATE", "USER")
        End If
            
        With .Cells(.Rows.Count, 1).End(xlUp)(2, 1)
            .Value = Target.Cells.Worksheet.Name
            .Offset(0, 1) = Target.Address
            .Offset(0, 2) = vOldVal(i)
            With .Offset(0, 3)
                If bBold = True Then
                    .ClearComments
                    AddComment.Text Text:="Bold values are the result of formulas"
                End If
                           
                If Target = "" Or IsNull(Target) Then
                    .Value = "(null)"
                Else
                    .Value = Target
                End If
                          
                .Font.Bold = bBold
            End With
                    
            .Offset(0, 4) = Time
            .Offset(0, 5) = Date
            .Offset(0, 6) = Environ("USERNAME")
        End With
                
        .Cells.Columns.AutoFit
        .Protect Password:="newpass"
    End With
        
    vOldVal(i) = vbNullString
End Sub

