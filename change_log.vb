'Copyright 2011 Steven Oliver <oliver.steven@gmail.com>
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

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If Target.Cells.Count < UBound(vOldValR) Then
        ReDim vOldValR(0 To Target.Cells.Count - 1) As Range
        ReDim vOldVal(0 To Target.Cells.Count - 1) As String
    Else
        Exit Sub
    End If
    
    If Target.Cells.Count > 1 Then
        multi = True
        Dim i As Integer
        i = 0
        For Each lRange In Target.Cells
            Set vOldValR(i) = lRange
            vOldVal(i) = lRange.Value
            i = i + 1
        Next
    ElseIf IsNull(Target.Text) Then
        multi = False
        Set vOldValR(0) = Target
        vOldVal(0) = vbNullString
    Else
        multi = False
        Set vOldValR(0) = Target
        vOldVal(i) = Target.Text
    End If
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim i As Integer
    
    i = 0
    If multi Then
        For Each lArr In vOldValR()
            Call Write_Change(lArr, i)
            i = i + 1
        Next lArr
    Else
        Call Write_Change(Target, 0)
    End If
End Sub

Private Sub Write_Change(ByVal Target As Range, i As Integer)
Dim bBold As Boolean

On Error Resume Next
    With Application
         .ScreenUpdating = False
         .EnableEvents = False
    End With

    If vOldVal(i) = "" Or IsNull(vOldVal(i)) Then vOldVal(i) = "(null)"
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
                          .AddComment.Text Text:="Bold values are the result of formulas"
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
    With Application
         .ScreenUpdating = True
         .EnableEvents = True
    End With
On Error GoTo 0
End Sub

