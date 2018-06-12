'title: VBA_Standard_Template # update for current workbook
'desc: workbook purpose # update for current workbook
'inputs: data sources # update for current workbook

'author: Michael J. Armstrong
'contact: marmstrong310@gmail.com
'© 2017, 2018 Michael J. Armstrong, Toronto, Ontario, Canada
'Distributed under the terms of the GNU General Public License v3.0

'This program is free software: you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation, either version 3 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program.  If not, see <http://www.gnu.org/licenses/>.
'(see "GNU GPL License.txt"
'
Option Explicit
'Dim global_variables

Sub global_nom()
'sets all commonly used references
'for use through all sub processes
'allows for easier reference
'
'<---Call global_nom---> should be added as the first line of all subs
'
'Worksheet references
    'Set XXX = Sheets("Sheet Name")
    
'Range references
    'Set YYY_ref = XXX.Range("Range:Ref")
    
'Miscellaneous references
    'common_variable = value

'Column references, as integers for use in cell references
'<-----Worksheet Name----->
    '<--data-->

End Sub

Sub sub_template()
'sub-process description
'<-----CALL GLOBAL VARIABLES----->
    Call global_nom
'<-----SET PROCESS SPECIFIC VARIABLES----->
    Dim StartTime As Double, SecondsElapsed As Double 'timer variables
    Dim cntr As Double: cntr = 0 'counting variables
    Dim i As Double
	Dim nxt_r As Double: nxt_r = 
    Dim lst_r As Double: lst_r = XX.Cells(Rows.count, col_ref).End(xlUp).Row 'set end of iteration (last row)
'<-----PROCESS OPTIMIZATION/DEACTIVATE----->
    OptimizeOn
'<-----RECORD START TIME OF PROCESS----->
    StartTime = Timer
'<-----ERROR HANDLING----->
    On Error Resume Next
'<::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::>
'<-----BEGIN PROCESS CODING----->

'<------END PROCESS CODING------>
'<::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::>
'<-----PROCESS OPTIMIZATION/REACTIVATE----->
    OptimizeOff
'<-----CALCULATE PROCESS RUN TIME----->
    SecondsElapsed = Round(Timer - StartTime, 2)
'<-----DISPLAY PROCESS RUN TIME----->
    MsgBox (CounterResults(cntr, SecondsElapsed))
End Sub

Sub clr_data()
Dim i As Double: i = InputBox("!!!WARNING!!!" & vbNewLine & _
                            "Cleared data cannot be recovered!" & vbNewLine & _
                            "Enter row number to start clearing")
Dim nxt_r As Double: nxt_r = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row + 1
Call global_nom
'Deletes all rows with data from row i through the end.
'Spans check across the first 10 columns in the active sheet.
        If nxt_r > i Then
            Rows(i & ":" & nxt_r).EntireRow.Delete
        ElseIf ActiveSheet.Cells(Rows.count, 2).End(xlUp).Row + 1 > i Then
            Rows(i & ":" & nxt_r).EntireRow.Delete
        ElseIf ActiveSheet.Cells(Rows.count, 3).End(xlUp).Row + 1 > i Then
            Rows(i & ":" & nxt_r).EntireRow.Delete
        ElseIf ActiveSheet.Cells(Rows.count, 4).End(xlUp).Row + 1 > i Then
            Rows(i & ":" & nxt_r).EntireRow.Delete
        ElseIf ActiveSheet.Cells(Rows.count, 5).End(xlUp).Row + 1 > i Then
            Rows(i & ":" & nxt_r).EntireRow.Delete
        ElseIf ActiveSheet.Cells(Rows.count, 6).End(xlUp).Row + 1 > i Then
            Rows(i & ":" & nxt_r).EntireRow.Delete
        ElseIf ActiveSheet.Cells(Rows.count, 7).End(xlUp).Row + 1 > i Then
            Rows(i & ":" & nxt_r).EntireRow.Delete
        ElseIf ActiveSheet.Cells(Rows.count, 8).End(xlUp).Row + 1 > i Then
            Rows(i & ":" & nxt_r).EntireRow.Delete
        ElseIf ActiveSheet.Cells(Rows.count, 9).End(xlUp).Row + 1 > i Then
            Rows(i & ":" & nxt_r).EntireRow.Delete
        ElseIf ActiveSheet.Cells(Rows.count, 10).End(xlUp).Row + 1 > i Then
            Rows(i & ":" & nxt_r).EntireRow.Delete
        Else
            MsgBox ("No data to clear")
        End If
'Return cursor to RiC1
    ActiveSheet.Cells(i, 1).Select

End Sub

Sub clr_data2()
Dim i As Byte: i = 2
Dim nxt_r As Double: nxt_r = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row + 1
Call global_nom
'Deletes all rows with data from row i through the end.
'Spans check across the first 10 columns in the active sheet.
        If nxt_r > i Then
            Rows(i & ":" & nxt_r).EntireRow.Delete
        ElseIf ActiveSheet.Cells(Rows.count, 2).End(xlUp).Row + 1 > i Then
            Rows(i & ":" & nxt_r).EntireRow.Delete
        ElseIf ActiveSheet.Cells(Rows.count, 3).End(xlUp).Row + 1 > i Then
            Rows(i & ":" & nxt_r).EntireRow.Delete
        ElseIf ActiveSheet.Cells(Rows.count, 4).End(xlUp).Row + 1 > i Then
            Rows(i & ":" & nxt_r).EntireRow.Delete
        ElseIf ActiveSheet.Cells(Rows.count, 5).End(xlUp).Row + 1 > i Then
            Rows(i & ":" & nxt_r).EntireRow.Delete
        ElseIf ActiveSheet.Cells(Rows.count, 6).End(xlUp).Row + 1 > i Then
            Rows(i & ":" & nxt_r).EntireRow.Delete
        ElseIf ActiveSheet.Cells(Rows.count, 7).End(xlUp).Row + 1 > i Then
            Rows(i & ":" & nxt_r).EntireRow.Delete
        ElseIf ActiveSheet.Cells(Rows.count, 8).End(xlUp).Row + 1 > i Then
            Rows(i & ":" & nxt_r).EntireRow.Delete
        ElseIf ActiveSheet.Cells(Rows.count, 9).End(xlUp).Row + 1 > i Then
            Rows(i & ":" & nxt_r).EntireRow.Delete
        ElseIf ActiveSheet.Cells(Rows.count, 10).End(xlUp).Row + 1 > i Then
            Rows(i & ":" & nxt_r).EntireRow.Delete
        Else
            MsgBox ("No data to clear")
        End If
'Return cursor to RiC1
    ActiveSheet.Cells(i, 1).Select

End Sub

Sub test()
Call global_nom
'used to test formulas or scripts

End Sub
