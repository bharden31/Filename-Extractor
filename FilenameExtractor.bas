Attribute VB_Name = "Module1"
' The MIT License (MIT)
'
' Copyright (c) 2016 R. Brock Harden
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.


Sub ExtractFileNames()
Attribute ExtractFileNames.VB_ProcData.VB_Invoke_Func = " \n14"
    '
    ' Takes the output of "dir" performed on a directory, and formats it such that
    ' only the relavent filenames remain in alphabetical order
    '
    ' Delete the header rows
    Rows("1:7").Select
    Selection.Delete Shift:=xlUp
    ' Use TextToColumns to trim all the unnecessary information from the sheet
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 9), Array(10, 9), Array(17, 9), Array(24, 9), Array(38, 2)) _
        , TrailingMinusNumbers:=True
    ' Search for "directoryList.txt" in the sheet and delete that row if present
    SeekAndDestroy ("directoryList.txt")
    ' Search for "ListDirectoryContents.bat" in the sheet and delete that row if present
    SeekAndDestroy ("ListDirectoryContents.bat")
    ' Delete the last two rows in the list
    Range("A1").End(xlDown).Select
    ActiveCell.Offset(-1, 0).Range("A1:A2").Select
    Selection.Delete Shift:=xlUp
    ' Ensure that list is in alphabetical order
    Columns("A:A").Sort key1:=Range("A1"), order1:=xlAscending
    Range("A1").Select
End Sub

Sub SeekAndDestroy(target As String)
    '
    ' Searches for the TARGET within the sheet, and deletes the row in which it resides
    ' If TARGET is not found, this subroutine does nothing
    '
    Range("A1").Select
    ' Perform a "find" to see if TARGET exists in the sheet
    Dim rng As Range
    Set rng = Cells.Find(What:=target, After:=ActiveCell, LookIn:= _
        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False)
    ' If TARGET was found, delete the cell and shift up
    If Not rng Is Nothing Then
        rng.Activate
        Selection.Delete Shift:=xlUp
    End If
End Sub
