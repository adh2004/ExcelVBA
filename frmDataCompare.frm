VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDataCompare 
   Caption         =   "Data Compare"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "frmDataCompare.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDataCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arr(25) As String

Sub LoadArray()

arr(0) = "A"
arr(1) = "B"
arr(2) = "C"
arr(3) = "D"
arr(4) = "E"
arr(5) = "F"
arr(6) = "G"
arr(7) = "H"
arr(8) = "I"
arr(9) = "J"
arr(10) = "K"
arr(11) = "L"
arr(12) = "M"
arr(13) = "N"
arr(14) = "O"
arr(15) = "P"
arr(16) = "Q"
arr(17) = "R"
arr(18) = "S"
arr(19) = "T"
arr(20) = "U"
arr(21) = "V"
arr(22) = "W"
arr(23) = "X"
arr(24) = "Y"
arr(25) = "Z"


End Sub
Function GetRangeColumn(val As String) As String

Dim i As Integer
Dim str As String

For i = 0 To UBound(arr)
    str = arr(i)
        If val = str Then
            GetRangeColumn = arr(i + 1)
            Exit For
        End If
Next i

End Function
Function ConvertedEnum(val As String) As Integer

Dim i As Integer
Dim str As String
Call LoadArray

For i = 0 To UBound(arr)
    str = arr(i)
        If val = str Then
            ConvertedEnum = i + 1
            Exit For
        End If
Next i
    
    
End Function


Private Sub CommandButton1_Click()

Dim i As Integer
Dim c As Range
Dim str As String
Dim rngStart As String
Dim rngEnd As String
Dim strRng As String
Dim matches As Integer
Dim searchedCells As Integer
Dim matchCheck As String
Dim sheetName As String
Dim colNumber As Integer
Dim matchColNum As Integer


matches = 0
searchedCells = 0
sheetName = Application.ActiveSheet.Name
colNumber = ConvertedEnum(UCase(MasterDSBox.Text))
matchColNum = colNumber + 1

Worksheets(sheetName).Cells(1, colNumber).Activate
ActiveCell.EntireColumn.offset(0, 1).Insert

rngStart = GetRangeColumn(UCase(ComparisonColBox.Text)) + CStr(1)
rngEnd = GetRangeColumn(UCase(ComparisonColBox.Text)) + RowNumBox.Text
strRng = rngStart + ":" + rngEnd


'Loop through cells to check for matches
Do Until IsEmpty(ActiveCell)
    
    matchCheck = ""
    i = ActiveCell.Row
    str = ActiveCell.Value
    
    For Each c In Range(strRng)
        If (ActiveCell.Value = c.Value) Then
            Cells(i, matchColNum).Value = "" & c.Row & ""
            matches = matches + 1
            matchCheck = "Found"
            Exit For
        End If
    Next c
    
    If (matchCheck <> "Found") Then
        Cells(i, matchColNum).Interior.ColorIndex = 3
    End If
        
    searchedCells = searchedCells + 1
    ActiveCell.offset(1, 0).Activate
    
Loop

    ActiveCell.EntireColumn.offset(0, 1).HorizontalAlignment = xlCenter
    
    
    'Populate caption to show number of matches
    lblMatchesFound.Caption = CStr(matches) + " matches found out of " + CStr(searchedCells) + " cells searched"

End Sub





