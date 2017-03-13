VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataControl 
   Caption         =   "CRUD"
   ClientHeight    =   3255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5535
   OleObjectBlob   =   "DataControl.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DataControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim con As Object
Dim catalog As Object
Dim adWrap As New clsAdoWrapper

Private Sub BtnInsert_Click()
    Call InsertData
End Sub

Private Sub ComboBox1_Change()
    Call UpDateColumnHeaders
End Sub

Private Sub UpDateColumnHeaders()
'To set the excel column headers based on selected table columns

    Dim i As Integer
    adWrap.DbConnect
    Rows(1).EntireRow.Clear
    
    Worksheets("Sheet1").Cells(1, 1).Activate
    For i = 0 To adWrap.catalog.Tables(Me.ComboBox1.Value).Columns.count - 1
        ActiveCell.Value = adWrap.catalog.Tables(Me.ComboBox1.Value).Columns(i).Name
        ActiveCell.BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
        ActiveCell.Offset(0, 1).Activate
    Next i
    
    adWrap.DbDisconnect
    'Add Update Column for updating particular records
    ActiveCell.Offset(0, 1).Value = "Update"
    ActiveCell.Offset(0, 1).BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
    ActiveCell.Offset(0, 2).Value = "Delete"
    ActiveCell.Offset(0, 2).BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
    
    Rows(1).EntireColumn.AutoFit
    Worksheets("Sheet1").Cells(1, 1).Activate
    
End Sub

Private Sub CommandButton2_Click()
    Call DeleteData
End Sub

Private Sub DeleteData()
    Dim i As Integer: i = 1
    Dim query As String: query = "DELETE FROM " & Me.ComboBox1.Text & " WHERE "
    Dim rowCount As Long
    Dim colCount As Long
    
    
    Call PopulateHeaderCollection
    rowCount = Cells(Rows.count, 1).End(xlUp).Row
    colCount = Cells(1, Columns.count).End(xlToLeft).Column - 1
    
    adWrap.DbConnect
    On Error GoTo ErrHandler
    
    Worksheets("Sheet1").Cells(2, colCount - 1).Activate
    Do Until (ActiveCell.Row = rowCount + 1)
        If ActiveCell.Value = "X" Then
            query = query + adWrap.tableCol.Item(i) + " = " + "'" + Cells(ActiveRow, 1) + "'"
            adWrap.con.Execute query, , adCmdText
            ActiveCell.Offset(1, 0).Activate
            
    Loop
    adWrap.DbDisconnect
    Exit Sub
ErrHandler:
            MsgBox "Error, Database disconnected" + Err.Description
End Sub

Private Sub CommandButton4_Click()

    Dim TableName As String: TableName = Me.ComboBox1.Value
    
    Worksheets("Sheet1").Cells(2, 1).Activate
    
    Do Until IsEmpty(ActiveCell)
        ActiveCell.EntireRow.Clear
        ActiveCell.Offset(1, 0).Activate
    Loop
    Rows(1).EntireColumn.AutoFit
    Worksheets("Sheet1").Cells(1, 1).Activate
    
    adWrap.DbConnect
    adWrap.RetrieveData (TableName)
    Sheet1.Range("A2").CopyFromRecordset adWrap.rs
    adWrap.DbDisconnect
    Sheet1.Rows.EntireColumn.AutoFit
    
End Sub

Private Sub CommandButton5_Click()
    Call ClearContents
End Sub
Public Sub ClearContents()
'Clear all contents except the table headers

    Worksheets("Sheet1").Cells(2, 1).Activate
    
    Do Until IsEmpty(ActiveCell)
        ActiveCell.EntireRow.Clear
        ActiveCell.Offset(1, 0).Activate
    Loop
    Worksheets("Sheet1").Cells(2, 1).Activate
    
End Sub

Private Sub UserForm_Initialize()

    Dim i As Integer
    adWrap.DbConnect
    
    For i = 0 To adWrap.catalog.Tables.count - 1
        If adWrap.catalog.Tables.Item(i).Type = "TABLE" Then
            Me.ComboBox1.AddItem (adWrap.catalog.Tables.Item(i).Name)
        End If
        
    Next
    adWrap.DbDisconnect
    
End Sub

Public Sub InsertData()
'Inserts Data into the Access Database. Called from the Insert button

    Dim sFields As String
    Dim count As Integer
    Dim sLength As String
    Dim endStr As String
    Dim endLength As Integer
    Dim query As String
    Dim insertVal As String
    
    
    sFields = sBuilder("Insert")
    sLength = Len(sFields)
    endLength = sLength - 2
    endStr = Mid(sFields, endLength, 3)
    
    adWrap.DbConnect
    
    On Error GoTo ErrHandler
    
    If Mid(endStr, 1, 1) = "," Then
        count = CInt(Right(endStr, 2))
        sFields = Replace(sFields, endStr, "")
    Else
        count = CInt(Right(endStr, 1))
        endStr = Mid(endStr, 2, 2)
        sFields = Replace(sFields, endStr, "")
    End If
    
    Worksheets("Sheet1").Cells(2, 1).Activate
    Do Until IsEmpty(ActiveCell)
        Do Until IsEmpty(ActiveCell)
            If ActiveCell.Offset(0, 1).Value = "Update" Then
                insertVal = insertVal + "'" + CStr(ActiveCell.Value) + "'"
            Else
                insertVal = insertVal + "'" + CStr(ActiveCell.Value) + "'" + ","
            End If
            ActiveCell.Offset(0, 1).Activate
        Loop
        query = "INSERT INTO  " & Me.ComboBox1.Value & "(" & sFields & ") VALUES(" & insertVal & ")"
        adWrap.con.Execute query, , adCmdText
        Worksheets("Sheet1").Cells(ActiveCell.Row + 1, 1).Activate
    Loop
    adWrap.DbDisconnect
    
    Exit Sub
    
    
     
ErrHandler:
            adWrap.DbDisconnect
            MsgBox "Error, database disconnected" + Err.Description
End Sub

Private Sub CommandButton1_Click()
    Call UpdateData
End Sub

Public Sub UpdateData()
'This sub checks each row in the UPDATE column to see if it needs to be updated.
'If the row in question is marked with an X then the row is considered to need updating
'Each data column in the row is coupled with its column header and used to build the UPDATE query
'The update query is run once the query is built.  After all rows are checked the database is disconnected.

    Dim query As String: query = "UPDATE " & Me.ComboBox1.Value & " SET "
    Dim whereClause As String
    Dim rowCount As Long
    Dim colCount As Long
    Dim i As Integer: i = 2
    Call PopulateHeaderCollection
    
    rowCount = Cells(Rows.count, 1).End(xlUp).Row
    colCount = Cells(1, Columns.count).End(xlToLeft).Column - 1
    adWrap.DbConnect
    
    On Error GoTo ErrHandler
    
    Worksheets("Sheet1").Cells(2, colCount).Activate
    Do Until (ActiveCell.Row = rowCount + 1)
        If ActiveCell.Value = "X" Then
            Worksheets("Sheet1").Cells(ActiveCell.Row, 2).Activate
            Do Until IsEmpty(ActiveCell)
                If ActiveCell.Offset(0, 1).Value = "" Then
                    query = query + adWrap.tableCol.Item(i) + " = " + "'" + ActiveCell.Value + "'"
                Else
                    query = query + adWrap.tableCol.Item(i) + " = " + "'" + ActiveCell.Value + "'" + ","
                End If
                ActiveCell.Offset(0, 1).Activate
                i = i + 1
            Loop
            whereClause = " WHERE AID = " & Worksheets("Sheet1").Cells(ActiveCell.Row, 1).Value & ""
            query = query + whereClause
            adWrap.con.Execute query, , adCmdText
            Worksheets("Sheet1").Cells(ActiveCell.Row + 1, colCount).Activate
        End If
        ActiveCell.Offset(1, 0).Activate
    Loop
    adWrap.DbDisconnect
    Exit Sub
    
ErrHandler:
            adWrap.DbDisconnect
            MsgBox "Error, Database Disconnected " + Err.Description
End Sub
Public Sub PopulateHeaderCollection()

    Worksheets("Sheet1").Cells(1, 1).Activate
    Do Until IsEmpty(ActiveCell)
        adWrap.tableCol.Add (ActiveCell.Value)
        ActiveCell.Offset(0, 1).Activate
    Loop
    
End Sub
Public Function sBuilder(ByVal Action As String) As String

    Dim count As Integer
    Worksheets("Sheet1").Cells(1, 1).Activate
    
    Select Case (Action)
        
        Case "Insert"
            Do Until IsEmpty(ActiveCell)
                If ActiveCell.Offset(0, 1).Value = "" Then
                    sBuilder = sBuilder + ActiveCell.Value
                Else
                    sBuilder = sBuilder + ActiveCell.Value + ","
                End If
                ActiveCell.Offset(0, 1).Activate
                count = count + 1
            Loop
            sBuilder = sBuilder + "," + CStr(count)
            
        Case "Update"
            Do Until IsEmpty(ActiveCell)
    End Select
    
End Function
