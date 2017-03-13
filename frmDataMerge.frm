VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDataMerge 
   Caption         =   "Merge Data"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4125
   OleObjectBlob   =   "frmDataMerge.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDataMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim colCollection As Collection
Dim cbArr() As CheckBox

Public Sub GetCbValues()


Dim cb As CheckBox
Dim cCon As Control
Dim cbCol As Collection
Set cbCol = New Collection
Set colCollection = New Collection
Dim i As Integer
Dim o As Object

i = 0

For Each cCon In Me.Controls
        If cCon.Name Like "cb*" Then
            If cCon.Value = True Then
                colCollection.Add (cCon.Caption)
            End If
        End If
       
Next

End Sub

Private Sub btnMerge_Click()

Call MergeCells
    
End Sub

Public Sub MergeCells()

Dim c As Range
Dim i As Integer
Dim v As Variant
Dim str As String
Dim r As Range

Dim rngCol As Collection
Set rngCol = New Collection
Set r = ActiveCell

i = ActiveCell.Row
Call GetCbValues

For Each v In colCollection
    str = CStr(v) + CStr(i)
    rngCol.Add (str)
Next

For Each v In rngCol
    str = Replace(r.Address, "$", "")
    If Range(str) <> Range(CStr(v)) Then
        ActiveCell.Value = ActiveCell.Value + Range(CStr(v)).Value
        Range(CStr(v)).Clear
    End If
Next
    
tbPreview.Text = ActiveCell.Value


End Sub

Private Sub btnMergeAll_Click()

Do Until IsEmpty(ActiveCell.Value)

    Call MergeCells
    ActiveCell.offset(1, 0).Activate
Loop

End Sub

