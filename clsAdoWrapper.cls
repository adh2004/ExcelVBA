VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAdoWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private sCon As String:
Public con As Object
Public catalog As Object
Public rs As Object
Public tableCol As Collection

Private Sub Class_Initialize()
    'Set object variables and connection string
    
    sCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " _
    & "C:\Users\andre\Desktop\WBL_Project\WBL_ProjectDB.mdb;Persist Security Info=False;"
    Set catalog = CreateObject("ADOX.Catalog")
    Set con = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.RecordSet")
    con.ConnectionString = sCon
    Set tableCol = New Collection
    
End Sub
Public Sub DbConnect()
    'Check for connection to database and connect if connection is closed
    
On Error GoTo ErrHandler
    If con.State = adStateClosed Then
        con.Open
        catalog.ActiveConnection = con
    Else
        Exit Sub
    End If
Exit Sub
    
ErrHandler:
    MsgBox "Error connecting to database " + Err.Description
End Sub

Public Sub DbDisconnect()
    'Check for connections to database and disconnect if connection is open
    
On Error GoTo ErrHandler
    If con.State = adStateOpen Then
        con.Close
    Else
        Exit Sub
    End If
Exit Sub

ErrHandler:
    MsgBox "Error disconnecting from database " + Err.Description
End Sub
Public Sub RetrieveData(ByVal TableName As String)
   Dim query As String: query = "SELECT * FROM " & TableName & ""
        
        rs.Open query, con
        
End Sub

Private Sub Class_Terminate()
    Set catalog = Nothing
    Set con = Nothing
End Sub
