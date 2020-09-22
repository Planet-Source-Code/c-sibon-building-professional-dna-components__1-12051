VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBussinessLogic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Copyright (C) by Cândido G. Sibon
'candido_sibon@hetnet.nl
Option Explicit

Const strConn = "Provider=SQLOLEDB.1; Persist Security Info=False; User ID=sa;" & _
                "Initial Catalog=CDWinkel; Data Source=TNTTech"

Public Function GetBesteling_ByKlant_ID(ByVal strKlant_ID As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
        strSQL = "SELECT * FROM BESTELING JOIN CDS ON BESTELING.CD_ID = CDS.CD_ID WHERE K_ID = '" & strKlant_ID & "' ORDER BY CDS.CD_NAME"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        Set GetBesteling_ByKlant_ID = rs
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
End Function

Public Function AddBesteling(ByVal strKlant_ID As String, ByVal strCD_ID As String) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
    Dim objUID As GUIDGENOBJLib.UniqueID
    Set objUID = New GUIDGENOBJLib.UniqueID
    Dim strUID As String
        strUID = objUID.NewID
        strSQL = "SELECT * FROM BESTELING"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        With rs
            .AddNew
            !Bestel_ID = strUID
            !K_ID = strKlant_ID
            !CD_ID = strCD_ID
            .Update
        End With
        AddBesteling = strUID
        rs.Close
    Set rs.ActiveConnection = Nothing
    Set objUID = Nothing
    Set rs = Nothing
End Function

Public Function DeleteRecord(ByVal strID As String)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
        strSQL = "SELECT * FROM BESTELING WHERE BESTEL_ID = '" & strID & "'"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        With rs
            .Delete
        End With
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
End Function

Public Function GetBesteling_ByKlant_CDNumber(ByVal strKlant_ID As String) As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim Teller As Integer
    Dim strSQL As String
        strSQL = "SELECT * FROM BESTELING JOIN CDS ON BESTELING.CD_ID = CDS.CD_ID WHERE K_ID = '" & strKlant_ID & "' ORDER BY CDS.CD_NAME"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        Teller = rs.RecordCount
        GetBesteling_ByKlant_CDNumber = Teller
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
End Function

Public Function ListAllOrders() As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
        strSQL = "select k.k_id, k.k_name, count(k.k_name) as aantal from klanten k , besteling b where k.k_id = b.k_id group by k.k_name, k.k_id"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        Set ListAllOrders = rs
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
End Function
