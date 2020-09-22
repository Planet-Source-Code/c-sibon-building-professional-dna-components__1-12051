VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Copyright (C) by Cândido G. Sibon
'candido_sibon@hetnet.nl
Option Explicit

Const strConn = "Provider=SQLOLEDB.1; Persist Security Info=False; User ID=sa;" & _
                "Initial Catalog=CDWinkel; Data Source=TNTTECH"

Public Function ListAll() As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
        strSQL = "SELECT * FROM KLANTEN ORDER BY KLANTEN.K_NAME ASC"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
    Set ListAll = rs
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
End Function

Public Function Nieuwe_Klant(ByVal strNieuweKlantNaam As String) As String
    Dim strID As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
    Dim objUID As GUIDGENOBJLib.UniqueID
    Set objUID = New GUIDGENOBJLib.UniqueID
    Dim strUID As String
        strUID = objUID.NewID
        strSQL = "SELECT * FROM KLANTEN"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        With rs
            .AddNew
            !K_ID = strUID
            !K_NAME = strNieuweKlantNaam
            .Update
        End With
        Nieuwe_Klant = strUID
        rs.Close
    Set rs.ActiveConnection = Nothing
    Set objUID = Nothing
    Set rs = Nothing
End Function

Public Function Edit_Klant(ByVal strKlantID As String, ByVal strKlantNieuweNaam As String) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
        strSQL = "SELECT * FROM KLANTEN WHERE K_ID = '" & strKlantID & "'"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        With rs
            .Fields("K_NAME").Value = strKlantNieuweNaam
            .Update
        End With
        Edit_Klant = strKlantID
        rs.Close
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
End Function

Public Function Delete_Klant(ByVal strKlantID As String) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
        strSQL = "SELECT * FROM KLANTEN WHERE K_ID = '" & strKlantID & "'"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        With rs
            .Delete
        End With
        Delete_Klant = strKlantID
        rs.Close
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
End Function

Public Function GetKlant_ByNaam(ByVal strKlantNaam As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
        strSQL = "SELECT * FROM KLANTEN WHERE K_NAME LIKE '%" & strKlantNaam & "%'"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        Set GetKlant_ByNaam = rs
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
End Function

Public Function GetKlant_ByID(ByVal strKlantID As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
        strSQL = "SELECT * FROM KLANTEN WHERE K_NAME = '" & strKlantID & "'"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        Set GetKlant_ByID = rs
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
End Function
