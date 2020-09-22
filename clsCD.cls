VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Copyright (C) by Cândido G. Sibon
'candido_sibon@hetnet.nl
Option Explicit

Const strConn = "Provider=SQLOLEDB.1; Persist Security Info=False; User ID=sa;" & _
                "Initial Catalog=CDWinkel; Data Source=TNTTech"
                
Public Function ListAll() As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
        strSQL = "SELECT * FROM CDS ORDER BY CDS.CD_NAME ASC"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
    Set ListAll = rs
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
End Function

Public Function Nieuwe_CD(ByVal strNieuweCDNaam As String) As String
    Dim strID As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
    Dim objUID As GUIDGENOBJLib.UniqueID
    Set objUID = New GUIDGENOBJLib.UniqueID
    Dim strUID As String
        strUID = objUID.NewID
        strSQL = "SELECT * FROM CDS"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        With rs
            .AddNew
            !CD_ID = strUID
            !CD_NAME = strNieuweCDNaam
            .Update
        End With
        Nieuwe_CD = strUID
        rs.Close
    Set rs.ActiveConnection = Nothing
    Set objUID = Nothing
    Set rs = Nothing
End Function

Public Function Edit_CD(ByVal strCDID As String, ByVal strCDNieuweNaam As String) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
        strSQL = "SELECT * FROM CDS WHERE CD_ID = '" & strCDID & "'"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        With rs
            .Fields("CD_NAME").Value = strCDNieuweNaam
            .Update
        End With
        Edit_CD = strCDID
        rs.Close
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
End Function

Public Function Delete_CD(ByVal strCDID As String) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
        strSQL = "SELECT * FROM CDS WHERE CD_ID = '" & strCDID & "'"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        With rs
            .Delete
        End With
        Delete_CD = strCDID
        rs.Close
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
End Function

Public Function GetCD_ByNaam(ByVal strCDNaam As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
        strSQL = "SELECT * FROM CDS WHERE CD_NAME LIKE '%" & strCDNaam & "%'"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        Set GetCD_ByNaam = rs
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
End Function

Public Function GetCD_ByID(ByVal strCDID As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
        strSQL = "SELECT * FROM CDS WHERE CD_ID = '" & strCDID & "'"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        Set GetCD_ByID = rs
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
End Function
