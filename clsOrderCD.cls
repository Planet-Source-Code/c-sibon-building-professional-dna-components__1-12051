VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrderCD"
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
        strSQL = "SELECT * FROM BESTELING_CDS ORDER BY BESTELING_CDS.BESTELCD_ID ASC"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
    Set ListAll = rs
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
End Function

Public Function Nieuwe_CDBestling(ByVal strCD_ID As String, ByVal strBestel_ID As String) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
    Dim objUID As GUIDGENOBJLib.UniqueID
    Set objUID = New GUIDGENOBJLib.UniqueID
    Dim strUID As String
        strUID = objUID.NewID
        strSQL = "SELECT * FROM BESTELING_CDS"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        With rs
            .AddNew
            !BestelCD_ID = strUID
            !CD_ID = strCD_ID
            !Bestel_ID = strBestel_ID
            .Update
        End With
        Nieuwe_CDBestling = strUID
        rs.Close
    Set rs.ActiveConnection = Nothing
    Set objUID = Nothing
    Set rs = Nothing
End Function

Public Function Edit_CDBesteling(ByVal strBestelCD_ID As String, ByVal strCD_ID As String, ByVal strBestel_ID As String) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
        strSQL = "SELECT * FROM BESTELING_CDS WHERE CD_ID = '" & strBestelCD_ID & "'"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        With rs
            .Fields("CD_ID").Value = strCD_ID
            .Fields("Bestel_ID").Value = strBestel_ID
            .Update
        End With
        Edit_CDBesteling = strBestelCD_ID
        rs.Close
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
End Function

Public Function Delete_CD_Besteling(ByVal strBestelCD_ID As String) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
        strSQL = "SELECT * FROM BESTELING_CDs WHERE CD_ID = '" & strBestelCD_ID & "'"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        With rs
            .Delete
        End With
        Delete_CD_Besteling = strBestelCD_ID
        rs.Close
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
End Function

Public Function GetCDBesteling_ByCD_ID(ByVal strBestelCD_ID As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
        strSQL = "SELECT * FROM BESTELING_CDs WHERE CD_NAME LIKE '%" & strBestelCD_ID & "%'"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        Set GetCDBesteling_ByCD_ID = rs
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
End Function

Public Function GetCD_ByBesteling_ID(ByVal strBestelCD_ID As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
        strSQL = "SELECT * FROM BESTELING_CDS WHERE CD_ID = '" & strBestelCD_ID & "'"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        Set GetCD_ByBesteling_ID = rs
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
End Function
