VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCDContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Copyright (C) by Cândido G. Sibon
'candido_sibon@hetnet.nl
Option Explicit

Const strConn = "Provider=SQLOLEDB.1; Persist Security Info=False; User ID=sa;" & _
                "Initial Catalog=MBase; Data Source=TNTTech"

Public Function ListAll(ByVal strOrderBy As String, ByVal strASCorDESC As Boolean) As ADODB.Recordset
    Dim strType As String
    If strOrderBy = "0" Then
        strOrderBy = "SONG_ARTIST"
    End If
    If strASCorDESC = True Then
        strType = "ASC"
    Else
        strType = "DESC"
    End If
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
        strSQL = "SELECT * FROM SONGS ORDER BY " & strOrderBy & " " & strType & ""
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
    Set ListAll = rs
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
End Function

Public Function Nieuwe_CDContent(ByVal strCDNameID As String, ByVal intTrackNumber As Integer, ByVal strTrackName As String) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
    Dim objUID As GUIDGENOBJLib.UniqueID
    Set objUID = New GUIDGENOBJLib.UniqueID
    Dim strUID As String
        strUID = objUID.NewID
        strSQL = "SELECT * FROM CDCONTENT"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        With rs
            .AddNew
            !CDContents_ID = strUID
            !CD_ID = strCDNameID
            !Track_Number = intTrackNumber
            !Track_Naam = strTrackName
            .Update
        End With
        Nieuwe_CDContent = strUID
        rs.Close
    Set rs.ActiveConnection = Nothing
    Set objUID = Nothing
    Set rs = Nothing
End Function

Public Function Edit_CDContent(ByVal strCDContentID As String, Optional ByVal intTrackNumber As Integer, Optional ByVal strTrackName As String) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
        strSQL = "SELECT * FROM CDCONTENT WHERE CDCONTENTS_ID = '" & strCDContentID & "'"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        With rs
            .Fields("Track_Number").Value = intTrackNumber
            .Fields("Track_Naam").Value = strTrackName
            .Update
        End With
        Edit_CDContent = strCDContentID
        rs.Close
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
End Function

Public Function Delete_CDContent(ByVal strCDContentID As String) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
        strSQL = "SELECT * FROM CDCONTENT WHERE CDCONTENTS_ID = '" & strCDContentID & "'"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        With rs
            .Delete
        End With
        Delete_CDContent = strCDContentID
        rs.Close
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
End Function

Public Function GetCDContent_ByNaam(Optional ByVal strCDTrackName As String, Optional ByVal intTrackNumber As Integer) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
        strSQL = "SELECT * FROM CDCONTENT WHERE TRACK_1 LIKE '%" & strCDTrackName & "%' OR TRACK_2 LIKE '%" & intTrackNumber & "%'"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        Set GetCDContent_ByNaam = rs
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
End Function

Public Function GetCDContent_ByCD_ID(Optional ByVal strCD_ID As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
        strSQL = "SELECT * FROM CDCONTENT WHERE CD_ID = '" & strCD_ID & "' ORDER BY TRACK_NUMBER"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        Set GetCDContent_ByCD_ID = rs
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
End Function

Public Function GetCDContent_ByCDContents_ID(Optional ByVal strCDContents_ID As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim strSQL As String
        strSQL = "SELECT * FROM CDCONTENT WHERE CDContents_ID = '" & strCDContents_ID & "'"
        rs.CursorLocation = adUseClient
        rs.Open strSQL, strConn, adOpenForwardOnly, adLockOptimistic
        Set GetCDContent_ByCDContents_ID = rs
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
End Function
