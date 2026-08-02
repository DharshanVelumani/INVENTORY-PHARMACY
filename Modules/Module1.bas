Attribute VB_Name = "Module1"
Option Explicit

Public con As New ADODB.Connection
Public rs As New ADODB.Recordset
Public rs1 As New ADODB.Recordset
Public rs2 As New ADODB.Recordset
Public rsbill As New ADODB.Recordset
Public Constring As String
Public AdminTargetForm As String

' Windows Crypto API declarations for SHA-256
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, ByVal pbData As Any, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long

Private Const PROV_RSA_AES As Long = 24
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
Private Const ALG_CLASS_HASH As Long = 32768
Private Const ALG_TYPE_ANY As Long = 0
Private Const ALG_SID_SHA_256 As Long = 12
Private Const CALG_SHA_256 As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_256)
Private Const HP_HASHVAL As Long = 2

' PlaySound declaration (left for compatibility)
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Sub Main()
    Dim dbPath As String
    
    ' Detect Access database path dynamically, prioritizing the Database folder copy
    dbPath = App.Path & "\Database\Pharmacy1.mdb" ' Compiled EXE in root folder
    If Dir(dbPath) = "" Then
        dbPath = App.Path & "\..\Database\Pharmacy1.mdb" ' IDE run (relative to Project folder)
        If Dir(dbPath) = "" Then
            dbPath = App.Path & "\Pharmacy1.mdb" ' Compiled EXE in Database folder
            If Dir(dbPath) = "" Then
                dbPath = App.Path & "\Pharmacy1.mdb" ' fallback default
            End If
        End If
    End If
    
    Constring = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & dbPath
End Sub

' Safely opens database connection if closed
Public Function SafeOpenConnection() As Boolean
    On Error Resume Next
    If con.State = adStateOpen Then
        ' Test connection health with a dummy query
        Dim testRs As New ADODB.Recordset
        testRs.Open "SELECT Count(*) FROM Login", con, adOpenForwardOnly, adLockReadOnly
        If Err.Number = 0 Then
            SafeOpenConnection = True
            Exit Function
        Else
            con.Close
        End If
        Err.Clear
    End If
    
    con.ConnectionString = Constring
    con.Open
    If Err.Number = 0 Then
        SafeOpenConnection = True
    Else
        SafeOpenConnection = False
    End If
End Function

' Executes an Insert/Update/Delete SQL statement safely with automatic reconnection
Public Function ExecuteSQL(ByVal sql As String) As Boolean
    On Error GoTo ErrHandler
    If Not SafeOpenConnection() Then
        ExecuteSQL = False
        Exit Function
    End If
    
    con.Execute sql
    ExecuteSQL = True
    Exit Function
    
ErrHandler:
    MsgBox "Database Execution Error: " & Err.Description, vbCritical, "Database Error"
    ExecuteSQL = False
End Function

' Opens and returns a Recordset safely
Public Function GetRecordset(ByVal sql As String) As ADODB.Recordset
    On Error GoTo ErrHandler
    Dim rsTemp As New ADODB.Recordset
    
    If Not SafeOpenConnection() Then
        Set GetRecordset = Nothing
        Exit Function
    End If
    
    rsTemp.Open sql, con, adOpenKeyset, adLockOptimistic
    Set GetRecordset = rsTemp
    Exit Function
    
ErrHandler:
    MsgBox "Database Query Error: " & Err.Description, vbCritical, "Database Error"
    Set GetRecordset = Nothing
End Function

' Helper to escape single quotes to prevent SQL injection
Public Function EscapeSQL(ByVal val As String) As String
    EscapeSQL = Replace(val, "'", "''")
End Function

' Hashing algorithm using native Windows SHA-256 (Crypto API)
Public Function HashPassword(ByVal password As String) As String
    On Error GoTo ErrHandler
    Dim hProv As Long
    Dim hHash As Long
    Dim lLen As Long
    Dim sBuf As String
    Dim i As Long
    Dim sByte As String
    Dim sHash As String
    
    HashPassword = ""
    If password = "" Then Exit Function
    
    ' Acquire context
    If CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_AES, CRYPT_VERIFYCONTEXT) <> 0 Then
        ' Create hash object
        If CryptCreateHash(hProv, CALG_SHA_256, 0, 0, hHash) <> 0 Then
            ' Feed data into hash
            If CryptHashData(hHash, password, Len(password), 0) <> 0 Then
                ' Get hash value length (should be 32 bytes for SHA-256)
                lLen = 32
                sBuf = String(lLen, 0)
                ' Retrieve hash value parameters
                If CryptGetHashParam(hHash, HP_HASHVAL, sBuf, lLen, 0) <> 0 Then
                    ' Convert bytes to hex string
                    sHash = ""
                    For i = 1 To lLen
                        sByte = Hex(Asc(Mid(sBuf, i, 1)))
                        If Len(sByte) = 1 Then sByte = "0" & sByte
                        sHash = sHash & LCase(sByte)
                    Next i
                    HashPassword = sHash
                End If
            End If
            CryptDestroyHash hHash
        End If
        CryptReleaseContext hProv, 0
    End If
    Exit Function
    
ErrHandler:
    HashPassword = ""
End Function
