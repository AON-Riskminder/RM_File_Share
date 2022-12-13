Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports System.Data
Imports System.Net.Mail
Imports System.Security.Cryptography
Imports System
Imports System.Text
Imports System.IO
Imports System.Data.OleDb
Imports System.Net
Imports System.Web.UI
Imports System.Reflection
Imports System.Xml
Imports System.Xml.XPath
Imports System.Web.HttpUtility
Imports System.Convert
Imports System.Web
Imports System.Text.RegularExpressions
Imports System.Globalization
Imports System.Collections
Imports System.Linq

Public Module Functions

#Region "Crypt"
    Public Function Crypt(ByVal OrigTxt As String, Optional ByVal Action As Byte = 0) As String
        'OrigTxt --> text to be crypted (I've tried to load into OrigTxt a 3Mb binary file and it worked fine)
        'CryptKey --> key to be used to crypt OrigTxt
        'Action (Optional with deafult=0) --> 0:Code  Any other number:Decode
        Dim CryptKey As String = "7bM4zU1hyAEkAbAAMjgzOGVkMjMtOTBlNy00YzJlLWFlYmYtMGI5YTFhZGQy1ThjREUOdx9Hg9vu6aL9cj6p7N3SJIs1"
        'Dim CryptKey As String = "7bM4zU1hyAEkAAAAMjgzOGVkMjMtOTBlNy00YzJlLWFlYmYtMGI5YTFhZGQyYThjREUOdx9Hg9vu6aL9cj6p7N3SJIs1"
        Dim I As Integer, J As Integer, CT As String, Codif As String
        Dim A_Max As Integer, A_Min As Integer, A_Key As Integer, A_Char As Integer, Num As Long, Pos As Long
        I = 1
        Codif = ""
        ' A_Min and A_Max specify the first and last ASCII code to be used for the Base Alphabet
        A_Min = 0
        A_Max = 255
        For J = 1 To Len(OrigTxt)
            'Adjust I value so it can correctly loop through characters of CryptKey
            I = ((I - 1) Mod Len(CryptKey)) + 1
            'Find the ascii code of the current character of CryptKey
            A_Key = Asc(Mid(CryptKey, I, 1))
            'Extract the character to be coded from OrigTxt (based on J value)
            CT = Mid(OrigTxt, J, 1)
            'Obtain the ascii code of the character just extracted
            A_Char = Asc(CT)
            If A_Char > A_Max Then
                Num = A_Char
            Else
                Select Case Action  ' 0=Code  AnyOther=Decode
                    Case 0 'Coding phase
                        Pos = A_Char - A_Min + 1
                        Num = A_Key - 1 + Pos
                        If Num > A_Max Then
                            Num = A_Min - 1 + Num - A_Max
                        End If
                    Case Else 'Decoding phase
                        Pos = A_Char - A_Key + 1
                        If Pos < 0 Then
                            Pos = (A_Max - A_Key + 1) + (A_Char - A_Min + 1)
                        End If
                        Num = A_Min + Pos - 1
                End Select
            End If
            Mid(OrigTxt, J, 1) = Chr(Num)
            'Move to the next character of CryptKey
            I = I + 1
        Next
        Crypt = OrigTxt
    End Function

    Public Function SimpleCrypt(
       ByVal Text As String) As String
        ' Encrypts/decrypts the passed string using 
        ' a simple ASCII value-swapping algorithm
        Dim strTempChar As String = "", i As Integer
        For i = 1 To Len(Text)
            If Asc(Mid$(Text, i, 1)) < 128 Then
                strTempChar =
                CType(Asc(Mid$(Text, i, 1)) + 128, String)
            ElseIf Asc(Mid$(Text, i, 1)) > 128 Then
                strTempChar =
                CType(Asc(Mid$(Text, i, 1)) - 128, String)
            End If
            Mid$(Text, i, 1) = Chr(CType(strTempChar, Integer))
        Next i
        Return Text
    End Function

    Private URLcryptKey As String = "sDfjk324kDfjG234oPsk"

    Public Function URLcrypt(ByVal CryptString As String, Optional ByVal Action As Integer = 0, Optional EncodeURL As Boolean = True, Optional decodeUrl As Boolean = False) As String
        Dim byKey() As Byte = {}
        Dim IV() As Byte = {&H12, &H34, &H56, &H78, &H90, &HAB, &HCD, &HEF}
        If Action = 0 Then
            'Encrypt
            Try
                byKey = System.Text.Encoding.UTF8.GetBytes(Left(URLcryptKey, 8))
                Dim des As New DESCryptoServiceProvider()
                Dim inputByteArray() As Byte = Encoding.UTF8.GetBytes(CryptString)
                Dim ms As New MemoryStream()
                Dim cs As New CryptoStream(ms, des.CreateEncryptor(byKey, IV), CryptoStreamMode.Write)
                cs.Write(inputByteArray, 0, inputByteArray.Length)
                cs.FlushFinalBlock()
                If EncodeURL Then
                    Return UrlEncode(Convert.ToBase64String(ms.ToArray()))
                Else
                    Return Convert.ToBase64String(ms.ToArray())
                End If
            Catch ex As Exception
                Return ex.Message
            End Try
        Else
            'Decrypt
            If decodeUrl Then CryptString = UrlDecode(CryptString)
            Dim inputByteArray(CryptString.Length) As Byte
            Try
                byKey = System.Text.Encoding.UTF8.GetBytes(Left(URLcryptKey, 8))
                Dim des As New DESCryptoServiceProvider()
                inputByteArray = Convert.FromBase64String(CryptString)
                Dim ms As New MemoryStream()
                Dim cs As New CryptoStream(ms, des.CreateDecryptor(byKey, IV), CryptoStreamMode.Write)
                cs.Write(inputByteArray, 0, inputByteArray.Length)
                cs.FlushFinalBlock()
                Dim encoding As System.Text.Encoding = System.Text.Encoding.UTF8
                Return encoding.GetString(ms.ToArray())
            Catch ex As Exception
                Return ex.Message
            End Try
        End If
    End Function

    Private vstrEncryptionKey As String = "4500¤%/Uupf45%2rtåpf8--GgGgløåzxcÆØZer"

    Public Function EncryptString128Bit(ByVal vstrTextToBeEncrypted As String) As String
        Dim bytValue() As Byte
        Dim bytKey() As Byte
        Dim bytEncoded() As Byte = Nothing
        Dim bytIV() As Byte = {121, 241, 10, 1, 132, 74, 11, 39, 255, 91, 45, 78, 14, 211, 22, 62}
        Dim intLength As Integer
        Dim intRemaining As Integer
        Dim objMemoryStream As New MemoryStream()
        Dim objCryptoStream As CryptoStream
        Dim objRijndaelManaged As RijndaelManaged


        '   **********************************************************************
        '   ******  Strip any null character from string to be encrypted    ******
        '   **********************************************************************

        vstrTextToBeEncrypted = StripNullCharacters(vstrTextToBeEncrypted)

        '   **********************************************************************
        '   ******  Value must be within ASCII range (i.e., no DBCS chars)  ******
        '   **********************************************************************

        bytValue = Encoding.ASCII.GetBytes(vstrTextToBeEncrypted.ToCharArray)

        intLength = Len(vstrEncryptionKey)

        '   ********************************************************************
        '   ******   Encryption Key must be 256 bits long (32 bytes)      ******
        '   ******   If it is longer than 32 bytes it will be truncated.  ******
        '   ******   If it is shorter than 32 bytes it will be padded     ******
        '   ******   with upper-case Xs.                                  ****** 
        '   ********************************************************************

        If intLength >= 32 Then
            vstrEncryptionKey = Strings.Left(vstrEncryptionKey, 32)
        Else
            intLength = Len(vstrEncryptionKey)
            intRemaining = 32 - intLength
            vstrEncryptionKey = vstrEncryptionKey & Strings.StrDup(intRemaining, "X")
        End If

        bytKey = Encoding.ASCII.GetBytes(vstrEncryptionKey.ToCharArray)

        objRijndaelManaged = New RijndaelManaged()

        '   ***********************************************************************
        '   ******  Create the encryptor and write value to it after it is   ******
        '   ******  converted into a byte array                              ******
        '   ***********************************************************************

        Try
            objCryptoStream = New CryptoStream(objMemoryStream,
              objRijndaelManaged.CreateEncryptor(bytKey, bytIV),
              CryptoStreamMode.Write)
            objCryptoStream.Write(bytValue, 0, bytValue.Length)
            objCryptoStream.FlushFinalBlock()
            bytEncoded = objMemoryStream.ToArray
            objMemoryStream.Close()
            objCryptoStream.Close()
        Catch
        End Try

        '   ***********************************************************************
        '   ******   Return encryptes value (converted from  byte Array to   ******
        '   ******   a base64 string).  Base64 is MIME encoding)             ******
        '   ***********************************************************************

        Return Convert.ToBase64String(bytEncoded)
    End Function

    Public Function DecryptString128Bit(ByVal vstrStringToBeDecrypted As String) As String

        Dim bytDataToBeDecrypted() As Byte
        Dim bytTemp() As Byte
        Dim bytIV() As Byte = {121, 241, 10, 1, 132, 74, 11, 39, 255, 91, 45, 78, 14, 211, 22, 62}
        Dim objRijndaelManaged As New RijndaelManaged()
        Dim objMemoryStream As MemoryStream
        Dim objCryptoStream As CryptoStream
        Dim bytDecryptionKey() As Byte

        Dim intLength As Integer
        Dim intRemaining As Integer
        'Dim intCtr As Integer
        Dim strReturnString As String = String.Empty
        'Dim achrCharacterArray() As Char
        'Dim intIndex As Integer


        bytDataToBeDecrypted = Convert.FromBase64String(vstrStringToBeDecrypted)

        intLength = Len(vstrEncryptionKey)

        If intLength >= 32 Then
            vstrEncryptionKey = Strings.Left(vstrEncryptionKey, 32)
        Else
            intLength = Len(vstrEncryptionKey)
            intRemaining = 32 - intLength
            vstrEncryptionKey = vstrEncryptionKey & Strings.StrDup(intRemaining, "X")
        End If

        bytDecryptionKey = Encoding.ASCII.GetBytes(vstrEncryptionKey.ToCharArray)

        ReDim bytTemp(bytDataToBeDecrypted.Length)

        objMemoryStream = New MemoryStream(bytDataToBeDecrypted)

        '   ***********************************************************************
        '   ******  Create the decryptor and write value to it after it is   ******
        '   ******  converted into a byte array                              ******
        '   ***********************************************************************

        Try

            objCryptoStream = New CryptoStream(objMemoryStream,
               objRijndaelManaged.CreateDecryptor(bytDecryptionKey, bytIV),
               CryptoStreamMode.Read)

            objCryptoStream.Read(bytTemp, 0, bytTemp.Length)

            objCryptoStream.FlushFinalBlock()
            objMemoryStream.Close()
            objCryptoStream.Close()

        Catch

        End Try

        '   *****************************************
        '   ******   Return decypted value     ******
        '   *****************************************

        Return StripNullCharacters(Encoding.ASCII.GetString(bytTemp))

    End Function

    Public Function StripNullCharacters(ByVal vstrStringWithNulls As String) As String

        Dim intPosition As Integer
        Dim strStringWithOutNulls As String

        intPosition = 1
        strStringWithOutNulls = vstrStringWithNulls

        Do While intPosition > 0
            intPosition = InStr(intPosition, vstrStringWithNulls, vbNullChar)

            If intPosition > 0 Then
                strStringWithOutNulls = Left$(strStringWithOutNulls, intPosition - 1) &
                   Right$(strStringWithOutNulls, Len(strStringWithOutNulls) - intPosition)
            End If

            If intPosition > strStringWithOutNulls.Length Then
                Exit Do
            End If
        Loop

        Return strStringWithOutNulls

    End Function

    Private SupergrosCKey As String = "sAfjk356kAfjA253oAsk"

    Public Function SuperGrossURLcrypt(ByVal CryptString As String, Optional ByVal Action As Integer = 0) As String
        Dim byKey() As Byte = {}
        Dim IV() As Byte = {&H12, &H34, &H56, &H78, &H90, &HAB, &HCD, &HEF}
        If Action = 0 Then
            Try
                byKey = System.Text.Encoding.UTF8.GetBytes(Left(SupergrosCKey, 8))
                Dim des As New DESCryptoServiceProvider()
                Dim inputByteArray() As Byte = Encoding.UTF8.GetBytes(CryptString)
                Dim ms As New MemoryStream()
                Dim cs As New CryptoStream(ms, des.CreateEncryptor(byKey, IV), CryptoStreamMode.Write)
                cs.Write(inputByteArray, 0, inputByteArray.Length)
                cs.FlushFinalBlock()
                Return Convert.ToBase64String(ms.ToArray())
            Catch ex As Exception
                Return ex.Message
            End Try
        Else
            Dim inputByteArray(CryptString.Length) As Byte
            Try
                byKey = System.Text.Encoding.UTF8.GetBytes(Left(SupergrosCKey, 8))
                Dim des As New DESCryptoServiceProvider()
                inputByteArray = Convert.FromBase64String(CryptString)
                Dim ms As New MemoryStream()
                Dim cs As New CryptoStream(ms, des.CreateDecryptor(byKey, IV), CryptoStreamMode.Write)
                cs.Write(inputByteArray, 0, inputByteArray.Length)
                cs.FlushFinalBlock()
                Dim encoding As System.Text.Encoding = System.Text.Encoding.UTF8
                Return encoding.GetString(ms.ToArray())
            Catch ex As Exception
                Return ex.Message
            End Try
        End If
    End Function

#End Region

#Region "SQL functions"
    'Public Shared Sub DBwrite(sqlcmd As String)
    '    Try
    '        Dim cn_Exec As SqlConnection
    '        cn_Exec = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("cnDB").ConnectionString())
    '        Dim exec_cmd As SqlCommand = New SqlCommand("--" & sqlcmd, cn_Exec)
    '        cn_Exec.Open()
    '        exec_cmd.ExecuteScalar()
    '        cn_Exec.Close()
    '    Catch ex As Exception
    '        DBexecute("--" & ex.Message)
    '    End Try
    'End Sub

    Public Sub DBexecute(ByVal sqlcmd As String, DbConn As String)
        Try
            Dim cn_Exec As SqlConnection
            'cn_Exec = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings(HttpContext.Current.Session("cnDB")).ConnectionString())
            cn_Exec = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings(DbConn).ConnectionString())
            Dim exec_cmd As SqlCommand = New SqlCommand(sqlcmd, cn_Exec)
            cn_Exec.Open()
            exec_cmd.ExecuteScalar()
            cn_Exec.Close()
        Catch ex As Exception
            'DBexecute("--Errorhandler: " & ex.Message)
            'If ex.Message.ToString.IndexOf("Incorrect syntax near") > -1 Then
            '    HttpContext.Current.Session.Clear()
            '    HttpContext.Current.Response.Redirect("/default.aspx")
            'End If
            'ErrorHandler(ex, sqlcmd)
        End Try
    End Sub

    Public Function DBreturnDS(ByVal sqlcmd As String, DbConn As String) As DataSet
        Dim ds As New DataSet
        Try
            Dim cn_Exec As SqlConnection
            'cn_Exec = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings(HttpContext.Current.Session("cnDB")).ConnectionString())
            cn_Exec = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings(DbConn).ConnectionString())
            Dim exec_cmd As SqlCommand = New SqlCommand(sqlcmd, cn_Exec)
            cn_Exec.Open()
            Dim dbAdapter As New System.Data.SqlClient.SqlDataAdapter(sqlcmd, cn_Exec)
            dbAdapter.SelectCommand.CommandTimeout = 240
            dbAdapter.Fill(ds)
            dbAdapter.Dispose()
            cn_Exec.Close()
        Catch ex As Exception
            'DBexecute("-->1" & ex.Message.ToString.IndexOf("Incorrect syntax near"))
            'If ex.Message.ToString.IndexOf("Incorrect syntax near") > -1 Then
            '    HttpContext.Current.Session.Clear()
            '    HttpContext.Current.Response.Redirect("/default.aspx")
            'End If
            'ErrorHandler(ex, sqlcmd)
        End Try
        Return ds
    End Function

    Public Function DBreturnDT(ByVal sqlcmd As String, DbConn As String) As DataTable
        Dim dt As New DataTable
        Try
            Dim cn_Exec As New SqlConnection
            'cn_Exec = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings(HttpContext.Current.Session("cnDB")).ConnectionString())
            cn_Exec = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings(DbConn).ConnectionString())
            Dim exec_cmd As SqlCommand = New SqlCommand(sqlcmd, cn_Exec)
            cn_Exec.Open()
            Dim dbAdapter As New System.Data.SqlClient.SqlDataAdapter(sqlcmd, cn_Exec)
            dbAdapter.SelectCommand.CommandTimeout = 240
            dbAdapter.Fill(dt)
            dbAdapter.Dispose()
            cn_Exec.Close()
        Catch ex As Exception
            'DBexecute("-->1" & ex.Message.ToString.IndexOf("Incorrect syntax near"))
            'If ex.Message.ToString.IndexOf("Incorrect syntax near") > -1 Then
            '    HttpContext.Current.Session.Clear()
            '    HttpContext.Current.Response.Redirect("/default.aspx")
            'End If
            'ErrorHandler(ex, sqlcmd)
        End Try
        Return dt
    End Function

#End Region


    'Public Shared Function GetFileExtensionImage(ByVal FileExt As String) As String
    '    Dim result As String = ""
    '    Select Case FileExt.ToLower
    '        Case ".jpg", ".gif", ".png", ".bmp"
    '            result = "imageicon.png"
    '        Case ".docx", ".doc", ".dot"
    '            result = "word.png"
    '        Case ".xls", ".xlsx"
    '            result = "excel.png"
    '        Case ".pdf"
    '            result = "pdf.png"
    '        Case "fupload"
    '            result = "fupload.png"
    '        Case Else
    '            result = "unknown.png"
    '    End Select
    '    Return "/images/shared/" & result
    'End Function

    'Public Shared Function GetHexColor(ByVal colorObj As System.Drawing.Color) As String
    '    Return "#" & Hex(colorObj.R) & Hex(colorObj.G) & Hex(colorObj.B)
    'End Function

    'Public Shared Sub WriteToLog(ByVal LogTable As String, ByVal ID As Integer, ByVal LogDesc As String, ByVal Change As String)
    '    DBexecute( _
    '     " INSERT " & LogTable & " (CreatedBy,ChangeID,LogComment,SQLstatement)" & _
    '     " VALUES (" & HttpContext.Current.Session("BrugerID") & "," & ID & ",'" & LogDesc & "','" & Change & "')")
    'End Sub

    'Public Shared Function UserHasAccessTo(ByVal UserAccessString As String, ByVal SecID As Integer) As Boolean
    '    If InStr(UserAccessString, "," & SecID & ",") Then
    '        Return True
    '    Else
    '        Return False
    '    End If
    'End Function

    'Public Shared Function SimpleEncrypt(ByVal TheText As String) As String
    '    Dim tempChar As String = Nothing
    '    Dim i As Integer = 0
    '    For i = 1 To TheText.Length
    '        If ToInt32(TheText.Chars(i - 1)) < 128 Then
    '            tempChar = System.Convert.ToString(ToInt32(TheText.Chars(i - 1)) + 100)
    '        ElseIf ToInt32(TheText.Chars(i - 1)) > 128 Then
    '            tempChar = System.Convert.ToString(ToInt32(TheText.Chars(i - 1)) - 100)
    '        End If
    '        TheText = TheText.Remove(i - 1, 1).Insert(i - 1, (CChar(ChrW(tempChar))).ToString())
    '    Next i
    '    Return TheText
    'End Function

    'Public Shared Function CleanSQLText(ByVal Text As String) As String
    '    Dim tmpText As String = Text
    '    If tmpText <> "" Then
    '        tmpText = Replace(tmpText, "'", "''")
    '        tmpText = Replace(tmpText, "--", "")
    '    End If
    '    Return tmpText
    'End Function

    'Public Shared Function GetHost(ByVal strHost As String) As String
    '    Select Case LCase(strHost)
    '        Case "devrecord.riskminder.dk"
    '            Return "https://devrecord.riskminder.dk/"
    '        Case "devlive.riskminder.dk"
    '            Return "https://devlive.riskminder.dk/"
    '        Case "demorecord.riskminder.dk"
    '            Return "https://demorecord.riskminder.dk/"
    '        Case "supergros.riskminder.dk"
    '            Return "https://supergros.riskminder.dk/"
    '        Case "pisiffik.riskminder.dk"
    '            Return "https://pisiffik.riskminder.dk/"
    '        Case "dsg.riskminder.dk"
    '            Return "https://dsg.riskminder.dk/"
    '        Case "altadiscount.riskminder.dk"
    '            Return "https://altadiscount.riskminder.dk/"
    '        Case "kiwidanmark.riskminder.dk"
    '            Return "https://kiwidanmark.riskminder.dk/"
    '        Case "beta.riskminder.dk"
    '            Return "https://beta.riskminder.dk/"
    '        Case "scaniarecord.riskminder.dk"
    '            Return "https://beta.riskminder.dk/"
    '        Case "devrecord2.riskminder.dk"
    '            Return "https://beta.riskminder.dk/"
    '        Case "demormrecord.riskminder.dk"
    '            Return "https://demormrecord.riskminder.dk/"
    '        Case Else
    '            Return "https://rmrecord.riskminder.dk/"
    '    End Select
    'End Function

    ''Public Shared Function GetCnString(ByVal strHost As String) As String
    ''	Select Case LCase(strHost)
    ''		Case "devrecord.riskminder.dk", "devlive.riskminder.dk"
    ''			Return "DevConnectionStrMaster"
    ''		Case "demorecord.riskminder.dk"
    ''			Return "DemoConnectionStrMaster"
    ''		Case "dsg.riskminder.dk"
    ''			Return "DSGConnectionStrMaster"
    ''		Case "supergros.riskminder.dk", "pisiffik.riskminder.dk"
    ''			Return "ConnectionStrMaster"
    ''		Case "demormrecord.riskminder.dk"
    ''			Return "ConnectionStrMaster"
    ''		Case "altadiscount.riskminder.dk"
    ''			Return "ConnectionStrMaster"
    ''		Case "kiwidanmark.riskminder.dk"
    ''			Return "ConnectionStrMaster"
    ''		Case "rmrecord.riskminder.dk"
    ''			Return "ConnectionStrMaster"
    ''		Case "beta.riskminder.dk"
    ''			Return "DevConnectionStrMaster"
    ''		Case "scaniarecord.riskminder.dk"
    ''			Return "ConnectionStrMaster"
    ''		Case "devrecord2.riskminder.dk"
    ''			Return "ConnectionStrMaster"
    ''		Case Else
    ''			Return ""
    ''	End Select
    ''End Function


    ''Public Shared Function SiteMapXmlFile(ByVal OrgId As String, ByVal XmlFile As String, ByVal OriginalPath As String) As String
    ''    If File.Exists("c\Inetpub\rmrecord.riskminder.dk\Include\ContextMenuer\" & OrgId & "-" & XmlFile) Then
    ''        SiteMapXmlFile = "/ContextMenuer/" & OrgId & "-" & XmlFile
    ''    Else
    ''        SiteMapXmlFile = OriginalPath & XmlFile
    ''    End If

    ''End Function

    'Private Shared Function ApplicationPath() As String
    '    Return _
    '    Path.GetDirectoryName( _
    '    [Assembly].GetExecutingAssembly().Location _
    '    )
    'End Function

    ''Public Shared Function Translate(ByVal Originaltxt As String, ByVal translateID As String, Optional ByVal ShowNumbers As String = "False-0") As String

    ''    Dim MySplit As Array
    ''    MySplit = Split(ShowNumbers, "-")

    ''    Translate = "Error"
    ''    If IsDBNull(translateID) Or Len(translateID) < 1 Then
    ''        If MySplit(0) = "True" Then
    ''            Translate = "(" & MySplit(1) & ") " & Originaltxt
    ''        Else
    ''            Translate = Originaltxt
    ''        End If
    ''        Exit Function
    ''    Else
    ''        If MySplit(0) = "True" Then
    ''            Translate = "(" & MySplit(1) & ") " & translateID
    ''        Else
    ''            Translate = translateID
    ''        End If

    ''    End If

    ''End Function

    'Public Shared Function AddON(ByVal orgid As Integer, ByVal myAddon As String, ByVal strConnection As String) As Boolean

    '    If Right(strConnection, 6) = "server" Then
    '        strConnection = strConnection
    '    Else
    '        strConnection = System.Configuration.ConfigurationManager.AppSettings(strConnection)
    '    End If

    '    Dim myConnection As New SqlConnection(strConnection)
    '    Dim myCommand As SqlCommand = _
    '     New SqlCommand( _
    '     "Select * from dbo.tbl_Org_AddOns where TilkøbsID = '" & myAddon & "' And OrgId = '" & orgid & "'", myConnection)

    '    Dim drTop As SqlDataReader
    '    Try
    '        myConnection.Open()
    '        drTop = myCommand.ExecuteReader
    '        If drTop.HasRows Then
    '            Return True
    '        Else
    '            Return False
    '        End If
    '        drTop.Close()

    '    Catch ex As Exception
    '        Return False
    '    End Try
    'End Function


End Module

