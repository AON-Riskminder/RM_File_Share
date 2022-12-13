Imports System.Data
Imports System.Reflection.Emit

Public Class file_download
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Init
        '0:linkdate,1:linkid, 2:orgid, 3:formid, 4:svarid, 5:createdby
        ''22-09-2014 11:57:00#135#22#0#-999#0

        Dim DBcon As String = "cnDB_DEV"
        Dim arr() As String = Split(URLcrypt(Request.QueryString(0).ToString, 1), "#")

        If HttpContext.Current.Request.ServerVariables("HTTP_HOST").ToLower = "devshared.riskminder.dk" Then
            DBcon = "cnDB_DEV"
        Else
            Select Case arr(2)
                Case "10"
                    DBcon = "cnDB_DS"
                Case "22"
                    DBcon = "cnDB_DSK"
                Case Else
                    DBcon = "cnDB_SHARED"
            End Select
        End If

        DBexecute("---->" & DBcon & ":::---::::" & arr(2), "cnDB_DEV")
        Try



            Dim ds As DataSet = DBreturnDS("exec [sp_external_validate_and_download] @orgid=" & arr(2) & ",@formularid=" & arr(3) & ",@svarid=" & arr(4) & ",@linkid=" & arr(1) & ",@fileid=" & arr(6), DBcon)
            If ds.Tables.Count = 2 Then
                Dim ts As TimeSpan = Date.Parse(ds.Tables(0).Rows(0).Item("linkexpires")) - Now
                If ts.TotalSeconds > 0 Then
                    download_file(ds.Tables(1))
                Else
                    lblText.Text = "<b>Vi beklager</b><br><br>Dette link er desværre upløbet og kan ikke længere anvendes"
                End If
            Else
                'Link invalid (does not exist)
                lblText.Text = "<b>Vi beklager</b><br><br>Dette link er desværre upløbet og kan ikke længere anvendes"
            End If
        Catch ex As Exception
            DBexecute("----ERROR>" & ex.Message.ToString, "cnDB_DEV")
        End Try
        'Dim dt As DataTable = DBreturnDT( _
        '    " select * from tbl_userfiles_master a " & _
        '    "   inner join tbl_userfiles b ON a.FileID = b.fileID" & _
        '    " where a.FileID = " & arr(0))
    End Sub

    Protected Sub download_file(mydt As DataTable)
        Dim b() As Byte = mydt.Rows(0).Item("FileData")
        If Request.Browser.Browser = "IE" Then
            Response.AppendHeader("Accept-Ranges", "none")
            Response.ContentType = "application/please-download-me"
            Response.AddHeader("Content-Disposition", "attachment; filename=" & mydt.Rows(0).Item("Userfilename").ToString)
            'Response.ContentType = dt.Rows(0).Item("ContentType").ToString
            'Select Case dt.Rows(0).Item("ContentType").ToString.ToLower
            '    Case "application/pdf"
            '        Response.AddHeader("Content-Disposition", "inline; filename=" & dt.Rows(0).Item("Userfilename").ToString)
            '    Case Else
            '        'Response.AddHeader("Content-Disposition", "attachment; filename=" & dt.Rows(0).Item("VersionFileName").ToString)
            'End Select
            'If (Request.Browser.Version = "7.0" OrElse Request.Browser.Version = "8.0") Then
            'End If
        Else


            Response.Clear()
            Response.ClearContent()
            Response.ClearHeaders()
            'Response.ContentType = mydt.Rows(0).Item("ContentType").ToString
            Response.ContentType = "application/octet-stream"
            Dim MyFilename As String = Replace(Replace(mydt.Rows(0).Item("Userfilename").ToString, ",", ""), ".", "")
            Dim MyFileExstension = mydt.Rows(0).Item("UserFileExt").ToString
            MyFilename = MyFilename.Substring(0, Len(MyFilename) - (Len(MyFileExstension) - 1)) & MyFileExstension
            Response.AddHeader("Content-Disposition", "attachment; filename=" & MyFilename)

        End If
        mydt.Dispose()
        Response.BinaryWrite(b)
        Response.End()
    End Sub
End Class
