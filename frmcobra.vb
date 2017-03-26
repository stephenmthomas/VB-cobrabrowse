Imports System.ComponentModel

Public Class cmdGenPrev

    Private Sub GetSource()
        txtSource.Text = New System.Net.WebClient().DownloadString(txtURL.Text)

    End Sub
    Private Sub MakeIMGPage()
        Dim strFile As String = Application.StartupPath & "\preview.html"
        My.Computer.FileSystem.WriteAllText(strFile, txtSource.Text, False)
    End Sub
    Private Sub OpenIMGPage()
        WB.Navigate(Application.StartupPath & "\preview.html")
        txtURL.Text = ""
    End Sub
    Private Sub TrimSource()
        Dim strSrc As String = txtSource.Text 'Search string
        Dim strDelStart As String = "<blockquote class=""postcontent restore "">" 'Entry delimiter
        Dim strDelEnd As String = "</blockquote>" 'End delimiter
        Dim nIndexStart As Integer = strSrc.IndexOf(strDelStart) 'nIndexStart is the first occurance of the entry delimiter
        Dim nIndexEnd As Integer = strSrc.IndexOf(strDelEnd) 'nIndexEnd is the first occurance of the ending delimiter

        If nIndexStart > -1 AndAlso nIndexEnd > -1 Then 'if the nIndexs have a value greater than 1, they occur in the string
            Dim res As String = Strings.Mid(strSrc, nIndexStart + strDelStart.Length + 1, nIndexEnd - nIndexStart - strDelStart.Length) 'removes the text between delimiters
            txtSource.Text = res
        Else
            MessageBox.Show("Delimiters not found, wrong website may have been specified!")
        End If
    End Sub
    Private Sub TrimMore()
        Dim strRep As String = Replace(txtSource.Text, "><", ">" & "[SPLIT]" & "<")
        Dim SplitSrc() As String = Split(strRep, "[SPLIT]")

        txtSource.Text = ""

        For Each chunk In SplitSrc
            If chunk.Contains("img src=") Then
                txtSource.Text = txtSource.Text & Replace(chunk, " border=""0"" alt="""" /", "") & vbNewLine
            End If
        Next

        txtSource.Text = Replace(txtSource.Text, "<img src=""", "")
        txtSource.Text = Replace(txtSource.Text, """>", "")

    End Sub
    Private Sub GenREP(target As String, rep As String)
        'replace target with rep in img codes
        Dim imgarray() As String = txtSource.Lines
        Dim href As String
        Dim src As String

        txtSource.Text = ""

        For Each img In imgarray
            href = "<a href=""" & Replace(img, target, rep) & """>"
            src = "<img src=""" & img & """>"
            txtSource.Text = txtSource.Text & href & " " & src & vbNewLine
        Next
    End Sub
    Private Sub GenREPSR(target As String, rep As String)
        'replace target with rep in img codes
        'only works if target exists in imgarray imgs
        Dim imgarray() As String = txtSource.Lines

        Dim found As Boolean

        Dim href As String
        Dim src As String

        Dim newSource As String

        newSource = ""
        found = False

        For Each img In imgarray
            If InStr(img, target) = 0 Then

            ElseIf InStr(img, target) > 0 Then
                found = True
                href = "<a href=""" & Replace(img, target, rep) & """>"
                src = "<img src=""" & img & """>"

                newSource = newSource & href & " " & src & vbNewLine
            End If
        Next

        If found = True Then
            txtSource.Text = newSource
        End If
    End Sub
    Private Sub GenPrev(targ As String, repl As String)
        'generate a preview using target and replace strings 
        GetSource()
        TrimSource()
        TrimMore()
        GenREP(targ, repl)
        MakeIMGPage()
        OpenIMGPage()
    End Sub

    Private Sub cmdBack_Click(sender As Object, e As EventArgs) Handles cmdBack.Click
        WB.GoBack()

    End Sub

    Private Sub cmdFor_Click(sender As Object, e As EventArgs) Handles cmdFor.Click
        WB.GoForward()

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles cmdOpen.Click
        Dim webAddress As String = Application.StartupPath & "\preview.html"
        Process.Start(webAddress)
    End Sub

    Private Sub cmdAuto_Click(sender As Object, e As EventArgs) Handles cmdAuto.Click
        GetSource()
        TrimSource()
        TrimMore()
        GenREPSR("/t/", "/i/")
        GenREPSR("/small/", "/big/")
        GenREPSR("/th/", "/i/")
        GenREPSR("_t", "")
        GenREPSR("://t", "://i")
        MakeIMGPage()
        OpenIMGPage()
    End Sub
End Class
