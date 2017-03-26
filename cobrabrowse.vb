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
        Dim strSource As String = txtSource.Text 'String that is being searched
        Dim strDelStart As String = "<blockquote class=""postcontent restore "">" 'First delimiting word
        Dim strDelEnd As String = "</blockquote>" 'Second delimiting word
        Dim nIndexStart As Integer = strSource.IndexOf(strDelStart) 'Find the first occurrence of f1
        Dim nIndexEnd As Integer = strSource.IndexOf(strDelEnd) 'Find the first occurrence of f2

        If nIndexStart > -1 AndAlso nIndexEnd > -1 Then '-1 means the word was not found.
            Dim res As String = Strings.Mid(strSource, nIndexStart + strDelStart.Length + 1, nIndexEnd - nIndexStart - strDelStart.Length) 'Crop the text between
            txtSource.Text = res
        Else
            MessageBox.Show("One or both of the delimiting words were not found!")
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
