Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.IO
Imports System.Globalization

Public Class MainForm

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim folderBrowserDialog1 As FolderBrowserDialog = New FolderBrowserDialog
        Dim result As DialogResult = folderBrowserDialog1.ShowDialog()

        If result = DialogResult.OK Then
            DisplayList(folderBrowserDialog1.SelectedPath)
        End If

    End Sub

    Private Sub DisplayList(folderName As String)

        Dim files() As String = IO.Directory.GetFiles(folderName)
        Dim fileList As List(Of KeyValuePair(Of String, Long)) = New List(Of KeyValuePair(Of String, Long))
        Dim msg As New System.Text.StringBuilder

        For Each file As String In files
            Dim myFile As New FileInfo(file)
            fileList.Add(New KeyValuePair(Of String, Long)(myFile.Name, myFile.Length))
        Next

        fileList.Sort(Function(kv1, kv2) kv1.Value.CompareTo(kv2.Value))

        Dim amountOfFiles As Integer = fileList.Count
        If amountOfFiles > 3 Then amountOfFiles = 3

        msg.Append(fileList.Count.ToString() & " files processed. " & vbCrLf & vbCrLf & vbCrLf)

        If amountOfFiles = 1 Then
            msg.Append("The biggest file is: " & vbCrLf & vbCrLf)
        Else
            If amountOfFiles > 1 Then
                msg.Append("The biggest files are: " & vbCrLf & vbCrLf)
            End If
        End If

        For i = fileList.Count - 1 To fileList.Count - amountOfFiles Step -1
            msg.Append(fileList.Item(i).Key & " at " & fileList.Item(i).Value.ToString("N", CultureInfo.InvariantCulture) & " bytes." & vbCrLf & vbCrLf)
        Next

        MessageBox.Show(msg.ToString)

    End Sub

End Class
