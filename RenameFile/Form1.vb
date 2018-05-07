Imports System.IO

Public Class Form1
    Dim sSourceFolderpath As String
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim FolderBrowserDialog As New FolderBrowserDialog()
        FolderBrowserDialog.RootFolder = Environment.SpecialFolder.Desktop
        FolderBrowserDialog.SelectedPath = Environment.SpecialFolder.Desktop
        FolderBrowserDialog.Description = "Select Folder"
        If FolderBrowserDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            sSourceFolderpath = FolderBrowserDialog.SelectedPath
            ComboBox1.Enabled = True
            If ComboBox1.Items.Count = 0 Then
                ComboBox1.Items.Add("Last Modified Date")
                ComboBox1.Items.Add("Created Date")
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Dim nCount As Integer = 0
        Dim sFile As String() = System.IO.Directory.GetFiles(sSourceFolderpath)

        Using writer As StreamWriter = File.AppendText(sSourceFolderpath & "\Logging.txt")
            writer.WriteLine("---------------------------")
            writer.Flush()
            writer.Close()
        End Using

        For Each zfile In sFile

            Dim sLastTimetmp As Date
            If ComboBox1.SelectedItem = "Last Modified Date" Then
                sLastTimetmp = My.Computer.FileSystem.GetFileInfo(zfile).LastWriteTime
            Else
                sLastTimetmp = My.Computer.FileSystem.GetFileInfo(zfile).CreationTime
            End If

            Dim sLastTime As String = CustomDate(sLastTimetmp)
            Dim sExtension As String = My.Computer.FileSystem.GetFileInfo(zfile).Extension

            Try

    'Do Nishant Pandey
                My.Computer.FileSystem.RenameFile(zfile, sLastTime & sExtension)
                'Loop While (Not File.Exists("My.Computer.FileSystem.GetFileInfo(zfile).Directory.FullName & " \ " & sLastTime & sExtension"))

                Using writer As StreamWriter = File.AppendText(My.Computer.FileSystem.GetFileInfo(zfile).DirectoryName & "\Logging.txt")
                    writer.WriteLine(nCount + 1 & "." & "Old File Name ====> " & My.Computer.FileSystem.GetFileInfo(zfile).FullName)
                    writer.WriteLine(nCount + 1 & "." & "New File Name ====> " & sLastTime & sExtension)
                    writer.Flush()
                    writer.Close()
                End Using

                nCount = nCount + 1

            Catch ex As Exception
                MessageBox.Show(ex.Message)
                Using writer As StreamWriter = File.AppendText(My.Computer.FileSystem.GetFileInfo(zfile).DirectoryName & "\Logging.txt")
                    writer.WriteLine(nCount + 1 & "." & "Error in the file Name ====> " & My.Computer.FileSystem.GetFileInfo(zfile).FullName)
                    writer.Flush()
                    writer.Close()
                End Using

                nCount = nCount + 1
            End Try

        Next

        ComboBox1.Enabled = False

    End Sub

    Private Function CustomDate(ByVal sLastTimetmp As Date) As String
        Dim sLastTime As String

        Dim sYear As String = sLastTimetmp.Year
        Dim sMonth As String = sLastTimetmp.Month
        If sMonth.Length = 2 Then
        Else
            sMonth = "0" & sMonth
        End If
        Dim sDay As String = sLastTimetmp.Day
        If sDay.Length = 2 Then
        Else
            sDay = "0" & sDay
        End If
        Dim sHr As String = sLastTimetmp.Hour
        If sHr.Length = 2 Then
        Else
            sHr = "0" & sHr
        End If
        Dim sMin As String = sLastTimetmp.Minute
        If sMin.Length = 2 Then
        Else
            sMin = "0" & sMin
        End If
        Dim sSec As String = sLastTimetmp.Second
        If sSec.Length = 2 Then
        Else
            sSec = "0" & sSec
        End If

        sLastTime = sYear & sMonth & sDay & sHr & sMin & sSec

        Return sLastTime
    End Function

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Button2.Enabled = True
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Button2.Enabled = False
        ComboBox1.Enabled = False
    End Sub
End Class
