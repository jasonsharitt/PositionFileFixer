Imports System
Imports System.IO
Imports System.Text

Public Class Form1
    Dim fd As New OpenFileDialog()
    Dim sd As New FolderBrowserDialog()
    Dim posfilename As String = ""
    Dim savefilefolder As String = ""
    Dim filename As String = ""
    Dim filereader As StreamReader
    Dim filemaker As StreamWriter
    Dim filemaker2 As StreamWriter
    Dim filemaker3 As StreamWriter
    Dim filemaker4 As StreamWriter
    Dim filemaker5 As StreamWriter
    Dim filemaker6 As StreamWriter
    Dim line As String
    Dim savefilename As String
    Dim savefilename2 As String
    Dim savefilename3 As String
    Dim savefilename4 As String
    Dim savefilename5 As String
    Dim savefilename6 As String
    Dim linesplit() As String
    Dim newlinenewdate As String
    Dim totalrowcount As Integer = 1
    Dim initialdate As String = ""
    Dim datedictionary As New Dictionary(Of String, List(Of String))
    Dim datenoquotes As String
    Dim hour As Integer
    Dim filelist As New List(Of String)
    Dim badfilelist As New List(Of String)
  
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'For running contigual track reports, 4 hour file length is highly recommended. 
        'A message box will appear to remind the user.
        If RadioButton1.Checked = False And RadioButton2.Checked = False Then
            MsgBox("You must choose a length of time")
        Else
            fd.Title = "Choose Position Files"
            fd.InitialDirectory = "I:\"
            fd.Filter = "Position Files (*.csv)|*.csv|Position Files (*.csv)|*.csv"
            fd.FilterIndex = 2
            fd.RestoreDirectory = True
            fd.Multiselect = True
            If fd.ShowDialog() = Windows.Forms.DialogResult.OK Then
                posfilename = fd.FileName
            End If
            With sd
                .Description = "Choose folder to save position files"
                .RootFolder = Environment.SpecialFolder.Desktop
                .SelectedPath = "I:\"
                .ShowNewFolderButton = True
                If sd.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    savefilefolder = sd.SelectedPath
                End If
            End With
            If posfilename <> "" Then
                '****************Create Progress Bar*********************************
                'This progress bar does not calculate correctly, need to fix
                ProgressBar1.Minimum = 0
                ProgressBar1.Maximum = fd.FileNames.Count * 2
                '**********************************************************************


                '************Create a dictionary that lists all the dates the position files contain**************

                For Each Me.posfilename In fd.FileNames
                    filereader = My.Computer.FileSystem.OpenTextFileReader(Me.posfilename)
                    filename = Path.GetFileName(Me.posfilename)
                    Do Until filereader.EndOfStream()
                        line = filereader.ReadLine()
                        'There are many variables possible here that make a line ineligible for date extraction
                        'I've added a few conditions to cover the errors I've come across.
                        If line <> "" Then
                            If line.Contains(",") And line.Length > 50 Then
                                linesplit = line.Split(New [Char]() {","})
                                If linesplit.Count = 8 Then
                                    If IsNumeric(line.Substring(0, 2)) And linesplit(2).Length >= 19 Then
                                        If linesplit(2).Length > 19 Then
                                            initialdate = linesplit(2).Substring(1, 10)
                                        ElseIf linesplit(2).Length = 19 Then
                                            initialdate = linesplit(2).Substring(0, 10)
                                        End If
                                        If Not datedictionary.ContainsKey(initialdate) Then
                                            datedictionary.Add(initialdate, New List(Of String)(New String() {posfilename}))
                                        End If
                                        Dim dictkeys = datedictionary(initialdate)
                                        If Not dictkeys.Contains(posfilename) Then
                                            dictkeys.Add(posfilename)
                                        End If
                                    Else
                                        If Not badfilelist.Contains(Me.posfilename) Then
                                            badfilelist.Add(Me.posfilename)
                                        End If
                                    End If
                                Else
                                    If Not badfilelist.Contains(Me.posfilename) Then
                                        badfilelist.Add(Me.posfilename)
                                    End If
                                End If
                            Else
                                If Not badfilelist.Contains(Me.posfilename) Then
                                    badfilelist.Add(Me.posfilename)
                                End If
                            End If

                        End If

                    Loop
                    filereader.Close()
                    ProgressBar1.Increment(1)
                Next
                If badfilelist.Count > 1 Then
                    Dim badfiles As New StringBuilder()
                    For Each item In badfilelist
                        badfiles.AppendLine(item)
                    Next
                    MsgBox(badfilelist.ToString(), , "These files contained errors")
                End If

                ' **********************************************************************************
                If RadioButton1.Checked = True Then
                    For Each strKey In datedictionary.Keys()
                        savefilename = savefilefolder + "\" + strKey + "FullDay.csv"
                        filemaker = New StreamWriter(savefilename, False)
                        Dim keylist = datedictionary(strKey)
                        For Each strfilename In keylist
                            filereader = My.Computer.FileSystem.OpenTextFileReader(strfilename)
                            Do Until filereader.EndOfStream()
                                line = filereader.ReadLine()
                                If line <> "" Then
                                    If line.Length > 20 Then
                                        linesplit = line.Split(New [Char]() {","})
                                        If linesplit(2).Length > 19 Then
                                            datenoquotes = linesplit(2).ToString.Substring(1, 19)
                                        Else
                                            datenoquotes = linesplit(2)
                                        End If
                                        If strKey = datenoquotes.Substring(0, 10) Then
                                            newlinenewdate = linesplit(0) + "," + linesplit(1) + "," + datenoquotes + "," + linesplit(3) + "," + linesplit(4) + "," + linesplit(5) + "," + linesplit(6) + "," + linesplit(7)
                                            filemaker.WriteLine(newlinenewdate)
                                        End If
                                    End If
                                End If
                            Loop
                            filereader.Close()
                        Next
                        ProgressBar1.Increment(1)
                        filemaker.Close()
                        filereader.Close()
                        filelist.Add(savefilename)
                    Next
                    filereader.Close()
                    filemaker.Close()
                End If
                If RadioButton2.Checked = True Then
                    For Each strKey In datedictionary.Keys()
                        savefilename = savefilefolder + "\" + strKey + " 0AM to 4AM.csv"
                        savefilename2 = savefilefolder + "\" + strKey + " 4AM to 8AM.csv"
                        savefilename3 = savefilefolder + "\" + strKey + " 8AM to 12PM.csv"
                        savefilename4 = savefilefolder + "\" + strKey + " 12PM to 16PM.csv"
                        savefilename5 = savefilefolder + "\" + strKey + " 16PM to 20PM.csv"
                        savefilename6 = savefilefolder + "\" + strKey + " 20PM to 24PM.csv"
                        filemaker = New StreamWriter(savefilename, False)
                        filemaker2 = New StreamWriter(savefilename2, False)
                        filemaker3 = New StreamWriter(savefilename3, False)
                        filemaker4 = New StreamWriter(savefilename4, False)
                        filemaker5 = New StreamWriter(savefilename5, False)
                        filemaker6 = New StreamWriter(savefilename6, False)
                        Dim keylist = datedictionary(strKey)
                        For Each strfilename In keylist
                            filereader = My.Computer.FileSystem.OpenTextFileReader(strfilename)
                            Do Until filereader.EndOfStream()
                                line = filereader.ReadLine()
                                If line <> "" Then
                                    If IsNumeric(line.Substring(0, 2)) Then
                                        If line.Length > 20 Then
                                            linesplit = line.Split(New [Char]() {","})
                                            If linesplit.Count = 8 Then
                                                If linesplit(2).Length > 19 Then
                                                    datenoquotes = linesplit(2).ToString.Substring(1, 19)
                                                Else
                                                    datenoquotes = linesplit(2)
                                                End If
                                                hour = CInt(datenoquotes.Substring(11, 2))
                                                If strKey = datenoquotes.Substring(0, 10) Then
                                                    If hour >= 0 AndAlso hour < 4 Then
                                                        newlinenewdate = linesplit(0) + "," + linesplit(1) + "," + datenoquotes + "," + linesplit(3) + "," + linesplit(4) + "," + linesplit(5) + "," + linesplit(6) + "," + linesplit(7)
                                                        filemaker.WriteLine(newlinenewdate)
                                                    ElseIf hour >= 4 AndAlso hour < 8 Then
                                                        newlinenewdate = linesplit(0) + "," + linesplit(1) + "," + datenoquotes + "," + linesplit(3) + "," + linesplit(4) + "," + linesplit(5) + "," + linesplit(6) + "," + linesplit(7)
                                                        filemaker2.WriteLine(newlinenewdate)
                                                    ElseIf hour >= 8 AndAlso hour < 12 Then
                                                        newlinenewdate = linesplit(0) + "," + linesplit(1) + "," + datenoquotes + "," + linesplit(3) + "," + linesplit(4) + "," + linesplit(5) + "," + linesplit(6) + "," + linesplit(7)
                                                        filemaker3.WriteLine(newlinenewdate)
                                                    ElseIf hour >= 12 AndAlso hour < 16 Then
                                                        newlinenewdate = linesplit(0) + "," + linesplit(1) + "," + datenoquotes + "," + linesplit(3) + "," + linesplit(4) + "," + linesplit(5) + "," + linesplit(6) + "," + linesplit(7)
                                                        filemaker4.WriteLine(newlinenewdate)
                                                    ElseIf hour >= 16 AndAlso hour < 20 Then
                                                        newlinenewdate = linesplit(0) + "," + linesplit(1) + "," + datenoquotes + "," + linesplit(3) + "," + linesplit(4) + "," + linesplit(5) + "," + linesplit(6) + "," + linesplit(7)
                                                        filemaker5.WriteLine(newlinenewdate)
                                                    ElseIf hour >= 20 AndAlso hour < 24 Then
                                                        newlinenewdate = linesplit(0) + "," + linesplit(1) + "," + datenoquotes + "," + linesplit(3) + "," + linesplit(4) + "," + linesplit(5) + "," + linesplit(6) + "," + linesplit(7)
                                                        filemaker6.WriteLine(newlinenewdate)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Loop
                            filereader.Close()
                        Next
                        ProgressBar1.Increment(1)
                        filemaker.Close()
                        filemaker2.Close()
                        filemaker3.Close()
                        filemaker4.Close()
                        filemaker5.Close()
                        filemaker6.Close()
                        filereader.Close()
                        filelist.Add(savefilename)
                        filelist.Add(savefilename2)
                        filelist.Add(savefilename3)
                        filelist.Add(savefilename4)
                        filelist.Add(savefilename5)
                        filelist.Add(savefilename6)
                    Next
                    filereader.Close()
                    filemaker.Close()
                    filemaker2.Close()
                    filemaker3.Close()
                    filemaker4.Close()
                    filemaker5.Close()
                    filemaker6.Close()
                End If
            End If
        End If
        'This ForEach works, I don't know why it won't work in the Report Generator application though.
        'There is a note about it in that application.
        For Each g In filelist
            Dim h As New FileInfo(g)
            Dim glength As Long = h.Length
            If glength < 1 Then
                My.Computer.FileSystem.DeleteFile(g)
            End If
        Next
        My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Exclamation)
        MsgBox("Reports Complete")
    End Sub
End Class
