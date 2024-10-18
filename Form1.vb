Imports System.IO
Imports System.Xml

Public Class Form1

    Dim SS As New Collection
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Icon = My.Resources.Windows

        SS.Add(Label_CPU_SS)
        SS.Add(Label_RAM_SS)
        SS.Add(Label_GRAPH_SS)
        SS.Add(Label_GAME_SS)
        SS.Add(Label_DISK_SS)

        XML()

    End Sub

    Private Sub Form1_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        TableLayoutPanel1.Width = Me.Width - 40
    End Sub

    Sub XML()
        Dim path = Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\Performance\WinSAT\DataStore\"
        Dim highDir = ""
        Dim lastHigh = New DateTime(1900, 1, 1)
        For Each subdir In Directory.GetFiles(path, "*.xml")
            Dim fi = New FileInfo(subdir)
            Dim created = fi.LastWriteTime
            If subdir.IndexOf("Formal.Assessment") <> -1 Then
                If created > lastHigh Then
                    highDir = subdir
                    lastHigh = created
                End If
            End If
        Next
        Dim xml = New XmlDocument()
        If Not String.IsNullOrWhiteSpace(highDir) Then
            xml.Load(highDir)
            Dim xnList = xml.SelectNodes("/WinSAT/WinSPR")
            Dim Score(5) As Decimal
            Dim MiniScore As Decimal = 10
            For Each xnNode In xnList
                If Decimal.TryParse(xnNode("CpuScore").InnerText, Score(1)) Then
                    If Score(1) < MiniScore Then
                        MiniScore = Score(1)
                    End If
                End If
                If Decimal.TryParse(xnNode("MemoryScore").InnerText, Score(2)) Then
                    If Score(2) < MiniScore Then
                        MiniScore = Score(2)
                    End If
                End If
                If Decimal.TryParse(xnNode("GraphicsScore").InnerText, Score(3)) Then
                    If Score(3) < MiniScore Then
                        MiniScore = Score(3)
                    End If
                End If
                If Decimal.TryParse(xnNode("GamingScore").InnerText, Score(4)) Then
                    If Score(4) < MiniScore Then
                        MiniScore = Score(4)
                    End If
                End If
                If Decimal.TryParse(xnNode("DiskScore").InnerText, Score(5)) Then
                    If Score(5) < MiniScore Then
                        MiniScore = Score(5)
                    End If
                End If
            Next
            For i = 1 To 5
                SS(i).Text = Score(i)
                If Score(i) = MiniScore Then
                    SS(i).BackColor = Color.Gainsboro
                Else
                    SS(i).BackColor = Color.WhiteSmoke
                End If
            Next
            Label_Score.Text = (Math.Floor(MiniScore * 10) / 10D).ToString("0.0")
            xnList = xml.SelectNodes("/WinSAT/SystemConfig")
            For Each xnNode In xnList
                Label_CPU.Text = xnNode.SelectSingleNode("Processor/Instance/ProcessorName").InnerText
                Dim ramType As String = xnNode.SelectSingleNode("Memory/DIMM/PartNumber").InnerText & " "
                Select Case xnNode.SelectSingleNode("Memory/DIMM/MemoryType").InnerText
                    Case "22"
                        ramType &= "DDR2"
                    Case "24"
                        ramType &= "DDR3"
                    Case "26"
                        ramType &= "DDR4"
                    Case "28"
                        ramType &= "DDR5"
                End Select
                Label_RAM.Text = ramType & " " & xnNode.SelectSingleNode("Memory/TotalPhysical/Size").InnerText
                Label_GRAPH.Text = xnNode.SelectSingleNode("Graphics/AdapterDescription").InnerText
                Label_GAME.Text = ""
                Label_DISK.Text = xnNode.SelectSingleNode("Disk/SystemDisk/Vendor").InnerText & xnNode.SelectSingleNode("Disk/SystemDisk/Model").InnerText & " " & (CLng(xnNode.SelectSingleNode("Disk/SystemDisk/Size").InnerText) / (1024 ^ 3)).ToString("0.00") & "GB"
                Label_WinVer.Text = xnNode.SelectSingleNode("OSVersion/OSName").InnerText & " (Build: " & xnNode.SelectSingleNode("OSVersion/BuildLab").InnerText & ")"
                Label_time.Text = "Assessment Time: " & lastHigh.ToString()
            Next
        End If
    End Sub

    Private Async Sub startButton_Click(sender As Object, e As EventArgs) Handles startButton.Click
        For i = 1 To 5
            SS(i).Text = 0
            SS(i).BackColor = Color.WhiteSmoke
        Next
        Label_Score.Text = (Math.Floor(0 * 10) / 10D).ToString("0.0")
        startButton.Enabled = False
        Await asyncTest()
        XML()
        startButton.Enabled = True
    End Sub

    Private Function asyncTest() As Task
        Return Task.Factory.StartNew(Sub()
                                         startTest()
                                     End Sub)
    End Function

    Sub startTest()
        Dim Process = New Process()
        Process.StartInfo.FileName = "WinSAT.exe"
        Process.StartInfo.Arguments = "formal"
        Process.StartInfo.WorkingDirectory = "C:\"
        Process.StartInfo.UseShellExecute = True
        Process.StartInfo.Verb = "runas"
        Try
            Process.Start()
            Process.WaitForExit()
        Catch ex As Exception
            ' Handle the exception (e.g., user canceled the UAC prompt)
            MessageBox.Show("The operation was canceled or failed: " & ex.Message & Process.StartInfo.FileName)
        End Try
    End Sub

End Class
