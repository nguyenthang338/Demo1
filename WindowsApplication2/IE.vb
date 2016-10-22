Module Ie
    Sub OpenIe()      
        'Click link 
        Dim number As Integer
        Dim Rand As New Random()
        number = Rand.Next(1, 3)
        Dim count1 As Integer
        count1 = 1
        While (count1 <= number)
            count1 = count1 + 1
            Ie_adfly()
            IE_Sledgehammer()
        End While
        ' Open popup
        Dim count As Integer
        count = 1
        While (count <= 6)
            count = count + 1
            Ie()
            IE_Sledgehammer()
        End While
        IE_Sledgehammer()
    End Sub
    Sub Ie()
        Dim ie As Object
        Dim x As Object
        Dim arrayLink As String
        Dim random As New Random
        Try
            x = CreateObject("wscript.shell")
            ie = CreateObject("InternetExplorer.Application")
            arrayLink = "triparoundtheworldnow.blogspot.com"
            ie.Navigate(arrayLink)
            ie.Toolbar = 1
            ie.StatusBar = 1
            ie.Height = Screen.PrimaryScreen.Bounds.Height
            ie.Width = Screen.PrimaryScreen.Bounds.Width
            ie.Top = 0
            ie.Left = 0
            ie.Visible = 1

            Dim Rand As New Random()
            Dim time As Integer
            Dim a, b As Integer
            a = random.Next(100, 700)
            b = random.Next(10, 800)
            time = Rand.Next(17000, 22000)
            Threading.Thread.Sleep(time)

            'AuTo click Random 
            Dim number As Integer
            number = Rand.Next(1, 4)
            If (number = 1) Then
                ie.Document.elementFromPoint(a, b).Click()
                Threading.Thread.Sleep(Rand.Next(7000, 10000))
            End If
        Catch ex As Exception
        End Try
    End Sub
    Sub Ie_adfly()
        Dim ie As Object
        Dim x As Object

        Dim arrayLink As Array
        Dim random As New Random
        Try
            x = CreateObject("wscript.shell")
            ie = CreateObject("InternetExplorer.Application")
            Dim randomnumber As New Random()
            arrayLink = {"http://adf.ly/NuIHa", "http://adf.ly/NvXQZ", "http://adf.ly/NvXQY", "http://adf.ly/1a94HE", "http://adf.ly/1a94HK",
                         "http://q.gs/AASV0", "http://q.gs/AASUx", "http://q.gs/AASUy", "http://q.gs/AASUz", "http://q.gs/AASUw",
                         "http://adf.ly/NvXQW", "http://adf.ly/NvXQb", "http://adf.ly/NvXQe", "http://q.gs/AASV2", "http://q.gs/AASV3",
                         "http://adf.ly/NvXQd", "http://adf.ly/NvXQc", "http://adf.ly/NvXQS", "http://q.gs/AASV4", "http://q.gs/AASV1",
                         "http://adf.ly/NvXQU", "http://adf.ly/NuIHb", "http://adf.ly/NuIHd"}
            ie.Navigate(arrayLink(random.Next(0, arrayLink.Length) - 1))


            ie.Toolbar = 1
            ie.StatusBar = 1
            ie.Height = Screen.PrimaryScreen.Bounds.Height
            ie.Width = Screen.PrimaryScreen.Bounds.Width
            ie.Top = 0
            ie.Left = 0

            ie.Visible = 1

            Threading.Thread.Sleep(15000)
        Catch ex As Exception
        End Try
        Try
            ie.Document.getElementById("skip_ad_button").Click()
            Threading.Thread.Sleep(4000)
        Catch ex As Exception

        End Try

    End Sub
 
    Sub IE_Sledgehammer()
        Try
            Dim objWMI As Object, objProcess As Object, objProcesses As Object
            objWMI = GetObject("winmgmts://.")
            objProcesses = objWMI.ExecQuery( _
                "SELECT * FROM Win32_Process WHERE Name = 'iexplore.exe'")
            For Each objProcess In objProcesses
                Call objProcess.Terminate()
            Next
            objProcesses = Nothing : objWMI = Nothing
        Catch ex As Exception

        End Try

    End Sub
End Module
