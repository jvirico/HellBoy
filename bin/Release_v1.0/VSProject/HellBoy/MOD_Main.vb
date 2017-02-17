
Module MOD_Main

    Sub Main()

        Dim sConfig_URL As String

        Dim sWindowTitle As String
        Dim sButtonCaption As String
        Dim iSleepInSeconds As Integer
        Dim sTTLifeInSecs As String

        Try

            If Environment.GetCommandLineArgs.Length = 5 Then
                sWindowTitle = Environment.GetCommandLineArgs(1)
                sButtonCaption = Environment.GetCommandLineArgs(2)
                iSleepInSeconds = Environment.GetCommandLineArgs(3)
                sTTLifeInSecs = Environment.GetCommandLineArgs(4)
            Else
                sConfig_URL = System.AppDomain.CurrentDomain.BaseDirectory & "HellBoy.config"

                sWindowTitle = GetConfigItem(sConfig_URL, "WindowTitle")
                sButtonCaption = GetConfigItem(sConfig_URL, "ButtonCaption")
                iSleepInSeconds = GetConfigItem(sConfig_URL, "SleepSeconds")
                sTTLifeInSecs = GetConfigItem(sConfig_URL, "TTLifeInSecs")
            End If

            Console.WriteLine("HELLBOY MONITORING" & vbNewLine)
            Console.WriteLine("------------------" & vbNewLine)
            Console.WriteLine("Window Title:   " & sWindowTitle & vbNewLine)
            Console.WriteLine("Button Caption: " & sButtonCaption & vbNewLine)
            Console.WriteLine("Every:          " & iSleepInSeconds & " seconds" & vbNewLine)
            Console.WriteLine("During:         " & sTTLifeInSecs & " seconds" & vbNewLine)
            Console.WriteLine(vbNewLine)

            If sTTLifeInSecs = "" Or Not IsNumeric(sTTLifeInSecs) Then sTTLifeInSecs = 0

            HellBoy(sWindowTitle, sButtonCaption, iSleepInSeconds, CInt(sTTLifeInSecs))

        Catch ex As Exception

        End Try

    End Sub

    Private Function GetNow() As Double

        GetNow = CLng(DateTime.Now.Subtract(New DateTime(1970, 1, 1)).TotalMilliseconds)

    End Function

    Private Function GetDead(iTTL As Integer) As Double

        GetDead = CLng(DateTime.Now.Subtract(New DateTime(1970, 1, 1)).TotalMilliseconds) + iTTL * 1000

    End Function


    Public Sub HellBoy(sWindowTitle As String, sButtonCaption As String, iSleepInSeconds As Integer, iTTL As Integer)

        'Dim iTimeStampNow As Double
        Dim iTimeStampDead As Double
        Dim dDeadTime As Date

        If iTTL = 0 Then
            Do
                'If ClickWindowButton(sWindowTitle, sButtonCaption) Then Console.WriteLine(Now & vbNewLine)
                ClickWindowButton(sWindowTitle, sButtonCaption)
                Threading.Thread.Sleep(iSleepInSeconds * 1000)
            Loop
        Else
            'iTimeStampNow = GetNow()
            iTimeStampDead = GetDead(iTTL)

            Do While GetNow() < iTimeStampDead
                'If ClickWindowButton(sWindowTitle, sButtonCaption) Then Console.WriteLine(Now & vbNewLine)
                ClickWindowButton(sWindowTitle, sButtonCaption)
                Threading.Thread.Sleep(iSleepInSeconds * 1000)
            Loop
        End If

    End Sub


End Module
