Module MOD_WindowsHandler

    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
(ByVal lpClassName As String, ByVal lpWindowName As String) As Integer

    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
    (ByVal hWnd1 As Integer, ByVal hWnd2 As Integer, ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Integer

    Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hwnd As Integer, ByVal lpString As String, ByVal cch As Integer) As Integer

    Private Declare Function GetWindowTextLength Lib "user32" Alias _
    "GetWindowTextLengthA" (ByVal hwnd As Integer) As Integer

    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer

    Const WM_SETTEXT = &HC
    Const BM_CLICK = &HF5



    Public Function ClickWindowButton(sWinTitle As String, sButtonText As String) As Boolean

        Dim Ret As Integer, ChildRet As Integer, OpenRet As Integer
        Dim strBuff As String, ButCap As String
        Dim retorno As Boolean

        'Threading.Thread.Sleep(1500) '<~~ We need this so that there ample time for the VBA window to launch

        '~~> Get the handle of the "VBAProject Password" Window
        Ret = FindWindow(vbNullString, sWinTitle)

        If Ret <> 0 Then
            'MsgBox "VBAProject Password Window Found"

            ''~~> Get the handle of the TextBox Window where we need to type the password
            ChildRet = FindWindowEx(Ret, 0, "Button", vbNullString)

            'If ChildRet <> 0 Then
            'MsgBox "TextBox's Window Found"
            '~~> This is where we send the password to the Text Window
            'SendMess(MyPassword, ChildRet)

            '~~> Get the handle of the Button's "Window"
            'ChildRet = FindWindowEx(Ret, 0, "Button", vbNullString)

            '~~> Check if we found it or not
            If ChildRet <> 0 Then
                'MsgBox "Button's Window Found"

                '~~> Get the caption of the child window
                strBuff = New String(Chr(0), GetWindowTextLength(ChildRet) + 1)
                GetWindowText(ChildRet, strBuff, Len(strBuff))
                ButCap = strBuff

                '~~> Loop through all child windows
                Do While ChildRet <> 0
                    '~~> Check if the caption has the word "OK"
                    If InStr(1, ButCap, sButtonText) Then
                        '~~> If this is the button we are looking for then exit
                        OpenRet = ChildRet
                        Exit Do
                    End If

                    '~~> Get the handle of the next child window
                    ChildRet = FindWindowEx(Ret, ChildRet, "Button", vbNullString)
                    '~~> Get the caption of the child window
                    strBuff = New String(Chr(0), GetWindowTextLength(ChildRet) + 1)
                    GetWindowText(ChildRet, strBuff, Len(strBuff))
                    ButCap = strBuff
                Loop

                '~~> Check if we found it or not
                If OpenRet <> 0 Then
                    '~~> Click the OK Button
                    SendMessage(ChildRet, BM_CLICK, 0, vbNullString)

                Else
                    'MsgBox("The Handle of OK Button was not found")
                End If
            Else
                'MsgBox("Button's Window Not Found")
            End If
            'Else
            '    MsgBox("The Edit Box was not found")
            'End If
            'Else
            'MsgBox("VBAProject Password Window was not Found")
        End If

    End Function


    Sub SendMess(ByVal Message As String, ByVal hwnd As Integer)
        Call SendMessage(hwnd, WM_SETTEXT, False, Message)
    End Sub



End Module
