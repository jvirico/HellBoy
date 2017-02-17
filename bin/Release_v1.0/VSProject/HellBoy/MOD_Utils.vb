Imports System
Imports System.IO

Module MOD_Utils
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sFileConfig_URL"></param>
    ''' <param name="sSettingName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetConfigItem(ByVal sFileConfig_URL As String, ByVal sSettingName As String) As String

        Dim S As New IO.StreamReader(sFileConfig_URL) : Dim Result As String = ""

        Try

            Do While (S.Peek <> -1)
                Dim Line As String = S.ReadLine
                If Line.ToLower.StartsWith(sSettingName.ToLower & ":") Then
                    If Line.Length > sSettingName.Length + 2 Then
                        Result = Line.Substring(sSettingName.Length + 2)
                        Exit Do
                    Else
                        Result = ""
                    End If

                End If
            Loop
            Return Result

        Catch ex As Exception
            Return ""
        End Try
    End Function
End Module
