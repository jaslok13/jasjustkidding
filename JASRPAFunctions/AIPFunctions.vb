Public Class AIPFunctions
    Const AIPLabelGUIDUnrestricted As String = ""
    Const AIPLabelGUIDInternal As String = ""
    Const AIPLabelGUIDConfidential As String = ""
    Const strSpace As String = " "
    Public Shared Function GetAIPLabel(strFilePath As String) As String
        Dim strAIPLabel As String = String.Empty
        Dim strArg As String = "-command" & strSpace & Microsoft.VisualBasic.Strings.Chr(34) & "get-aipfilestatus" & strSpace & "-path" & strSpace & "'" & strFilePath & "'|" & strSpace & "foreach" & strSpace & "{$_.MainLabelName}" & Microsoft.VisualBasic.Strings.Chr(34)
        If System.IO.File.Exists(strFilePath) Then
            Dim psi As System.Diagnostics.ProcessStartInfo = New System.Diagnostics.ProcessStartInfo(fileName:="powershell", arguments:=strArg)
            psi.UseShellExecute = False
            psi.CreateNoWindow = True
            psi.RedirectStandardError = True
            psi.RedirectStandardOutput = True

            Dim p As New System.Diagnostics.Process
            p.StartInfo = psi
            p.Start()
            Dim output = p.StandardOutput.ReadToEndAsync
            p.WaitForExit()
            Dim intWaitedSoFar = 0
            While (Not output.IsCompleted) And intWaitedSoFar <= 60
                System.Threading.Thread.Sleep(1000)
                intWaitedSoFar = intWaitedSoFar + 1
            End While

            If output.IsCompleted Then
                strAIPLabel = output.Result
            Else
                Throw New System.Exception("Timed out waiting for powershell command to complete")
            End If
        Else
            Throw New System.Exception("File not found: " & strFilePath)
        End If
        Return strAIPLabel
    End Function
    Public Shared Sub SetAIPLabel(strFilePath As String, strAIPLabel As String)
        If System.IO.File.Exists(strFilePath) Then
            If Not String.IsNullOrEmpty(strAIPLabel) Then
                Dim strAIPLabelGUID As String = String.Empty
                If strAIPLabel.Trim.ToUpper.Equals("UNRESTRICTED") Then
                    strAIPLabelGUID = AIPLabelGUIDUnrestricted
                ElseIf strAIPLabel.Trim.ToUpper.Equals("CONFIDENTIAL") Then
                    strAIPLabelGUID = AIPLabelGUIDConfidential
                ElseIf strAIPLabel.Trim.ToUpper.Equals("INTERNAL") Then
                    strAIPLabelGUID = AIPLabelGUIDInternal
                Else
                    Throw New System.Exception("Invalid AIP Label: " & strAIPLabel)
                End If

                Dim strPSCommand As String = "set-aipfilelabel"
                Dim strArg As String = "-command" & strSpace & Microsoft.VisualBasic.Strings.Chr(34) & strPSCommand & strSpace & "-path" & strSpace & "'" & strFilePath & "'" & strSpace & "-labelid" & strSpace & "'" & strAIPLabelGUID & "'|format-list" & strSpace & Microsoft.VisualBasic.Strings.Chr(34)
                Dim psi As System.Diagnostics.ProcessStartInfo = New System.Diagnostics.ProcessStartInfo(fileName:="powershell", arguments:=strArg)
                psi.UseShellExecute = False
                psi.CreateNoWindow = True
                psi.RedirectStandardError = True
                psi.RedirectStandardOutput = True

                Dim p As New System.Diagnostics.Process
                p.StartInfo = psi
                p.Start()
                Dim output = p.StandardOutput.ReadToEndAsync
                p.WaitForExit()
                Dim intWaitedSoFar = 0
                While (Not output.IsCompleted) And intWaitedSoFar <= 60
                    System.Threading.Thread.Sleep(1000)
                    intWaitedSoFar = intWaitedSoFar + 1
                End While

                If output.IsCompleted Then
                    Dim strStatus As String = ""
                    Dim strComment As String = ""
                    For Each strOutputLine As String In output.Result.Split(vbCrLf)
                        strOutputLine = strOutputLine.Trim
                        If strOutputLine.StartsWith("Status") Then
                            strStatus = strOutputLine.Substring(startIndex:=strOutputLine.IndexOf(":") + 1).Trim
                        ElseIf strOutputLine.StartsWith("Comment") Then
                            strComment = strOutputLine.Substring(startIndex:=strOutputLine.IndexOf(":") + 1).Trim
                        ElseIf Not String.IsNullOrEmpty(strOutputLine) Then
                            strComment = strComment & strSpace & strOutputLine
                        End If
                    Next
                    If strStatus.ToUpper.Trim.Equals("SUCCESS") Then
                        'SetAIPLabel successful, do nothing
                    Else
                        Throw New System.Exception(strComment)
                    End If
                Else
                    Throw New System.Exception("Timed out waiting for powershell command to complete")
                End If
            Else
                Throw New System.Exception("Blank value passed for AIP Label")
            End If
        Else
            Throw New System.Exception("File not found: " & strFilePath)
        End If
    End Sub
End Class
