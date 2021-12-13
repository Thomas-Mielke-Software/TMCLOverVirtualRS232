Module Module1

    Sub Main()
        Dim TMCM As TMCLOverVirtualRS232.SerialCOMPort
        TMCM = New TMCLOverVirtualRS232.SerialCOMPort
        Dim ports As String()
        ports = TMCM.AvailablePortNames
        For Each port In ports
            If port.Substring(0, 3) <> "COM" Then Continue For
            If port.Substring(3, 1) < "1" Or port.Substring(3, 1) > "9" Then Continue For
            Console.WriteLine("TMCM module found on " + port)
        Next

        If ports.Length > 0 Then
            If TMCM.SetSelectedPort(0) Then  ' just open the first port that was found
                TMCM.OpenPort()
            End If
        End If

        ' TMCM.

        Console.WriteLine("")
        If TMCM.SendTMCMCommand(9, 1, 1, 2, 1) Then
            Console.WriteLine("SendTMCMCommand returned true")
        Else
            Console.WriteLine("SendTMCMCommand returned false")
        End If
        Console.WriteLine("Press Enter to end program.")
        Console.ReadLine()

    End Sub

End Module
