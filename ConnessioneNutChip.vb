Private Sub Command1_Click()
' Configura la porta seriale
MSComm1.CommPort = 1 ' Imposta la porta COM1
MSComm1.Settings = "9600,N,8,1" ' Baud rate 9600, no parità, 8 bit, 1 stop bit
MSComm1.PortOpen = True ' Apre la porta seriale

' Invia il comando per accendere il LED
Dim comando As String
comando = "A" ' Comando che corrisponde al pin I1 del Nutchip
MSComm1.Output = comando ' Invia il comando

' Chiudi la porta seriale
MSComm1.PortOpen = False
End Sub
