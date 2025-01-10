Private Sub Command1_Click() ' Pulsante per attivare l'antifurto
MSComm1.CommPort = 1 ' Porta COM1
MSComm1.Settings = "9600,N,8,1"
MSComm1.PortOpen = True

MSComm1.Output = "A" ' Comando per attivare l'antifurto
Label1.Caption = "Antifurto Attivato"

MSComm1.PortOpen = False
End Sub

Private Sub Command2_Click() ' Pulsante per disattivare l'antifurto
MSComm1.CommPort = 1
MSComm1.Settings = "9600,N,8,1"
MSComm1.PortOpen = True

MSComm1.Output = "D" ' Comando per disattivare l'antifurto
Label1.Caption = "Antifurto Disattivato"

MSComm1.PortOpen = False
End Sub

Private Sub MSComm1_OnComm()
Dim inputData As String
inputData = MSComm1.Input ' Riceve dati dal Nutchip

If inputData = "ALARM" Then
MsgBox "Allarme! Sensore Attivato!"
End If
End Sub
