Private Sub Form_Load()
' Configura la porta seriale
MSComm1.CommPort = 1 ' Porta COM1
MSComm1.Settings = "9600,N,8,1" ' 9600 baud rate
MSComm1.PortOpen = True ' Apre la porta seriale
End Sub

Private Sub Timer1_Timer()
' Legge lo stato del sensore dal Nutchip
Dim inputData As String
If MSComm1.InBufferCount > 0 Then
inputData = MSComm1.Input ' Legge i dati ricevuti
If inputData = "SENSOR_ON" Then
Label1.Caption = "Sensore Attivato"
' Puoi aggiungere logica per attivare un dispositivo
ElseIf inputData = "SENSOR_OFF" Then
Label1.Caption = "Sensore Disattivato"
End If
End If
End Sub

Private Sub Command1_Click()
' Pulsante per accendere manualmente il dispositivo
MSComm1.Output = "1" ' Invia comando per attivare Q1
Label2.Caption = "Dispositivo Acceso"
End Sub

Private Sub Command2_Click()
' Pulsante per spegnere manualmente il dispositivo
MSComm1.Output = "2" ' Invia comando per disattivare Q1
Label2.Caption = "Dispositivo Spento"
End Sub

Private Sub Form_Unload(Cancel As Integer)
MSComm1.PortOpen = False ' Chiude la porta seriale all'uscita
End Sub
