Private Sub Form_Load()
' Configurazione iniziale della porta seriale
MSComm1.CommPort = 1 ' Imposta la porta COM1
MSComm1.Settings = "9600,N,8,1" ' Baud rate 9600, no parità, 8 bit, 1 stop bit
MSComm1.PortOpen = True ' Apre la porta seriale
End Sub

Private Sub Command1_Click() ' Pulsante per accendere dispositivo 1
MSComm1.Output = "1" ' Invia comando per attivare Q1
Label1.Caption = "Dispositivo 1 Acceso"
End Sub

Private Sub Command2_Click() ' Pulsante per spegnere dispositivo 1
MSComm1.Output = "2" ' Invia comando per disattivare Q1
Label1.Caption = "Dispositivo 1 Spento"
End Sub

Private Sub Command3_Click() ' Pulsante per accendere dispositivo 2
MSComm1.Output = "3" ' Invia comando per attivare Q2
Label2.Caption = "Dispositivo 2 Acceso"
End Sub

Private Sub Command4_Click() ' Pulsante per spegnere dispositivo 2
MSComm1.Output = "4" ' Invia comando per disattivare Q2
Label2.Caption = "Dispositivo 2 Spento"
End Sub

Private Sub Form_Unload(Cancel As Integer)
MSComm1.PortOpen = False ' Chiude la porta seriale all'uscita
End Sub
