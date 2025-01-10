Private Sub Form_Load()
' Configura la porta seriale
MSComm1.CommPort = 1 ' Usa la porta COM1
MSComm1.Settings = "9600,N,8,1"
MSComm1.PortOpen = True ' Apre la comunicazione seriale
End Sub

Private Sub Timer1_Timer()
' Legge la temperatura dal Nutchip
Dim inputData As String
If MSComm1.InBufferCount > 0 Then
inputData = MSComm1.Input ' Dati ricevuti
Label1.Caption = "Temperatura: " & inputData & " °C"

' Esempio: Attiva un dispositivo se temperatura > 30°C
If Val(inputData) > 30 Then
Label2.Caption = "Attivazione dispositivo!"
Else
Label2.Caption = "Dispositivo disattivato"
End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
MSComm1.PortOpen = False ' Chiude la porta seriale all'uscita
End Sub
