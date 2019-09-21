Imports System.ServiceModel
Imports ADVL_Message_Service_1.ServiceReference1

Public Class HostAppMsgServiceCallback
    Implements ServiceReference1.IMsgServiceCallback

    Public Sub OnSendMessage(message As String) Implements ServiceReference1.IMsgServiceCallback.OnSendMessage
        'Throw New NotImplementedException()
        'A message has been received.
        'Set the InstrReceived property value to the XMessage. This will also apply the instructions in the XMessage.
        Main.InstrReceived = message
    End Sub

    Public Function OnSendMessageCheck() As String Implements ServiceReference1.IMsgServiceCallback.OnSendMessageCheck
        Return "OK"
    End Function
End Class
