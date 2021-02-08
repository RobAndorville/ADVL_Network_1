Imports System.ServiceModel
Public Class MsgServiceCallBack
    'Implements ServiceReference1.IMsgServiceCallback
    Implements IMsgServiceCallback
    'Public Sub OnSendMessage(message As String) Implements ServiceReference1.IMsgServiceCallback.OnSendMessage
    Public Sub OnSendMessage(message As String) Implements IMsgServiceCallback.OnSendMessage
        'A message has been received.
        'Set the InstrReceived property value to the XMessage. This will also apply the instructions in the XMessage.
        Try
            Main.InstrReceived = message
        Catch ex As Exception
            Main.Message.AddWarning("Callback message error: " & ex.Message & vbCrLf)
        End Try

    End Sub

    'Public Function OnSendMessageCheck() As String Implements IMsgServiceCallback.OnSendMessageCheck
    '    'If IsNothing(Main) Then
    '    '    Return "Failed"
    '    'Else
    '    '    Return "OK"
    '    'End If
    '    Return "OK"
    'End Function
End Class
