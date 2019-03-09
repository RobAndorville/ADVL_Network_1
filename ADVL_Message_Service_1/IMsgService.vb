Imports System.ServiceModel
<ServiceContract(CallbackContract:=GetType(IMsgServiceCallback))>
Public Interface IMsgService
    <OperationContract()>
    Function Connect(ByVal appNetName As String, ByVal appName As String, ByVal connectionName As String, ByVal projectName As String, ByVal projectDescription As String, ByVal projectType As ADVL_Utilities_Library_1.Project.Types, ByVal projectPath As String, ByVal getAllWarnings As Boolean, ByVal getAllMessages As Boolean) As String
    'Function Connect(ByVal appName As String, ByVal connectionName As String, ByVal projectName As String, ByVal projectDescription As String, ByVal settingsLocnType As ADVL_Utilities_Library_1.FileLocation.Types, ByVal settingsLocnPath As String, ByVal appType As clsConnection.AppTypes, ByVal getAllWarnings As Boolean, ByVal getAllMessages As Boolean) As String

    <OperationContract()>
    Function ConnectionAvailable(ByVal AppNetName As String, ByVal ConnName As String) As Boolean

    <OperationContract()>
    Sub SendMessage(ByVal appNetName As String, ByVal connName As String, ByVal message As String)
    'Sub SendMessage(ByVal connName As String, ByVal message As String)

    <OperationContract()>
    Sub SendAllMessage(ByVal message As String, ByVal SenderName As String)

    <OperationContract()>
    Sub SendMainNodeMessage(ByVal message As String)

    <OperationContract()>
    Sub GetConnectionList()

    <OperationContract()>
    Sub GetApplicationList()

    <OperationContract()>
    Sub GetApplicationInfo(ByVal appName As String)

    <OperationContract()>
    Sub GetMessageServiceAppInfo()

    <OperationContract()>
    Function Disconnect(ByVal appNetName As String, ByVal connName As String) As Boolean
    'Function Disconnect(ByVal connName As String) As Boolean

    <OperationContract()>
    Function IsAlive() As Boolean

    <OperationContract()>
    Function AppNetNameUsed(ByVal AppNetName As String) As Boolean

End Interface

Public Interface IMsgServiceCallback

    <OperationContract(IsOneWay:=True)>
    Sub OnSendMessage(ByVal message As String)
End Interface
