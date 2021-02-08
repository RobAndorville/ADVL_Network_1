
Imports System.ServiceModel
<CallbackBehavior(ConcurrencyMode:=ConcurrencyMode.Multiple, UseSynchronizationContext:=False)>
Public Class MsgService
    Implements IMsgService

    Private Shared ReadOnly connections As New List(Of clsConnection)()
    'connections is a list of connection information.
    'Each item contains the application name and the callback used to send a message to the application, a GetAllWarningsflag and a GetAllMessages flag.

    'Private Shared ReadOnly adminConnections As New List(Of clsConnection)()
    'adminConnections is a list of Admin Connection information.
    'Each item contains the application name, the callback used to send a message to the application, a GetWarnings flag and a GetAllMessages flag.

    'The Main Node is an application associated with this message service. It has a user interface that displays a list of connected applications.
    'UPDATE 6-Oct-2018 - ADVL_Message_Service_1 now hosts the Message Service. This is a separate application to the ADVL_Application_Network AppNet). - (COULD NOT COMMUNICATE BETWEEN AppNet and other apps when the Message Service was hosted from it.)
    Private Shared _mainNodeName As String = ""
    Property MainNodeName As String
        Get
            Return _mainNodeName
        End Get
        Set(value As String)
            _mainNodeName = value
        End Set
    End Property

    'The Main Node callback - used to send a message to the Main Node application.
    Private Shared _mainNodeCallback As IMsgServiceCallback
    Property MainNodeCallback As IMsgServiceCallback
        Get
            Return _mainNodeCallback
        End Get
        Set(value As IMsgServiceCallback)
            _mainNodeCallback = value
        End Set
    End Property


    Public Function Connect(ByVal proNetName As String, ByVal appName As String, ByVal connectionName As String, ByVal projectName As String, ByVal projectDescription As String, ByVal projectType As ADVL_Utilities_Library_1.Project.Types, ByVal projectPath As String, ByVal getAllWarnings As Boolean, ByVal getAllMessages As Boolean) As String Implements IMsgService.Connect
        'The Connect function adds a connection to the connections list.

        Try
            Dim callback As IMsgServiceCallback = OperationContext.Current.GetCallbackChannel(Of IMsgServiceCallback)()

            ''Check if connection is already on the connections list:
            Dim conn As clsConnection

            If ConnectionAvailable(proNetName, connectionName) Then
                Dim Connection As New clsConnection(proNetName, appName, connectionName, projectName, projectDescription, projectType, projectPath, callback, getAllWarnings, getAllMessages)
                connections.Add(Connection)

                'Add the new connection information to the data grid:
                If connectionName <> "MessageService" Then
                    If Main.ConnectionNameAvailable(proNetName, connectionName) Then

                        'Dont show it on Main.dgvConnections!

                        Main.dgvConnections.Rows.Add()
                        Dim CurrentRow As Integer = Main.dgvConnections.Rows.Count - 1 'Last blank connections row removed: 2 changed to 1

                        Main.dgvConnections.Rows(CurrentRow).Cells(0).Value = proNetName 'New connection ProNet Name

                        Main.dgvConnections.Rows(CurrentRow).Cells(1).Value = appName 'New connection App Name
                        Main.dgvConnections.Rows(CurrentRow).Cells(2).Value = connectionName 'New connection Name 
                        Main.dgvConnections.Rows(CurrentRow).Cells(3).Value = projectName 'New Project Name 

                        Main.dgvConnections.Rows(CurrentRow).Cells(4).Value = projectType.ToString 'New Project Type
                        Main.dgvConnections.Rows(CurrentRow).Cells(5).Value = projectPath 'New Project Path

                        Select Case getAllWarnings
                            Case True
                                Main.dgvConnections.Rows(CurrentRow).Cells(6).Value = "True" 'New connection GetAllWarnings is True
                            Case False
                                Main.dgvConnections.Rows(CurrentRow).Cells(6).Value = "False" 'New connection GetAllWarnings is False
                        End Select
                        Select Case getAllMessages
                            Case True
                                Main.dgvConnections.Rows(CurrentRow).Cells(7).Value = "True" 'New connection GetAllMessages is True
                            Case False
                                Main.dgvConnections.Rows(CurrentRow).Cells(7).Value = "False" 'New connection GetAllMessages is False
                        End Select
                        Main.dgvConnections.Rows(CurrentRow).Cells(8).Value = callback.GetHashCode 'New connection Callback hash code
                        Main.dgvConnections.Rows(CurrentRow).Cells(9).Value = Format(Now, "d-MMM-yyyy H:mm:ss") 'New connection start time
                        Main.dgvConnections.Rows(CurrentRow).Cells(10).Value = "0" 'New connection duration
                        Main.dgvConnections.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                        Main.dgvConnections.AutoResizeColumns()
                    Else
                        'Connection App Name not available.
                    End If
                End If

                Main.Message.Add("Connection added: [" & proNetName & "]." & connectionName & vbCrLf & vbCrLf)

                Connect = connectionName
            Else
                'Connection is already on the list.
                SendMainNodeMessage("WARNING: Connection failed because [" & proNetName & "]." & connectionName & " is already on the connections list." & vbCrLf & vbCrLf)
                Return False
            End If
            'End If
        Catch ex As Exception
            SendMainNodeMessage("WARNING: Connection failed: " & ex.Message & vbCrLf & vbCrLf)
            Return False
        End Try
    End Function


    Private Function NewConnName(ByVal ProNetName As String, ByVal reqConnName As String) As String
        'Return an available connection name based in the requested name: reqConnName.

        'Check if reqConnName is already on the connections list:
        Dim conn As clsConnection
        conn = connections.Find(Function(item As clsConnection)
                                    Return item.ConnectionName = reqConnName And item.ProNetName = ProNetName
                                End Function)
        If IsNothing(conn) Then 'reqConnName is not already on the list
            Return reqConnName 'Return reqConnName as an available connection name.
        Else
            Dim Imax As Integer = connections.Count + 1
            Dim tryConnName As String = ""
            Dim I As Integer
            For I = 1 To Imax
                tryConnName = reqConnName & "-" & I
                conn = connections.Find(Function(item As clsConnection)
                                            Return item.ConnectionName = tryConnName And item.ProNetName = ProNetName
                                        End Function)
                If IsNothing(conn) Then 'tryConnName is not already on the list
                    'tryConnName is not on the list. It can be used for a new connection.
                    Exit For
                Else 'tryConnName is already on the connection list.
                    tryConnName = ""
                End If
            Next
            Return tryConnName
        End If
    End Function


    Public Function ConnectionAvailable(ByVal ProNetName As String, ByVal ConnName As String) As Boolean Implements IMsgService.ConnectionAvailable
        'Return True if a Connection named ConnName is available for use in the Application Network named AppNetName.

        Dim conn As clsConnection
        conn = connections.Find(Function(item As clsConnection)
                                    Return item.ConnectionName = ConnName And item.ProNetName = ProNetName
                                End Function)
        If IsNothing(conn) Then 'ConnName is available for use in the Project Network named ProNetName.
            Return True
        Else 'ConnName is already in use in the Project Network named ProNetName.
            Return False
        End If

    End Function

    Public Function ConnectionExists(ByVal ProNetName As String, ByVal ConnName As String) As Boolean Implements IMsgService.ConnectionExists
        'Return True if a Connection named ConnName exists in the Application Network named AppNetName. (Opposite to ConnectionAvailable function.)

        Dim conn As clsConnection
        conn = connections.Find(Function(item As clsConnection)
                                    Return item.ConnectionName = ConnName And item.ProNetName = ProNetName
                                End Function)
        If IsNothing(conn) Then 'ConnName does not exist in the Project Network named ProNetName.
            Return False
        Else 'ConnName exists in the Project Network named ProNetName.
            Return True
        End If
    End Function


    Public Sub SendMessage(ByVal proNetName As String, ByVal connName As String, ByVal message As String) Implements IMsgService.SendMessage
        'Send the message to the application with the connection name appName.

        'FOR DEBUGGING:
        'Main.Message.Add("Executing SendMessage(AppNetName = " & appNetName & " , connName = " & connName & " )" & vbCrLf)
        'Main.Message.Add("Executing SendMessage(ProNetName = " & proNetName & " , connName = " & connName & " )" & vbCrLf)

        'Find the connection for the application corresponding to appName:
        Dim conn As clsConnection
        conn = connections.Find(Function(item As clsConnection)
                                    Return item.ConnectionName = connName And item.ProNetName = proNetName
                                End Function)
        If IsNothing(conn) Then
            'The connection is not on the list!

            If connName = "MessageService" Then
                Main.InstrReceived = message
            Else
                Main.Message.Add("Connection name: " & connName & " not found." & vbCrLf)
            End If
        Else
            If DirectCast(conn.Callback, ICommunicationObject).State = CommunicationState.Opened Then
                'Send a message showing the callers callback:
                Dim callback As IMsgServiceCallback = OperationContext.Current.GetCallbackChannel(Of IMsgServiceCallback)()
                Dim SenderName As String
                Dim connMatch = From conn2 In connections Where conn2.Callback.GetHashCode = callback.GetHashCode
                If connMatch.Count > 0 Then
                    'conn.Callback.OnSendMessage("The message was sent from: " & connMatch(0).AppName & vbCrLf)
                    'conn.Callback.OnSendMessage(connMatch(0).AppName & "> ")
                Else
                    conn.Callback.OnSendMessage("The sender is not on the connection list " & vbCrLf)
                End If
                'conn.Callback.OnSendMessage(message) 'The following error was returned when a large list of projected coordinate reference system names was sent by the Coordinates app:
                'An exception of type 'System.ServiceModel.ProtocolException' occurred in
                'mscorlib.dll but was not handled in user code.
                'Additional onformation: The remote server returned an unexpected response:
                '(413) Request Entity Too Large.

                'If DirectCast(conn.Callback, IContextChannel).State = CommunicationState.Faulted Then
                '    Debug.Print("Faulted ...")
                'End If

                'Debug.Print("State: " & DirectCast(conn.Callback, IContextChannel).State.ToString)
                If IsNothing(DirectCast(conn.Callback, IContextChannel).SessionId) Then
                    Debug.Print("SessionId is Nothing")
                Else
                    Debug.Print("SessionId: " & DirectCast(conn.Callback, IContextChannel).SessionId.ToString)
                End If

                Debug.Print("LocalAddress: " & DirectCast(conn.Callback, IContextChannel).LocalAddress.ToString)

                If IsNothing(conn.Callback) Then
                    Debug.Print("IsNothing")
                End If

                Try
                    conn.Callback.OnSendMessage(message) 'If the application receiving the message has crashed, a timeout error is raised.
                Catch ex As Exception

                End Try

            Else
                connections.Remove(conn)
            End If
        End If
    End Sub

    Public Function CheckConnection(ByVal proNetName As String, ByVal connName As String) As String Implements IMsgService.CheckConnection
        'Check the connection with the specified Project Network Name and Connection Name.

        'METHOD FOR CHECKING THE CONNECTION NOT YET FOUND!!!

        'Find the connection fwith the specified Project Network Name and Connection Name:
        Dim conn As clsConnection
        conn = connections.Find(Function(item As clsConnection)
                                    Return item.ConnectionName = connName And item.ProNetName = proNetName
                                End Function)
        If IsNothing(conn) Then
            Return "None"
        Else
            ''NOTE: The following code always returns Opened!!!
            'If DirectCast(conn.Callback, ICommunicationObject).State = CommunicationState.Opened Then
            '        Return "Opened"
            '    ElseIf DirectCast(conn.Callback, ICommunicationObject).State = CommunicationState.Closed Then
            '        Return "Closed"
            '    ElseIf DirectCast(conn.Callback, ICommunicationObject).State = CommunicationState.Closing Then
            '        Return "Closing"
            '    ElseIf DirectCast(conn.Callback, ICommunicationObject).State = CommunicationState.Created Then
            '        Return "Created"
            '    ElseIf DirectCast(conn.Callback, ICommunicationObject).State = CommunicationState.Faulted Then
            '        Return "Faulted"
            '    ElseIf DirectCast(conn.Callback, ICommunicationObject).State = CommunicationState.Opening Then
            '        Return "Opening"
            '    Else
            '        Return "Unknown"
            '    End If

            If DirectCast(conn.Callback, ICommunicationObject).State = CommunicationState.Opened Then
                'DirectCast(conn.Callback, ICommunicationObject).EndPoint.SendTimeout = New System.TimeSpan(0, 0, 1)
                'conn.Callback.OnSendMessage("Test")
                'Return "Opened" 'THIS LINE DOES NOT LOCK THE NETWORK APP.
                'Return conn.Callback.OnSendMessageCheck()
            End If

        End If

    End Function


    Public Sub SendAllMessage(ByVal message As String, ByVal SenderConnName As String) Implements IMsgService.SendAllMessage
        'Send the message to all connections in the connections list.
        Dim I As Integer 'Loop index
        For I = 1 To connections.Count
            If connections(I - 1).ConnectionName = SenderConnName Then
                'Dont send the message back to the sender.
            Else
                If DirectCast(connections(I - 1).Callback, ICommunicationObject).State = CommunicationState.Opened Then
                    connections(I - 1).Callback.OnSendMessage(message)
                Else

                End If
            End If

        Next
    End Sub


    Public Sub SendMainNodeMessage(ByVal message As String) Implements IMsgService.SendMainNodeMessage
        'Send the message to the Main Node.
        'ADD TRY ... CATCH ------------------------------------------------------------
        If MainNodeName <> "" Then
            MainNodeCallback.OnSendMessage(message)
        End If
    End Sub


    Public Sub GetConnectionList() Implements IMsgService.GetConnectionList
        Dim callback As IMsgServiceCallback = OperationContext.Current.GetCallbackChannel(Of IMsgServiceCallback)() 'The connection list will be sent back to the requesting connection.

        Dim decl As New XDeclaration("1.0", "utf-8", "yes")
        Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
        Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
        Dim connectionList As New XElement("ConnectionList")

        For Each item In connections
            Dim connectionInfo As New XElement("Connection")

            Dim proNetName As New XElement("ProNetName", item.ProNetName)
            connectionInfo.Add(proNetName)

            Dim name As New XElement("Name", item.ConnectionName)
            connectionInfo.Add(name)
            Dim appName As New XElement("ApplicationName", item.AppName)
            connectionInfo.Add(appName)

            Dim getAllMessages As New XElement("GetAllMessages", item.GetAllMessages)
            connectionInfo.Add(getAllMessages)
            Dim getAllMWarnings As New XElement("GetAllWarnings", item.GetAllWarnings)
            connectionInfo.Add(getAllMWarnings)
            Dim projectName As New XElement("ProjectName", item.ProjectName)
            connectionInfo.Add(projectName)
            Dim projectDescription As New XElement("ProjectDescription", item.ProjectDescription)
            connectionInfo.Add(projectDescription)

            Dim projectType As New XElement("ProjectType", item.ProjectType)
            connectionInfo.Add(projectType)
            Dim projectPath As New XElement("ProjectPath", item.ProjectPath)
            connectionInfo.Add(projectPath)

            connectionList.Add(connectionInfo)
        Next

        xmessage.Add(connectionList)
        doc.Add(xmessage)

        callback.OnSendMessage(doc.ToString)

    End Sub


    Public Sub GetApplicationList(ByVal ClientLocn As String) Implements IMsgService.GetApplicationList
        'Get the list of applications from the Message Service.
        Dim callback As IMsgServiceCallback = OperationContext.Current.GetCallbackChannel(Of IMsgServiceCallback)() 'The application list will be sent back to the requesting conection.

        Dim decl As New XDeclaration("1.0", "utf-8", "yes")
        Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
        Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
        Dim applicationList As New XElement("ApplicationList")

        'Main.Message.Add("Executing GetApplicationList()" & vbCrLf) 'For testing 'For testing

        For Each item In Main.App.List
            Dim applicationInfo As New XElement("Application")
            Dim name As New XElement("Name", item.Name)
            applicationInfo.Add(name)
            'Main.Message.Add("NApplication name: " & item.Name & vbCrLf) 'For testing
            Dim description As New XElement("Description", item.Description)
            applicationInfo.Add(description)
            Dim directory As New XElement("Directory", item.Directory)
            applicationInfo.Add(directory)
            Dim executablePath As New XElement("ExecutablePath", item.ExecutablePath)
            applicationInfo.Add(executablePath)
            applicationList.Add(applicationInfo)
            'Exit For ' For testing - Exit after the first item!
        Next

        If Trim(ClientLocn) = "" Then
            xmessage.Add(applicationList)
        Else
            Dim location As New XElement(ClientLocn)
            location.Add(applicationList)
            xmessage.Add(location)
        End If
        doc.Add(xmessage)

        'Main.Message.Add("Sending Message:" & vbCrLf & doc.ToString & vbCrLf) 'For testing 'For testing

        callback.OnSendMessage(doc.ToString)

    End Sub

    Public Sub GetApplicationInfo(ByVal appName As String) Implements IMsgService.GetApplicationInfo
        'Get information about an application.
        Dim callback As IMsgServiceCallback = OperationContext.Current.GetCallbackChannel(Of IMsgServiceCallback)() 'The application information will be sent back to the requesting conection.

        Dim decl As New XDeclaration("1.0", "utf-8", "yes")
        Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
        Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
        Dim applicationInfo As New XElement("ApplicationInfo")

        Dim newName As String = ""
        Dim newDescription As String = ""
        Dim newDirectory As String = ""
        Dim newExePath As String = ""

        Dim Count As Integer = 0

        For Each item In Main.App.List
            If item.Name = appName Then 'The appName has been found in the Application List.
                Count += 1 'Increment the cout of found names. (There should only be one application with this name found.)
                newName = item.Name
                newDescription = item.Description
                newDirectory = item.Directory
                newExePath = item.ExecutablePath
            End If
        Next

        If Count = 0 Then
            Dim name As New XElement("Name", newName)
            applicationInfo.Add(name)
            Dim description As New XElement("Description", "")
            applicationInfo.Add(description)
            Dim directory As New XElement("Directory", "")
            applicationInfo.Add(directory)
            Dim executablePath As New XElement("ExecutablePath", "")
            applicationInfo.Add(executablePath)
            Dim status As New XElement("Status", "Application not found in Message Service list.")
            applicationInfo.Add(status)
            xmessage.Add(applicationInfo)
            doc.Add(xmessage)
            callback.OnSendMessage(doc.ToString)
        ElseIf Count > 1 Then
            Dim name As New XElement("Name", newName)
            applicationInfo.Add(name)
            Dim description As New XElement("Description", "")
            applicationInfo.Add(description)
            Dim directory As New XElement("Directory", "")
            applicationInfo.Add(directory)
            Dim executablePath As New XElement("ExecutablePath", "")
            applicationInfo.Add(executablePath)
            Dim status As New XElement("Status", "More than one Application name matches found in Message Service list.")
            applicationInfo.Add(status)
            xmessage.Add(applicationInfo)
            doc.Add(xmessage)
            callback.OnSendMessage(doc.ToString)
        Else 'Single application found with the name appName.
            Dim name As New XElement("Name", newName)
            applicationInfo.Add(name)
            Dim description As New XElement("Description", newDescription)
            applicationInfo.Add(description)
            Dim directory As New XElement("Directory", newDirectory)
            applicationInfo.Add(directory)
            Dim executablePath As New XElement("ExecutablePath", newExePath)
            applicationInfo.Add(executablePath)
            Dim status As New XElement("Status", "Application information found in Message service list.")
            applicationInfo.Add(status)
            xmessage.Add(applicationInfo)
            doc.Add(xmessage)
            callback.OnSendMessage(doc.ToString)
        End If

    End Sub


    Public Sub GetProjectList(ByVal ClientLocn As String) Implements IMsgService.GetProjectList
        'Get the list of Projects from the Message Service.
        'Send the list to the Location in the Client. (The list will be contained in the <ClientLocn></ClientLocn> element.)

        Dim callback As IMsgServiceCallback = OperationContext.Current.GetCallbackChannel(Of IMsgServiceCallback)() 'The application list will be sent back to the requesting conection.

        Dim decl As New XDeclaration("1.0", "utf-8", "yes")
        Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
        Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
        Dim projectList As New XElement("ProjectList")

        'Main.Message.Add("Executing GetApplicationList()" & vbCrLf) 'For testing 'For testing

        'For Each item In Main.App.List
        For Each item In Main.Proj.List
            Dim applicationInfo As New XElement("Project")
            Dim name As New XElement("Name", item.Name)
            applicationInfo.Add(name)
            'Main.Message.Add("NApplication name: " & item.Name & vbCrLf) 'For testing
            Dim description As New XElement("Description", item.Description)
            applicationInfo.Add(description)
            Dim proNetName As New XElement("ProjectNetworkName", item.ProNetName)
            applicationInfo.Add(proNetName)
            Dim iD As New XElement("ID", item.ID)
            applicationInfo.Add(iD)
            Dim projType As New XElement("Type", item.Type.ToString)
            applicationInfo.Add(projType)
            Dim projPath As New XElement("Path", item.Path)
            applicationInfo.Add(projPath)
            Dim appName As New XElement("ApplicationName", item.ApplicationName)
            applicationInfo.Add(appName)
            Dim parentProjName As New XElement("ParentProjectName", item.ParentProjectName)
            applicationInfo.Add(parentProjName)
            Dim parentProjID As New XElement("ParentProjectID", item.ParentProjectID)
            applicationInfo.Add(parentProjID)
            projectList.Add(applicationInfo)
            'Exit For ' For testing - Exit after the first item!
        Next

        If Trim(ClientLocn) = "" Then
            xmessage.Add(projectList)
        Else
            Dim location As New XElement(ClientLocn)
            location.Add(projectList)
            xmessage.Add(location)

        End If

        doc.Add(xmessage)

        'Main.Message.Add("Sending Message:" & vbCrLf & doc.ToString & vbCrLf) 'For testing 'For testing

        callback.OnSendMessage(doc.ToString)
    End Sub


    Public Sub GetAdvlNetworkAppInfo() Implements IMsgService.GetAdvlNetworkAppInfo
        'Get information about the Andorville™ Network application.
        Dim callback As IMsgServiceCallback = OperationContext.Current.GetCallbackChannel(Of IMsgServiceCallback)() 'The application information will be sent back to the requesting conection.

        Dim decl As New XDeclaration("1.0", "utf-8", "yes")
        Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
        Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
        Dim applicationInfo As New XElement("AdvlNetworkAppInfo")

        Dim applicationName As New XElement("Name", Main.ApplicationInfo.Name)
        applicationInfo.Add(applicationName)
        Dim applicationPath As New XElement("Path", Main.ApplicationInfo.ApplicationDir)
        applicationInfo.Add(applicationPath)
        Dim applicationExePath As New XElement("ExePath", Main.ApplicationInfo.ExecutablePath)
        applicationInfo.Add(applicationExePath)

        xmessage.Add(applicationInfo)
        doc.Add(xmessage)

        callback.OnSendMessage(doc.ToString)

    End Sub


    Public Function IsAlive() As Boolean Implements IMsgService.IsAlive
        'Returns True if the service is running
        Return True
    End Function


    Public Function Disconnect(ByVal proNetName As String, ByVal connName As String) As Boolean Implements IMsgService.Disconnect 'UPDATED 2Feb19
        'The Disconnect function removes a connection from the connections list.
        'Find the connection for the application corresponding to connName: 'Find the connection for the application corresponding to appName: 'UPDATED 12May18
        Dim conn As clsConnection
        conn = connections.Find(Function(item As clsConnection)
                                    Return item.ConnectionName = connName And item.ProNetName = proNetName
                                End Function)
        If IsNothing(conn) Then
            'The connection is not on the list!
            Main.Message.AddWarning("WARNING: Disconnection failed because proNetName = " & proNetName & " and connName = " & connName & " is not on the connections list." & vbCrLf & vbCrLf)

            'Show the connections in the list:
            Dim I As Integer
            Dim NConn As Integer = connections.Count
            Main.Message.Add("Number of connections: " & NConn & vbCrLf)
            For I = 0 To NConn - 1
                Main.Message.Add(I & "  ProNetName: " & connections(I).ProNetName & "  Connection Name: " & connections(I).ConnectionName & vbCrLf)
            Next
            'Run this in case the connection is still listed in dgvConnections:
            Main.RemoveConnectionWithName(proNetName, connName)
            Return False
        Else
            connections.Remove(conn)
            Main.Message.Add("Connection removed: [" & proNetName & "]." & connName & vbCrLf & vbCrLf)
            Main.RemoveConnectionWithName(proNetName, connName)
            Return True
        End If
    End Function


    Public Function ProNetNameUsed(ByVal ProNetName As String) As Boolean Implements IMsgService.ProNetNameUsed
        'Return True if the specified Application Network Name is used in the list of Connections.

        Dim conn As clsConnection
        conn = connections.Find(Function(item As clsConnection)
                                    Return item.ProNetName = ProNetName
                                End Function)
        If IsNothing(conn) Then 'No connection found using the Project Network Name ProNetName.
            Return False
        Else 'A connection was found using the Project Network Name ProNetName.
            Return True
        End If

    End Function

    Public Sub StartProjectAtPath(ByVal ProjectPath As String, ConnectionName As String) Implements IMsgService.StartProjectAtPath
        'Start the project at the specified Project Path using the specified Connection Name.
        Main.StartProject(ProjectPath, ConnectionName)
    End Sub

    Public Sub StartProjectWithName(ByVal ProjectName As String, ByVal ProNetName As String, ByVal AppName As String, ByVal ConnName As String) Implements IMsgService.StartProjectWithName
        Main.StartProject(ProjectName, ProNetName, AppName, ConnName)
    End Sub

    Public Function ProjectOpen(ByVal ProjectPath As String) As Boolean Implements IMsgService.ProjectOpen
        'Return True if the Project at the specified Path is open.

        Dim conn As clsConnection
        conn = connections.Find(Function(item As clsConnection)
                                    Return item.ProjectPath = ProjectPath
                                End Function)
        If IsNothing(conn) Then 'No connection found using a Project at the specified Path.
            Return False
        Else 'A connection was found using the Project at the specified Path.
            Return True
        End If
    End Function

    Public Function ConnNameFromProjPath(ByVal ProjectPath As String) As String Implements IMsgService.ConnNameFromProjPath
        'Return the Connection Name corresponding to the specified Project Path if it is open.
        'Return "" if the Project is not open.

        Dim conn As clsConnection
        conn = connections.Find(Function(item As clsConnection)
                                    Return item.ProjectPath = ProjectPath
                                End Function)
        If IsNothing(conn) Then 'No connection found using a Project at the specified Path.
            Return ""
        Else 'A connection was found using the Project at the specified Path.
            Return conn.ConnectionName
        End If
    End Function

    Public Function ConnNameFromProjName(ByVal ProjectName As String, ByVal ProNetName As String, ByVal AppName As String) As String Implements IMsgService.ConnNameFromProjName
        'Return the Connection Name corresponding to the specified Project Name, Project Network Name and Application Name if it is open.
        'Return "" if the Project is not open.

        Dim conn As clsConnection
        conn = connections.Find(Function(item As clsConnection)
                                    Return item.ProjectName = ProjectName And item.ProNetName = ProNetName And item.AppName = AppName
                                End Function)
        If IsNothing(conn) Then 'No connection found using a the specified Project.
            Return ""
        Else 'A connection was found using the specified Project.
            Return conn.ConnectionName
        End If
    End Function

End Class
