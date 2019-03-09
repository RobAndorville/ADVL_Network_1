Public Class clsConnection
    'This class specifies the Connection items in the Connection list.
    'These items contain the AppName, Callback, GetAllWarnings and GetAllMessages fields.
    'AppName is the name of the application with the Admin Connection.
    'Callback is the connection used to send the message.
    'GetAllWarnings is a flag that indicates if all warning messages are to be received.
    'GetAllMessages is a flag that indicates if all messages are to be received.

    'ADDED 2Feb19
    Private _appNetName As String = "" 'The name of the Application Network containing the connection.
    Friend Property AppNetName As String
        Get
            Return _appNetName
        End Get
        Set(value As String)
            _appNetName = value
        End Set
    End Property

    Private _appName As String 'The name of the application being connected.
    Friend Property AppName As String
        Get
            Return _appName
        End Get
        Set(value As String)
            _appName = value
        End Set
    End Property

    'Private _reqConnName As String 'The requested connection name. If not available, a new name will be used and the new name returned to the application.
    Private _connectionName As String 'The connection name. If not available, a new name will be used and the new name returned to the application.
    'Friend Property ReqConnName As String
    Friend Property ConnectionName As String
        Get
            Return _connectionName
        End Get
        Set(value As String)
            _connectionName = value
        End Set
    End Property

    Private _projectName As String 'The name of the project that the application will open.
    Friend Property ProjectName As String
        Get
            Return _projectName
        End Get
        Set(value As String)
            _projectName = value
        End Set
    End Property

    Private _projectDescription As String 'A description of the project.
    Friend Property ProjectDescription As String
        Get
            Return _projectDescription
        End Get
        Set(value As String)
            _projectDescription = value
        End Set
    End Property

    'Private _settingsLocnType As ADVL_Utilities_Library_1.FileLocation.Types 'The type of location used to store the project settings (Directory or Archive)
    'Friend Property SettingsLocnType As ADVL_Utilities_Library_1.FileLocation.Types
    '    Get
    '        Return _settingsLocnType
    '    End Get
    '    Set(value As ADVL_Utilities_Library_1.FileLocation.Types)
    '        _settingsLocnType = value
    '    End Set
    'End Property

    'ADDED 2Feb19
    Private _projectType As ADVL_Utilities_Library_1.Project.Types 'The type of Project (Directory, Archive or Hybrid)
    Friend Property ProjectType As ADVL_Utilities_Library_1.Project.Types
        Get
            Return _projectType
        End Get
        Set(value As ADVL_Utilities_Library_1.Project.Types)
            _projectType = value
        End Set
    End Property

    'Private _settingsLocnPath As String 'The path to the project settings location.
    ''Friend Property ProjectPath As String
    'Friend Property SettingsLocnPath As String
    '    Get
    '        Return _settingsLocnPath
    '    End Get
    '    Set(value As String)
    '        _settingsLocnPath = value
    '    End Set
    'End Property

    'ADDED 2Feb19
    Private _projectPath As String 'The path to the project location.
    'Friend Property ProjectPath As String
    Friend Property ProjectPath As String
        Get
            Return _projectPath
        End Get
        Set(value As String)
            _projectPath = value
        End Set
    End Property

    'REMOVED 2Feb19
    ''Public Enum enumAppType
    'Public Enum AppTypes 'List of connection types.
    '    Application
    '    MainNode
    '    Node
    'End Enum

    ''Private _appType As enumAppType = enumAppType.Application
    'Private _appType As AppTypes = AppTypes.Application 'The type of connection (Application, MainNode or Node).
    'Friend Property AppType As AppTypes
    '    Get
    '        Return _appType
    '    End Get
    '    Set(value As AppTypes)
    '        _appType = value
    '    End Set
    'End Property

    Private _callback As IMsgServiceCallback
    Friend Property Callback As IMsgServiceCallback
        Get
            Return _callback
        End Get
        Set(value As IMsgServiceCallback)
            _callback = value
        End Set
    End Property

    Private _getAllWarnings As Boolean = False 'If True, this connection will receive all warnings.
    Friend Property GetAllWarnings As Boolean
        Get
            Return _getAllWarnings
        End Get
        Set(value As Boolean)
            _getAllWarnings = value
        End Set
    End Property

    Private _getAllMessages As Boolean = False 'If True, this connection will receive all messages.
    Friend Property GetAllMessages As Boolean
        Get
            Return _getAllMessages
        End Get
        Set(value As Boolean)
            _getAllMessages = value
        End Set
    End Property

    'Friend Sub New(ByVal newAppName As String, ByRef newAppType As AppTypes, ByRef newCallback As IMsgServiceCallback, ByVal newGetAllWarnings As Boolean, ByVal newGetAllMessages As Boolean)
    'Friend Sub New(ByVal newAppName As String, ByVal newReqConnName As String, ByVal newProjectName As String, ByVal newProjectPath As String, ByRef newAppType As AppTypes, ByRef newCallback As IMsgServiceCallback, ByVal newGetAllWarnings As Boolean, ByVal newGetAllMessages As Boolean)
    'Friend Sub New(ByVal newAppName As String, ByVal newConnName As String, ByVal newProjectName As String, ByVal newProjectPath As String, ByRef newAppType As AppTypes, ByRef newCallback As IMsgServiceCallback, ByVal newGetAllWarnings As Boolean, ByVal newGetAllMessages As Boolean)
    'Friend Sub New(ByVal newAppName As String, ByVal newConnName As String, ByVal newProjectName As String, ByVal newProjectDescription As String, ByVal newSettingsLocnType As ADVL_Utilities_Library_1.FileLocation.Types, ByVal newSettingsLocnPath As String, ByRef newAppType As AppTypes, ByRef newCallback As IMsgServiceCallback, ByVal newGetAllWarnings As Boolean, ByVal newGetAllMessages As Boolean)
    'Friend Sub New(ByVal newAppName As String, ByVal newConnName As String, ByVal newProjectName As String, ByVal newProjectDescription As String, ByVal newProjectType As ADVL_Utilities_Library_1.Project.Types, ByVal newProjectPath As String, ByRef newCallback As IMsgServiceCallback, ByVal newGetAllWarnings As Boolean, ByVal newGetAllMessages As Boolean)
    Friend Sub New(ByVal newAppNetName As String, ByVal newAppName As String, ByVal newConnName As String, ByVal newProjectName As String, ByVal newProjectDescription As String, ByVal newProjectType As ADVL_Utilities_Library_1.Project.Types, ByVal newProjectPath As String, ByRef newCallback As IMsgServiceCallback, ByVal newGetAllWarnings As Boolean, ByVal newGetAllMessages As Boolean) 'UPDATED 3Feb19
        AppNetName = newAppNetName 'ADDED 3Feb19
        AppName = newAppName
        ConnectionName = newConnName
        ProjectName = newProjectName
        ProjectDescription = newProjectDescription
        'SettingsLocnType = newSettingsLocnType
        ProjectType = newProjectType
        'SettingsLocnPath = newSettingsLocnPath
        ProjectPath = newProjectPath
        'AppType = newAppType
        Callback = newCallback
        GetAllWarnings = newGetAllWarnings
        GetAllMessages = newGetAllMessages
    End Sub
End Class
