'==============================================================================================================================================================================================
'Executing SendMessage(ProNetN
'Copyright 2018 Signalworks Pty Ltd, ABN 26 066 681 598

'Licensed under the Apache License, Version 2.0 (the "License");
'you may not use this file except in compliance with the License.
'You may obtain a copy of the License at
'
'http://www.apache.org/licenses/LICENSE-2.0
'
'Unless required by applicable law or agreed to in writing, software
'distributed under the License is distributed on an "AS IS" BASIS,
''WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'See the License for the specific language governing permissions and
'limitations under the License.
'
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Imports System.ServiceModel
Imports System.ServiceModel.Description

Imports System.Security.Permissions
Imports System.ComponentModel

Imports System.IO
Imports System.IO.Compression

<PermissionSet(SecurityAction.Demand, Name:="FullTrust")>
<System.Runtime.InteropServices.ComVisibleAttribute(True)>
Public Class Main
    'The ADVL_Message_Service application hosts the message service.
    'This is used by Andorville™ software applications To exchange information.

#Region " Coding Notes - Notes on the code used in this class." '==============================================================================================================================

    'ADD THE SYSTEM UTILITIES REFERENCE: ==========================================================================================
    'The following references are required by this software: 
    'ADVL_Utilities_Library_1.dll
    'To add the reference, press Project \ Add Reference... 
    '  Select the Browse option then press the Browse button
    '  Find the ADVL_Utilities_Library_1.dll file (it should be located in the directory ...\Projects\ADVL_Utilities_Library_1\ADVL_Utilities_Library_1\bin\Debug\)
    '  Press the Add button. Press the OK button.
    'The Utilities Library is used for Project Management, Archive file management, running XSequence files and running XMessage files.
    'If there are problems with a reference, try deleting it from the references list and adding it again.

    'References required:
    'System.ServiceModel
    'System.Runtime.Serialization
    '------------------------------------------------------------------------------------------------------------------------------
    '
    'Calling JavaScript from VB.NET:
    'The following Imports statement and permissions are required for the Main form:
    'Imports System.Security.Permissions
    '<PermissionSet(SecurityAction.Demand, Name:="FullTrust")> _
    '<System.Runtime.InteropServices.ComVisibleAttribute(True)> _
    'NOTE: the line continuation characters (_) will disappear form the code view after they have been typed!
    '------------------------------------------------------------------------------------------------------------------------------
    'Calling VB.NET from JavaScript
    'Add the following line to the Main.Load method:
    '  Me.WebBrowser1.ObjectForScripting = Me
    '------------------------------------------------------------------------------------------------------------------------------
    '
    'Using the XmlHtmDisplay control.
    '  The ADVL_Utilities_Library_1 project was added to this solution before the XmlHtmDisplay control appeared in the ToolBox.
    '    File \ Add \ Existing Project ...
    '------------------------------------------------------------------------------------------------------------------------------



#End Region 'Coding Notes ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Variable Declarations - All the variables and class objects used in this form and this application." '===============================================================================

    Public WithEvents ApplicationInfo As New ADVL_Utilities_Library_1.ApplicationInfo 'This object is used to store application information.
    Public WithEvents Project As New ADVL_Utilities_Library_1.Project 'This object is used to store Project information.
    Public WithEvents Message As New ADVL_Utilities_Library_1.Message 'This object is used to display messages in the Messages window.
    Public WithEvents ApplicationUsage As New ADVL_Utilities_Library_1.Usage 'This object stores application usage information.

    'Declare Forms used by the application:
    Public WithEvents WebPageList As frmWebPageList
    Public WithEvents ProjectArchive As frmArchive 'Form used to view the files in a Project archive
    Public WithEvents SettingsArchive As frmArchive 'Form used to view the files in a Settings archive
    Public WithEvents DataArchive As frmArchive 'Form used to view the files in a Data archive
    Public WithEvents SystemArchive As frmArchive 'Form used to view the files in a System archive

    Public WithEvents NewHtmlDisplay As frmHtmlDisplay
    Public HtmlDisplayFormList As New ArrayList 'Used for displaying multiple HtmlDisplay forms.

    Public WithEvents NewWebPage As frmWebPage
    Public WebPageFormList As New ArrayList 'Used for displaying multiple WebView forms.

    Public WithEvents ConnectionTools As frmConnectionTools 'Used to check the connection to the Message Service.


    'Project \ Add Reference \ Assemblies \ Framework \ System.ServiceModel
    Private Shared myHost As ServiceHost
    Dim smb As ServiceMetadataBehavior


    'Declare objects used to connect to the Message Service:
    Public client As ServiceReference1.MsgServiceClient
    'NOTE: This connection is for the Message Service form to communication with other applications.
    Public ConnectionName As String = "ADVL_Network_1" 'The name of the connection used to connect this application to the ComNet (Message Service).

    'Variables used to check the connection to the Message Service Form:
    Private ConnectionCheck As Boolean = False 'Used to check if the connection is working.
    Private ConnectionCheckStatus As String = "Passed"   'The status of the connection check: Waiting, Passed or Failed.
    Private ConnectionCheckStart As Date = Now      'Records the time each connection check is started.

    Public WithEvents XMsg As New ADVL_Utilities_Library_1.XMessage
    Dim XDoc As New System.Xml.XmlDocument
    Public Status As New System.Collections.Specialized.StringCollection
    Dim ClientAppName As String 'The name of the client 
    Dim MessageText As String 'The text of a message sent through the MessageExchange
    Dim MessageDest As String 'The destination of a message sent through the MessageExchange.

    Dim ClientProNetName As String = "" 'The name of the client Project Network requesting service. THIS MAY NEVER BE USED BY ADVL_NETWORK???
    Dim ClientConnName As String = "" 'The name of the client connection requesting service
    Dim MessageXDoc As System.Xml.Linq.XDocument
    Dim xmessage As XElement 'This will contain the message. It will be added to MessageXDoc.
    Dim xlocns As New List(Of XElement) 'A list of locations. Each location forms part of the reply message. The information in the reply message will be sent to the specified location in the client application.

    'Dim CompletionInstruction As String = "Stop" 'The last instruction returned on completion of the processing of an XMessage.
    Public OnCompletionInstruction As String = "Stop" 'The last instruction returned on completion of the processing of an XMessage.
    Public EndInstruction As String = "Stop" 'Another method of specifying the last instruction. This is processed in the EndOfSequence section of XMsg.Instructions.


    'Variable for local processing of an XMessage:
    Public WithEvents XMsgLocal As New ADVL_Utilities_Library_1.XMessage
    Dim XDocLocal As New System.Xml.XmlDocument
    Public StatusLocal As New System.Collections.Specialized.StringCollection

    'Main.Load variables:
    Dim ProjectSelected As Boolean = False 'If True, a project has been selected using Command Arguments. Used in Main.Load.

    'The following variables are used to run JavaScript in Web Pages loaded into the Document View: -------------------
    Public WithEvents XSeq As New ADVL_Utilities_Library_1.XSequence
    Private XStatus As New System.Collections.Specialized.StringCollection

    'Flags used for adding new connections or applications: ---------------------------------------------------------------------------
    Dim AddNewConnection As Boolean = False 'If True, a new connection can be added to the connection list.
    Dim AddNewApplication As Boolean = False 'If True, a new application can be added to the application list.
    Dim ApplicationNo As Integer 'The index number of an application that has been found in the App list.
    'If an application name is already on the application list, AddNewApplication is set to False.
    '----------------------------------------------------------------------------------------------------------------------------------

    'Variables used to start a new application: ---------------------------------------------------------------------------------------
    Dim StartAppName As String = ""
    Dim StartAppConnName As String = ""
    Dim StartAppProjectName As String = "" 'For starting an application with a specific project name.
    Dim StartAppProjectID As String = ""   'For starting an application with a specific project ID.
    Dim StartAppProjectPath As String = "" 'For starting an application with a specific project path.
    '
    Dim SelectedProNetName 'The selected Project Network Name - Used when starting a new application or removing a connection.
    '----------------------------------------------------------------------------------------------------------------------------------

    'Application List: ----------------------------------------------------------------------------------------------------------------
    Public App As New App 'App contains a list of all applications. App also contains methods to read, add and save the list.
    'The list is read from an xml file on startup and saved to an xml file on exit.
    'The list is displayed in the dgvApplications datagridview - in the Application List tab.

    'Project List: ----------------------------------------------------------------------------------------------------------------
    Public Proj As New Proj 'Proj contains a list of all projects. Proj also contains methods to read, add and save the list.
    'The list is read from an xml file on startup and saved to an xml file on exit.
    'The list is displayed in the dgvProjects datagridview - in the Project List tab.

    'Application Dictionary: ----------------------------------------------------------------------------------------------------------
    Dim AddNewApp As Boolean = False 'If True, a new application can be added to the AppInfo dictionary.
    Dim AppName As String = ""    'The name of the new App. (This is also the key for the AppInfo dictionary.)
    Dim AppText As String = ""    'The text of the new App (Displayed on the AppTree).
    Dim AppInfo As New Dictionary(Of String, clsAppInfo) 'Dictionary of information about each application shown in the AppTreeView.

    'Project Dictionary: --------------------------------------------------------------------------------------------------------------
    Dim AddNewProject As Boolean = False 'If True, a new project is being added to the ProjectInfo dictionary.
    Dim ProjectName As String = ""       'The name of the new project. (This is also the key for the ProjectInfo dictionary.) 
    Dim ProjectText As String = ""       'The text of the new project (displayed on the AppTree).
    Dim ProjInfo As New Dictionary(Of String, clsProjInfo) 'Dictionary of information about each project shown in the AppTreeView.
    'The dictionary key is the ID and ".Proj"

    Dim NApplicationIcons As Integer = 0 'The number of application icons.
    Dim NProjectIcons As Integer = 8 'The number of Project icons. (These 8 icons are stored in ProjectIconImageList and added to AppTreeImageList when an App Tree is opened.)

    Dim LastExitAttempt As DateTime = Now

    'StartProject variables:
    Private StartProject_AppName As String  'The application name
    Private StartProject_ConnName As String 'The connection name
    Private StartProject_ProNetName As String 'The Project Network name
    Private StartProject_ProjID As String   'The project ID
    Private StartProject_ProjName As String ' The project name

    Private WithEvents bgwComCheck As New System.ComponentModel.BackgroundWorker 'Used to perform communication checks on a separate thread.

    Private WithEvents bgwAppComCheck As New System.ComponentModel.BackgroundWorker 'Used to perform communication checks of other connections on a separate thread.

    Public WithEvents bgwSendMessage As New System.ComponentModel.BackgroundWorker 'Used to send a message through the Message Service.
    Dim SendMessageParams As New clsSendMessageParams 'This hold the Send Message parameters: .ProjectNetworkName, .ConnectionName & .Message

    'Alternative SendMessage background worker - needed to send a message while instructions are being processed.
    Public WithEvents bgwSendMessageAlt As New System.ComponentModel.BackgroundWorker 'Used to send a message through the Message Service - alternative backgound worker.
    Dim SendMessageParamsAlt As New clsSendMessageParams 'This hold the Send Message parameters: .ProjectNetworkName, .ConnectionName & .Message - for the alternative background worker.

    Private AppComCheckStatus As String = ""

    Private ACC_ProNetName As String = "" 'Application Communication Check - Project Network Name
    Private ACC_ConnName As String = "" 'Application Communication Check - Connection Name

    Dim Searching As Boolean = False 'If searching IP Addresses then Searching is True.
    Dim SearchThread As System.Threading.Thread

#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Properties - All the properties used in this form and this application" '============================================================================================================

    Private _connectedToComNet As Boolean = False  'True if the application is connected to the Communication Network (Message Service).
    Property ConnectedToComNet As Boolean
        Get
            Return _connectedToComNet
        End Get
        Set(value As Boolean)
            _connectedToComNet = value
        End Set
    End Property

    Private _instrReceived As String = "" 'Contains Instructions received via the Message Service.
    Property InstrReceived As String
        Get
            Return _instrReceived
        End Get
        Set(value As String)
            If value = Nothing Then
                Message.Add("Empty message received!")
            Else
                _instrReceived = value
                ProcessInstructions(_instrReceived)
            End If
        End Set
    End Property

    Private Sub ProcessInstructions(ByVal Instructions As String)
        'Process the XMessage instructions.

        Dim MsgType As String
        If Instructions.StartsWith("<XMsg>") Then
            MsgType = "XMsg"
            If ShowXMessages Then
                'Add the message header to the XMessages window:
                Message.XAddText("Message received: " & vbCrLf, "XmlReceivedNotice")
            End If
        ElseIf Instructions.StartsWith("<XSys>") Then
            MsgType = "XSys"
            If ShowSysMessages Then
                'Add the message header to the XMessages window:
                Message.XAddText("System Message received: " & vbCrLf, "XmlReceivedNotice")
            End If
        Else
            MsgType = "Unknown"
        End If

        'If ShowXMessages Then
        '    'Add the message header to the XMessages window:
        '    Message.XAddText("Message received: " & vbCrLf, "XmlReceivedNotice")
        'End If

        'If Instructions.StartsWith("<XMsg>") Then 'This is an XMessage set of instructions.
        'If Instructions.StartsWith("<XMsg>") Or Instructions.StartsWith("<XSys>") Then 'This is an XMessage or XSystem set of instructions.
        If MsgType = "XMsg" Or MsgType = "XSys" Then 'This is an XMessage or XSystem set of instructions.
            Try
                'Inititalise the reply message:
                ClientProNetName = ""
                ClientConnName = ""
                ClientAppName = ""
                xlocns.Clear() 'Clear the list of locations in the reply message. 

                Dim Decl As New XDeclaration("1.0", "utf-8", "yes")
                MessageXDoc = New XDocument(Decl, Nothing) 'Reply message - this will be sent to the Client App.
                'xmessage = New XElement("XMsg")
                xmessage = New XElement(MsgType)
                xlocns.Add(New XElement("Main")) 'Initially set the location in the Client App to Main.

                'Run the received message:
                Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
                XDoc.LoadXml(XmlHeader & vbCrLf & Instructions.Replace("&", "&amp;")) 'Replace "&" with "&amp:" before loading the XML text.
                'If ShowXMessages Then
                '    Message.XAddXml(XDoc)  'Add the message to the XMessages window.
                '    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                'End If
                If (MsgType = "XMsg") And ShowXMessages Then
                    Message.XAddXml(XDoc)  'Add the message to the XMessages window.
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                ElseIf (MsgType = "XSys") And ShowSysMessages Then
                    Message.XAddXml(XDoc)  'Add the message to the XMessages window.
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                End If
                XMsg.Run(XDoc, Status)
            Catch ex As Exception
                Message.Add("Error running XMsg: " & ex.Message & vbCrLf)
            End Try

            'XMessage has been run.
            'Reply to this message:
            'Add the message reply to the XMessages window:
            'Complete the MessageXDoc:
            xmessage.Add(xlocns(xlocns.Count - 1)) 'Add the last location reply instructions to the message.
            MessageXDoc.Add(xmessage)
            MessageText = MessageXDoc.ToString

            If ClientConnName = "" Then
                'No client to send a message to - process the message locally.
                'If ShowXMessages Then
                '    Message.XAddText("Message processed locally:" & vbCrLf, "XmlSentNotice")
                '    Message.XAddXml(MessageText)
                '    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                'End If
                If (MsgType = "XMsg") And ShowXMessages Then
                    Message.XAddText("Message processed locally:" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(MessageText)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                ElseIf (MsgType = "XSys") And ShowSysMessages Then
                    Message.XAddText("System Message processed locally:" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(MessageText)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                End If
                ProcessLocalInstructions(MessageText)
            Else
                'If ShowXMessages Then
                '    Message.XAddText("Message sent to [" & ClientProNetName & "]." & ClientConnName & ":" & vbCrLf, "XmlSentNotice")   'NOTE: There is no SendMessage code in the Message Service application!
                '    Message.XAddXml(MessageText)
                '    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                'End If
                If (MsgType = "XMsg") And ShowXMessages Then
                    Message.XAddText("Message sent to [" & ClientProNetName & "]." & ClientConnName & ":" & vbCrLf, "XmlSentNotice")   'NOTE: There is no SendMessage code in the Message Service application!
                    Message.XAddXml(MessageText)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                ElseIf (MsgType = "XSys") And ShowSysMessages Then
                    Message.XAddText("System Message sent to [" & ClientProNetName & "]." & ClientConnName & ":" & vbCrLf, "XmlSentNotice")   'NOTE: There is no SendMessage code in the Message Service application!
                    Message.XAddXml(MessageText)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                End If

                'Send Message on a new thread:
                SendMessageParams.ProjectNetworkName = ClientProNetName
                SendMessageParams.ConnectionName = ClientConnName
                SendMessageParams.Message = MessageText
                If bgwSendMessage.IsBusy Then
                    Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                Else
                    bgwSendMessage.RunWorkerAsync(SendMessageParams)
                End If

            End If
        Else 'This is not an XMessage!
            If Instructions.StartsWith("<XMsgBlk>") Then 'This is an XMessageBlock.
                'Process the received message:
                Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
                XDoc.LoadXml(XmlHeader & vbCrLf & Instructions.Replace("&", "&amp;")) 'Replace "&" with "&amp:" before loading the XML text.
                'NOTE: The message is an <XMsgBlk> - use the ShowXMessages property to determine if the message is shown:
                If ShowXMessages Then
                    Message.XAddXml(XDoc)   'Add the message to the XMessages window.
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                End If
                'If (MsgType = "XMsg") And ShowXMessages Then
                '    Message.XAddXml(XDoc)   'Add the message to the XMessages window.
                '    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                'ElseIf (MsgType = "XSys") And ShowSysMessages Then
                '    Message.XAddXml(XDoc)   'Add the message to the XMessages window.
                '    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                'End If

                'Process the XMessageBlock:
                Dim XMsgBlkLocn As String
                XMsgBlkLocn = XDoc.GetElementsByTagName("ClientLocn")(0).InnerText
                Select Case XMsgBlkLocn
                    Case "TestLocn" 'Replace this with the required location name.
                        Dim XInfo As Xml.XmlNodeList = XDoc.GetElementsByTagName("XInfo") 'Get the XInfo node list
                        Dim InfoXDoc As New Xml.Linq.XDocument 'Create an XDocument to hold the information contained in XInfo 
                        InfoXDoc = XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>" & vbCrLf & XInfo(0).InnerXml) 'Read the information into InfoXDoc
                        'Add processing instructions here - The information in the InfoXDoc is usually stored in an XDocument in the application or as an XML file in the project.

                    Case Else
                        Message.AddWarning("Unknown XInfo Message location: " & XMsgBlkLocn & vbCrLf)
                End Select
            Else
                Message.XAddText("The message is not an XMessage or XMessageBlock: " & vbCrLf & Instructions & vbCrLf & vbCrLf, "Normal")
            End If
            'Message.XAddText("The message is not an XMessage: " & Instructions & vbCrLf, "Normal")
        End If
    End Sub

    Private Sub ProcessLocalInstructions(ByVal Instructions As String)
        'Process the XMessage instructions locally.

        'If Instructions.StartsWith("<XMsg>") Then 'This is an XMessage set of instructions.
        If Instructions.StartsWith("<XMsg>") Or Instructions.StartsWith("<XSys>") Then 'This is an XMessage set of instructions.
            'Run the received message:
            Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
            XDocLocal.LoadXml(XmlHeader & vbCrLf & Instructions)
            XMsgLocal.Run(XDocLocal, StatusLocal)
        Else 'This is not an XMessage!
            Message.XAddText("The message is not an XMessage: " & Instructions & vbCrLf, "Normal")
        End If
    End Sub

    Private _showXMessages As Boolean = True 'If True, XMessages that are sent or received will be shown in the Messages window.
    Property ShowXMessages As Boolean
        Get
            Return _showXMessages
        End Get
        Set(value As Boolean)
            _showXMessages = value
        End Set
    End Property

    Private _showSysMessages As Boolean = True 'If True, System messages that are sent or received will be shown in the messages window.
    Property ShowSysMessages As Boolean
        Get
            Return _showSysMessages
        End Get
        Set(value As Boolean)
            _showSysMessages = value
        End Set
    End Property

    Private _closedFormNo As Integer 'Temporarily holds the number of the form that is being closed. 
    Property ClosedFormNo As Integer
        Get
            Return _closedFormNo
        End Get
        Set(value As Integer)
            _closedFormNo = value
        End Set
    End Property

    'Private _startPageFileName As String = "" 'The file name of the html document displayed in the Start Page tab.
    'Public Property StartPageFileName As String
    '    Get
    '        Return _startPageFileName
    '    End Get
    '    Set(value As String)
    '        _startPageFileName = value
    '    End Set
    'End Property

    Private _workflowFileName As String = "" 'The file name of the html document displayed in the Workflow tab.
    Public Property WorkflowFileName As String
        Get
            Return _workflowFileName
        End Get
        Set(value As String)
            _workflowFileName = value
        End Set
    End Property

    Private _zipFilePath As String = "" 'The path of the zip file dragged into the View Zip Archive tab.
    Property ZipFilePath As String
        Get
            Return _zipFilePath
        End Get
        Set(value As String)
            _zipFilePath = value
        End Set
    End Property



#End Region 'Properties -----------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Process XML Files - Read and write XML files." '=====================================================================================================================================

    Private Sub SaveFormSettings()
        'Save the form settings in an XML document.
        Dim settingsData = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <!--Form settings for Main form.-->
                           <FormSettings>
                               <Left><%= Me.Left %></Left>
                               <Top><%= Me.Top %></Top>
                               <Width><%= Me.Width %></Width>
                               <Height><%= Me.Height %></Height>
                               <ShowXMessages><%= ShowXMessages %></ShowXMessages>
                               <ShowSysMessages><%= ShowSysMessages %></ShowSysMessages>
                               <WorkFlowFileName><%= WorkflowFileName %></WorkFlowFileName>
                               <!---->
                               <SelectedTabIndex><%= TabControl1.SelectedIndex %></SelectedTabIndex>
                               <ZipFilePath><%= ZipFilePath %></ZipFilePath>
                               <!---->
                               <ConnectionApplicationNameColumnWidth><%= dgvConnections.Columns(0).Width %></ConnectionApplicationNameColumnWidth>
                               <ConnectionTypeColumnWidth><%= dgvConnections.Columns(1).Width %></ConnectionTypeColumnWidth>
                               <ConnectionCallbackHashcodeColumnWidth><%= dgvConnections.Columns(2).Width %></ConnectionCallbackHashcodeColumnWidth>
                               <ConnectionStartTimeColumnWidth><%= dgvConnections.Columns(3).Width %></ConnectionStartTimeColumnWidth>
                               <!---->
                               <ApplicationNameColumnWidth><%= dgvApplications.Columns(0).Width %></ApplicationNameColumnWidth>
                               <ApplicationDescriptionColumnWidth><%= dgvApplications.Columns(1).Width %></ApplicationDescriptionColumnWidth>
                               <!---->
                               <ConnectAppToNetwork><%= chkConnect.Checked %></ConnectAppToNetwork>
                               <Connect2AppToNetwork><%= chkConnect2.Checked %></Connect2AppToNetwork>
                               <!---->
                               <AppTreeTabSplitDistance><%= SplitContainer1.SplitterDistance %></AppTreeTabSplitDistance>
                               <ShowMessages><%= chkShowMessages.Checked %></ShowMessages>
                               <ShowApplication><%= chkShowApp.Checked %></ShowApplication>
                           </FormSettings>

        '<AlignMessageWindows><%= chkAlignMessages.Checked %></AlignMessageWindows>

        'Add code to include other settings to save after the comment line <!---->

        'Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & "_" & Me.Text & ".xml"
        Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & " - Main.xml"
        Project.SaveXmlSettings(SettingsFileName, settingsData)
    End Sub

    Private Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        'Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & "_" & Me.Text & ".xml"
        Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & " - Main.xml"

        If Project.SettingsFileExists(SettingsFileName) Then
            Dim Settings As System.Xml.Linq.XDocument
            Project.ReadXmlSettings(SettingsFileName, Settings)

            If IsNothing(Settings) Then 'There is no Settings XML data.
                Exit Sub
            End If

            'Restore form position and size:
            If Settings.<FormSettings>.<Left>.Value <> Nothing Then Me.Left = Settings.<FormSettings>.<Left>.Value
            If Settings.<FormSettings>.<Top>.Value <> Nothing Then Me.Top = Settings.<FormSettings>.<Top>.Value
            If Settings.<FormSettings>.<Height>.Value <> Nothing Then Me.Height = Settings.<FormSettings>.<Height>.Value
            If Settings.<FormSettings>.<Width>.Value <> Nothing Then Me.Width = Settings.<FormSettings>.<Width>.Value

            If Settings.<FormSettings>.<ShowXMessages>.Value <> Nothing Then ShowXMessages = Settings.<FormSettings>.<ShowXMessages>.Value
            If Settings.<FormSettings>.<ShowSysMessages>.Value <> Nothing Then ShowSysMessages = Settings.<FormSettings>.<ShowSysMessages>.Value

            If Settings.<FormSettings>.<WorkFlowFileName>.Value <> Nothing Then WorkflowFileName = Settings.<FormSettings>.<WorkFlowFileName>.Value

            'Add code to read other saved setting here:
            If Settings.<FormSettings>.<SelectedTabIndex>.Value <> Nothing Then TabControl1.SelectedIndex = Settings.<FormSettings>.<SelectedTabIndex>.Value
            If Settings.<FormSettings>.<ZipFilePath>.Value <> Nothing Then
                ZipFilePath = Settings.<FormSettings>.<ZipFilePath>.Value
                txtZipFileDir.Text = System.IO.Path.GetDirectoryName(ZipFilePath)
                txtZipFileName.Text = System.IO.Path.GetFileName(ZipFilePath)
            End If

            If Settings.<FormSettings>.<ConnectionApplicationNameColumnWidth>.Value <> Nothing Then dgvConnections.Columns(0).Width = Settings.<FormSettings>.<ConnectionApplicationNameColumnWidth>.Value
            If Settings.<FormSettings>.<ConnectionTypeColumnWidth>.Value <> Nothing Then dgvConnections.Columns(1).Width = Settings.<FormSettings>.<ConnectionTypeColumnWidth>.Value
            If Settings.<FormSettings>.<ConnectionCallbackHashcodeColumnWidth>.Value <> Nothing Then dgvConnections.Columns(2).Width = Settings.<FormSettings>.<ConnectionCallbackHashcodeColumnWidth>.Value
            If Settings.<FormSettings>.<ConnectionStartTimeColumnWidth>.Value <> Nothing Then dgvConnections.Columns(3).Width = Settings.<FormSettings>.<ConnectionStartTimeColumnWidth>.Value

            If Settings.<FormSettings>.<ApplicationNameColumnWidth>.Value <> Nothing Then dgvApplications.Columns(0).Width = Settings.<FormSettings>.<ApplicationNameColumnWidth>.Value
            If Settings.<FormSettings>.<ApplicationDescriptionColumnWidth>.Value <> Nothing Then dgvApplications.Columns(1).Width = Settings.<FormSettings>.<ApplicationDescriptionColumnWidth>.Value

            If Settings.<FormSettings>.<ConnectAppToNetwork>.Value = Nothing Then
                'Leave at default value.
            Else
                If Settings.<FormSettings>.<ConnectAppToNetwork>.Value = True Then
                    chkConnect.Checked = True
                Else
                    chkConnect.Checked = False
                End If
            End If

            If Settings.<FormSettings>.<Connect2AppToNetwork>.Value = Nothing Then
                'Leave at default value.
            Else
                If Settings.<FormSettings>.<Connect2AppToNetwork>.Value = True Then
                    chkConnect2.Checked = True
                Else
                    chkConnect2.Checked = False
                End If
            End If

            If Settings.<FormSettings>.<AppTreeTabSplitDistance>.Value <> Nothing Then SplitContainer1.SplitterDistance = Settings.<FormSettings>.<AppTreeTabSplitDistance>.Value

            'If Settings.<FormSettings>.<AlignMessageWindows>.Value <> Nothing Then
            '    If Settings.<FormSettings>.<AlignMessageWindows>.Value = True Then
            '        chkAlignMessages.Checked = True
            '    Else
            '        chkAlignMessages.Checked = False
            '    End If
            'End If

            If Settings.<FormSettings>.<ShowMessages>.Value <> Nothing Then
                If Settings.<FormSettings>.<ShowMessages>.Value = True Then
                    chkShowMessages.Checked = True
                Else
                    chkShowMessages.Checked = False
                End If
            End If

            If Settings.<FormSettings>.<ShowApplication>.Value <> Nothing Then
                If Settings.<FormSettings>.<ShowApplication>.Value = True Then
                    chkShowApp.Checked = True
                Else
                    chkShowApp.Checked = False
                End If
            End If

            CheckFormPos()
        End If
    End Sub

    Private Sub CheckFormPos()
        'Check that the form can be seen on a screen.

        'Dim MinWidthVisible As Integer = 48 'Minimum number of X pixels visible. The form will be moved if this many form pixels are not visible.
        'Dim MinHeightVisible As Integer = 48 'Minimum number of Y pixels visible. The form will be moved if this many form pixels are not visible.
        Dim MinWidthVisible As Integer = 192 'Minimum number of X pixels visible. The form will be moved if this many form pixels are not visible.
        Dim MinHeightVisible As Integer = 64 'Minimum number of Y pixels visible. The form will be moved if this many form pixels are not visible.

        Dim FormRect As New Rectangle(Me.Left, Me.Top, Me.Width, Me.Height)
        Dim WARect As Rectangle = Screen.GetWorkingArea(FormRect) 'The Working Area rectangle - the usable area of the screen containing the form.

        ''Check if the top of the form is less than zero:
        'If Me.Top < 0 Then Me.Top = 0

        'Check if the top of the form is above the top of the Working Area:
        If Me.Top < WARect.Top Then
            Me.Top = WARect.Top
        End If

        'Check if the top of the form is too close to the bottom of the Working Area:
        If (Me.Top + MinHeightVisible) > (WARect.Top + WARect.Height) Then
            Me.Top = WARect.Top + WARect.Height - MinHeightVisible
        End If

        'Check if the left edge of the form is too close to the right edge of the Working Area:
        If (Me.Left + MinWidthVisible) > (WARect.Left + WARect.Width) Then
            Me.Left = WARect.Left + WARect.Width - MinWidthVisible
        End If

        'Check if the right edge of the form is too close to the left edge of the Working Area:
        If (Me.Left + Me.Width - MinWidthVisible) < WARect.Left Then
            Me.Left = WARect.Left - Me.Width + MinWidthVisible
        End If

    End Sub

    Private Sub ReadApplicationInfo()
        'Read the Application Information.

        If ApplicationInfo.FileExists Then
            ApplicationInfo.ReadFile()
        Else
            'There is no Application_Info.xml file.
            DefaultAppProperties() 'Create a new Application Info file with default application properties:
            ApplicationInfo.WriteFile() 'Write the file now. The file information may be used by other applications.
        End If
    End Sub

    Private Sub DefaultAppProperties()
        'These properties will be saved in the Application_Info_ADVL_2.xml file in the application directory.
        'If this file is deleted, it will be re-created using these default application properties.

        'Change this to show your application Name, Description and Creation Date.
        'ApplicationInfo.Name = "ADVL_Message_Service_1"
        ApplicationInfo.Name = "ADVL_Network_1"

        'ApplicationInfo.ApplicationDir is set when the application is started.
        ApplicationInfo.ExecutablePath = Application.ExecutablePath

        'The ADVL_Message_Service application hosts the message service.
        'This is used by Andorville™ software applications To exchange information.

        'ApplicationInfo.Description = "The Message Service application hosts the Message Service. This is used by Andorville™ software applications to exchange information."
        ApplicationInfo.Description = "The Andorville™ Network application hosts the Message Service. This is used by Andorville™ software applications to exchange information."
        ApplicationInfo.CreationDate = "6-Oct-2016 12:00:00"

        'Author -----------------------------------------------------------------------------------------------------------
        'Change this to show your Name, Description and Contact information.
        ApplicationInfo.Author.Name = "Signalworks Pty Ltd"
        ApplicationInfo.Author.Description = "Signalworks Pty Ltd" & vbCrLf &
            "Australian Proprietary Company" & vbCrLf &
            "ABN 26 066 681 598" & vbCrLf &
            "Registration Date 05/10/1994"

        ApplicationInfo.Author.Contact = "http://www.andorville.com.au/"

        'File Associations: -----------------------------------------------------------------------------------------------
        'Add any file associations here.
        'The file extension and a description of files that can be opened by this application are specified.
        'The example below specifies a coordinate system parameter file type with the file extension .ADVLCoord.
        'Dim Assn1 As New ADVL_System_Utilities.FileAssociation
        'Assn1.Extension = "ADVLCoord"
        'Assn1.Description = "Andorville™ software coordinate system parameter file"
        'ApplicationInfo.FileAssociations.Add(Assn1)

        'Version ----------------------------------------------------------------------------------------------------------
        ApplicationInfo.Version.Major = My.Application.Info.Version.Major
        ApplicationInfo.Version.Minor = My.Application.Info.Version.Minor
        ApplicationInfo.Version.Build = My.Application.Info.Version.Build
        ApplicationInfo.Version.Revision = My.Application.Info.Version.Revision

        'Copyright --------------------------------------------------------------------------------------------------------
        'Add your copyright information here.
        ApplicationInfo.Copyright.OwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        ApplicationInfo.Copyright.PublicationYear = "2018"

        'Trademarks -------------------------------------------------------------------------------------------------------
        'Add your trademark information here.
        Dim Trademark1 As New ADVL_Utilities_Library_1.Trademark
        Trademark1.OwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        Trademark1.Text = "Andorville"
        Trademark1.Registered = False
        Trademark1.GenericTerm = "software"
        ApplicationInfo.Trademarks.Add(Trademark1)
        Dim Trademark2 As New ADVL_Utilities_Library_1.Trademark
        Trademark2.OwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        Trademark2.Text = "AL-H7"
        Trademark2.Registered = False
        Trademark2.GenericTerm = "software"
        ApplicationInfo.Trademarks.Add(Trademark2)
        Dim Trademark3 As New ADVL_Utilities_Library_1.Trademark
        Trademark3.OwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        Trademark3.Text = "AL-M7"
        Trademark3.Registered = False
        Trademark3.GenericTerm = "software"
        ApplicationInfo.Trademarks.Add(Trademark3)

        'License -------------------------------------------------------------------------------------------------------
        'Add your license information here.
        ApplicationInfo.License.CopyrightOwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        ApplicationInfo.License.PublicationYear = "2018"

        'License Links:
        'http://choosealicense.com/
        'http://www.apache.org/licenses/
        'http://opensource.org/

        'Apache License 2.0 ---------------------------------------------
        ApplicationInfo.License.Code = ADVL_Utilities_Library_1.License.Codes.Apache_License_2_0
        ApplicationInfo.License.Notice = ApplicationInfo.License.ApacheLicenseNotice 'Get the pre-defined Aapche license notice.
        ApplicationInfo.License.Text = ApplicationInfo.License.ApacheLicenseText     'Get the pre-defined Apache license text.

        'Code to use other pre-defined license types is shown below:

        'GNU General Public License, version 3 --------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.GNU_GPL_V3_0
        'ApplicationInfo.License.Notice = 'Add the License Notice to ADVL_Utilities_Library_1 License class.
        'ApplicationInfo.License.Text = 'Add the License Text to ADVL_Utilities_Library_1 License class.

        'The MIT License ------------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.MIT_License
        'ApplicationInfo.License.Notice = ApplicationInfo.License.MITLicenseNotice
        'ApplicationInfo.License.Text = ApplicationInfo.License.MITLicenseText

        'No License Specified -------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.None
        'ApplicationInfo.License.Notice = ""
        'ApplicationInfo.License.Text = ""

        'The Unlicense --------------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.The_Unlicense
        'ApplicationInfo.License.Notice = ApplicationInfo.License.UnLicenseNotice
        'ApplicationInfo.License.Text = ApplicationInfo.License.UnLicenseText

        'Unknown License ------------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.Unknown
        'ApplicationInfo.License.Notice = ""
        'ApplicationInfo.License.Text = ""

        'Source Code: --------------------------------------------------------------------------------------------------
        'Add your source code information here if required.
        'THIS SECTION WILL BE UPDATED TO ALLOW A GITHUB LINK.
        ApplicationInfo.SourceCode.Language = "Visual Basic 2015"
        ApplicationInfo.SourceCode.FileName = ""
        ApplicationInfo.SourceCode.FileSize = 0
        ApplicationInfo.SourceCode.FileHash = ""
        ApplicationInfo.SourceCode.WebLink = ""
        ApplicationInfo.SourceCode.Contact = ""
        ApplicationInfo.SourceCode.Comments = ""

        'ModificationSummary: -----------------------------------------------------------------------------------------
        'Add any source code modification here is required.
        ApplicationInfo.ModificationSummary.BaseCodeName = ""
        ApplicationInfo.ModificationSummary.BaseCodeDescription = ""
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Major = 0
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Minor = 0
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Build = 0
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Revision = 0
        ApplicationInfo.ModificationSummary.Description = "This is the first released version of the application. No earlier base code used."

        'Library List: ------------------------------------------------------------------------------------------------
        'Add the ADVL_Utilties_Library_1 library:
        Dim NewLib As New ADVL_Utilities_Library_1.LibrarySummary
        NewLib.Name = "ADVL_System_Utilities"
        NewLib.Description = "System Utility classes used in Andorville™ software development system applications"
        NewLib.CreationDate = "7-Jan-2016 12:00:00"
        NewLib.LicenseNotice = "Copyright 2016 Signalworks Pty Ltd, ABN 26 066 681 598" & vbCrLf &
                               vbCrLf &
                               "Licensed under the Apache License, Version 2.0 (the ""License"");" & vbCrLf &
                               "you may not use this file except in compliance with the License." & vbCrLf &
                               "You may obtain a copy of the License at" & vbCrLf &
                               vbCrLf &
                               "http://www.apache.org/licenses/LICENSE-2.0" & vbCrLf &
                               vbCrLf &
                               "Unless required by applicable law or agreed to in writing, software" & vbCrLf &
                               "distributed under the License is distributed on an ""AS IS"" BASIS," & vbCrLf &
                               "WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied." & vbCrLf &
                               "See the License for the specific language governing permissions and" & vbCrLf &
                               "limitations under the License." & vbCrLf

        NewLib.CopyrightNotice = "Copyright 2016 Signalworks Pty Ltd, ABN 26 066 681 598"

        NewLib.Version.Major = 1
        NewLib.Version.Minor = 0
        NewLib.Version.Build = 1
        NewLib.Version.Revision = 0

        NewLib.Author.Name = "Signalworks Pty Ltd"
        NewLib.Author.Description = "Signalworks Pty Ltd" & vbCrLf &
            "Australian Proprietary Company" & vbCrLf &
            "ABN 26 066 681 598" & vbCrLf &
            "Registration Date 05/10/1994"

        NewLib.Author.Contact = "http://www.andorville.com.au/"

        Dim NewClass1 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass1.Name = "ZipComp"
        NewClass1.Description = "The ZipComp class is used to compress files into and extract files from a zip file."
        NewLib.Classes.Add(NewClass1)
        Dim NewClass2 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass2.Name = "XSequence"
        NewClass2.Description = "The XSequence class is used to run an XML property sequence (XSequence) file. XSequence files are used to record and replay processing sequences in Andorville™ software applications."
        NewLib.Classes.Add(NewClass2)
        Dim NewClass3 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass3.Name = "XMessage"
        NewClass3.Description = "The XMessage class is used to read an XML Message (XMessage). An XMessage is a simplified XSequence used to exchange information between Andorville™ software applications."
        NewLib.Classes.Add(NewClass3)
        Dim NewClass4 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass4.Name = "Location"
        NewClass4.Description = "The Location class consists of properties and methods to store data in a location, which is either a directory or archive file."
        NewLib.Classes.Add(NewClass4)
        Dim NewClass5 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass5.Name = "Project"
        NewClass5.Description = "An Andorville™ software application can store data within one or more projects. Each project stores a set of related data files. The Project class contains properties and methods used to manage a project."
        NewLib.Classes.Add(NewClass5)
        Dim NewClass6 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass6.Name = "ProjectSummary"
        NewClass6.Description = "ProjectSummary stores a summary of a project."
        NewLib.Classes.Add(NewClass6)
        Dim NewClass7 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass7.Name = "DataFileInfo"
        NewClass7.Description = "The DataFileInfo class stores information about a data file."
        NewLib.Classes.Add(NewClass7)
        Dim NewClass8 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass8.Name = "Message"
        NewClass8.Description = "The Message class contains text properties and methods used to display messages in an Andorville™ software application."
        NewLib.Classes.Add(NewClass8)
        Dim NewClass9 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass9.Name = "ApplicationSummary"
        NewClass9.Description = "The ApplicationSummary class stores a summary of an Andorville™ software application."
        NewLib.Classes.Add(NewClass9)
        Dim NewClass10 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass10.Name = "LibrarySummary"
        NewClass10.Description = "The LibrarySummary class stores a summary of a software library used by an application."
        NewLib.Classes.Add(NewClass10)
        Dim NewClass11 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass11.Name = "ClassSummary"
        NewClass11.Description = "The ClassSummary class stores a summary of a class contained in a software library."
        NewLib.Classes.Add(NewClass11)
        Dim NewClass12 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass12.Name = "ModificationSummary"
        NewClass12.Description = "The ModificationSummary class stores a summary of any modifications made to an application or library."
        NewLib.Classes.Add(NewClass12)
        Dim NewClass13 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass13.Name = "ApplicationInfo"
        NewClass13.Description = "The ApplicationInfo class stores information about an Andorville™ software application."
        NewLib.Classes.Add(NewClass13)
        Dim NewClass14 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass14.Name = "Version"
        NewClass14.Description = "The Version class stores application, library or project version information."
        NewLib.Classes.Add(NewClass14)
        Dim NewClass15 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass15.Name = "Author"
        NewClass15.Description = "The Author class stores information about an Author."
        NewLib.Classes.Add(NewClass15)
        Dim NewClass16 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass16.Name = "FileAssociation"
        NewClass16.Description = "The FileAssociation class stores the file association extension and description. An application can open files on its file association list."
        NewLib.Classes.Add(NewClass16)
        Dim NewClass17 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass17.Name = "Copyright"
        NewClass17.Description = "The Copyright class stores copyright information."
        NewLib.Classes.Add(NewClass17)
        Dim NewClass18 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass18.Name = "License"
        NewClass18.Description = "The License class stores license information."
        NewLib.Classes.Add(NewClass18)
        Dim NewClass19 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass19.Name = "SourceCode"
        NewClass19.Description = "The SourceCode class stores information about the source code for the application."
        NewLib.Classes.Add(NewClass19)
        Dim NewClass20 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass20.Name = "Usage"
        NewClass20.Description = "The Usage class stores information about application or project usage."
        NewLib.Classes.Add(NewClass20)
        Dim NewClass21 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass21.Name = "Trademark"
        NewClass21.Description = "The Trademark class stored information about a trademark used by the author of an application or data."
        NewLib.Classes.Add(NewClass21)

        ApplicationInfo.Libraries.Add(NewLib)

        'Add other library information here: --------------------------------------------------------------------------

    End Sub

    'Save the form settings if the form is being minimised:
    Protected Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = &H112 Then 'SysCommand
            If m.WParam.ToInt32 = &HF020 Then 'Form is being minimised
                SaveFormSettings()
            End If
        End If
        MyBase.WndProc(m)
    End Sub

    Private Sub SaveProjectSettings()
        'Save the project settings in an XML file.
        'Add any Project Settings to be saved into the settingsData XDocument.
        Dim settingsData = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <!--Project settings for ADVL_Coordinates_1 application.-->
                           <ProjectSettings>
                           </ProjectSettings>

        Dim SettingsFileName As String = "ProjectSettings_" & ApplicationInfo.Name & "_" & ".xml"
        Project.SaveXmlSettings(SettingsFileName, settingsData)

    End Sub

    Private Sub RestoreProjectSettings()
        'Restore the project settings from an XML document.

        Dim SettingsFileName As String = "ProjectSettings_" & ApplicationInfo.Name & "_" & ".xml"

        If Project.SettingsFileExists(SettingsFileName) Then
            Dim Settings As System.Xml.Linq.XDocument
            Project.ReadXmlSettings(SettingsFileName, Settings)

            If IsNothing(Settings) Then 'There is no Settings XML data.
                Exit Sub
            End If

            'Restore a Project Setting example:
            If Settings.<ProjectSettings>.<Setting1>.Value = Nothing Then
                'Project setting not saved.
                'Setting1 = ""
            Else
                'Setting1 = Settings.<ProjectSettings>.<Setting1>.Value
            End If

            'Continue restoring saved settings.

        End If

    End Sub

#End Region 'Process XML Files ----------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Display Methods - Code used to display this form." '============================================================================================================================

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Loading the Main form.

        'Set the Application Directory path: ------------------------------------------------
        Project.ApplicationDir = My.Application.Info.DirectoryPath.ToString

        'Read the Application Information file: ---------------------------------------------
        ApplicationInfo.ApplicationDir = My.Application.Info.DirectoryPath.ToString 'Set the Application Directory property
        ''Get the Application Version Information:
        'ApplicationInfo.Version.Major = My.Application.Info.Version.Major
        'ApplicationInfo.Version.Minor = My.Application.Info.Version.Minor
        'ApplicationInfo.Version.Build = My.Application.Info.Version.Build
        'ApplicationInfo.Version.Revision = My.Application.Info.Version.Revision
        'UPDATED VERSION OF THS CODE IS AT THE END OF THE METHOD.

        If ApplicationInfo.ApplicationLocked Then
            MessageBox.Show("The application is locked. If the application is not already in use, remove the 'Application_Info.lock file from the application directory: " & ApplicationInfo.ApplicationDir, "Notice", MessageBoxButtons.OK)
            Dim dr As System.Windows.Forms.DialogResult
            dr = MessageBox.Show("Press 'Yes' to unlock the application", "Notice", MessageBoxButtons.YesNo)
            If dr = System.Windows.Forms.DialogResult.Yes Then
                ApplicationInfo.UnlockApplication()
            Else
                Application.Exit()
                Exit Sub
            End If
        End If

        ReadApplicationInfo()
        ApplicationInfo.LockApplication() 'ALWAYS LOCK THE MESSAGE SERVICE!!! THERE CAN BE ONLY ONE INSTANCE RUNNING ON A SINGLE COMPUTER.
        'THE APPLICATION IS NOW ONLY LOCKED WHEN THE APPLICATION INFO FILE IS BEING UPDATED.

        'Read the Application Usage information: --------------------------------------------
        ApplicationUsage.StartTime = Now
        ApplicationUsage.SaveLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        ApplicationUsage.SaveLocn.Path = Project.ApplicationDir
        ApplicationUsage.RestoreUsageInfo()

        'Restore Project information: -------------------------------------------------------
        Project.Application.Name = ApplicationInfo.Name

        'Set up Message object:
        Message.ApplicationName = ApplicationInfo.Name

        'Set up a temporary initial settings location:
        Dim TempLocn As New ADVL_Utilities_Library_1.FileLocation
        TempLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        TempLocn.Path = ApplicationInfo.ApplicationDir
        Message.SettingsLocn = TempLocn

        Me.Show() 'Show this form before showing the Message form - This will show the App icon on top in the TaskBar.

        'Start showing messages here - Message system is set up.
        'Message.AddText("------------------- Starting Application: ADVL Message Service ---------------------- " & vbCrLf, "Heading")
        Message.AddText("-------------- Starting Application: ADVL Network / Message Service ----------------- " & vbCrLf, "Heading")
        'Message.AddText("Application usage: Total duration = " & Format(ApplicationUsage.TotalDuration.TotalHours, "#.##") & " hours" & vbCrLf, "Normal")
        Dim TotalDuration As String = ApplicationUsage.TotalDuration.Days.ToString.PadLeft(5, "0"c) & "d:" &
                           ApplicationUsage.TotalDuration.Hours.ToString.PadLeft(2, "0"c) & "h:" &
                           ApplicationUsage.TotalDuration.Minutes.ToString.PadLeft(2, "0"c) & "m:" &
                           ApplicationUsage.TotalDuration.Seconds.ToString.PadLeft(2, "0"c) & "s"
        Message.AddText("Application usage: Total duration = " & TotalDuration & vbCrLf, "Normal")


        'https://msdn.microsoft.com/en-us/library/z2d603cy(v=vs.80).aspx#Y550
        'Process any command line arguments:
        Try
            For Each s As String In My.Application.CommandLineArgs
                Message.Add("Command line argument: " & vbCrLf)
                Message.AddXml(s & vbCrLf & vbCrLf)
                InstrReceived = s
            Next
        Catch ex As Exception
            Message.AddWarning("Error processing command line arguments: " & ex.Message & vbCrLf)
        End Try

        If ProjectSelected = False Then
            'Read the Settings Location for the last project used:
            Project.ReadLastProjectInfo()
            'The Last_Project_Info.xml file contains:
            '  Project Name and Description. Settings Location Type and Settings Location Path.
            Message.Add("Last project info has been read." & vbCrLf)
            'Message.Add("Project.Type.ToString  " & Project.Type.ToString & vbCrLf)
            Message.Add("Project type: " & Project.Type.ToString & vbCrLf)
            'Message.Add("Project.Path  " & Project.Path & vbCrLf)
            Message.Add("Project path: " & Project.Path & vbCrLf)

            'At this point read the application start arguments, if any.
            'The selected project may be changed here.

            'Check if the project is locked:
            If Project.ProjectLocked Then
                Message.AddWarning("The project is locked: " & Project.Name & vbCrLf)
                Dim dr As System.Windows.Forms.DialogResult
                dr = MessageBox.Show("Press 'Yes' to unlock the project", "Notice", MessageBoxButtons.YesNo)
                If dr = System.Windows.Forms.DialogResult.Yes Then
                    Project.UnlockProject()
                    Message.AddWarning("The project has been unlocked: " & Project.Name & vbCrLf)
                    'Read the Project Information file: -------------------------------------------------
                    Message.Add("Reading project info." & vbCrLf)
                    Project.ReadProjectInfoFile() 'Read the file in the SettingsLocation: ADVL_Project_Info.xml
                    Project.LockProject() 'Lock the project while it is open in this application.
                    'Set the project start time. This is used to track project usage.
                    Project.Usage.StartTime = Now
                    ApplicationInfo.SettingsLocn = Project.SettingsLocn
                    'Set up the Message object:
                    Message.SettingsLocn = Project.SettingsLocn
                    Message.Show() 'Added 18May19
                Else
                    'Continue without any project selected.
                    Project.Name = ""
                    Project.Type = ADVL_Utilities_Library_1.Project.Types.None
                    Project.Description = ""
                    Project.SettingsLocn.Path = ""
                    Project.DataLocn.Path = ""
                End If
            Else
                'Read the Project Information file: -------------------------------------------------
                Message.Add("Reading project info." & vbCrLf)
                Project.ReadProjectInfoFile()  'Read the file in the SettingsLocation: ADVL_Project_Info.xml
                Project.LockProject() 'Lock the project while it is open in this application.
                'Set the project start time. This is used to track project usage.
                Project.Usage.StartTime = Now
                ApplicationInfo.SettingsLocn = Project.SettingsLocn
                'Set up the Message object:
                Message.SettingsLocn = Project.SettingsLocn
                Message.Show() 'Added 18May19
            End If
        Else
            Project.LockProject() 'Lock the project while it is open in this application.
            ProjectSelected = False 'Reset the Project Selected flag.
        End If

        'START Initialise the form: ===============================================================

        'Set up dgvConnections
        dgvConnections.ColumnHeadersDefaultCellStyle.Font = New Font(dgvConnections.Font, FontStyle.Bold) 'Use bold font for the column headers

        dgvConnections.ColumnCount = 12 'Column added to show connection Status
        dgvConnections.Columns(0).HeaderText = "Project Network Name"
        dgvConnections.Columns(1).HeaderText = "Application Name"
        dgvConnections.Columns(2).HeaderText = "Connection Name"
        dgvConnections.Columns(3).HeaderText = "Project Name"
        dgvConnections.Columns(4).HeaderText = "Project Type"
        dgvConnections.Columns(5).HeaderText = "Project Path"
        dgvConnections.Columns(5).DefaultCellStyle.WrapMode = DataGridViewTriState.True

        dgvConnections.Columns(6).HeaderText = "Get All Warnings"
        dgvConnections.Columns(7).HeaderText = "Get All Messages"
        dgvConnections.Columns(8).HeaderText = "Callback HashCode"
        dgvConnections.Columns(9).HeaderText = "Connection Start Time"
        dgvConnections.Columns(10).HeaderText = "Duration d:h:m:s"
        dgvConnections.Columns(11).HeaderText = "Status"

        dgvConnections.Rows.Clear()
        dgvConnections.AutoResizeColumns()
        dgvConnections.AutoResizeRows()
        dgvConnections.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells

        dgvConnections.AllowUserToAddRows = False 'This stops the last blank row from showing.
        dgvConnections.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgvConnections.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvConnections.AllowUserToResizeColumns = True
        dgvConnections.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        dgvConnections.AutoResizeRows()

        'Set up dgvApplications:
        'Columns in the DataGridView are: Application Name, Description
        dgvApplications.ColumnHeadersDefaultCellStyle.Font = New Font(dgvApplications.Font, FontStyle.Bold) 'Use bold font for the column headers
        dgvApplications.ColumnCount = 2
        dgvApplications.Columns(0).HeaderText = "Application Name"
        dgvApplications.Columns(1).HeaderText = "Description"
        dgvApplications.Columns(1).DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgvApplications.Rows.Clear()
        dgvApplications.AutoResizeColumns()
        dgvApplications.AutoResizeRows()
        dgvApplications.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells

        'Set up dgvProjects:
        dgvProjects.ColumnHeadersDefaultCellStyle.Font = New Font(dgvProjects.Font, FontStyle.Bold) 'Use bold font for the column headers
        dgvProjects.ColumnCount = 6
        'dgvProjects.Columns(0).HeaderText = "Name"
        dgvProjects.Columns(0).HeaderText = "Project Name"
        dgvProjects.Columns(1).HeaderText = "Project Network"
        dgvProjects.Columns(2).HeaderText = "Type"
        dgvProjects.Columns(3).HeaderText = "ID"
        dgvProjects.Columns(4).HeaderText = "Application Name"
        dgvProjects.Columns(5).HeaderText = "Project Description"
        dgvProjects.Rows.Clear()
        dgvProjects.AutoResizeColumns()
        dgvProjects.AutoResizeRows()
        dgvProjects.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells

        'Set up DataGridView2 used to display the contents of a Zip archive:
        'DataGridView2.ColumnCount = 4
        'DataGridView2.ColumnCount = 5 'Adding percent column
        DataGridView2.ColumnCount = 6 'Adding Directory column
        DataGridView2.ColumnHeadersDefaultCellStyle.Font = New Font(DataGridView1.Font, FontStyle.Bold) 'Use bold font for the column headers
        'DataGridView2.Columns(0).HeaderText = "File Name"
        DataGridView2.Columns(0).HeaderText = "Directory"
        DataGridView2.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        'DataGridView2.Columns(1).HeaderText = "Directory"
        DataGridView2.Columns(1).HeaderText = "File Name"
        DataGridView2.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        DataGridView2.Columns(2).HeaderText = "Date Modified"
        DataGridView2.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DataGridView2.Columns(3).HeaderText = "Size"
        DataGridView2.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        'DataGridView2.Columns(4).HeaderText = "Compressed"
        DataGridView2.Columns(4).HeaderText = "Compr"
        DataGridView2.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView2.Columns(5).HeaderText = "%"
        DataGridView2.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView2.Columns(5).DefaultCellStyle.Format = "N2"
        DataGridView2.AllowUserToAddRows = False
        DataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView2.AllowDrop = True

        bgwSendMessage.WorkerReportsProgress = True
        bgwSendMessage.WorkerSupportsCancellation = True

        Me.WebBrowser1.ObjectForScripting = Me

        InitialiseForm() 'Initialise the form for a new project.

        SetUpHost()

        'END   Initialise the form: ---------------------------------------------------------------

        RestoreFormSettings() 'Restore the form settings
        OpenStartPage()
        Message.ShowXMessages = ShowXMessages
        Message.ShowSysMessages = ShowSysMessages
        RestoreProjectSettings() 'Restore the Project settings

        ReadApplicationList() 'The list of all Applications Stored in the Application Directory.
        ReadGlobalProjectList() 'This is the list of all projects.

        ShowProjectInfo() 'Show the project information.

        Message.AddText("------------------- Started OK -------------------------------------------------------------------------- " & vbCrLf & vbCrLf, "Heading")

        ConnectToComNet()

        'Get the Application Version Information:
        If System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed Then
            'Application is network deployed.
            ApplicationInfo.Version.Number = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
            ApplicationInfo.Version.Major = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.Major
            ApplicationInfo.Version.Minor = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.Minor
            ApplicationInfo.Version.Build = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.Build
            ApplicationInfo.Version.Revision = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.Revision
            ApplicationInfo.Version.Source = "Publish"
            Message.Add("Application version: " & System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString & vbCrLf)
        Else
            'Application is not network deployed.
            ApplicationInfo.Version.Number = My.Application.Info.Version.ToString
            ApplicationInfo.Version.Major = My.Application.Info.Version.Major
            ApplicationInfo.Version.Minor = My.Application.Info.Version.Minor
            ApplicationInfo.Version.Build = My.Application.Info.Version.Build
            ApplicationInfo.Version.Revision = My.Application.Info.Version.Revision
            ApplicationInfo.Version.Source = "Assembly"
            'Message.Add("The Application is in Debug mode. The Version information may be incorrect." & vbCrLf)
            Message.Add("Application version: " & My.Application.Info.Version.ToString & vbCrLf)
        End If

    End Sub

    Private Sub InitialiseForm()
        'Initialise the form for a new project.

        'OpenStartPage()

        AppTreeImageList.Images.Clear()
        'AppTreeImageList.TransparentColor = Color.White 'This sets the color to treat as transparent. The AppTree does not support transparent colors. Use white instead of transparent colors in icons.
        trvAppTree.ImageList = AppTreeImageList
        OpenAppTree()

        XmlHtmDisplay1.AllowDrop = True

        XmlHtmDisplay1.WordWrap = False

        XmlHtmDisplay1.Settings.ClearAllTextTypes()
        'Default message display settings:
        XmlHtmDisplay1.Settings.AddNewTextType("Warning")
        XmlHtmDisplay1.Settings.TextType("Warning").FontName = "Arial"
        XmlHtmDisplay1.Settings.TextType("Warning").Bold = True
        XmlHtmDisplay1.Settings.TextType("Warning").Color = Color.Red
        XmlHtmDisplay1.Settings.TextType("Warning").PointSize = 12

        XmlHtmDisplay1.Settings.AddNewTextType("Default")
        XmlHtmDisplay1.Settings.TextType("Default").FontName = "Arial"
        XmlHtmDisplay1.Settings.TextType("Default").Bold = False
        XmlHtmDisplay1.Settings.TextType("Default").Color = Color.Black
        XmlHtmDisplay1.Settings.TextType("Default").PointSize = 10

        XmlHtmDisplay1.Settings.XValue.Bold = True

        XmlHtmDisplay1.Settings.UpdateFontIndexes()
        XmlHtmDisplay1.Settings.UpdateColorIndexes()

        XmlHtmDisplay1.Settings.XmlLargeFileSizeLimit = 100000

        XmlHtmDisplay2.AllowDrop = True

        XmlHtmDisplay2.WordWrap = False

        XmlHtmDisplay2.Settings.ClearAllTextTypes()
        'Default message display settings:
        XmlHtmDisplay2.Settings.AddNewTextType("Warning")
        XmlHtmDisplay2.Settings.TextType("Warning").FontName = "Arial"
        XmlHtmDisplay2.Settings.TextType("Warning").Bold = True
        XmlHtmDisplay2.Settings.TextType("Warning").Color = Color.Red
        XmlHtmDisplay2.Settings.TextType("Warning").PointSize = 12

        XmlHtmDisplay2.Settings.AddNewTextType("Default")
        XmlHtmDisplay2.Settings.TextType("Default").FontName = "Arial"
        XmlHtmDisplay2.Settings.TextType("Default").Bold = False
        XmlHtmDisplay2.Settings.TextType("Default").Color = Color.Black
        XmlHtmDisplay2.Settings.TextType("Default").PointSize = 10

        XmlHtmDisplay2.Settings.XValue.Bold = True

        XmlHtmDisplay2.Settings.UpdateFontIndexes()
        XmlHtmDisplay2.Settings.UpdateColorIndexes()

    End Sub


    Private Sub ShowProjectInfo()
        'Show the project information:

        txtProjectName.Text = Project.Name
        txtProjectDescription.Text = Project.Description
        Select Case Project.Type
            Case ADVL_Utilities_Library_1.Project.Types.Directory
                txtProjectType.Text = "Directory"
            Case ADVL_Utilities_Library_1.Project.Types.Archive
                txtProjectType.Text = "Archive"
            Case ADVL_Utilities_Library_1.Project.Types.Hybrid
                txtProjectType.Text = "Hybrid"
            Case ADVL_Utilities_Library_1.Project.Types.None
                txtProjectType.Text = "None"
        End Select
        txtCreationDate.Text = Format(Project.Usage.FirstUsed, "d-MMM-yyyy H:mm:ss")
        txtLastUsed.Text = Format(Project.Usage.LastUsed, "d-MMM-yyyy H:mm:ss")

        txtProjectPath2.Text = Project.Path

        Select Case Project.SettingsLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtSettingsLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtSettingsLocationType.Text = "Archive"
        End Select
        txtSettingsLocationPath.Text = Project.SettingsLocn.Path

        Select Case Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtDataLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtDataLocationType.Text = "Archive"
        End Select
        txtDataLocationPath.Text = Project.DataLocn.Path

        Select Case Project.SystemLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtSystemLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtSystemLocationType.Text = "Archive"
        End Select
        txtSystemLocationPath.Text = Project.SystemLocn.Path

        'txtTotalDuration.Text = Project.Usage.TotalDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
        '                        Project.Usage.TotalDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
        '                        Project.Usage.TotalDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
        '                        Project.Usage.TotalDuration.Seconds.ToString.PadLeft(2, "0"c)

        'txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
        '                          Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
        '                          Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
        '                          Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c)

        txtTotalDuration.Text = Project.Usage.TotalDuration.Days.ToString.PadLeft(5, "0"c) & "d:" &
                        Project.Usage.TotalDuration.Hours.ToString.PadLeft(2, "0"c) & "h:" &
                        Project.Usage.TotalDuration.Minutes.ToString.PadLeft(2, "0"c) & "m:" &
                        Project.Usage.TotalDuration.Seconds.ToString.PadLeft(2, "0"c) & "s"

        txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & "d:" &
                                  Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & "h:" &
                                  Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & "m:" &
                                  Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c) & "s"


    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Application

        'Check first if there are open connections other than the ADVL_Network_1 connection:
        If dgvConnections.Rows.Count > 1 Then
            Dim Duration As TimeSpan = Now - LastExitAttempt
            If Duration.Seconds < 10 Then
                'This is the second attempt to exit within 10 seconds!
                Dim dr As System.Windows.Forms.DialogResult
                dr = MessageBox.Show("Are you sure you want to exit the application?", "Notice", MessageBoxButtons.YesNo)
                If dr = System.Windows.Forms.DialogResult.No Then
                    LastExitAttempt = Now
                    Exit Sub
                End If
            Else
                'Dont exit because one or more connections are open.
                Beep()
                Message.AddWarning("There are connections still open!" & vbCrLf)
                Message.AddWarning("Close these connections before closing the Message Service." & vbCrLf)
                LastExitAttempt = Now
                Exit Sub
            End If
        ElseIf dgvConnections.Rows.Count = 0 Then 'Last blank row now longer showing: 1 changed to 0
            'OK to exit - there are no connections open - not even ADVL_Network_1
        ElseIf dgvConnections.Rows.Count = 1 Then 'Last blank row now longer showing: 2 changed to 1
            'There is one connection open - check if it is ADVL_Network_1:
            If dgvConnections.Rows(0).Cells(2).Value = "ADVL_Network_1" Then
                'OK to exit - ADVL_Network_1 connection gets closed later.
            Else
                Dim Duration As TimeSpan = Now - LastExitAttempt
                If Duration.Seconds < 10 Then
                    'This is the second attempt to exit within 10 seconds!
                    Dim dr As System.Windows.Forms.DialogResult
                    dr = MessageBox.Show("Are you sure you want to exit the application?", "Notice", MessageBoxButtons.YesNo)
                    If dr = System.Windows.Forms.DialogResult.No Then
                        LastExitAttempt = Now
                        Exit Sub
                    End If
                Else
                    'Dont exit because one or more connections are open.
                    Beep()
                    Message.AddWarning("There are " & dgvConnections.Rows.Count & " connections still open!" & vbCrLf)
                    Message.AddWarning("Close these connections before closing the Message Service." & vbCrLf)
                    LastExitAttempt = Now
                    Exit Sub
                End If
            End If
        End If

        SaveAppTree()

        ApplicationInfo.WriteFile() 'Update the Application Information file.

        Project.SaveLastProjectInfo() 'Save information about the last project used.

        Project.Usage.SaveUsageInfo() 'Save Project usage information.

        Project.UnlockProject() 'Unlock the project.

        ApplicationUsage.SaveUsageInfo() 'Save Application usage information.
        ApplicationInfo.UnlockApplication()

        WriteApplicationListAdvl_2() 'List of Application stored in the Application Directory.

        WriteGlobalProjectListAdvl_2()

        DisconnectFromComNet()
        'myHost.Close() 'This takes a while.
        myHost.Abort()

        Application.Exit()

    End Sub


    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        'Save the form settings if the form state is normal. (A minimised form will have the incorrect size and location.)
        If WindowState = FormWindowState.Normal Then
            SaveFormSettings()
        End If
    End Sub

#End Region 'Form Display Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Open and Close Forms - Code used to open and close other forms." '-------------------------------------------------------------------------------------------------------------------


    Private Sub btnMessages_Click(sender As Object, e As EventArgs) Handles btnMessages.Click
        'Show the Messages form.
        Message.ApplicationName = ApplicationInfo.Name
        Message.SettingsLocn = Project.SettingsLocn
        Message.Show()
        Message.ShowXMessages = ShowXMessages
        Message.MessageForm.BringToFront()
    End Sub

    Private Sub btnWebPages_Click(sender As Object, e As EventArgs) Handles btnWebPages.Click
        'Open the Web Pages form.

        If IsNothing(WebPageList) Then
            WebPageList = New frmWebPageList
            WebPageList.Show()
        Else
            WebPageList.Show()
            WebPageList.BringToFront()
        End If
    End Sub

    Private Sub WebPageList_FormClosed(sender As Object, e As FormClosedEventArgs) Handles WebPageList.FormClosed
        WebPageList = Nothing
    End Sub

    Public Function OpenNewWebPage() As Integer
        'Open a new HTML Web View window, or reuse an existing one if avaiable.
        'The new forms index number in WebViewFormList is returned.

        NewWebPage = New frmWebPage
        If WebPageFormList.Count = 0 Then
            WebPageFormList.Add(NewWebPage)
            WebPageFormList(0).FormNo = 0
            WebPageFormList(0).Show
            Return 0 'The new HTML Display is at position 0 in WebViewFormList()
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To WebPageFormList.Count - 1 'Check if there are closed forms in WebViewFormList. They can be re-used.
                If IsNothing(WebPageFormList(I)) Then
                    WebPageFormList(I) = NewWebPage
                    WebPageFormList(I).FormNo = I
                    WebPageFormList(I).Show
                    FormAdded = True
                    Return I 'The new Html Display is at position I in WebViewFormList()
                    Exit For
                End If
            Next
            If FormAdded = False Then 'Add a new form to WebViewFormList
                Dim FormNo As Integer
                WebPageFormList.Add(NewWebPage)
                FormNo = WebPageFormList.Count - 1
                WebPageFormList(FormNo).FormNo = FormNo
                WebPageFormList(FormNo).Show
                Return FormNo 'The new WebPage is at position FormNo in WebPageFormList()
            End If
        End If
    End Function

    Public Sub WebPageFormClosed()
        'This subroutine is called when the Web Page form has been closed.
        'The subroutine is usually called from the FormClosed event of the WebPage form.
        'The WebPage form may have multiple instances.
        'The ClosedFormNumber property should contains the number of the instance of the WebPage form.
        'This property should be updated by the WebPage form when it is being closed.
        'The ClosedFormNumber property value is used to determine which element in WebPageList should be set to Nothing.

        If WebPageFormList.Count < ClosedFormNo + 1 Then
            'ClosedFormNo is too large to exist in WebPageFormList
            Exit Sub
        End If

        If IsNothing(WebPageFormList(ClosedFormNo)) Then
            'The form is already set to nothing
        Else
            WebPageFormList(ClosedFormNo) = Nothing
        End If
    End Sub

    Public Function OpenNewHtmlDisplayPage() As Integer
        'Open a new HTML display window, or reuse an existing one if avaiable.
        'The new forms index number in HtmlDisplayFormList is returned.

        NewHtmlDisplay = New frmHtmlDisplay
        If HtmlDisplayFormList.Count = 0 Then
            HtmlDisplayFormList.Add(NewHtmlDisplay)
            HtmlDisplayFormList(0).FormNo = 0
            HtmlDisplayFormList(0).Show
            Return 0 'The new HTML Display is at position 0 in HtmlDisplayFormList()
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To HtmlDisplayFormList.Count - 1 'Check if there are closed forms in HtmlDisplayFormList. They can be re-used.
                If IsNothing(HtmlDisplayFormList(I)) Then
                    HtmlDisplayFormList(I) = NewHtmlDisplay
                    HtmlDisplayFormList(I).FormNo = I
                    HtmlDisplayFormList(I).Show
                    FormAdded = True
                    Return I 'The new Html Display is at position I in HtmlDisplayFormList()
                    Exit For
                End If
            Next
            If FormAdded = False Then 'Add a new form to HtmlDisplayFormList
                Dim FormNo As Integer
                HtmlDisplayFormList.Add(NewHtmlDisplay)
                FormNo = HtmlDisplayFormList.Count - 1
                HtmlDisplayFormList(FormNo).FormNo = FormNo
                HtmlDisplayFormList(FormNo).Show
                Return FormNo 'The new HtmlDisplay is at position FormNo in HtmlDisplayFormList()
            End If
        End If
    End Function

    Private Sub btnConnectionTools_Click(sender As Object, e As EventArgs) Handles btnConnectionTools.Click
        'Open the Connection Tools form.

        If IsNothing(ConnectionTools) Then
            ConnectionTools = New frmConnectionTools
            ConnectionTools.Show()
        Else
            ConnectionTools.Show()
            ConnectionTools.BringToFront()
        End If
    End Sub

    Private Sub ConnectionTools_FormClosed(sender As Object, e As FormClosedEventArgs) Handles ConnectionTools.FormClosed
        ConnectionTools = Nothing
    End Sub


#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Methods - The main actions performed by this form." '===========================================================================================================================

    Private Sub SetUpHost()
        'Set Up Host:

        'Code Source:
        'https://msdn.microsoft.com/en-us/library/ms731758(v=vs.110).aspx?cs-save-lang=1&cs-lang=vb#code-snippet-4


        Try
            'Dim baseAddress As Uri = New Uri("http://localhost:8733/ADVLService")
            Dim baseAddress As Uri = New Uri("http://localhost:8734/ADVLService")
            myHost = New ServiceModel.ServiceHost(GetType(MsgService), baseAddress)

            smb = New ServiceMetadataBehavior()
            smb.HttpGetEnabled = True

            smb.MetadataExporter.PolicyVersion = PolicyVersion.Policy15
            myHost.Description.Behaviors.Add(smb)

            Dim binding As New WSDualHttpBinding
            binding.ReceiveTimeout = New TimeSpan(1, 0, 0) '1 hour, 0 minutes, 0 seconds
            binding.OpenTimeout = New TimeSpan(1, 0, 0) '1 hour, 0 minutes, 0 seconds
            binding.SendTimeout = New TimeSpan(1, 0, 0) '1 hour, 0 minutes, 0 seconds
            binding.ReceiveTimeout = New TimeSpan(1, 0, 0) '1 hour, 0 minutes, 0 seconds
            binding.MaxReceivedMessageSize = 2147483647
            binding.MaxBufferPoolSize = 2147483647
            binding.BypassProxyOnLocal = True
            binding.MessageEncoding = WSMessageEncoding.Text
            binding.ReaderQuotas.MaxArrayLength = 2147483647          'Reference to System.Runtime.Serialzation required.
            binding.ReaderQuotas.MaxStringContentLength = 2147483647  'Reference to System.Runtime.Serialzation required.
            binding.ReaderQuotas.MaxBytesPerRead = 2147483647         'Reference to System.Runtime.Serialzation required.
            binding.ReaderQuotas.MaxDepth = 2147483647                'Reference to System.Runtime.Serialzation required.
            binding.ReaderQuotas.MaxNameTableCharCount = 2147483647   'Reference to System.Runtime.Serialzation required.
            binding.ReliableSession.InactivityTimeout = New TimeSpan(1, 0, 0) '1 hour, 0 minutes, 0 seconds

            myHost.AddServiceEndpoint(GetType(IMsgService), binding, baseAddress)

            myHost.Open() 'Additional information: Contract requires Duplex, but Binding 'BasicHttpBinding' doesn't support it or isn't configured properly to support it.
        Catch ex As Exception
            Message.AddWarning("Error setting up the Message Service host:" & vbCrLf & ex.Message & vbCrLf)
        End Try

        'https://stackoverflow.com/questions/6070078/can-i-call-a-method-in-a-self-hosted-wcf-service-locally
        'https://stackoverflow.com/questions/15205337/current-operationcontext-is-null-in-wcf-windows-service/15270541#15270541

    End Sub

#Region " Connect to ComNet - Code used to connect to the Communication Network (Message Service)" '===========================================================================================

    Private Sub ConnectToComNet()
        'Connect to the Message Service. (ComNet)

        client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New HostAppMsgServiceCallback))

        Try
            client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 16) 'Temporarily set the send timeaout to 16 seconds
            client.ConnectAsync("", ApplicationInfo.Name, ConnectionName, Project.Name, Project.Description, Project.Type, Project.Path, False, False)

            ConnectedToComNet = True
            Message.Add("Connected to the Andorville™ Network with Connection Name: []." & ConnectionName & vbCrLf)

            client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour

            bgwComCheck.WorkerReportsProgress = True
            bgwComCheck.WorkerSupportsCancellation = True
            If bgwComCheck.IsBusy Then
                'The ComCheck thread is already running.
            Else
                bgwComCheck.RunWorkerAsync() 'Start the ComCheck thread.
            End If

        Catch ex As System.TimeoutException
            'Message.Add("Timeout error. Check if the Message Service is running." & vbCrLf)
            Message.Add("ConnectToComNet: Timeout error. Check if the Message Service is running." & vbCrLf)
        Catch ex As Exception
            'Message.Add("Error message: " & ex.Message & vbCrLf)
            Message.Add("ConnectToComNet: Error message: " & ex.Message & vbCrLf)
            client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
        End Try

    End Sub

    Private Sub DisconnectFromComNet()
        'Disconnect from the Communication Network (Message Service).

        If IsNothing(client) Then
            Message.Add("Already disconnected from the Message Service." & vbCrLf)
            ConnectedToComNet = False
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted." & vbCrLf)
                ConnectionName = ""
            Else
                Try
                    client.DisconnectAsync("", ConnectionName)
                    ConnectedToComNet = False
                    Message.Add("Disconnected from the Message Service." & vbCrLf)
                    If bgwComCheck.IsBusy Then
                        bgwComCheck.CancelAsync()
                    End If

                Catch ex As Exception
                    Message.AddWarning("Error disconnecting from Message Service: " & ex.Message & vbCrLf)
                End Try
            End If
        End If
    End Sub




#End Region 'Connect to ComNet ----------------------------------------------------------------------------------------------------------------------------------------------------------------


    Public Function ConnectionNameAvailable(ByVal ProNetName As String, ByVal ConnName As String) As Boolean
        'If AppNetName-ConnName is already on the dgvConnections list, the name is not available for a new connection and the function returns False.

        Dim NameFound As Boolean = False
        Dim I As Integer 'Loop index
        For I = 0 To dgvConnections.Rows.Count - 1
            If dgvConnections.Rows(I).Cells(2).Value = ConnName Then
                If dgvConnections.Rows(I).Cells(0).Value = ProNetName Then
                    'The Connection named ConnName has been found in the Project Network named ProNetName. 
                    NameFound = True
                    Exit For
                End If
            End If
        Next

        If NameFound = True Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Sub RemoveConnectionWithName(ByVal ProNetName As String, ByVal ConnName As String)
        'Remove the connection entry from dgvConnections with the Project Network Name = ProNetName and Connection Name = ConnName.

        Dim I As Integer 'Loop index
        For I = 0 To dgvConnections.Rows.Count - 1
            If dgvConnections.Rows(I).Cells(2).Value = ConnName Then
                If dgvConnections.Rows(I).Cells(0).Value = ProNetName Then
                    'The Connection named ConnName has been found in the Project Network named ProNetName.
                    dgvConnections.Rows.Remove(dgvConnections.Rows(I))
                    Exit For
                End If
            End If
        Next

    End Sub

    Private Sub UpdateApplicationGrid()
        'Update dgvApplications with the contents of App.List

        dgvApplications.Rows.Clear()

        Dim NApps As Integer = App.List.Count

        If NApps = 0 Then
            Exit Sub
        End If

        Dim Index As Integer

        For Index = 0 To NApps - 1
            dgvApplications.Rows.Add()
            dgvApplications.Rows(Index).Cells(0).Value = App.List(Index).Name
            dgvApplications.Rows(Index).Cells(1).Value = App.List(Index).Description
        Next

        dgvApplications.AutoResizeColumns()
        dgvApplications.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells 'Just resize the Name column. - The Description column is multi-line.

    End Sub

    Private Sub SaveAppTree()
        'Save the Application Tree.
        'This is named AppTree.Lib
        'This is stored in the Project Data Location.

        Dim decl As New XDeclaration("1.0", "utf-8", "yes")
        Dim XDoc As New XDocument(decl, Nothing)
        XDoc.Add(New XComment(""))
        XDoc.Add(New XComment("Application Tree Information"))

        Dim myAppTree As New XElement("ApplicationTree")

        Dim AppTreeImageListCount As New XElement("NApplicationIcons", AppTreeImageList.Images.Count - NProjectIcons)
        myAppTree.Add(AppTreeImageListCount)
        SaveAppTreeImageList()

        SaveAppTreeNode(myAppTree, "", trvAppTree.Nodes)

        XDoc.Add(myAppTree)

        Project.SaveXmlData("AppTree.Lib", XDoc)

    End Sub

    Private Sub SaveAppTreeImageList()
        'Save all of the images in AppTreeImageList
        Dim NImages As Integer = AppTreeImageList.Images.Count
        Dim I As Integer
        For I = NProjectIcons To NImages - 1
            Try
                Dim imageData As New IO.MemoryStream
                AppTreeImageList.Images(I).Save(imageData, Imaging.ImageFormat.Bmp)
                imageData.Position = 0
                Project.SaveData("AppTreeImage" & I & ".bmp", imageData)
            Catch ex As Exception
                Message.AddWarning("Error saving AppTree image no: " & I & " " & ex.Message & vbCrLf)
            End Try
        Next
    End Sub

    Private Sub SaveAppTreeNode(ByRef myElement As XElement, Parent As String, ByRef tnc As TreeNodeCollection)
        'Save the nodes in the TreeNodeCollection in the XElement.
        'This method calls itself recursively to save all nodes in trvAppTree.

        Dim I As Integer

        If tnc.Count = 0 Then 'Leaf

        Else
            For I = 0 To tnc.Count - 1
                Dim NodeKey As String = tnc(I).Name
                Dim myNode As New XElement(System.Xml.XmlConvert.EncodeName(NodeKey)) 'A space character os not allowed in an XElement name. Replace spaces with &sp characters.
                Dim myNodeText As New XElement("Text", tnc(I).Text)
                myNode.Add(myNodeText)

                If NodeKey = "ADVL_Application_Network_1" Then 'This the root node of the Application Tree.
                    'Save:
                    '  Description
                    '  ExecutablePath
                    '  Directory
                    '  IconNumber
                    '  OpenIconNumber
                    Dim myAppDescr As New XElement("Description", AppInfo(NodeKey).Description)
                    myNode.Add(myAppDescr)
                    Dim myAppExePath As New XElement("ExecutablePath", AppInfo(NodeKey).ExecutablePath)
                    myNode.Add(myAppExePath)
                    Dim myAppDirectory As New XElement("Directory", AppInfo(NodeKey).Directory)
                    myNode.Add(myAppDirectory)
                    Dim myAppIconNumber As New XElement("IconNumber", AppInfo(NodeKey).IconNumber)
                    myNode.Add(myAppIconNumber)
                    Dim myAppOpenIconNumber As New XElement("OpenIconNumber", AppInfo(NodeKey).OpenIconNumber)
                    myNode.Add(myAppOpenIconNumber)

                Else  'Non-root node.
                    If tnc(I).Nodes.Count > 0 Then
                        Dim isExpanded As New XElement("IsExpanded", tnc(I).IsExpanded)
                        myNode.Add(isExpanded)
                    End If

                    If NodeKey.EndsWith(".Proj") Then 'Project Node
                        'Save:
                        '  Name
                        '  Description
                        '  Type
                        '  SettingsLocnType NO LONGER SAVED
                        '  SettingsLocnPath NO LONGER SAVED
                        '  DataLocnType     NO LONGER SAVED
                        '  DataLocnPath     NO LONGER SAVED
                        '  Path
                        '  ID
                        '  ApplicationName
                        '  ApplicationDir   NO LONGER SAVED
                        '  ParentProjectName
                        '  ParentProjectID
                        '  RelativePath  (The project path relative to the Parent Project path)
                        '  IconNumber
                        '  OpenIconNumber

                        Dim myProjName As New XElement("Name", ProjInfo(NodeKey).Name)
                        myNode.Add(myProjName)

                        Dim myProjCreationDate As New XElement("CreationDate", Format(ProjInfo(NodeKey).CreationDate, "d-MMM-yyyy H:mm:ss"))
                        myNode.Add(myProjCreationDate)

                        Dim myProjDescr As New XElement("Description", ProjInfo(NodeKey).Description)
                        myNode.Add(myProjDescr)
                        Dim myProjType As New XElement("Type", ProjInfo(NodeKey).Type)
                        myNode.Add(myProjType)

                        Dim myProjPath As New XElement("Path", ProjInfo(NodeKey).Path)
                        myNode.Add(myProjPath)
                        Dim myProjRelativePath As New XElement("RelativePath", ProjInfo(NodeKey).RelativePath)
                        myNode.Add(myProjRelativePath)

                        Dim myProjID As New XElement("ID", ProjInfo(NodeKey).ID)
                        myNode.Add(myProjID)

                        Dim myProjAppName As New XElement("ApplicationName", ProjInfo(NodeKey).ApplicationName)
                        myNode.Add(myProjAppName)

                        Dim myProjParentProjName As New XElement("ParentProjectName", ProjInfo(NodeKey).ParentProjectName)
                        myNode.Add(myProjParentProjName)
                        Dim myProjParentProjID As New XElement("ParentProjectID", ProjInfo(NodeKey).ParentProjectID)
                        myNode.Add(myProjParentProjID)

                        Dim myProjParentProjPath As New XElement("ParentProjectPath", ProjInfo(NodeKey).ParentProjectPath)
                        myNode.Add(myProjParentProjPath)

                        Dim myProjIconNo As New XElement("IconNumber", ProjInfo(NodeKey).IconNumber)
                        myNode.Add(myProjIconNo)
                        Dim myProjOpenIconNo As New XElement("OpenIconNumber", ProjInfo(NodeKey).OpenIconNumber)
                        myNode.Add(myProjOpenIconNo)

                    Else 'Application Node
                        'Save:
                        '  Description
                        '  ExecutablePath
                        '  Directory
                        '  IconNumber
                        '  OpenIconNumber
                        Dim myAppDescr As New XElement("Description", AppInfo(NodeKey).Description)
                        myNode.Add(myAppDescr)
                        Dim myAppExePath As New XElement("ExecutablePath", AppInfo(NodeKey).ExecutablePath)
                        myNode.Add(myAppExePath)
                        Dim myAppDirectory As New XElement("Directory", AppInfo(NodeKey).Directory)
                        myNode.Add(myAppDirectory)
                        Dim myAppIconNumber As New XElement("IconNumber", AppInfo(NodeKey).IconNumber)
                        myNode.Add(myAppIconNumber)
                        Dim myAppOpenIconNumber As New XElement("OpenIconNumber", AppInfo(NodeKey).OpenIconNumber)
                        myNode.Add(myAppOpenIconNumber)
                    End If
                End If
                SaveAppTreeNode(myNode, tnc(I).Name, tnc(I).Nodes)
                myElement.Add(myNode)
            Next
        End If
    End Sub

    Private Sub OpenAppTree()
        'Open the Application Tree.
        'This is named AppTree.Lib
        'This is stored in the Project Data Location.
        'If the file is not found, trvAppTree is shown with just the Application Network.

        trvAppTree.Nodes.Clear()
        AppInfo.Clear()
        ProjInfo.Clear()

        If Project.DataFileExists("AppTree.Lib") Then
            Dim XTree As XDocument
            Project.ReadXmlData("AppTree.Lib", XTree)

            If XTree.<ApplicationTree>.<NApplicationIcons>.Value = Nothing Then
                NApplicationIcons = 0
            Else
                NApplicationIcons = XTree.<ApplicationTree>.<NApplicationIcons>.Value
                OpenAppTreeImageList()
            End If

            OpenXTree(XTree)
        Else
            LoadProjectIcons()
            'Get the Icon for the Message Service:
            Dim myIcon = System.Drawing.Icon.ExtractAssociatedIcon(Me.ApplicationInfo.ExecutablePath)
            AppTreeImageList.Images.Add(myIcon)
            trvAppTree.ImageList = AppTreeImageList
            trvAppTree.Nodes.Add("ADVL_Message_Service_1", "Message Service", 8, 8) 'Key, Text, ImageIndex, SelectedImageIndex.
            AppInfo.Add("ADVL_Message_Service_1", New clsAppInfo)
            AppInfo("ADVL_Message_Service_1").Description = ApplicationInfo.Description
            AppInfo("ADVL_Message_Service_1").Directory = ApplicationInfo.ApplicationDir
            AppInfo("ADVL_Message_Service_1").ExecutablePath = ApplicationInfo.ExecutablePath
            AppInfo("ADVL_Message_Service_1").IconNumber = 8
            AppInfo("ADVL_Message_Service_1").OpenIconNumber = 8
        End If
    End Sub

    Private Sub LoadProjectIcons()
        'Load the Project icons into AppTreeImageList:

        AppTreeImageList.Images.Clear() 'Clear all existing images in the AppTreeImageList

        AppTreeImageList.Images.Add(ProjectIconImageList.Images(0)) 'Default Project icon
        AppTreeImageList.Images.Add(ProjectIconImageList.Images(1)) 'Open Default Project icon
        AppTreeImageList.Images.Add(ProjectIconImageList.Images(2)) 'Directory Project icon
        AppTreeImageList.Images.Add(ProjectIconImageList.Images(3)) 'Open Directory Project icon
        AppTreeImageList.Images.Add(ProjectIconImageList.Images(4)) 'Archive Project icon
        AppTreeImageList.Images.Add(ProjectIconImageList.Images(5)) 'Open Archive Project icon
        AppTreeImageList.Images.Add(ProjectIconImageList.Images(6)) 'Hybrid Project icon
        AppTreeImageList.Images.Add(ProjectIconImageList.Images(7)) 'Open Hybrid Project icon
    End Sub

    Private Sub OpenAppTreeImageList()
        'Open all of the images in AppTreeImageList

        AppTreeImageList.Images.Clear()

        If NApplicationIcons = 0 Then
            LoadProjectIcons()
            'There are no Application icons to load.
        Else
            LoadProjectIcons()
            Dim I As Integer
            For I = NProjectIcons To NApplicationIcons + NProjectIcons - 1
                Dim imageData As New IO.MemoryStream
                Project.ReadData("AppTreeImage" & I & ".bmp", imageData)
                imageData.Position = 0
                AppTreeImageList.Images.Add(Bitmap.FromStream(imageData))
            Next
        End If
    End Sub

    Private Sub OpenXTree(ByRef XTree As XDocument)
        'Open the Application Tree stored in XTree.

        Dim I As Integer

        'Need to convert the XDocument to an XmlDocument:
        Dim XDoc As New System.Xml.XmlDocument
        XDoc.LoadXml(XTree.ToString)

        ProcessChildNode(XDoc.DocumentElement, trvAppTree.Nodes, "", True)
    End Sub

    Private Sub ProcessChildNode(ByVal xml_Node As System.Xml.XmlNode, ByVal tnc As TreeNodeCollection, ByVal Spaces As String, ByVal ParentNodeIsExpanded As Boolean)
        'Opening the AppTree.Lib file containing the Application Tree.
        'This subroutine calls itself to process the child node branches.

        Dim NodeInfo As System.Xml.XmlNode
        Dim NodeText As String = ""
        Dim NodeKey As String = ""
        Dim IsExpanded As Boolean = True
        Dim HasNodes As Boolean = True

        For Each ChildNode As System.Xml.XmlNode In xml_Node.ChildNodes
            Dim myNodeText As System.Xml.XmlNode
            myNodeText = ChildNode.SelectSingleNode("Text")
            If IsNothing(myNodeText) Then

            Else
                Dim myNodeTextValue As String = myNodeText.InnerText 'This is the text displayed next to the node in the tree view.
                If ChildNode.Name = "ADVL_Application_Network_1" Then 'This the root node of the Application Tree.
                    NodeKey = System.Xml.XmlConvert.DecodeName(ChildNode.Name)
                    If AppInfo.ContainsKey(NodeKey) Then
                        Message.AddWarning("The Application Network node is already listed in the AppInfo dictionary: " & NodeKey & vbCrLf)
                    Else
                        AppInfo.Add(NodeKey, New clsAppInfo) 'Add the App name to the AppInfo dictionary.
                        'Read the App Description:
                        NodeInfo = ChildNode.SelectSingleNode("Description")
                        If IsNothing(NodeInfo) Then
                            AppInfo(NodeKey).Description = ""
                        Else
                            AppInfo(NodeKey).Description = NodeInfo.InnerText
                        End If
                        'Read the App Executable Path:
                        NodeInfo = ChildNode.SelectSingleNode("ExecutablePath")
                        If IsNothing(NodeInfo) Then
                            AppInfo(NodeKey).ExecutablePath = ""
                        Else
                            AppInfo(NodeKey).ExecutablePath = NodeInfo.InnerText
                        End If
                        'Read the App Directory:
                        NodeInfo = ChildNode.SelectSingleNode("Directory")
                        If IsNothing(NodeInfo) Then
                            AppInfo(NodeKey).Directory = ""
                        Else
                            AppInfo(NodeKey).Directory = NodeInfo.InnerText
                        End If
                        'Read the App IconNumber:
                        NodeInfo = ChildNode.SelectSingleNode("IconNumber")
                        If IsNothing(NodeInfo) Then
                            AppInfo(NodeKey).IconNumber = ""
                        Else
                            AppInfo(NodeKey).IconNumber = NodeInfo.InnerText
                        End If
                        'Read the App OpenIconNumber:
                        NodeInfo = ChildNode.SelectSingleNode("OpenIconNumber")
                        If IsNothing(NodeInfo) Then
                            AppInfo(NodeKey).OpenIconNumber = ""
                        Else
                            AppInfo(NodeKey).OpenIconNumber = NodeInfo.InnerText
                        End If
                        'Read Node IsExpanded:
                        NodeInfo = ChildNode.SelectSingleNode("IsExpanded")
                        If IsNothing(NodeInfo) Then
                            IsExpanded = True
                        Else
                            IsExpanded = NodeInfo.InnerText
                        End If

                        Dim new_Node As TreeNode = tnc.Add(NodeKey, myNodeTextValue, AppInfo(NodeKey).IconNumber, AppInfo(NodeKey).OpenIconNumber)

                        ProcessChildNode(ChildNode, new_Node.Nodes, Spaces, IsExpanded)

                        If IsExpanded Then
                            new_Node.Expand()
                        End If

                    End If
                ElseIf ChildNode.Name.EndsWith(".Proj") Then 'Project node.
                    NodeKey = System.Xml.XmlConvert.DecodeName(ChildNode.Name)
                    If ProjInfo.ContainsKey(NodeKey) Then
                        Message.AddWarning("The Project node is already listed in the ProjectInfo dictionary: " & NodeKey & vbCrLf)
                    Else
                        ProjInfo.Add(NodeKey, New clsProjInfo) 'Add the Project Name to the ProjectInfo dictionary.
                        'Read the Project Name:
                        NodeInfo = ChildNode.SelectSingleNode("Name")
                        If IsNothing(NodeInfo) Then
                            ProjInfo(NodeKey).Name = ""
                        Else
                            ProjInfo(NodeKey).Name = NodeInfo.InnerText
                        End If

                        'Read the Project Creation Date:
                        NodeInfo = ChildNode.SelectSingleNode("CreationDate")
                        If IsNothing(NodeInfo) Then
                            ProjInfo(NodeKey).CreationDate = "1-Jan-2000 12:00:00"
                        Else
                            ProjInfo(NodeKey).CreationDate = NodeInfo.InnerText
                        End If

                        'Read the Project Description:
                        NodeInfo = ChildNode.SelectSingleNode("Description")
                        If IsNothing(NodeInfo) Then
                            ProjInfo(NodeKey).Description = ""
                        Else
                            ProjInfo(NodeKey).Description = NodeInfo.InnerText
                        End If
                        'Read the Project Type:
                        NodeInfo = ChildNode.SelectSingleNode("Type")
                        If IsNothing(NodeInfo) Then
                            ProjInfo(NodeKey).Type = ADVL_Utilities_Library_1.Project.Types.None
                        Else
                            Select Case NodeInfo.InnerText
                                Case "None"
                                    ProjInfo(NodeKey).Type = ADVL_Utilities_Library_1.Project.Types.None
                                Case "Directory"
                                    ProjInfo(NodeKey).Type = ADVL_Utilities_Library_1.Project.Types.Directory
                                Case "Archive"
                                    ProjInfo(NodeKey).Type = ADVL_Utilities_Library_1.Project.Types.Archive
                                Case "Hybrid"
                                    ProjInfo(NodeKey).Type = ADVL_Utilities_Library_1.Project.Types.Hybrid
                                Case Else
                                    ProjInfo(NodeKey).Type = ADVL_Utilities_Library_1.Project.Types.None
                            End Select
                        End If
                        'Read the Project Path:
                        NodeInfo = ChildNode.SelectSingleNode("Path")
                        If IsNothing(NodeInfo) Then
                            ProjInfo(NodeKey).Path = ""
                        Else
                            ProjInfo(NodeKey).Path = NodeInfo.InnerText
                        End If

                        'Read the Relative Path (The Project Path relative to the Parent Path.)
                        NodeInfo = ChildNode.SelectSingleNode("RelativePath")
                        If IsNothing(NodeInfo) Then
                            ProjInfo(NodeKey).RelativePath = ""
                        Else
                            ProjInfo(NodeKey).RelativePath = NodeInfo.InnerText
                        End If

                        'Read the Project ID:
                        NodeInfo = ChildNode.SelectSingleNode("ID")
                        If IsNothing(NodeInfo) Then
                            ProjInfo(NodeKey).ID = ""
                        Else
                            ProjInfo(NodeKey).ID = NodeInfo.InnerText
                        End If

                        'Read the Application Name:
                        NodeInfo = ChildNode.SelectSingleNode("ApplicationName")
                        If IsNothing(NodeInfo) Then
                            ProjInfo(NodeKey).ApplicationName = ""
                        Else
                            ProjInfo(NodeKey).ApplicationName = NodeInfo.InnerText
                        End If

                        'Read the Parent Project Name:
                        'Legacy code version: (In case an old file version contains <HostProjectName>)
                        NodeInfo = ChildNode.SelectSingleNode("HostProjectName")
                        If IsNothing(NodeInfo) Then
                            ProjInfo(NodeKey).ParentProjectName = ""
                        Else
                            ProjInfo(NodeKey).ParentProjectName = NodeInfo.InnerText
                        End If
                        'Updated code version:
                        NodeInfo = ChildNode.SelectSingleNode("ParentProjectName")
                        If IsNothing(NodeInfo) Then
                            'ProjInfo(NodeKey).ParentProjectName = "" 'DONT CHANGE THIS - THE CODE ABOVE WILL HAVE SET THE CORRECT VALUE
                        Else
                            ProjInfo(NodeKey).ParentProjectName = NodeInfo.InnerText
                        End If

                        'Read the Parent Project ID:
                        'Legacy code version: (In case an old file version contains <HostProjectID>)
                        NodeInfo = ChildNode.SelectSingleNode("HostProjectID")
                        If IsNothing(NodeInfo) Then
                            ProjInfo(NodeKey).ParentProjectID = ""
                        Else
                            ProjInfo(NodeKey).ParentProjectID = NodeInfo.InnerText
                        End If
                        'Updated code version:
                        NodeInfo = ChildNode.SelectSingleNode("ParentProjectID")
                        If IsNothing(NodeInfo) Then
                            'ProjInfo(NodeKey).ParentProjectID = "" 'DONT CHANGE THIS - THE CODE ABOVE WILL HAVE SET THE CORRECT VALUE
                        Else
                            ProjInfo(NodeKey).ParentProjectID = NodeInfo.InnerText
                        End If

                        'Read the ParentProject Path:
                        NodeInfo = ChildNode.SelectSingleNode("ParentProjectPath")
                        If IsNothing(NodeInfo) Then
                            ProjInfo(NodeKey).ParentProjectPath = ""
                        Else
                            ProjInfo(NodeKey).ParentProjectPath = NodeInfo.InnerText
                        End If

                        'Read the Icon Number
                        NodeInfo = ChildNode.SelectSingleNode("IconNumber")
                        If IsNothing(NodeInfo) Then
                            ProjInfo(NodeKey).IconNumber = ""
                        Else
                            ProjInfo(NodeKey).IconNumber = NodeInfo.InnerText
                        End If
                        'Read the Open Icon Number:
                        NodeInfo = ChildNode.SelectSingleNode("OpenIconNumber")
                        If IsNothing(NodeInfo) Then
                            ProjInfo(NodeKey).OpenIconNumber = ""
                        Else
                            ProjInfo(NodeKey).OpenIconNumber = NodeInfo.InnerText
                        End If

                        'Read Node IsExpanded:
                        NodeInfo = ChildNode.SelectSingleNode("IsExpanded")
                        If IsNothing(NodeInfo) Then
                            IsExpanded = True
                        Else
                            IsExpanded = NodeInfo.InnerText
                        End If

                        Dim new_Node As TreeNode = tnc.Add(NodeKey, myNodeTextValue, ProjInfo(NodeKey).IconNumber, ProjInfo(NodeKey).OpenIconNumber)

                        ProcessChildNode(ChildNode, new_Node.Nodes, Spaces, IsExpanded)

                        If IsExpanded Then
                            new_Node.Expand()
                        End If
                    End If
                Else 'Application node.
                    NodeKey = System.Xml.XmlConvert.DecodeName(ChildNode.Name)
                    If AppInfo.ContainsKey(NodeKey) Then
                        Message.AddWarning("The Application node is already listed in the AppInfo dictionary: " & NodeKey & vbCrLf)
                    Else
                        AppInfo.Add(NodeKey, New clsAppInfo) 'Add the App name to the AppInfo dictionary.
                        'Read the App Description:
                        NodeInfo = ChildNode.SelectSingleNode("Description")
                        If IsNothing(NodeInfo) Then
                            AppInfo(NodeKey).Description = ""
                        Else
                            AppInfo(NodeKey).Description = NodeInfo.InnerText
                        End If
                        'Read the App Executable Path:
                        NodeInfo = ChildNode.SelectSingleNode("ExecutablePath")
                        If IsNothing(NodeInfo) Then
                            AppInfo(NodeKey).ExecutablePath = ""
                        Else
                            AppInfo(NodeKey).ExecutablePath = NodeInfo.InnerText
                        End If
                        'Read the App Directory:
                        NodeInfo = ChildNode.SelectSingleNode("Directory")
                        If IsNothing(NodeInfo) Then
                            AppInfo(NodeKey).Directory = ""
                        Else
                            AppInfo(NodeKey).Directory = NodeInfo.InnerText
                        End If
                        'Read the App IconNumber:
                        NodeInfo = ChildNode.SelectSingleNode("IconNumber")
                        If IsNothing(NodeInfo) Then
                            AppInfo(NodeKey).IconNumber = ""
                        Else
                            AppInfo(NodeKey).IconNumber = NodeInfo.InnerText
                        End If
                        'Read the App OpenIconNumber:
                        NodeInfo = ChildNode.SelectSingleNode("OpenIconNumber")
                        If IsNothing(NodeInfo) Then
                            AppInfo(NodeKey).OpenIconNumber = ""
                        Else
                            AppInfo(NodeKey).OpenIconNumber = NodeInfo.InnerText
                        End If
                        'Read Node IsExpanded:
                        NodeInfo = ChildNode.SelectSingleNode("IsExpanded")
                        If IsNothing(NodeInfo) Then
                            IsExpanded = True
                        Else
                            IsExpanded = NodeInfo.InnerText
                        End If

                        Dim new_Node As TreeNode = tnc.Add(NodeKey, myNodeTextValue, AppInfo(NodeKey).IconNumber, AppInfo(NodeKey).OpenIconNumber)

                        ProcessChildNode(ChildNode, new_Node.Nodes, Spaces, IsExpanded)

                        If IsExpanded Then
                            new_Node.Expand()
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub UpdateWebPage(ByVal FileName As String)
        'Update the web page in WebPageFormList if the Web file name is FileName.

        Dim NPages As Integer = WebPageFormList.Count
        Dim I As Integer

        Try
            For I = 0 To NPages - 1
                If IsNothing(WebPageFormList(I)) Then
                    'Web page has been deleted!
                Else
                    If WebPageFormList(I).FileName = FileName Then
                        WebPageFormList(I).OpenDocument
                    End If
                End If
            Next
        Catch ex As Exception
            Message.AddWarning(ex.Message & vbCrLf)
        End Try

    End Sub


#Region " Start Page Code" '=========================================================================================================================================

    Public Sub OpenStartPage()
        'Open the workflow page:

        If Project.DataFileExists(WorkflowFileName) Then
            'Note: WorkflowFileName should have been restored when the application started.
            DisplayWorkflow()
        ElseIf Project.DataFileExists("StartPage.html") Then
            WorkflowFileName = "StartPage.html"
            DisplayWorkflow()
        Else
            CreateStartPage()
            WorkflowFileName = "StartPage.html"
            DisplayWorkflow()
        End If


        ''Open the StartPage.html file and display in the Start Page tab.

        'If Project.DataFileExists("StartPage.html") Then
        '    'StartPageFileName = "StartPage.html"
        '    WorkflowFileName = "StartPage.html"
        '    DisplayStartPage()
        'Else
        '    CreateStartPage()
        '    'StartPageFileName = "StartPage.html"
        '    WorkflowFileName = "StartPage.html"
        '    DisplayStartPage()
        'End If

    End Sub

    'Public Sub DisplayStartPage()
    '    'Display the StartPage.html file in the Start Page tab.

    '    If Project.DataFileExists(WorkflowFileName) Then
    '        Dim rtbData As New IO.MemoryStream
    '        Project.ReadData(WorkflowFileName, rtbData)
    '        rtbData.Position = 0
    '        Dim sr As New IO.StreamReader(rtbData)
    '        WebBrowser1.DocumentText = sr.ReadToEnd()
    '    Else
    '        Message.AddWarning("Web page file not found: " & WorkflowFileName & vbCrLf)
    '    End If
    'End Sub

    Public Sub DisplayWorkflow()
        'Display the StartPage.html file in the Start Page tab.

        If Project.DataFileExists(WorkflowFileName) Then
            Dim rtbData As New IO.MemoryStream
            Project.ReadData(WorkflowFileName, rtbData)
            rtbData.Position = 0
            Dim sr As New IO.StreamReader(rtbData)
            WebBrowser1.DocumentText = sr.ReadToEnd()
            'WebBrowser2.DocumentText = sr.ReadToEnd()
        Else
            Message.AddWarning("Web page file not found: " & WorkflowFileName & vbCrLf)
        End If
    End Sub

    Private Sub CreateStartPage()
        'Create a new default StartPage.html file.

        Dim htmData As New IO.MemoryStream
        Dim sw As New IO.StreamWriter(htmData)
        sw.Write(AppInfoHtmlString("Application Information")) 'Create a web page providing information about the application.
        sw.Flush()
        Project.SaveData("StartPage.html", htmData)
    End Sub

    Public Function AppInfoHtmlString(ByVal DocumentTitle As String) As String
        'Create an Application Information Web Page.

        'This function should be edited to provide a brief description of the Application.

        Dim sb As New System.Text.StringBuilder

        sb.Append("<!DOCTYPE html>" & vbCrLf)
        sb.Append("<html>" & vbCrLf)
        sb.Append("<head>" & vbCrLf)
        sb.Append("<title>" & DocumentTitle & "</title>" & vbCrLf)
        sb.Append("<meta name=""description"" content=""Application information."">" & vbCrLf)
        sb.Append("</head>" & vbCrLf)

        sb.Append("<body style=""font-family:arial;"">" & vbCrLf & vbCrLf)

        sb.Append("<h2>" & "Andorville&trade; Network" & "</h2>" & vbCrLf & vbCrLf) 'Add the page title.
        sb.Append("<hr>") 'Add a horizontal divider line.
        sb.Append("<p>The Andorville&trade; Network is used by Andorville&trade; applications to exchange information.</p>" & vbCrLf) 'Add an application description.
        sb.Append("<p style=""line-height:1.5;""><b>The application form contains the following tabs:</b><br>" & vbCrLf)
        sb.Append("<font size = ""2.5"">" & vbCrLf)
        sb.Append("<b>Workflow</b> - This page, containing a brief description of the application or a user defined workflow.<br>" & vbCrLf)
        sb.Append("<b>Connections</b> - A list of all the applications connected to the Message Service.<br>" & vbCrLf)
        sb.Append("<b>Application Tree</b> - A tree view of all the applications that have been connected and their associated projects.<br>" & vbCrLf)
        sb.Append("<b>Application List</b> - A list of all the applications that have been connected.<br>" & vbCrLf)
        sb.Append("<b>Project List</b> - A list of all the projects that have been connected.<br>" & vbCrLf)
        sb.Append("<b>Project Information</b> - Information about the selected Message Service project.</p>" & vbCrLf)
        sb.Append("</font>" & vbCrLf)
        sb.Append("<hr>" & vbCrLf & vbCrLf) 'Add a horizontal divider line.

        sb.Append(DefaultJavaScriptString)

        sb.Append("</body>" & vbCrLf)
        sb.Append("</html>" & vbCrLf)

        Return sb.ToString

    End Function

    Public Function DefaultJavaScriptString() As String
        'Generate the default JavaScript section of an Andorville(TM) Workflow Web Page.

        Dim sb As New System.Text.StringBuilder

        'Add JavaScript section:
        sb.Append("<script>" & vbCrLf & vbCrLf)

        'START: User defined JavaScript functions ==========================================================================
        'Add functions to implement the main actions performed by this web page.
        sb.Append("//START: User defined JavaScript functions ==========================================================================" & vbCrLf)
        sb.Append("//  Add functions to implement the main actions performed by this web page." & vbCrLf & vbCrLf)

        sb.Append("//END:   User defined JavaScript functions __________________________________________________________________________" & vbCrLf & vbCrLf & vbCrLf)
        'END:   User defined JavaScript functions --------------------------------------------------------------------------


        'START: User modified JavaScript functions ==========================================================================
        'Modify these function to save all required web page settings and process all expected XMessage instructions.
        sb.Append("//START: User modified JavaScript functions ==========================================================================" & vbCrLf)
        sb.Append("//  Modify these function to save all required web page settings and process all expected XMessage instructions." & vbCrLf & vbCrLf)

        'Add the SaveSettings function - This is used to save web page settings between sessions.
        sb.Append("//Save the web page settings." & vbCrLf)
        sb.Append("function SaveSettings() {" & vbCrLf)
        sb.Append("  var xSettings = ""<Settings>"" + "" \n"" ; //String containing the web page settings in XML format." & vbCrLf)
        sb.Append("  //Add xml lines to save each setting." & vbCrLf & vbCrLf)
        sb.Append("  xSettings +=    ""</Settings>"" + ""\n"" ; //End of the Settings element." & vbCrLf)
        sb.Append(vbCrLf)
        sb.Append("  //Save the settings as an XML file in the project." & vbCrLf)
        sb.Append("  window.external.SaveHtmlSettings(xSettings) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Process a single XMsg instruction (Information:Location pair)
        sb.Append("//Process an XMessage instruction:" & vbCrLf)
        sb.Append("function XMsgInstruction(Info, Locn) {" & vbCrLf)
        sb.Append("  switch(Locn) {" & vbCrLf)
        sb.Append("  //Insert case statements here." & vbCrLf)
        sb.Append("  default:" & vbCrLf)
        sb.Append("    window.external.AddWarning(""Unknown location: "" + Locn + ""\r\n"") ;" & vbCrLf)
        sb.Append("  }" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        sb.Append("//END:   User modified JavaScript functions __________________________________________________________________________" & vbCrLf & vbCrLf & vbCrLf)
        'END:   User modified JavaScript functions --------------------------------------------------------------------------

        'START: Required Document Library Web Page JavaScript functions ==========================================================================
        sb.Append("//START: Required Document Library Web Page JavaScript functions ==========================================================================" & vbCrLf & vbCrLf)

        'Add the AddText function - This sends a message to the message window using a named text type.
        sb.Append("//Add text to the Message window using a named txt type:" & vbCrLf)
        sb.Append("function AddText(Msg, TextType) {" & vbCrLf)
        sb.Append("  window.external.AddText(Msg, TextType) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the AddMessage function - This sends a message to the message window using default black text.
        sb.Append("//Add a message to the Message window using the default black text:" & vbCrLf)
        sb.Append("function AddMessage(Msg) {" & vbCrLf)
        sb.Append("  window.external.AddMessage(Msg) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the AddWarning function - This sends a red, bold warning message to the message window.
        sb.Append("//Add a warning message to the Message window using bold red text:" & vbCrLf)
        sb.Append("function AddWarning(Msg) {" & vbCrLf)
        sb.Append("  window.external.AddWarning(Msg) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the RestoreSettings function - This is used to restore web page settings.
        sb.Append("//Restore the web page settings." & vbCrLf)
        sb.Append("function RestoreSettings() {" & vbCrLf)
        sb.Append("  window.external.RestoreHtmlSettings() " & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'This line runs the RestoreSettings function when the web page is loaded.
        sb.Append("//Restore the web page settings when the page loads." & vbCrLf)
        sb.Append("window.onload = RestoreSettings; " & vbCrLf)
        sb.Append(vbCrLf)

        'Restores a single setting on the web page.
        sb.Append("//Restore a web page setting." & vbCrLf)
        sb.Append("  function RestoreSetting(FormName, ItemName, ItemValue) {" & vbCrLf)
        sb.Append("  document.forms[FormName][ItemName].value = ItemValue ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the RestoreOption function - This is used to add an option to a Select list.
        sb.Append("//Restore a Select control Option." & vbCrLf)
        sb.Append("function RestoreOption(SelectId, OptionText) {" & vbCrLf)
        sb.Append("  var x = document.getElementById(SelectId) ;" & vbCrLf)
        sb.Append("  var option = document.createElement(""Option"") ;" & vbCrLf)
        sb.Append("  option.text = OptionText ;" & vbCrLf)
        sb.Append("  x.add(option) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        sb.Append("//END:   Required Document Library Web Page JavaScript functions __________________________________________________________________________" & vbCrLf & vbCrLf)
        'END:   Required Document Library Web Page JavaScript functions --------------------------------------------------------------------------

        sb.Append("</script>" & vbCrLf & vbCrLf)

        Return sb.ToString

    End Function


    Public Function DefaultHtmlString(ByVal DocumentTitle As String) As String
        'Create a blank HTML Web Page.

        Dim sb As New System.Text.StringBuilder

        sb.Append("<!DOCTYPE html>" & vbCrLf)
        sb.Append("<html>" & vbCrLf)
        sb.Append("<!-- Andorville(TM) Workflow File -->" & vbCrLf)
        sb.Append("<!-- Application Name:    " & ApplicationInfo.Name & " -->" & vbCrLf)
        sb.Append("<!-- Application Version: " & My.Application.Info.Version.ToString & " -->" & vbCrLf)
        sb.Append("<!-- Creation Date:          " & Format(Now, "dd MMMM yyyy") & " -->" & vbCrLf)
        sb.Append("<head>" & vbCrLf)
        sb.Append("<title>" & DocumentTitle & "</title>" & vbCrLf)
        sb.Append("<meta name=""description"" content=""Workflow description."">" & vbCrLf)
        sb.Append("</head>" & vbCrLf)

        sb.Append("<body style=""font-family:arial;"">" & vbCrLf & vbCrLf)

        sb.Append("<h1>" & DocumentTitle & "</h1>" & vbCrLf & vbCrLf)

        sb.Append(DefaultJavaScriptString)

        sb.Append("</body>" & vbCrLf)
        sb.Append("</html>" & vbCrLf)

        Return sb.ToString

    End Function

#End Region 'Start Page ---------------------------------------------------------------------------------------------

    Private Sub ReadApplicationList()
        'Read the Application List.
        '  If the latest format version of the Application List is not present then convert an earlier version to the latest.

        If System.IO.File.Exists(ApplicationInfo.ApplicationDir & "\" & "Global_Application_List_ADVL_2.xml") Then 'Latest format version of the Application List found.
            ReadApplicationListAdvl_2()
        Else 'The Application List was found.
            Message.AddWarning("The Application List Xml document was not found." & vbCrLf)
        End If
    End Sub

    Private Sub ReadApplicationListAdvl_2()
        'Read the Application_List.xml file in the Application Directory. (ADVL_2 format version.)

        If System.IO.File.Exists(ApplicationInfo.ApplicationDir & "\Global_Application_List_ADVL_2.xml") Then
            Dim AppListXDoc As System.Xml.Linq.XDocument
            AppListXDoc = XDocument.Load(ApplicationInfo.ApplicationDir & "\Global_Application_List_ADVL_2.xml")

            Dim Apps = From item In AppListXDoc.<ApplicationList>.<Application>

            App.List.Clear()

            For Each item In Apps
                Dim NewApp As New AppSummary
                NewApp.Name = item.<Name>.Value
                NewApp.Description = item.<Description>.Value
                NewApp.Directory = item.<Directory>.Value
                NewApp.ExecutablePath = item.<ExecutablePath>.Value
                App.List.Add(NewApp)
            Next
            UpdateApplicationGrid()
        End If
    End Sub

    Private Sub WriteApplicationListAdvl_2()
        'Write the Application List in App.List() to the Application_List.xml file in the Application Directory.

        Dim ApplicationListXDoc = <?xml version="1.0" encoding="utf-8"?>
                                  <!---->
                                  <!--Application List File-->
                                  <ApplicationList>
                                      <FormatCode>ADVL_2</FormatCode>
                                      <%= From item In App.List
                                          Select
                                          <Application>
                                              <Name><%= item.Name %></Name>
                                              <Description><%= item.Description %></Description>
                                              <Directory><%= item.Directory %></Directory>
                                              <ExecutablePath><%= item.ExecutablePath %></ExecutablePath>
                                          </Application>
                                      %>
                                  </ApplicationList>

        ApplicationListXDoc.Save(ApplicationInfo.ApplicationDir & "\Global_Application_List_ADVL_2.xml")

    End Sub

    Private Sub ReadGlobalProjectList()
        'Read the Project List. 

        If System.IO.File.Exists(ApplicationInfo.ApplicationDir & "\" & "Global_Project_List_ADVL_2.xml") Then 'Latest format version of the Project List found.
            ReadGlobalProjectListAdvl_2()
        Else 'No versions of the Application List found.
            Message.AddWarning("No versions of the Global Project List Xml document were found." & vbCrLf)
        End If

    End Sub

    Private Sub ReadGlobalProjectListAdvl_2()
        'Read the Global_Project_List_ADVL_2.xml file in the Application Directory. (ADVL_2 format version.)

        If System.IO.File.Exists(ApplicationInfo.ApplicationDir & "\Global_Project_List_ADVL_2.xml") Then
            Dim ProjListXDoc As System.Xml.Linq.XDocument
            ProjListXDoc = XDocument.Load(ApplicationInfo.ApplicationDir & "\Global_Project_List_ADVL_2.xml")

            Dim Projects = From item In ProjListXDoc.<ProjectList>.<Project>

            Proj.List.Clear()

            For Each item In Projects
                Dim NewProj As New ProjSummary
                NewProj.Name = item.<Name>.Value

                If item.<ProNetName>.Value Is Nothing Then
                    'Check if the old AppNetName is used:
                    If item.<AppNetName>.Value Is Nothing Then
                        NewProj.ProNetName = ""
                    Else 'Use the old AppNetName value as the ProNetName:
                        NewProj.ProNetName = item.<AppNetName>.Value
                    End If
                Else
                    NewProj.ProNetName = item.<ProNetName>.Value
                End If

                NewProj.ID = item.<ID>.Value
                Select Case item.<Type>.Value
                    Case "None"
                        NewProj.Type = ADVL_Utilities_Library_1.Project.Types.None
                    Case "Directory"
                        NewProj.Type = ADVL_Utilities_Library_1.Project.Types.Directory
                    Case "Archive"
                        NewProj.Type = ADVL_Utilities_Library_1.Project.Types.Archive
                    Case "Hybrid"
                        NewProj.Type = ADVL_Utilities_Library_1.Project.Types.Hybrid
                    Case Else
                        Message.AddWarning("Unknown project type: " & item.<Type>.Value & vbCrLf)
                End Select
                NewProj.Path = item.<Path>.Value
                NewProj.Description = item.<Description>.Value
                NewProj.ApplicationName = item.<ApplicationName>.Value
                If item.<HostProjectName>.Value <> Nothing Then NewProj.ParentProjectName = item.<HostProjectName>.Value 'Legacy version - in case <HostProjectName> is used.
                If item.<ParentProjectName>.Value <> Nothing Then NewProj.ParentProjectName = item.<ParentProjectName>.Value 'Updated version.
                If item.<HostProjectID>.Value <> Nothing Then NewProj.ParentProjectID = item.<HostProjectID>.Value 'Legacy version - in case <HostProjectID> is used.
                If item.<ParentProjectID>.Value <> Nothing Then NewProj.ParentProjectID = item.<ParentProjectID>.Value 'Updated version.
                Proj.List.Add(NewProj)
            Next
            UpdateProjectGrid()
        End If
    End Sub

    Private Sub UpdateProjectGrid()
        'Update dgvProjects with the contents of App.List

        dgvProjects.Rows.Clear()

        Dim NProjects As Integer = Proj.List.Count

        If NProjects = 0 Then
            Exit Sub
        End If

        Dim Index As Integer

        For Index = 0 To NProjects - 1
            dgvProjects.Rows.Add()
            dgvProjects.Rows(Index).Cells(0).Value = Proj.List(Index).Name
            dgvProjects.Rows(Index).Cells(1).Value = Proj.List(Index).ProNetName
            dgvProjects.Rows(Index).Cells(2).Value = Proj.List(Index).Type
            dgvProjects.Rows(Index).Cells(3).Value = Proj.List(Index).ID
            dgvProjects.Rows(Index).Cells(4).Value = Proj.List(Index).ApplicationName
            dgvProjects.Rows(Index).Cells(5).Value = Proj.List(Index).Description
        Next
        dgvProjects.AutoResizeColumns()
        dgvProjects.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    End Sub

    Private Sub WriteGlobalProjectListAdvl_2()
        'Write the Global Project List in Proj.List() to the Global_Project_List_ADVL_2.xml file in the Application Directory.

        Dim ProjectListXDoc = <?xml version="1.0" encoding="utf-8"?>
                              <!---->
                              <!--Project List File-->
                              <ProjectList>
                                  <FormatCode>ADVL_2</FormatCode>
                                  <%= From item In Proj.List
                                      Select
                                          <Project>
                                              <Name><%= item.Name %></Name>
                                              <ProNetName><%= item.ProNetName %></ProNetName>
                                              <ID><%= item.ID %></ID>
                                              <Type><%= item.Type %></Type>
                                              <Path><%= item.Path %></Path>
                                              <Description><%= item.Description %></Description>
                                              <ApplicationName><%= item.ApplicationName %></ApplicationName>
                                              <ParentProjectName><%= item.ParentProjectName %></ParentProjectName>
                                              <ParentProjectID><%= item.ParentProjectID %></ParentProjectID>
                                          </Project>
                                  %>
                              </ProjectList>

        ProjectListXDoc.Save(ApplicationInfo.ApplicationDir & "\Global_Project_List_ADVL_2.xml")
    End Sub

    Private Sub btnAddDefaultProject_Click(sender As Object, e As EventArgs) Handles btnAddDefaultProject.Click
        'Add the Default project of an application to the Application Tree.

        If txtNodeKey.Text.EndsWith(".Proj") Then
            Message.AddWarning("Select an application node." & vbCrLf)
            Exit Sub
        End If

        Dim DefaultProjectPath As String = txtAppDirectory.Text & "\Default_Project"

        If System.IO.Directory.Exists(DefaultProjectPath) Then
            ProcessNewProject(DefaultProjectPath)
        Else
            'The Default project directory does not exist.
            Message.AddWarning("The Default project directory does not exist." & vbCrLf)
        End If
    End Sub

    Private Sub btnUpdateAppTreeIcon_Click(sender As Object, e As EventArgs) Handles btnUpdateAppTreeIcon.Click
        'Update the Application Icon

        If trvAppTree.SelectedNode Is Nothing Then
            'No node has been selected.
            Message.AddWarning("Please select an application node." & vbCrLf)
        Else
            Dim Node As TreeNode
            Node = trvAppTree.SelectedNode
            Dim NodeName As String = Node.Name
            If NodeName.EndsWith(".Proj") Then
                Message.AddWarning("Select an application node." & vbCrLf)
            Else
                'ShowAppInfo() 'FOR DEBUGGING.
                'ShowAppTreeImageListInfo() 'FOR DEBUGGING.

                'Message.Add("Updating Application Tree Icon:" & vbCrLf)  'FOR DEBUGGING.


                'Delete the application icons:
                If AppInfo(NodeName).IconNumber = AppInfo(NodeName).OpenIconNumber Then
                    'Message.Add("AppInfo(NodeName).IconNumber = AppInfo(NodeName).OpenIconNumber" & vbCrLf) 'FOR DEBUGGING.
                    'Message.Add("(The OpenIcon is the same as the Icon.)" & vbCrLf) 'For debugging.
                    'Delete the OpenIcon (same as Icon)
                    AppTreeImageList.Images.RemoveAt(AppInfo(NodeName).IconNumber) 'Remove the deleted node's icon.
                    'Message.Add("(Icon has been deleted. The Icon number was: " & AppInfo(NodeName).IconNumber & vbCrLf) 'FOR DEBUGGING.
                    'ShowAppInfo() 'Show the AppInfo data after the Icon deletion. 'FOR DEBUGGING.
                    'ShowAppTreeImageListInfo() 'FOR DEBUGGING.
                    Dim I As Integer
                    'Update the icon index numbers in AppInfo()
                    'Message.Add("Updating the Icon numbers." & vbCrLf) 'FOR DEBUGGING.
                    For I = 0 To AppInfo.Count - 1
                        If AppInfo(AppInfo.Keys(I)).IconNumber > AppInfo(NodeName).IconNumber Then
                            AppInfo(AppInfo.Keys(I)).IconNumber -= 1
                        End If
                        If AppInfo(AppInfo.Keys(I)).OpenIconNumber > AppInfo(NodeName).IconNumber Then
                            AppInfo(AppInfo.Keys(I)).OpenIconNumber -= 1
                        End If
                    Next
                    'ShowAppInfo() 'Show the AppInfo data after the Icon deletion.
                    'Update the icon index numbers in ProjectInfo()
                    For I = 0 To ProjInfo.Count - 1
                        If ProjInfo(ProjInfo.Keys(I)).IconNumber > AppInfo(NodeName).IconNumber Then
                            ProjInfo(ProjInfo.Keys(I)).IconNumber -= 1
                        End If
                        If ProjInfo(ProjInfo.Keys(I)).OpenIconNumber > AppInfo(NodeName).IconNumber Then
                            ProjInfo(ProjInfo.Keys(I)).OpenIconNumber -= 1
                        End If
                    Next
                ElseIf AppInfo(NodeName).IconNumber < AppInfo(NodeName).OpenIconNumber Then
                    'Message.Add("AppInfo(NodeName).IconNumber < AppInfo(NodeName).OpenIconNumber" & vbCrLf) 'FOR DEBUGGING.
                    'Delete the OpenIcon first. (Deleting the Icon will change the index numbers of following icons.)
                    AppTreeImageList.Images.RemoveAt(AppInfo(NodeName).OpenIconNumber)
                    AppTreeImageList.Images.RemoveAt(AppInfo(NodeName).IconNumber)
                    'Message.Add("(Icons have been deleted. The OpenIcon number was: " & AppInfo(NodeName).OpenIconNumber & " The Icon number was: " & AppInfo(NodeName).IconNumber & vbCrLf) 'FOR DEBUGGING.
                    'ShowAppInfo() 'Show the AppInfo data after the Icon deletion. 'FOR DEBUGGING.
                    'ShowAppTreeImageListInfo() 'FOR DEBUGGING.

                    'Message.Add("Updating the Icon numbers." & vbCrLf)
                    'Update the icon index numbers in AppInfo()
                    Dim I As Integer
                    Dim Shift As Integer = 0
                    For I = 0 To AppInfo.Count - 1
                        If AppInfo(AppInfo.Keys(I)).IconNumber > AppInfo(NodeName).IconNumber Then
                            Shift += 1
                        End If
                        If AppInfo(AppInfo.Keys(I)).IconNumber > AppInfo(NodeName).OpenIconNumber Then
                            Shift += 1
                        End If
                        AppInfo(AppInfo.Keys(I)).IconNumber -= Shift
                        Shift = 0
                        If AppInfo(AppInfo.Keys(I)).OpenIconNumber > AppInfo(NodeName).IconNumber Then
                            Shift += 1
                        End If
                        If AppInfo(AppInfo.Keys(I)).OpenIconNumber > AppInfo(NodeName).OpenIconNumber Then
                            Shift += 1
                        End If
                        AppInfo(AppInfo.Keys(I)).OpenIconNumber -= Shift
                        Shift = 0
                    Next
                    'ShowAppInfo() 'Show the AppInfo data after the Icon deletion. 'FOR DEBUGGING.
                    'ShowAppTreeImageListInfo() 'FOR DEBUGGING.
                    'Update the icon index numbers in ProjectInfo()
                    For I = 0 To ProjInfo.Count - 1
                        If ProjInfo(ProjInfo.Keys(I)).IconNumber > AppInfo(NodeName).IconNumber Then
                            Shift += 1
                        End If
                        If ProjInfo(ProjInfo.Keys(I)).IconNumber > AppInfo(NodeName).OpenIconNumber Then
                            Shift += 1
                        End If
                        ProjInfo(ProjInfo.Keys(I)).IconNumber -= Shift
                        Shift = 0
                        If ProjInfo(ProjInfo.Keys(I)).OpenIconNumber > AppInfo(NodeName).IconNumber Then
                            Shift += 1
                        End If
                        If ProjInfo(ProjInfo.Keys(I)).OpenIconNumber > AppInfo(NodeName).OpenIconNumber Then
                            Shift += 1
                        End If
                        ProjInfo(ProjInfo.Keys(I)).OpenIconNumber -= Shift
                    Next
                Else
                    'Message.Add("AppInfo(NodeName).IconNumber > AppInfo(NodeName).OpenIconNumber" & vbCrLf) 'FOR DEBUGGING.
                    'Message.Add("Delete the OpenIcon last." & vbCrLf) 'FOR DEBUGGING.
                    'Delete the OpenIcon last.
                    AppTreeImageList.Images.RemoveAt(AppInfo(NodeName).IconNumber)
                    'Message.Add("AppTreeImageList.Images.RemoveAt(AppInfo(NodeName).IconNumber): " & AppInfo(NodeName).IconNumber & vbCrLf) 'FOR DEBUGGING.
                    AppTreeImageList.Images.RemoveAt(AppInfo(NodeName).OpenIconNumber)
                    'Message.Add("AppTreeImageList.Images.RemoveAt(AppInfo(NodeName).OpenIconNumber): " & AppInfo(NodeName).OpenIconNumber & vbCrLf & vbCrLf) 'FOR DEBUGGING.

                    'Update the icon index numbers in AppInfo()
                    'Message.Add("Update the icon index numbers in AppInfo()" & vbCrLf) 'FOR DEBUGGING.
                    Dim I As Integer
                    Dim Shift As Integer = 0
                    For I = 0 To AppInfo.Count - 1
                        If AppInfo(AppInfo.Keys(I)).IconNumber > AppInfo(NodeName).IconNumber Then
                            Shift += 1
                        End If
                        If AppInfo(AppInfo.Keys(I)).IconNumber > AppInfo(NodeName).OpenIconNumber Then
                            Shift += 1
                        End If
                        AppInfo(AppInfo.Keys(I)).IconNumber -= Shift
                        Shift = 0
                        If AppInfo(AppInfo.Keys(I)).OpenIconNumber > AppInfo(NodeName).IconNumber Then
                            Shift += 1
                        End If
                        If AppInfo(AppInfo.Keys(I)).OpenIconNumber > AppInfo(NodeName).OpenIconNumber Then
                            Shift += 1
                        End If
                        AppInfo(I).OpenIconNumber -= Shift
                        Shift = 0
                    Next
                    'Update the icon index numbers in ProjectInfo()
                    For I = 0 To ProjInfo.Count - 1
                        If ProjInfo(ProjInfo.Keys(I)).IconNumber > AppInfo(NodeName).IconNumber Then
                            Shift += 1
                        End If
                        If ProjInfo(ProjInfo.Keys(I)).IconNumber > AppInfo(NodeName).OpenIconNumber Then
                            Shift += 1
                        End If
                        ProjInfo(ProjInfo.Keys(I)).IconNumber -= Shift
                        Shift = 0
                        If ProjInfo(ProjInfo.Keys(I)).OpenIconNumber > AppInfo(NodeName).IconNumber Then
                            Shift += 1
                        End If
                        If ProjInfo(ProjInfo.Keys(I)).OpenIconNumber > AppInfo(NodeName).OpenIconNumber Then
                            Shift += 1
                        End If
                        ProjInfo(ProjInfo.Keys(I)).OpenIconNumber -= Shift
                    Next
                End If

                'Get the new application icon:
                Dim myIcon = System.Drawing.Icon.ExtractAssociatedIcon(AppInfo(NodeName).ExecutablePath)

                AppTreeImageList.Images.Add(NodeName, myIcon) 'DO NOT NEED TO USE AN IMAGE KEY (NodeName).???
                ''AppTreeImageList.TransparentColor = Color.Black

                AppInfo(NodeName).IconNumber = AppTreeImageList.Images.IndexOfKey(NodeName)
                AppInfo(NodeName).OpenIconNumber = AppTreeImageList.Images.IndexOfKey(NodeName)
                'ShowAppTreeImageListInfo() 'FOR DEBUGGING.

                Message.Add("Updating AppTree Image Indexes." & vbCrLf) 'FOR DEBUGGING.

                'https://stackoverflow.com/questions/4520503/how-do-you-get-the-root-node-or-the-first-level-node-of-the-selected-node-in-a-t
                trvAppTree.Nodes(0).EnsureVisible()
                UpdateAppTreeImageIndexes(trvAppTree.TopNode) 'NOTE: THIS WORKS ONLY IF THE FIRST NODE IS FULLY VISIBLE - The TopNode is the first Fully Visible node!

            End If
        End If
    End Sub

    Private Sub ShowAppInfo()
        'Display the contents of AppInfo dictionary:

        Dim HeaderString As String
        Dim ValueString As String
        HeaderString = String.Format("{0,-64} {1,-24} {2,-24}", "Key", "Icon No.", "OpenIcon No.")
        Message.Add(HeaderString & vbCrLf)
        For Each kvp As KeyValuePair(Of String, clsAppInfo) In AppInfo
            ValueString = String.Format("{0,-64} {1,-24} {2,-24}", kvp.Key, kvp.Value.IconNumber, kvp.Value.OpenIconNumber)
            Message.Add(ValueString & vbCrLf)
        Next
        Message.Add(vbCrLf)
    End Sub

    Private Sub ShowAppTreeImageListInfo()
        'Display the contents of AppTreeImageList:

        Message.Add("List of images in AppTreeImageList:" & vbCrLf)
        For Each item In AppTreeImageList.Images.Keys
            Message.Add("Image key: " & item & " Index of Key: " & AppTreeImageList.Images.IndexOfKey(item) & vbCrLf)
        Next

    End Sub

    Private Sub btnProject_Click(sender As Object, e As EventArgs) Handles btnProject.Click
        Project.SelectProject()
    End Sub

    Private Sub btnOpenProject2_Click(sender As Object, e As EventArgs) Handles btnOpenProject2.Click
        If Project.Type = ADVL_Utilities_Library_1.Project.Types.Archive Then
            If IsNothing(ProjectArchive) Then
                ProjectArchive = New frmArchive
                ProjectArchive.Show()
                ProjectArchive.Title = "Project Archive"
                ProjectArchive.Path = Project.Path
            Else
                ProjectArchive.Show()
                ProjectArchive.BringToFront()
            End If
        Else
            Process.Start(Project.Path)
        End If
    End Sub

    Private Sub ProjectArchive_FormClosed(sender As Object, e As FormClosedEventArgs) Handles ProjectArchive.FormClosed
        ProjectArchive = Nothing
    End Sub

    Private Sub btnOpenSettings_Click(sender As Object, e As EventArgs) Handles btnOpenSettings.Click
        If Project.SettingsLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory Then
            Process.Start(Project.SettingsLocn.Path)
        ElseIf Project.SettingsLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Archive Then
            If IsNothing(SettingsArchive) Then
                SettingsArchive = New frmArchive
                SettingsArchive.Show()
                SettingsArchive.Title = "Settings Archive"
                SettingsArchive.Path = Project.SettingsLocn.Path
            Else
                SettingsArchive.Show()
                SettingsArchive.BringToFront()
            End If
        End If
    End Sub

    Private Sub SettingsArchive_FormClosed(sender As Object, e As FormClosedEventArgs) Handles SettingsArchive.FormClosed
        SettingsArchive = Nothing
    End Sub

    Private Sub btnOpenData_Click(sender As Object, e As EventArgs) Handles btnOpenData.Click
        If Project.DataLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory Then
            Process.Start(Project.DataLocn.Path)
        ElseIf Project.DataLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Archive Then
            If IsNothing(DataArchive) Then
                DataArchive = New frmArchive
                DataArchive.Show()
                DataArchive.Title = "Data Archive"
                DataArchive.Path = Project.DataLocn.Path
            Else
                DataArchive.Show()
                DataArchive.BringToFront()
            End If
        End If
    End Sub

    Private Sub DataArchive_FormClosed(sender As Object, e As FormClosedEventArgs) Handles DataArchive.FormClosed
        DataArchive = Nothing
    End Sub

    Private Sub btnOpenSystem_Click(sender As Object, e As EventArgs) Handles btnOpenSystem.Click
        If Project.SystemLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory Then
            Process.Start(Project.SystemLocn.Path)
        ElseIf Project.SystemLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Archive Then
            If IsNothing(SystemArchive) Then
                SystemArchive = New frmArchive
                SystemArchive.Show()
                SystemArchive.Title = "System Archive"
                SystemArchive.Path = Project.SystemLocn.Path
            Else
                SystemArchive.Show()
                SystemArchive.BringToFront()
            End If
        End If
    End Sub

    Private Sub SystemArchive_FormClosed(sender As Object, e As FormClosedEventArgs) Handles SystemArchive.FormClosed
        SystemArchive = Nothing
    End Sub

    Private Sub btnOpenAppDir_Click(sender As Object, e As EventArgs) Handles btnOpenAppDir.Click
        Process.Start(ApplicationInfo.ApplicationDir)
    End Sub

    Private Sub btnParameters_Click(sender As Object, e As EventArgs) Handles btnParameters.Click
        Project.ShowParameters()
    End Sub


#Region " Methods Called by JavaScript - A collection of methods that can be called by JavaScript in a web page shown in WebBrowser1" '========================================================
    'These methods are used to display HTML pages in the Workflow tab.
    'The same methods can be found in the WebView form, which displays web pages on seprate forms.

    'Display Messages ==============================================================================================

    Public Sub AddMessage(ByVal Msg As String)
        'Add a normal text message to the Message window.
        Message.Add(Msg)
    End Sub

    Public Sub AddWarning(ByVal Msg As String)
        'Add a warning text message to the Message window.
        Message.AddWarning(Msg)
    End Sub

    Public Sub AddTextTypeMessage(ByVal Msg As String, ByVal TextType As String)
        'Add a message with the specified Text Type to the Message window.
        Message.AddText(Msg, TextType)
    End Sub

    Public Sub AddXmlMessage(ByVal XmlText As String)
        'Add an Xml message to the Message window.
        Message.AddXml(XmlText)
    End Sub

    'END Display Messages ------------------------------------------------------------------------------------------


    'Run an XSequence ==============================================================================================

    Public Sub RunClipboardXSeq()
        'Run the XSequence instructions in the clipboard.

        Dim XDocSeq As System.Xml.Linq.XDocument
        Try
            XDocSeq = XDocument.Parse(My.Computer.Clipboard.GetText)
        Catch ex As Exception
            Message.AddWarning("Error reading Clipboard data. " & ex.Message & vbCrLf)
            Exit Sub
        End Try

        If IsNothing(XDocSeq) Then
            Message.Add("No XSequence instructions were found in the clipboard.")
        Else
            Dim XmlSeq As New System.Xml.XmlDocument
            Try
                XmlSeq.LoadXml(XDocSeq.ToString) 'Convert XDocSeq to an XmlDocument to process with XSeq.
                'Run the sequence:
                XSeq.RunXSequence(XmlSeq, Status)
            Catch ex As Exception
                Message.AddWarning("Error restoring HTML settings. " & ex.Message & vbCrLf)
            End Try
        End If
    End Sub

    Public Sub RunXSequence(ByVal XSequence As String)
        'Run the XMSequence
        Dim XmlSeq As New System.Xml.XmlDocument
        XmlSeq.LoadXml(XSequence)
        XSeq.RunXSequence(XmlSeq, Status)
    End Sub

    Private Sub XSeq_ErrorMsg(ErrMsg As String) Handles XSeq.ErrorMsg
        Message.AddWarning(ErrMsg & vbCrLf)
    End Sub

    Private Sub XSeq_Instruction(Data As String, Locn As String) Handles XSeq.Instruction
        'Execute each instruction produced by running the XSeq file.

        Select Case Locn

            'Start Project commands: ----------------------------------------------------
            Case "StartProject:AppName"
                StartProject_AppName = Data

            Case "StartProject:ConnectionName"
                StartProject_ConnName = Data

            Case "StartProject:ProNetName"
                StartProject_ProNetName = Data

            Case "StartProject:ProjectID"
                StartProject_ProjID = Data

            Case "StartProject:ProjectName"
                StartProject_ProjName = Data

            Case "StartProject:Command"
                Select Case Data
                    Case "Apply"
                        If StartProject_ProjName <> "" Then
                            StartApp_ProjectName(StartProject_AppName, StartProject_ProNetName, StartProject_ProjName, StartProject_ConnName)
                        ElseIf StartProject_ProjID <> "" Then

                        Else
                            Message.AddWarning("Project not specified. Project Name and Project ID are blank." & vbCrLf)
                        End If
                    Case Else
                        Message.AddWarning("Unknown Start Project command : " & Data & vbCrLf)
                End Select

            'END Start project commands ---------------------------------------------


            Case "Settings"

            Case "EndOfSequence"
                Message.Add("End of processing sequence" & Data & vbCrLf)
                'Clear the StartProject variables:
                StartProject_AppName = ""
                StartProject_ConnName = ""
                StartProject_ProNetName = ""
                StartProject_ProjID = ""
                StartProject_ProjName = ""

            Case Else
                Message.AddWarning("Unknown location: " & Locn & "  Data: " & Data & vbCrLf)

        End Select
    End Sub

    'END Run an XSequence ------------------------------------------------------------------------------------------


    'Run an XMessage ===============================================================================================

    Public Sub RunXMessage(ByVal XMsg As String)
        'Run the XMessage by sending it to InstrReceived.
        InstrReceived = XMsg
    End Sub

    Public Sub SendXMessage(ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMessage to the application with the connection name ConnName.
        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                If bgwSendMessage.IsBusy Then
                    Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                Else
                    Dim SendMessageParams As New clsSendMessageParams
                    'SendMessageParams.ProjectNetworkName = ProNetName
                    SendMessageParams.ProjectNetworkName = "" 'The Network Application is not in any Project Network.
                    SendMessageParams.ConnectionName = ConnName
                    SendMessageParams.Message = XMsg
                    bgwSendMessage.RunWorkerAsync(SendMessageParams)
                    If ShowXMessages Then
                        'Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                        Message.XAddText("Message sent to " & "[" & "" & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                        Message.XAddXml(XMsg)
                        Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub SendXMessageExt(ByVal ProNetName As String, ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMsg to the application with the connection name ConnName and Project Network Name ProNetname.
        'This version can send the XMessage to a connection external to the current Project Network.
        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                If bgwSendMessage.IsBusy Then
                    Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                Else
                    Dim SendMessageParams As New clsSendMessageParams
                    SendMessageParams.ProjectNetworkName = ProNetName
                    SendMessageParams.ConnectionName = ConnName
                    SendMessageParams.Message = XMsg
                    bgwSendMessage.RunWorkerAsync(SendMessageParams)
                    If ShowXMessages Then
                        Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                        Message.XAddXml(XMsg)
                        Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub SendXMessageWait(ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMsg to the application with the connection name ConnName.
        'Wait for the connection to be made.
        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            Try
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    Message.Add("client state is faulted. Message not sent!" & vbCrLf)
                Else
                    Dim StartTime As Date = Now
                    Dim Duration As TimeSpan
                    'Wait up to 16 seconds for the connection ConnName to be established
                    'While client.ConnectionExists(ProNetName, ConnName) = False 'Wait until the required connection is made.
                    While client.ConnectionExists("", ConnName) = False 'Wait until the required connection is made.

                        System.Threading.Thread.Sleep(1000) 'Pause for 1000ms
                        Duration = Now - StartTime
                        If Duration.Seconds > 16 Then Exit While
                    End While

                    'If client.ConnectionExists(ProNetName, ConnName) = False Then
                    If client.ConnectionExists("", ConnName) = False Then
                        'Message.AddWarning("Connection not available: " & ConnName & " in application network: " & ProNetName & vbCrLf)
                        Message.AddWarning("Connection not available: " & ConnName & " in application network: " & "" & vbCrLf)
                    Else
                        If bgwSendMessage.IsBusy Then
                            Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                        Else
                            Dim SendMessageParams As New clsSendMessageParams
                            'SendMessageParams.ProjectNetworkName = ProNetName
                            SendMessageParams.ProjectNetworkName = ""
                            SendMessageParams.ConnectionName = ConnName
                            SendMessageParams.Message = XMsg
                            bgwSendMessage.RunWorkerAsync(SendMessageParams)
                            If ShowXMessages Then
                                'Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                                Message.XAddText("Message sent to " & "[" & "" & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                                Message.XAddXml(XMsg)
                                Message.XAddText(vbCrLf, "Normal") 'Add extra line
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception
                Message.AddWarning(ex.Message & vbCrLf)
            End Try
        End If
    End Sub

    Public Sub SendXMessageExtWait(ByVal ProNetName As String, ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMsg to the application with the connection name ConnName and Project Network Name ProNetName.
        'Wait for the connection to be made.
        'This version can send the XMessage to a connection external to the current Project Network.
        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                Dim StartTime As Date = Now
                Dim Duration As TimeSpan
                'Wait up to 16 seconds for the connection ConnName to be established
                While client.ConnectionExists(ProNetName, ConnName) = False
                    System.Threading.Thread.Sleep(1000) 'Pause for 1000ms
                    Duration = Now - StartTime
                    If Duration.Seconds > 16 Then Exit While
                End While

                If client.ConnectionExists(ProNetName, ConnName) = False Then
                    Message.AddWarning("Connection not available: " & ConnName & " in application network: " & ProNetName & vbCrLf)
                Else
                    If bgwSendMessage.IsBusy Then
                        Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                    Else
                        Dim SendMessageParams As New clsSendMessageParams
                        SendMessageParams.ProjectNetworkName = ProNetName
                        SendMessageParams.ConnectionName = ConnName
                        SendMessageParams.Message = XMsg
                        bgwSendMessage.RunWorkerAsync(SendMessageParams)
                        If ShowXMessages Then
                            Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                            Message.XAddXml(XMsg)
                            Message.XAddText(vbCrLf, "Normal") 'Add extra line
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub XMsgInstruction(ByVal Info As String, ByVal Locn As String)
        'Send the XMessage Instruction to the JavaScript function XMsgInstruction for processing.
        Me.WebBrowser1.Document.InvokeScript("XMsgInstruction", New String() {Info, Locn})
    End Sub

    'END Run an XMessage -------------------------------------------------------------------------------------------


    'Get Information ===============================================================================================

    Public Function GetFormNo() As String
        'Return FormNo.ToString
        Return "-1"
    End Function

    Public Function GetParentFormNo() As String
        'Return the Form Number of the Parent Form (that called this form).
        'Return ParentWebPageFormNo.ToString
        Return "-1" 'The Main Form does not have a Parent Web Page.
    End Function

    Public Function GetConnectionName() As String
        'Return the Connection Name of the Project.
        Return ConnectionName
    End Function

    Public Function GetProNetName() As String
        'Return the Project Network Name of the Project.
        'Return ProNetName
        Return ""
    End Function

    Public Sub ParentProjectName(ByVal FormName As String, ByVal ItemName As String)
        'Return the Parent Project name:
        RestoreSetting(FormName, ItemName, Project.ParentProjectName)
    End Sub

    Public Sub ParentProjectPath(ByVal FormName As String, ByVal ItemName As String)
        'Return the Parent Project path:
        RestoreSetting(FormName, ItemName, Project.ParentProjectPath)
    End Sub

    Public Sub ParentProjectParameterValue(ByVal FormName As String, ByVal ItemName As String, ByVal ParameterName As String)
        'Return the specified Parent Project parameter value:
        RestoreSetting(FormName, ItemName, Project.ParentParameter(ParameterName).Value)
    End Sub

    Public Sub ProjectParameterValue(ByVal FormName As String, ByVal ItemName As String, ByVal ParameterName As String)
        'Return the specified Project parameter value:
        RestoreSetting(FormName, ItemName, Project.Parameter(ParameterName).Value)
    End Sub

    Public Sub ProjectNetworkName(ByVal FormName As String, ByVal ItemName As String)
        'Return the name of the Project Network:
        RestoreSetting(FormName, ItemName, Project.Parameter("ProNetName").Value)
    End Sub

    'END Get Information -------------------------------------------------------------------------------------------


    'Open a Web Page ===============================================================================================

    'Public Sub OpenWebPage(ByVal WebPageFileName As String)
    '    'Open a Web Page from the WebPageFileName.
    '    '  Pass the ParentName Property to the new web page. The is the name of this web page that is opening the new page.
    '    '  Pass the ParentWebPageFormNo Property to the new web page. This is the FormNo of this web page that is opening the new page.
    '    '    A hash code is generated from the ParentName. This is used to define a file name to save and restore the Web Page settings.
    '    '    The new web page can pass instructions back to the ParentWebPage using its ParentWebPageFormNo.

    '    Dim NewFormNo As Integer = OpenNewWebPage()

    '    'WebPageFormList(NewFormNo).ParentWebPageFileName = StartPageFileName 'Set the Parent Web Page property.
    '    WebPageFormList(NewFormNo).ParentWebPageFileName = WorkflowFileName 'Set the Parent Web Page property.
    '    WebPageFormList(NewFormNo).ParentWebPageFormNo = -1 'Set the Parent Form Number property.
    '    WebPageFormList(NewFormNo).Description = ""             'The web page description can be blank.
    '    WebPageFormList(NewFormNo).FileDirectory = ""           'Only Web files in the Project directory can be opened from another Web Page Form.
    '    WebPageFormList(NewFormNo).FileName = WebPageFileName  'Set the web page file name to be opened.
    '    WebPageFormList(NewFormNo).OpenDocument                'Open the web page file name.

    'End Sub

    Public Sub OpenWebPage(ByVal FileName As String)
        'Open the web page with the specified File Name.

        If FileName = "" Then

        Else
            'First check if the HTML file is already open:
            Dim FileFound As Boolean = False
            If WebPageFormList.Count = 0 Then

            Else
                Dim I As Integer
                For I = 0 To WebPageFormList.Count - 1
                    If WebPageFormList(I) Is Nothing Then

                    Else
                        If WebPageFormList(I).FileName = FileName Then
                            FileFound = True
                            WebPageFormList(I).BringToFront
                        End If
                    End If
                Next
            End If

            If FileFound = False Then
                Dim FormNo As Integer = OpenNewWebPage()
                WebPageFormList(FormNo).FileName = FileName
                WebPageFormList(FormNo).OpenDocument
                WebPageFormList(FormNo).BringToFront
            End If
        End If
    End Sub

    'END Open a Web Page -------------------------------------------------------------------------------------------


    'Open and Close Projects =======================================================================================

    Public Sub OpenProjectAtRelativePath(ByVal RelativePath As String, ByVal ConnectionName As String)
        'Open the Project at the specified Relative Path using the specified Connection Name.

        Dim ProjectPath As String
        If RelativePath.StartsWith("\") Then
            ProjectPath = Project.Path & RelativePath
            client.StartProjectAtPath(ProjectPath, ConnectionName)
        Else
            ProjectPath = Project.Path & "\" & RelativePath
            client.StartProjectAtPath(ProjectPath, ConnectionName)
        End If
    End Sub

    Public Sub CheckOpenProjectAtRelativePath(ByVal RelativePath As String, ByVal ConnectionName As String)
        'Check if the project at the specified Relative Path is open.
        'Open it if it is not already open.
        'Open the Project at the specified Relative Path using the specified Connection Name.

        Dim ProjectPath As String
        If RelativePath.StartsWith("\") Then
            ProjectPath = Project.Path & RelativePath
            If client.ProjectOpen(ProjectPath) Then
                'Project is already open.
            Else
                client.StartProjectAtPath(ProjectPath, ConnectionName)
            End If
        Else
            ProjectPath = Project.Path & "\" & RelativePath
            If client.ProjectOpen(ProjectPath) Then
                'Project is already open.
            Else
                client.StartProjectAtPath(ProjectPath, ConnectionName)
            End If
        End If
    End Sub

    Public Sub OpenProjectAtProNetPath(ByVal RelativePath As String, ByVal ConnectionName As String)
        'Open the Project at the specified Path (relative to the Project Network Path) using the specified Connection Name.

        Dim ProjectPath As String
        If RelativePath.StartsWith("\") Then
            If Project.ParameterExists("ProNetPath") Then
                ProjectPath = Project.GetParameter("ProNetPath") & RelativePath
                client.StartProjectAtPath(ProjectPath, ConnectionName)
            Else
                Message.AddWarning("The Project Network Path is not known." & vbCrLf)
            End If
        Else
            If Project.ParameterExists("ProNetPath") Then
                ProjectPath = Project.GetParameter("ProNetPath") & "\" & RelativePath
                client.StartProjectAtPath(ProjectPath, ConnectionName)
            Else
                Message.AddWarning("The Project Network Path is not known." & vbCrLf)
            End If
        End If
    End Sub

    Public Sub CheckOpenProjectAtProNetPath(ByVal RelativePath As String, ByVal ConnectionName As String)
        'Check if the project at the specified Path (relative to the Project Network Path) is open.
        'Open it if it is not already open.
        'Open the Project at the specified Path using the specified Connection Name.

        Dim ProjectPath As String
        If RelativePath.StartsWith("\") Then
            If Project.ParameterExists("ProNetPath") Then
                ProjectPath = Project.GetParameter("ProNetPath") & RelativePath
                'client.StartProjectAtPath(ProjectPath, ConnectionName)
                If client.ProjectOpen(ProjectPath) Then
                    'Project is already open.
                Else
                    client.StartProjectAtPath(ProjectPath, ConnectionName)
                End If
            Else
                Message.AddWarning("The Project Network Path is not known." & vbCrLf)
            End If
        Else
            If Project.ParameterExists("ProNetPath") Then
                ProjectPath = Project.GetParameter("ProNetPath") & "\" & RelativePath
                'client.StartProjectAtPath(ProjectPath, ConnectionName)
                If client.ProjectOpen(ProjectPath) Then
                    'Project is already open.
                Else
                    client.StartProjectAtPath(ProjectPath, ConnectionName)
                End If
            Else
                Message.AddWarning("The Project Network Path is not known." & vbCrLf)
            End If
        End If
    End Sub

    Public Sub CloseProjectAtConnection(ByVal ProNetName As String, ByVal ConnectionName As String)
        'Close the Project at the specified connection.

        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                'Create the XML instructions to close the application at the connection.
                Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class

                'NOTE: No reply expected. No need to provide the following client information(?)
                'Dim clientConnName As New XElement("ClientConnectionName", Me.ConnectionName)
                'xmessage.Add(clientConnName)

                Dim command As New XElement("Command", "Close")
                xmessage.Add(command)

                doc.Add(xmessage)

                'Show the message sent to AppNet:
                Message.XAddText("Message sent to: [" & ProNetName & "]." & ConnectionName & ":" & vbCrLf, "XmlSentNotice")
                Message.XAddXml(doc.ToString)
                Message.XAddText(vbCrLf, "Normal") 'Add extra line

                client.SendMessage(ProNetName, ConnectionName, doc.ToString)
            End If
        End If
    End Sub

    'END Open and Close Projects -----------------------------------------------------------------------------------


    'System Methods ================================================================================================

    Public Sub SaveHtmlSettings(ByVal xSettings As String, ByVal FileName As String)
        'Save the Html settings for a web page.

        'Convert the XSettings to XML format:
        Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
        Dim XDocSettings As New System.Xml.Linq.XDocument

        Try
            XDocSettings = System.Xml.Linq.XDocument.Parse(XmlHeader & vbCrLf & xSettings)
        Catch ex As Exception
            Message.AddWarning("Error saving HTML settings file. " & ex.Message & vbCrLf)
        End Try

        Project.SaveXmlData(FileName, XDocSettings)
    End Sub

    Public Sub RestoreHtmlSettings()
        'Restore the Html settings for a web page.

        Dim SettingsFileName As String = WorkflowFileName & "Settings"

        Dim XDocSettings As New System.Xml.Linq.XDocument
        Project.ReadXmlData(SettingsFileName, XDocSettings)

        If XDocSettings Is Nothing Then
            'Message.Add("No HTML Settings file : " & SettingsFileName & vbCrLf)
        Else
            Dim XSettings As New System.Xml.XmlDocument
            Try
                XSettings.LoadXml(XDocSettings.ToString)
                'Run the Settings file:
                XSeq.RunXSequence(XSettings, Status)
            Catch ex As Exception
                Message.AddWarning("Error restoring HTML settings. " & ex.Message & vbCrLf)
            End Try
        End If
    End Sub

    Public Sub RestoreSetting(ByVal FormName As String, ByVal ItemName As String, ByVal ItemValue As String)
        'Restore the setting value with the specified Form Name and Item Name.
        Me.WebBrowser1.Document.InvokeScript("RestoreSetting", New String() {FormName, ItemName, ItemValue})
    End Sub

    Public Sub RestoreOption(ByVal SelectId As String, ByVal OptionText As String)
        'Restore the Option text in the Select control with the Id SelectId.
        Me.WebBrowser1.Document.InvokeScript("RestoreOption", New String() {SelectId, OptionText})
    End Sub

    Private Sub SaveWebPageSettings()
        'Call the SaveSettings JavaScript function:
        Try
            Me.WebBrowser1.Document.InvokeScript("SaveSettings")
        Catch ex As Exception
            Message.AddWarning("Web page settings not saved: " & ex.Message & vbCrLf)
        End Try
    End Sub

    'END System Methods --------------------------------------------------------------------------------------------


    'Legacy Code (These methods should no longer be used) ==========================================================

    Public Sub JSMethodTest1()
        'Test method that is called from JavaScript.
        Message.Add("JSMethodTest1 called OK." & vbCrLf)
    End Sub

    Public Sub JSMethodTest2(ByVal Var1 As String, ByVal Var2 As String)
        'Test method that is called from JavaScript.
        Message.Add("Var1 = " & Var1 & " Var2 = " & Var2 & vbCrLf)
    End Sub

    Public Sub JSDisplayXml(ByRef XDoc As XDocument)
        Message.Add(XDoc.ToString & vbCrLf & vbCrLf)
    End Sub

    Public Sub ShowMessage(ByVal Msg As String)
        Message.Add(Msg)
    End Sub

    Public Sub AddText(ByVal Msg As String, ByVal TextType As String)
        Message.AddText(Msg, TextType)
    End Sub

    'END Legacy Code -----------------------------------------------------------------------------------------------


#End Region 'Methods Called by JavaScript -----------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Open Project or Application Code" '========================================================================

    Private Sub btnStartApp_Click(sender As Object, e As EventArgs) Handles btnStartApp.Click
        'Start the selected application

        If dgvApplications.SelectedRows.Count > 0 Then 'At least one Application has been selected.
            Dim AppNo As Integer = dgvApplications.SelectedRows(0).Index
            Dim ExePath As String

            Dim AppName As String

            ExePath = App.List(AppNo).ExecutablePath
            AppName = App.List(AppNo).Name

            If System.IO.File.Exists(ExePath) Then
                If chkConnect.Checked = True Then
                    Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                    Dim ConnectDoc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                    Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
                    Dim xConnName As New XElement("ConnectionName", AppName) 'Use the AppName as the Connection Name.
                    xmessage.Add(xConnName)
                    ConnectDoc.Add(xmessage)
                    'Start the application with the argument string containing the instruction to connect to the ComNet
                    Shell(Chr(34) & ExePath & Chr(34) & " " & Chr(34) & ConnectDoc.ToString & Chr(34), AppWinStyle.NormalFocus)
                Else
                    Shell(Chr(34) & ExePath & Chr(34), AppWinStyle.NormalFocus) 'Start the application with no argument
                End If
            Else
                Message.SetWarningStyle()
                Message.Add("Executable file not found: " & vbCrLf)
                Message.Add(ExePath & vbCrLf & vbCrLf)
                Message.SetNormalStyle()
            End If
        Else
            'No Application is selected.
        End If
    End Sub

    Private Function ApplicationNameAvailable(ByVal AppName As String) As Boolean
        'If AppName is not in the Application list, ApplicationNameAvailable is set to True.

        Dim NameFound As Boolean = False
        Dim I As Integer 'Loop index
        For I = 0 To dgvApplications.Rows.Count - 1
            If dgvApplications.Rows(I).Cells(0).Value = AppName Then
                NameFound = True
                ApplicationNo = I 'Save the index number of the Application that has been found.
                Exit For
            End If
        Next

        If NameFound = True Then
            Return False
        Else
            Return True
        End If
    End Function


    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        'Start the selected application or Project.
        StartAppOrProject()
    End Sub

    Private Sub StartAppOrProject()
        Dim AppName As String = trvAppTree.SelectedNode.Name
        If AppName.EndsWith(".Proj") Then
            'Project node.
            Dim StartAppName As String = txtApplicationName.Text
            Dim StartAppConnName As String = txtProjName.Text 'Connect the project using the Project Name as the Connection Name. (This will reduce the likelihood of common connnection names.)
            Dim StartAppProjectPath As String = txtProjPath.Text
            StartApp_ProjectPath(StartAppName, StartAppProjectPath, StartAppConnName)
        Else
            If AppName = "ADVL_Message_Service_1" Then
                Message.AddWarning("The Message Service is already running." & vbCrLf)
            Else
                If AppInfo.ContainsKey(AppName) Then
                    Dim ExePath As String = AppInfo(AppName).ExecutablePath
                    If System.IO.File.Exists(ExePath) Then
                        If chkConnect2.Checked = True Then
                            Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                            Dim ConnectDoc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                            Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
                            Dim xConnName As New XElement("ConnectionName", AppName) 'Use the AppName as the Connection Name.
                            xmessage.Add(xConnName)
                            ConnectDoc.Add(xmessage)
                            'Start the application with the argument string containing the instruction to connect to the ComNet
                            'Shell(Chr(34) & ExePath & Chr(34) & " " & Chr(34) & ConnectDoc.ToString & Chr(34), AppWinStyle.NormalFocus)
                            'Updated version using Process.Start:
                            Process.Start(ExePath, ConnectDoc.ToString)
                            'System.Diagnostics.Process.Start
                        Else
                            'Shell(Chr(34) & ExePath & Chr(34), AppWinStyle.NormalFocus) 'Start the application with no argument.
                            'Updated version using Process.Start:
                            Process.Start(ExePath)
                        End If
                    Else
                        Message.AddWarning("The application: " & AppName & " executable path: " & ExePath & " was not found." & vbCrLf)
                    End If
                Else
                    Message.AddWarning("The application: " & AppName & " was not found in the application list." & vbCrLf)
                End If
            End If
        End If
    End Sub

    Private Sub StartApp_ProjectPath(ByVal AppName As String, ByVal ProjectPath As String, ByVal ConnectionName As String)
        'Start the application with the name AppName.
        'If ProjectPath is not "" then open the specified project.
        'If ConnectionName is not "" then connect to the Application Network.

        'Look for AppName in Application List
        Dim AppInfo As AppSummary = App.FindName(AppName)
        If AppInfo.Name = "" Then
            'AppName not found in the Application List.
            Message.AddWarning("The application named " & AppName & " was not found in the application list." & vbCrLf)
        Else
            'Start the application:
            If ProjectPath = "" And ConnectionName = "" Then
                'No project selected and application will not be connected to the network.
                Shell(Chr(34) & AppInfo.ExecutablePath & Chr(34), AppWinStyle.NormalFocus) 'Start the application with no argument
            Else
                'Build the Application start message:
                Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                Dim ConnectDoc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
                If ProjectPath <> "" Then
                    Dim xproject As New XElement("ProjectPath", ProjectPath)
                    xmessage.Add(xproject)
                End If
                If ConnectionName <> "" Then
                    Dim xconnection As New XElement("ConnectionName", ConnectionName)
                    xmessage.Add(xconnection)
                End If
                ConnectDoc.Add(xmessage)
                Shell(Chr(34) & AppInfo.ExecutablePath & Chr(34) & " " & Chr(34) & ConnectDoc.ToString & Chr(34), AppWinStyle.NormalFocus)
            End If
        End If
    End Sub

    Private Sub StartProject()
        'Start the selected project (or application).

        Dim AppName As String = trvAppTree.SelectedNode.Name
        If AppName.EndsWith(".Proj") Then
            'Project node.
            Dim StartAppName As String = txtApplicationName.Text
            Dim StartAppConnName As String = txtApplicationName.Text
            Dim StartAppProjectPath As String = txtProjPath.Text
            StartApp_ProjectPath(StartAppName, StartAppProjectPath, StartAppConnName)
        Else
            If AppInfo.ContainsKey(AppName) Then
                Dim ExePath As String = AppInfo(AppName).ExecutablePath
                If System.IO.File.Exists(ExePath) Then
                    If chkConnect2.Checked = True Then
                        Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                        Dim ConnectDoc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                        Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
                        Dim xConnName As New XElement("ConnectionName", AppName) 'Use the AppName as the Connection Name.
                        xmessage.Add(xConnName)
                        ConnectDoc.Add(xmessage)
                        'Start the application with the argument string containing the instruction to connect to the ComNet
                        Shell(Chr(34) & ExePath & Chr(34) & " " & Chr(34) & ConnectDoc.ToString & Chr(34), AppWinStyle.NormalFocus)
                    Else
                        Shell(Chr(34) & ExePath & Chr(34), AppWinStyle.NormalFocus) 'Start the application with no argument.
                    End If
                Else
                    Message.AddWarning("The application: " & AppName & " executable path: " & ExePath & " was not found." & vbCrLf)
                End If
            Else
                Message.AddWarning("The application: " & AppName & " was not found in the application list." & vbCrLf)
            End If
        End If
    End Sub

    'Private Sub StartProject(ByVal ProjectName As String, ByVal ProjectNetworkName As String, ByVal AppName As String, ByVal ConnectionName As String)
    Public Sub StartProject(ByVal ProjectName As String, ByVal ProjectNetworkName As String, ByVal AppName As String, ByVal ConnectionName As String)
        'Open the project with the specified Project name, Project Network Name and Application Name with the specified Connection Name.

        'Find a matching project in dgvProjects:
        Dim NProjects As Integer = dgvProjects.RowCount
        Dim I As Integer 'Loop index.

        For I = 0 To NProjects - 1
            If dgvProjects.Rows(I).Cells(0).Value = ProjectName Then
                If dgvProjects.Rows(I).Cells(1).Value = ProjectNetworkName Then
                    If dgvProjects.Rows(I).Cells(4).Value = AppName Then
                        'Project found
                        Dim StartAppName As String = dgvProjects.Rows(I).Cells(4).Value
                        Dim StartAppProjectPath As String = Proj.List(I).Path
                        Dim ProNetName As String = Proj.List(I).ProNetName
                        If StartAppName = "ADVL_Project_Network_1" Then ProNetName = Proj.List(I).Name 'The Project Network Application uses its selected Project Name as the Project Network Name
                        If ConnectionNameAvailable(ProNetName, StartAppConnName) Then
                            StartApp_ProjectPath(StartAppName, StartAppProjectPath, StartAppConnName)
                        Else
                            Message.AddWarning("Connection name: " & StartAppConnName & " already used in the Project Network: " & ProNetName & vbCrLf)
                        End If
                    End If
                End If
            End If
        Next

    End Sub

    Private Sub btnOpenProject_Click(sender As Object, e As EventArgs) Handles btnOpenProject.Click
        'Start the selected project

        If dgvProjects.SelectedRows.Count = 0 Then
            Message.AddWarning("No project has been selected." & vbCrLf)
        ElseIf dgvProjects.SelectedRows.Count = 1 Then
            Dim SelRowNo As Integer = dgvProjects.SelectedRows(0).Index
            Dim StartAppName As String = dgvProjects.Rows(SelRowNo).Cells(4).Value
            Dim StartAppConnName As String = dgvProjects.Rows(SelRowNo).Cells(4).Value 'Use the AppName as the Connection Name. (The connection names can be duplicated as long as the ProNetNames are different.)
            Dim StartAppProjectPath As String = Proj.List(SelRowNo).Path

            Dim ProNetName As String = Proj.List(SelRowNo).ProNetName
            If StartAppName = "ADVL_Project_Network_1" Then ProNetName = Proj.List(SelRowNo).Name

            If ConnectionNameAvailable(ProNetName, StartAppConnName) Then
                StartApp_ProjectPath(StartAppName, StartAppProjectPath, StartAppConnName)
            Else
                Message.AddWarning("Connection name: " & StartAppConnName & " already used in the Project Network: " & ProNetName & vbCrLf)
            End If

        Else 'More than one project selected.
            Message.AddWarning("Two or more projects have been selected. Code to start these will be added later." & vbCrLf)
        End If

    End Sub

    Public Sub StartProject(ByVal ProjectPath As String, ByVal ConnectionName As String)
        'Open the project as the specified Project Path with the specified Connection Name.

        Dim myProject As New ADVL_Utilities_Library_1.Project
        myProject.Path = ProjectPath
        myProject.ReadProjectInfoFile()
        myProject.ReadParameters()


        Dim ApplicationName As String = myProject.Application.Name
        Dim ProNetname As String = myProject.GetParameter("ProNetName")

        If ApplicationName = "" Then
            Message.AddWarning("The project's application is not known." & vbCrLf)
        Else
            If ConnectionNameAvailable(ProNetname, ConnectionName) Then
                StartApp_ProjectPath(ApplicationName, ProjectPath, ConnectionName)
            Else
                Beep()
                Message.AddWarning("The project cannot be started. The connection name: " & ConnectionName & " is already used in the Project Network: " & ProNetname & vbCrLf)
            End If

        End If

    End Sub



#End Region 'Open Project or Application Code -----------------------------------------------------------------------

    Private Sub btnRemoveApp_Click(sender As Object, e As EventArgs) Handles btnRemoveApp.Click
        'Remove the selected application
        If dgvApplications.SelectedRows.Count > 0 Then
            Dim AppNo As Integer = dgvApplications.SelectedRows(0).Index

            App.List.RemoveAt(AppNo)

            UpdateApplicationGrid()
        Else
            'No Application is selected.
        End If
    End Sub

    Private Sub btnRemoveProject_Click(sender As Object, e As EventArgs) Handles btnRemoveProject.Click
        'Remove the selected project
        If dgvProjects.SelectedRows.Count > 0 Then
            Dim ProjNo As Integer = dgvProjects.SelectedRows(0).Index
            Proj.List.RemoveAt(ProjNo)
            UpdateProjectGrid()
        End If

    End Sub

    Private Sub dgvApplications_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvApplications.CellContentClick
        If e.RowIndex > -1 Then
            Dim RowNo As Integer = e.RowIndex
            dgvApplications.Rows(RowNo).Selected = True
            txtDirectory.Text = App.List(RowNo).Directory
            txtExePath.Text = App.List(RowNo).ExecutablePath
        End If
    End Sub

    Private Sub dgvProjects_SelectionChanged(sender As Object, e As EventArgs) Handles dgvProjects.SelectionChanged
        If dgvProjects.SelectedRows.Count > 0 Then
            Dim RowNo As Integer = dgvProjects.SelectedRows(0).Index
            If Proj.List.Count > RowNo Then
                txtProjectPath.Text = Proj.List(RowNo).Path
            End If
        End If
    End Sub

    Private Sub dgvApplications_Resize(sender As Object, e As EventArgs) Handles dgvApplications.Resize
        If dgvApplications.Columns.Count > 3 Then
            Dim DGVerticalScroll = dgvApplications.Controls.OfType(Of VScrollBar).SingleOrDefault.Visible

            If DGVerticalScroll Then
                dgvApplications.Columns(3).Width = dgvApplications.Width - dgvApplications.Columns(0).Width - dgvApplications.Columns(1).Width - dgvApplications.Columns(2).Width - dgvApplications.RowHeadersWidth - 22
                dgvApplications.AutoResizeRows()
            Else
                dgvApplications.Columns(3).Width = dgvApplications.Width - dgvApplications.Columns(0).Width - dgvApplications.Columns(1).Width - dgvApplications.Columns(2).Width - dgvApplications.RowHeadersWidth - 4
                dgvApplications.AutoResizeRows()
            End If
        Else
            'dgvAplications has not been configured with 4 columns yet.
        End If
    End Sub

    Private Sub TabPage3_Enter(sender As Object, e As EventArgs) Handles TabPage3.Enter
        dgvApplications.AutoResizeRows()
    End Sub

    Private Sub trvAppTree_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles trvAppTree.AfterSelect

        txtNodeKey.Text = e.Node.Name
        If e.Node.Name.EndsWith(".Proj") Then
            'Project node
            GroupBox1.Enabled = False 'Disable the Application GroupBox.
            GroupBox2.Enabled = True 'Enable the Project GroupBox.
            txtItemType.Text = "Project"
            txtItemDescription.Text = ProjInfo(e.Node.Name).Description
            txtIconNo.Text = ProjInfo(e.Node.Name).IconNumber
            txtOpenIconNo.Text = ProjInfo(e.Node.Name).OpenIconNumber
            pbIcon.Image = AppTreeImageList.Images(ProjInfo(e.Node.Name).IconNumber)
            pbOpenIcon.Image = AppTreeImageList.Images(ProjInfo(e.Node.Name).OpenIconNumber)

            txtExePath2.Text = ""
            txtAppDirectory.Text = ""

            txtProjName.Text = ProjInfo(e.Node.Name).Name
            txtProjType.Text = ProjInfo(e.Node.Name).Type.ToString
            txtProjPath.Text = ProjInfo(e.Node.Name).Path
            txtProjID.Text = ProjInfo(e.Node.Name).ID
            txtApplicationName.Text = ProjInfo(e.Node.Name).ApplicationName
            txtParentProjectName.Text = ProjInfo(e.Node.Name).ParentProjectName
            txtParentProjectID.Text = ProjInfo(e.Node.Name).ParentProjectID

            btnAddToProjTree.Enabled = True
        Else
            'Application node
            GroupBox1.Enabled = True 'Enable the Application GroupBox.
            GroupBox2.Enabled = False 'Disable the Project GroupBox.
            txtItemType.Text = "Application"
            txtItemDescription.Text = AppInfo(e.Node.Name).Description
            txtIconNo.Text = AppInfo(e.Node.Name).IconNumber
            txtOpenIconNo.Text = AppInfo(e.Node.Name).OpenIconNumber
            pbIcon.Image = AppTreeImageList.Images(AppInfo(e.Node.Name).IconNumber)
            pbOpenIcon.Image = AppTreeImageList.Images(AppInfo(e.Node.Name).OpenIconNumber)
            txtExePath2.Text = AppInfo(e.Node.Name).ExecutablePath
            txtAppDirectory.Text = AppInfo(e.Node.Name).Directory

            txtProjName.Text = ""
            txtProjType.Text = ""
            txtProjPath.Text = ""
            txtProjID.Text = ""
            txtApplicationName.Text = ""
            txtParentProjectName.Text = ""
            txtParentProjectID.Text = ""

            btnAddToProjTree.Enabled = False
        End If
    End Sub

    Private Sub trvAppTree_DoubleClick(sender As Object, e As EventArgs) Handles trvAppTree.DoubleClick
        'The Application Tree has been double clicked.
        'Start the Application or Project:
        StartAppOrProject()
    End Sub

    Private Sub btnDeleteNode_Click(sender As Object, e As EventArgs) Handles btnDeleteNode.Click
        'Delete the selected node.
        If trvAppTree.SelectedNode Is Nothing Then
            'No node has been selected.
        Else
            Dim Node As TreeNode
            Node = trvAppTree.SelectedNode
            Dim NodeName As String = Node.Name
            If Node.Nodes.Count > 0 Then
                Message.AddWarning("The selected node has child nodes. Delete the child nodes before deleting this node." & vbCrLf)
            Else
                If NodeName = "ADVL_Message_Service_1" Then
                    Message.AddWarning("The ADVL_Message_Service_1 node cannot be deleted." & vbCrLf)
                Else
                    If NodeName.EndsWith(".Proj") Then
                        'Deleting a Project node.
                        'Delete the ProjInfo entry:
                        ProjInfo.Remove(NodeName)
                        If Node.Parent Is Nothing Then
                            Node.Remove()
                        Else
                            Dim Parent As TreeNode = Node.Parent
                            Parent.Nodes.RemoveAt(Node.Index)
                        End If
                    Else
                        'Deleting an Application node.
                        'Delete the application icons:
                        If AppInfo(NodeName).IconNumber = AppInfo(NodeName).OpenIconNumber Then
                            'Delete the OpenIcon (same as Icon)
                            AppTreeImageList.Images.RemoveAt(AppInfo(NodeName).IconNumber) 'Remove the deleted node's icon.

                            Dim I As Integer
                            'Update the icon index numbers in AppInfo()
                            For I = 0 To AppInfo.Count - 1
                                If AppInfo(AppInfo.Keys(I)).IconNumber > AppInfo(NodeName).IconNumber Then
                                    AppInfo(AppInfo.Keys(I)).IconNumber -= 1
                                End If
                                If AppInfo(AppInfo.Keys(I)).OpenIconNumber > AppInfo(NodeName).IconNumber Then
                                    AppInfo(AppInfo.Keys(I)).OpenIconNumber -= 1
                                End If
                            Next
                            'Update the icon index numbers in ProjectInfo()
                            For I = 0 To ProjInfo.Count - 1
                                If ProjInfo(ProjInfo.Keys(I)).IconNumber > AppInfo(NodeName).IconNumber Then
                                    ProjInfo(ProjInfo.Keys(I)).IconNumber -= 1
                                End If
                                If ProjInfo(ProjInfo.Keys(I)).OpenIconNumber > AppInfo(NodeName).IconNumber Then
                                    ProjInfo(ProjInfo.Keys(I)).OpenIconNumber -= 1
                                End If
                            Next
                        ElseIf AppInfo(NodeName).IconNumber < AppInfo(NodeName).OpenIconNumber Then
                            'Delete the OpenIcon first. (Deleting the Icon will change the index numbers of following icons.)
                            AppTreeImageList.Images.RemoveAt(AppInfo(NodeName).OpenIconNumber)
                            AppTreeImageList.Images.RemoveAt(AppInfo(NodeName).IconNumber)

                            'Update the icon index numbers in AppInfo()
                            Dim I As Integer
                            Dim Shift As Integer = 0
                            For I = 0 To AppInfo.Count - 1
                                If AppInfo(AppInfo.Keys(I)).IconNumber > AppInfo(NodeName).IconNumber Then
                                    Shift += 1
                                End If
                                If AppInfo(AppInfo.Keys(I)).IconNumber > AppInfo(NodeName).OpenIconNumber Then
                                    Shift += 1
                                End If
                                AppInfo(AppInfo.Keys(I)).IconNumber -= Shift
                                Shift = 0
                                If AppInfo(AppInfo.Keys(I)).OpenIconNumber > AppInfo(NodeName).IconNumber Then
                                    Shift += 1
                                End If
                                If AppInfo(AppInfo.Keys(I)).OpenIconNumber > AppInfo(NodeName).OpenIconNumber Then
                                    Shift += 1
                                End If
                                AppInfo(AppInfo.Keys(I)).OpenIconNumber -= Shift
                                Shift = 0
                            Next
                            'Update the icon index numbers in ProjectInfo()
                            For I = 0 To ProjInfo.Count - 1
                                If ProjInfo(ProjInfo.Keys(I)).IconNumber > AppInfo(NodeName).IconNumber Then
                                    Shift += 1
                                End If
                                If ProjInfo(ProjInfo.Keys(I)).IconNumber > AppInfo(NodeName).OpenIconNumber Then
                                    Shift += 1
                                End If
                                ProjInfo(ProjInfo.Keys(I)).IconNumber -= Shift
                                Shift = 0
                                If ProjInfo(ProjInfo.Keys(I)).OpenIconNumber > AppInfo(NodeName).IconNumber Then
                                    Shift += 1
                                End If
                                If ProjInfo(ProjInfo.Keys(I)).OpenIconNumber > AppInfo(NodeName).OpenIconNumber Then
                                    Shift += 1
                                End If
                                ProjInfo(ProjInfo.Keys(I)).OpenIconNumber -= Shift
                            Next
                        Else
                            'Delete the OpenIcon last.
                            AppTreeImageList.Images.RemoveAt(AppInfo(NodeName).IconNumber)
                            AppTreeImageList.Images.RemoveAt(AppInfo(NodeName).OpenIconNumber)

                            'Update the icon index numbers in AppInfo()
                            Dim I As Integer
                            Dim Shift As Integer = 0
                            For I = 0 To AppInfo.Count - 1
                                If AppInfo(AppInfo.Keys(I)).IconNumber > AppInfo(NodeName).IconNumber Then
                                    Shift += 1
                                End If
                                If AppInfo(AppInfo.Keys(I)).IconNumber > AppInfo(NodeName).OpenIconNumber Then
                                    Shift += 1
                                End If
                                AppInfo(AppInfo.Keys(I)).IconNumber -= Shift
                                Shift = 0
                                If AppInfo(AppInfo.Keys(I)).OpenIconNumber > AppInfo(NodeName).IconNumber Then
                                    Shift += 1
                                End If
                                If AppInfo(AppInfo.Keys(I)).OpenIconNumber > AppInfo(NodeName).OpenIconNumber Then
                                    Shift += 1
                                End If
                                AppInfo(I).OpenIconNumber -= Shift
                                Shift = 0
                            Next
                            'Update the icon index numbers in ProjectInfo()
                            For I = 0 To ProjInfo.Count - 1
                                If ProjInfo(ProjInfo.Keys(I)).IconNumber > AppInfo(NodeName).IconNumber Then
                                    Shift += 1
                                End If
                                If ProjInfo(ProjInfo.Keys(I)).IconNumber > AppInfo(NodeName).OpenIconNumber Then
                                    Shift += 1
                                End If
                                ProjInfo(ProjInfo.Keys(I)).IconNumber -= Shift
                                Shift = 0
                                If ProjInfo(ProjInfo.Keys(I)).OpenIconNumber > AppInfo(NodeName).IconNumber Then
                                    Shift += 1
                                End If
                                If ProjInfo(ProjInfo.Keys(I)).OpenIconNumber > AppInfo(NodeName).OpenIconNumber Then
                                    Shift += 1
                                End If
                                ProjInfo(ProjInfo.Keys(I)).OpenIconNumber -= Shift
                            Next
                        End If

                        'Delete the AppInfo entry:
                        AppInfo.Remove(NodeName)

                        If Node.Parent Is Nothing Then
                            Node.Remove()
                        Else
                            Dim Parent As TreeNode = Node.Parent
                            Parent.Nodes.RemoveAt(Node.Index)
                        End If
                        'UpdateAppTreeImageIndexes(trvAppTree.TopNode) 'This is needed to update the TreeView node icons.
                        UpdateAppTreeImageIndexes(trvAppTree.Nodes(0)) 'This is needed to update the TreeView node icons.
                    End If

                End If
            End If
        End If
    End Sub

    Private Sub UpdateAppTreeImageIndexes(ByRef Node As TreeNode)
        'Update the AppTree images indexes.

        If Node.Name.EndsWith(".Proj") Then
            'Project node - The project icon indexes do not change.
        Else
            'Application node - update the icons.
            Node.ImageIndex = AppInfo(Node.Name).IconNumber
            'Message.Add("Node.ImageIndex = AppInfo(Node.Name).IconNumber: Node.Name = " & Node.Name & "AppInfo(Node.Name).IconNumber = " & AppInfo(Node.Name).IconNumber & vbCrLf)
            Node.SelectedImageIndex = AppInfo(Node.Name).OpenIconNumber
            'Message.Add("Node.SelectedImageIndex = AppInfo(Node.Name).OpenIconNumber: Node.Name = " & Node.Name & "AppInfo(Node.Name).OpenIconNumber = " & AppInfo(Node.Name).OpenIconNumber & vbCrLf & vbCrLf)

            'For Each ChildNode As TreeNode In Node.Nodes
            '    UpdateAppTreeImageIndexes(ChildNode)
            'Next
        End If

        For Each ChildNode As TreeNode In Node.Nodes
            UpdateAppTreeImageIndexes(ChildNode)
        Next

    End Sub

    Private Sub btnMoveUp_Click(sender As Object, e As EventArgs) Handles btnMoveUp.Click
        'Move the selected item up in the Application Tree.

        If trvAppTree.SelectedNode Is Nothing Then
            'No node has been selected.
        Else
            Dim Node As TreeNode
            Node = trvAppTree.SelectedNode
            Dim index As Integer = Node.Index
            If index = 0 Then
                'Already at the first node.
                Node.TreeView.Focus()
            Else
                Dim Parent As TreeNode = Node.Parent
                Parent.Nodes.RemoveAt(index)
                Parent.Nodes.Insert(index - 1, Node)
                trvAppTree.SelectedNode = Node
                Node.TreeView.Focus()
            End If
        End If
    End Sub

    Private Sub btnMoveDown_Click(sender As Object, e As EventArgs) Handles btnMoveDown.Click
        'Move the selected item down in the Application Tree.

        If trvAppTree.SelectedNode Is Nothing Then
            'No node has been selected.
        Else
            Dim Node As TreeNode
            Node = trvAppTree.SelectedNode
            Dim index As Integer = Node.Index
            Dim Parent As TreeNode = Node.Parent
            If index < Parent.Nodes.Count - 1 Then
                Parent.Nodes.RemoveAt(index)
                Parent.Nodes.Insert(index + 1, Node)
                trvAppTree.SelectedNode = Node
                Node.TreeView.Focus()
            Else
                'Already at the last node.
                Node.TreeView.Focus()
            End If
        End If
    End Sub

    Private Sub TabPage2_Enter(sender As Object, e As EventArgs) Handles TabPage2.Enter
        'Update the current duration:

        'txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
        '                          Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
        '                          Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
        '                          Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c)

        txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & "d:" &
                                   Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & "h:" &
                                   Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & "m:" &
                                   Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c) & "s"

        Timer2.Interval = 5000 '5 seconds
        Timer2.Enabled = True
        Timer2.Start()
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        'Update the current duration:

        'txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
        '                          Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
        '                          Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
        '                          Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c)

        txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & "d:" &
                           Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & "h:" &
                           Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & "m:" &
                           Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c) & "s"

    End Sub

    Private Sub TabPage2_Leave(sender As Object, e As EventArgs) Handles TabPage2.Leave
        Timer2.Enabled = False
    End Sub


#Region " Drag Drop Project Code" '==================================================================================

    Private Sub trvAppTree_DragDrop(sender As Object, e As DragEventArgs) Handles trvAppTree.DragDrop
        'DragDrop.

        Dim Path As String()
        Path = e.Data.GetData(DataFormats.FileDrop)

        Message.Add(vbCrLf & "------------------------------------------------------------------------------------------------------------ " & vbCrLf) 'Add separator line.
        Message.Add("Path.Count: " & Path.Count & vbCrLf)

        Dim I As Integer
        For I = 0 To Path.Count - 1
            Message.Add(vbCrLf & "Path(" & I & "): " & Path(I) & vbCrLf)
            ProcessNewProject(Path(I))
        Next
    End Sub

    Private Sub ProcessNewProject(ByVal ProjectPath As String)
        'Process a Project that has been dragged into the Application Tree View:

        'Message.Add(vbCrLf & "Processing Project:" & vbCrLf)
        'Message.Add("Project path: " & ProjectPath & vbCrLf)

        'Check if ProjectPath is a File or a Directory:
        Dim Attr As System.IO.FileAttributes = IO.File.GetAttributes(ProjectPath)
        If Attr.HasFlag(IO.FileAttributes.Directory) Then
            'Message.Add("Project path is a Directory." & vbCrLf)
            If System.IO.File.Exists(ProjectPath & "\Project_Info_ADVL_2.xml") Then
                'Message.Add("The directory is an Andorville(TM) project." & vbCrLf)
                ReadDragDropDirectoryProjectInfo(ProjectPath)
            ElseIf System.IO.File.Exists(ProjectPath & "\ADVL_Project_Info.xml") Then
                Message.Add("The directory is an Andorville(TM) project. (Old ADVL_1 format version.)" & vbCrLf)
                'Convert the ADVL_Project_Info.xml file into a Project_Info_ADVL_2.xml file:
                Dim ProjInfoConversion As New ADVL_Utilities_Library_1.FormatConvert.ProjectInfoFileConversion
                ProjInfoConversion.ProjectType = ADVL_Utilities_Library_1.FormatConvert.ProjectInfoFileConversion.ProjectTypes.Directory
                ProjInfoConversion.ProjectPath = ProjectPath
                ProjInfoConversion.InputFileName = "ADVL_Project_Info.xml"
                ProjInfoConversion.InputFormatCode = ADVL_Utilities_Library_1.FormatConvert.ProjectInfoFileConversion.FormatCodes.ADVL_1
                ProjInfoConversion.OutputFormatCode = ADVL_Utilities_Library_1.FormatConvert.ProjectInfoFileConversion.FormatCodes.ADVL_2
                ProjInfoConversion.Convert()
                If System.IO.File.Exists(ProjectPath & "\Project_Info_ADVL_2.xml") Then
                    ReadDragDropDirectoryProjectInfo(ProjectPath)
                Else
                    Message.AddWarning("The Project Information file could not be converted to the ADVL_2 format version." & vbCrLf)
                End If
            Else
                Message.Add("The directory is not an Andorville(TM) project." & vbCrLf)
            End If
        Else
            Message.Add("Project path is a File." & vbCrLf)
            If ProjectPath.EndsWith(".AdvlProject") Then
                'Message.Add("The file is an Andorville(TM) project." & vbCrLf)
                ReadDragDropArchiveProjectInfo(ProjectPath)
            Else
                Message.Add("The file is not an Andorville(TM) project." & vbCrLf)
            End If
        End If
    End Sub

    Private Sub ReadDragDropDirectoryProjectInfo(ByVal ProjectPath As String)
        'Read the Project Information from a Directory Project.

        Dim ProjectInfo As System.Xml.Linq.XDocument = XDocument.Load(ProjectPath & "\Project_Info_ADVL_2.xml")

        Dim ProjectNetworkName As String
        'If System.IO.File.Exists(ProjectPath & "\ProjectParams_ADVL2.xml") Then
        If System.IO.File.Exists(ProjectPath & "\Project_Params_ADVL_2.xml") Then
            'Dim ParameterInfo As System.Xml.Linq.XDocument = XDocument.Load(ProjectPath & "\ProjectParams_ADVL2.xml")
            Dim ParameterInfo As System.Xml.Linq.XDocument = XDocument.Load(ProjectPath & "\Project_Params_ADVL_2.xml")
            Dim ProNetNames = From names In ParameterInfo.<ProjectParameterList>.<Parameter>
                              Where names.<Name>.Value = "ProNetName"
                              Select names

            If ProNetNames.Count = 0 Then
                ProjectNetworkName = ""
                Message.Add("The Project Parameters file did not contain an ProNetName parameter." & vbCrLf)
            ElseIf ProNetNames.Count = 1 Then
                ProjectNetworkName = ProNetNames(0).<Value>.Value
            Else
                ProjectNetworkName = ProNetNames(0).<Value>.Value
                Message.Add("The Project Parameters file contained more than one ProNetName parameter." & vbCrLf)
            End If
        Else
            ProjectNetworkName = ""
            Message.Add("The project did not contain a Project Parameters file." & vbCrLf)
        End If

        'Message.Add(vbCrLf) 'Add a blank line.

        Dim ProjectName As String
        If ProjectInfo.<Project>.<Name>.Value = Nothing Then
            ProjectName = ""
        Else
            ProjectName = ProjectInfo.<Project>.<Name>.Value
        End If
        'Message.Add("Project Name = " & ProjectName & vbCrLf)

        Dim ProjectID As String
        If ProjectInfo.<Project>.<ID>.Value = Nothing Then
            ProjectID = ""
            Message.AddWarning("The Project ID is blank." & vbCrLf)
        Else
            ProjectID = ProjectInfo.<Project>.<ID>.Value
        End If
        'Message.Add("Project ID = " & ProjectID & vbCrLf)

        Dim ProjectType As String
        If ProjectInfo.<Project>.<Type>.Value = Nothing Then
            ProjectType = ""
        Else
            ProjectType = ProjectInfo.<Project>.<Type>.Value
        End If
        'Message.Add("Project Type = " & ProjectType & vbCrLf)

        'Message.Add("Project Path= " & ProjectPath & vbCrLf)

        Dim ProjectDescription As String
        If ProjectInfo.<Project>.<Description>.Value = Nothing Then
            ProjectDescription = ""
        Else
            ProjectDescription = ProjectInfo.<Project>.<Description>.Value
        End If
        'Message.Add("Project Description = " & ProjectDescription & vbCrLf)

        Dim ApplicationName As String
        If ProjectInfo.<Project>.<Application>.<Name>.Value = Nothing Then
            ApplicationName = ""
        Else
            ApplicationName = ProjectInfo.<Project>.<Application>.<Name>.Value
        End If
        'Message.Add("Application Name = " & ApplicationName & vbCrLf)

        Dim ParentProjectName As String
        'Legacy code version:
        If ProjectInfo.<Project>.<HostProject>.<Name>.Value = Nothing Then
            ParentProjectName = ""
        Else
            ParentProjectName = ProjectInfo.<Project>.<HostProject>.<Name>.Value
        End If

        'Updated code version:
        If ProjectInfo.<Project>.<ParentProject>.<Name>.Value = Nothing Then
            'ParentProjectName = ""  'NO NEED TO CHANGE THIS - THE CODE ABOVE SHOULD HAVE SET THE CORRECT VALUE.
        Else
            ParentProjectName = ProjectInfo.<Project>.<ParentProject>.<Name>.Value
        End If
        'Message.Add("Parent Project Name = " & ParentProjectName & vbCrLf)

        Dim ParentProjectID As String
        'Legacy code version:
        If ProjectInfo.<Project>.<HostProject>.<ID>.Value = Nothing Then
            ParentProjectID = ""
        Else
            ParentProjectID = ProjectInfo.<Project>.<HostProject>.<ID>.Value
        End If

        'Updated code version:
        If ProjectInfo.<Project>.<ParentProject>.<ID>.Value = Nothing Then
            'ParentProjectID = "" 'NO NEED TO CHANGE THIS - THE CODE ABOVE SHOULD HAVE SET THE CORRECT VALUE.
        Else
            ParentProjectID = ProjectInfo.<Project>.<ParentProject>.<ID>.Value
        End If
        'Message.Add("Parent Project ID = " & ParentProjectID & vbCrLf)

        'Add project to the Project List: ---------------------------------------------------
        'This is displayed in the Project List tab.
        If ProjectIdAvailable(ProjectID) Then
            'If ParentProjectID = "" Then
            'The Project is not a Child Project and can be added.
            dgvProjects.Rows.Add()
            Dim CurrentRow As Integer = dgvProjects.Rows.Count - 2
            dgvProjects.Rows(CurrentRow).Cells(0).Value = ProjectName
            dgvProjects.Rows(CurrentRow).Cells(1).Value = ProjectNetworkName
            dgvProjects.Rows(CurrentRow).Cells(2).Value = ProjectType
            dgvProjects.Rows(CurrentRow).Cells(3).Value = ProjectID
            dgvProjects.Rows(CurrentRow).Cells(4).Value = ApplicationName
            dgvProjects.Rows(CurrentRow).Cells(5).Value = ProjectDescription
            dgvProjects.AutoResizeColumns()

            Dim NewProjectInfo As New ProjSummary
            NewProjectInfo.Name = ProjectName
            NewProjectInfo.ProNetName = ProjectNetworkName
            NewProjectInfo.ID = ProjectID
            Select Case ProjectType
                Case "None"
                    NewProjectInfo.Type = ADVL_Utilities_Library_1.Project.Types.None
                Case "Directory"
                    NewProjectInfo.Type = ADVL_Utilities_Library_1.Project.Types.Directory
                Case "Archive"
                    NewProjectInfo.Type = ADVL_Utilities_Library_1.Project.Types.Archive
                Case "Hybrid"
                    NewProjectInfo.Type = ADVL_Utilities_Library_1.Project.Types.Hybrid
                Case Else
                    Message.AddWarning("Unknown project type: " & ProjectType & vbCrLf)
            End Select

            NewProjectInfo.Path = ProjectPath
            NewProjectInfo.Description = ProjectDescription
            NewProjectInfo.ApplicationName = ApplicationName
            NewProjectInfo.ParentProjectName = ParentProjectName
            NewProjectInfo.ParentProjectID = ParentProjectID

            Proj.List.Add(NewProjectInfo)
            Message.Add("Project added the list. Project ID = " & ProjectID & vbCrLf)

        Else
            'Message.Add("The Project is already in the list." & vbCrLf)
        End If


        If ParentProjectID = "" Then
            'Add project to the AppTree -----------------------------------------------------
            'This is displayed in the Applcation Tree tab.
            If ProjInfo.ContainsKey(ProjectID & ".Proj") Then
                'Message.Add("Project is already in the TreeView. Project ID = " & ProjectID & vbCrLf)
            Else
                ProjInfo.Add(ProjectID & ".Proj", New clsProjInfo)
                ProjInfo(ProjectID & ".Proj").Name = ProjectName
                ProjInfo(ProjectID & ".Proj").ProNetName = ProjectNetworkName
                ProjInfo(ProjectID & ".Proj").ID = ProjectID

                ProjInfo(ProjectID & ".Proj").Path = ProjectPath
                ProjInfo(ProjectID & ".Proj").Description = ProjectDescription
                ProjInfo(ProjectID & ".Proj").ApplicationName = ApplicationName
                ProjInfo(ProjectID & ".Proj").ParentProjectName = ParentProjectName
                ProjInfo(ProjectID & ".Proj").ParentProjectID = ParentProjectID

                Select Case ProjectType
                    Case "None"
                        ProjInfo(ProjectID & ".Proj").Type = ADVL_Utilities_Library_1.Project.Types.None
                        ProjInfo(ProjectID & ".Proj").IconNumber = 0
                        ProjInfo(ProjectID & ".Proj").OpenIconNumber = 1
                        Dim node As TreeNode()
                        If ApplicationName = trvAppTree.Nodes(0).Name Then
                            node = trvAppTree.Nodes.Find(ApplicationName, False)
                        Else
                            node = trvAppTree.Nodes(0).Nodes.Find(ApplicationName, False)
                        End If
                        If node Is Nothing Then
                            'Node not found.
                            Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                        ElseIf node.Length = 0 Then
                            'Node not found.
                            Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                        Else
                            trvAppTree.SelectedNode = node(0)
                            trvAppTree.SelectedNode.Nodes.Add(ProjectID & ".Proj", ProjectName, 0, 1) '0, 1 Default project icons.
                        End If
                    Case "Directory"
                        ProjInfo(ProjectID & ".Proj").Type = ADVL_Utilities_Library_1.Project.Types.Directory
                        ProjInfo(ProjectID & ".Proj").IconNumber = 2
                        ProjInfo(ProjectID & ".Proj").OpenIconNumber = 3
                        Dim node As TreeNode()
                        If ApplicationName = trvAppTree.Nodes(0).Name Then
                            node = trvAppTree.Nodes.Find(ApplicationName, False)
                        Else
                            node = trvAppTree.Nodes(0).Nodes.Find(ApplicationName, False)
                        End If
                        If node Is Nothing Then
                            'Node not found.
                            Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                        ElseIf node.Length = 0 Then
                            'Node not found.
                            Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                        Else
                            trvAppTree.SelectedNode = node(0)
                            trvAppTree.SelectedNode.Nodes.Add(ProjectID & ".Proj", ProjectName, 2, 3) '2, 3 Directory project icons.
                        End If
                    Case "Archive"
                        ProjInfo(ProjectID & ".Proj").Type = ADVL_Utilities_Library_1.Project.Types.Archive
                        ProjInfo(ProjectID & ".Proj").IconNumber = 4
                        ProjInfo(ProjectID & ".Proj").OpenIconNumber = 5
                        Dim node As TreeNode()
                        If ApplicationName = trvAppTree.Nodes(0).Name Then
                            node = trvAppTree.Nodes.Find(ApplicationName, False)
                        Else
                            node = trvAppTree.Nodes(0).Nodes.Find(ApplicationName, False)
                        End If
                        If node Is Nothing Then
                            'Node not found.
                            Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                        ElseIf node.Length = 0 Then
                            'Node not found.
                            Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                        Else
                            trvAppTree.SelectedNode = node(0)
                            trvAppTree.SelectedNode.Nodes.Add(ProjectID & ".Proj", ProjectName, 4, 5) '4, 5 Archive project icons.
                        End If
                    Case "Hybrid"
                        ProjInfo(ProjectID & ".Proj").Type = ADVL_Utilities_Library_1.Project.Types.Hybrid
                        ProjInfo(ProjectID & ".Proj").IconNumber = 6
                        ProjInfo(ProjectID & ".Proj").OpenIconNumber = 7

                        Dim node As TreeNode()

                        If ApplicationName = trvAppTree.Nodes(0).Name Then
                            node = trvAppTree.Nodes.Find(ApplicationName, False)
                        Else
                            node = trvAppTree.Nodes(0).Nodes.Find(ApplicationName, False)
                        End If

                        If node Is Nothing Then
                            'Node not found.
                            Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                        ElseIf node.Length = 0 Then
                            'Node not found.
                            Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                        Else
                            trvAppTree.SelectedNode = node(0)
                            trvAppTree.SelectedNode.Nodes.Add(ProjectID & ".Proj", ProjectName, 6, 7) '6, 7 Hybrid project icons.
                        End If
                    Case Else
                        Message.AddWarning("Unknown project type: " & ProjectType & vbCrLf)
                End Select
            End If
        Else
            'Message.AddWarning("This is a Child Project and cannot be added to the Application Tree. Parent Project ID = " & ParentProjectID & vbCrLf)
        End If
    End Sub

    Private Sub ReadDragDropDirectoryProjectInfo_OLD(ByVal ProjectPath As String)
        'Read the Project Information from a Directory Project.

        Dim ProjectInfo As System.Xml.Linq.XDocument = XDocument.Load(ProjectPath & "\Project_Info_ADVL_2.xml")

        'Dim ProjectAppNetName As String
        'If System.IO.File.Exists(ProjectPath & "\ProjectParams_ADVL2.xml") Then
        '    Dim ParameterInfo As System.Xml.Linq.XDocument = XDocument.Load(ProjectPath & "\ProjectParams_ADVL2.xml")
        '    Dim AppNetNames = From names In ParameterInfo.<ProjectParameterList>.<Parameter>
        '                      Where names.<Name>.Value = "AppNetName"
        '                      Select names

        '    If AppNetNames.Count = 0 Then
        '        ProjectAppNetName = ""
        '        Message.Add("The Project Parameters file did not contain an AppNetName parameter." & vbCrLf)
        '    ElseIf AppNetNames.Count = 1 Then
        '        ProjectAppNetName = AppNetNames(0).<Value>.Value
        '    Else
        '        ProjectAppNetName = AppNetNames(0).<Value>.Value
        '        Message.Add("The Project Parameters file contained more than one AppNetName parameter." & vbCrLf)
        '    End If
        'Else
        '    ProjectAppNetName = ""
        '    Message.Add("The project did not contain a Project Parameters file." & vbCrLf)
        'End If

        Dim ProjectNetworkName As String
        If System.IO.File.Exists(ProjectPath & "\ProjectParams_ADVL2.xml") Then
            Dim ParameterInfo As System.Xml.Linq.XDocument = XDocument.Load(ProjectPath & "\ProjectParams_ADVL2.xml")
            Dim ProNetNames = From names In ParameterInfo.<ProjectParameterList>.<Parameter>
                              Where names.<Name>.Value = "ProNetName"
                              Select names

            If ProNetNames.Count = 0 Then
                ProjectNetworkName = ""
                Message.Add("The Project Parameters file did not contain an ProNetName parameter." & vbCrLf)
            ElseIf ProNetNames.Count = 1 Then
                ProjectNetworkName = ProNetNames(0).<Value>.Value
            Else
                ProjectNetworkName = ProNetNames(0).<Value>.Value
                Message.Add("The Project Parameters file contained more than one ProNetName parameter." & vbCrLf)
            End If
        Else
            ProjectNetworkName = ""
            Message.Add("The project did not contain a Project Parameters file." & vbCrLf)
        End If

        Message.Add(vbCrLf) 'Add a blank line.

        Dim ProjectName As String
        If ProjectInfo.<Project>.<Name>.Value = Nothing Then
            ProjectName = ""
        Else
            ProjectName = ProjectInfo.<Project>.<Name>.Value
        End If
        'Message.Add("Project Name = " & ProjectName & vbCrLf)

        Dim ProjectID As String
        If ProjectInfo.<Project>.<ID>.Value = Nothing Then
            ProjectID = ""
            Message.AddWarning("The Project ID is blank." & vbCrLf)
        Else
            ProjectID = ProjectInfo.<Project>.<ID>.Value
        End If
        'Message.Add("Project ID = " & ProjectID & vbCrLf)

        Dim ProjectType As String
        If ProjectInfo.<Project>.<Type>.Value = Nothing Then
            ProjectType = ""
        Else
            ProjectType = ProjectInfo.<Project>.<Type>.Value
        End If
        'Message.Add("Project Type = " & ProjectType & vbCrLf)

        'Message.Add("Project Path= " & ProjectPath & vbCrLf)

        Dim ProjectDescription As String
        If ProjectInfo.<Project>.<Description>.Value = Nothing Then
            ProjectDescription = ""
        Else
            ProjectDescription = ProjectInfo.<Project>.<Description>.Value
        End If
        'Message.Add("Project Description = " & ProjectDescription & vbCrLf)

        Dim ApplicationName As String
        If ProjectInfo.<Project>.<Application>.<Name>.Value = Nothing Then
            ApplicationName = ""
        Else
            ApplicationName = ProjectInfo.<Project>.<Application>.<Name>.Value
        End If
        'Message.Add("Application Name = " & ApplicationName & vbCrLf)

        Dim ParentProjectName As String
        'Legacy code version:
        If ProjectInfo.<Project>.<HostProject>.<Name>.Value = Nothing Then
            ParentProjectName = ""
        Else
            ParentProjectName = ProjectInfo.<Project>.<HostProject>.<Name>.Value
        End If

        'Updated code version:
        If ProjectInfo.<Project>.<ParentProject>.<Name>.Value = Nothing Then
            'ParentProjectName = ""  'NO NEED TO CHANGE THIS - THE CODE ABOVE SHOULD HAVE SET THE CORRECT VALUE.
        Else
            ParentProjectName = ProjectInfo.<Project>.<ParentProject>.<Name>.Value
        End If

        'Message.Add("Parent Project Name = " & ParentProjectName & vbCrLf)

        Dim ParentProjectID As String
        'Legacy code version:
        If ProjectInfo.<Project>.<HostProject>.<ID>.Value = Nothing Then
            ParentProjectID = ""
        Else
            ParentProjectID = ProjectInfo.<Project>.<HostProject>.<ID>.Value
        End If

        'Updated code version:
        If ProjectInfo.<Project>.<ParentProject>.<ID>.Value = Nothing Then
            'ParentProjectID = "" 'NO NEED TO CHANGE THIS - THE CODE ABOVE SHOULD HAVE SET THE CORRECT VALUE.
        Else
            ParentProjectID = ProjectInfo.<Project>.<ParentProject>.<ID>.Value
        End If

        If ParentProjectID = "" Then

        Else
            Message.AddWarning("This is a Child Project and cannot be added to the list. Parent Project ID = " & ParentProjectID & vbCrLf)
            Exit Sub
        End If


        'Add project to the Project List: ---------------------------------------------------
        'This is displayed in the Project List tab.
        If ProjectIdAvailable(ProjectID) Then
            'If ParentProjectID = "" Then
            'The Project is not a Child Project and can be added.
            dgvProjects.Rows.Add()
            Dim CurrentRow As Integer = dgvProjects.Rows.Count - 2
            dgvProjects.Rows(CurrentRow).Cells(0).Value = ProjectName
            'dgvProjects.Rows(CurrentRow).Cells(1).Value = ProjectAppNetName 'ADDED 10Feb19
            dgvProjects.Rows(CurrentRow).Cells(1).Value = ProjectNetworkName
            dgvProjects.Rows(CurrentRow).Cells(2).Value = ProjectType
            dgvProjects.Rows(CurrentRow).Cells(3).Value = ProjectID
            dgvProjects.Rows(CurrentRow).Cells(4).Value = ApplicationName
            dgvProjects.Rows(CurrentRow).Cells(5).Value = ProjectDescription
            dgvProjects.AutoResizeColumns()

            Dim NewProjectInfo As New ProjSummary
            NewProjectInfo.Name = ProjectName
            'NewProjectInfo.AppNetName = ProjectAppNetName 'Added 10Feb19
            NewProjectInfo.ProNetName = ProjectNetworkName
            NewProjectInfo.ID = ProjectID
            Select Case ProjectType
                Case "None"
                    NewProjectInfo.Type = ADVL_Utilities_Library_1.Project.Types.None
                Case "Directory"
                    NewProjectInfo.Type = ADVL_Utilities_Library_1.Project.Types.Directory
                Case "Archive"
                    NewProjectInfo.Type = ADVL_Utilities_Library_1.Project.Types.Archive
                Case "Hybrid"
                    NewProjectInfo.Type = ADVL_Utilities_Library_1.Project.Types.Hybrid
                Case Else
                    Message.AddWarning("Unknown project type: " & ProjectType & vbCrLf)
            End Select

            NewProjectInfo.Path = ProjectPath
            NewProjectInfo.Description = ProjectDescription
            NewProjectInfo.ApplicationName = ApplicationName
            NewProjectInfo.ParentProjectName = ParentProjectName
            NewProjectInfo.ParentProjectID = ParentProjectID

            Proj.List.Add(NewProjectInfo)
            Message.Add("Project added the list. Project ID = " & ProjectID & vbCrLf)

        Else
            Message.Add("The Project is already in the list." & vbCrLf)
        End If

        'Add project to the AppTree -----------------------------------------------------
        'This is displayed in the Applcation Tree tab.
        If ProjInfo.ContainsKey(ProjectID & ".Proj") Then
            'Message.Add("Project is already in the TreeView. Project ID = " & ProjectID & vbCrLf)
        Else
            ProjInfo.Add(ProjectID & ".Proj", New clsProjInfo)
            ProjInfo(ProjectID & ".Proj").Name = ProjectName
            'ProjInfo(ProjectID & ".Proj").AppNetName = ProjectAppNetName
            ProjInfo(ProjectID & ".Proj").ProNetName = ProjectNetworkName
            ProjInfo(ProjectID & ".Proj").ID = ProjectID

            ProjInfo(ProjectID & ".Proj").Path = ProjectPath
            ProjInfo(ProjectID & ".Proj").Description = ProjectDescription
            ProjInfo(ProjectID & ".Proj").ApplicationName = ApplicationName
            ProjInfo(ProjectID & ".Proj").ParentProjectName = ParentProjectName
            ProjInfo(ProjectID & ".Proj").ParentProjectID = ParentProjectID

            Select Case ProjectType
                Case "None"
                    ProjInfo(ProjectID & ".Proj").Type = ADVL_Utilities_Library_1.Project.Types.None
                    ProjInfo(ProjectID & ".Proj").IconNumber = 0
                    ProjInfo(ProjectID & ".Proj").OpenIconNumber = 1
                    Dim node As TreeNode()
                    'If ApplicationName = trvAppTree.TopNode.Name Then
                    If ApplicationName = trvAppTree.Nodes(0).Name Then
                        node = trvAppTree.Nodes.Find(ApplicationName, False)
                    Else
                        'node = trvAppTree.TopNode.Nodes.Find(ApplicationName, False)
                        node = trvAppTree.Nodes(0).Nodes.Find(ApplicationName, False)
                    End If
                    If node Is Nothing Then
                        'Node not found.
                        Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                    ElseIf node.Length = 0 Then
                        'Node not found.
                        Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                    Else
                        trvAppTree.SelectedNode = node(0)
                        trvAppTree.SelectedNode.Nodes.Add(ProjectID & ".Proj", ProjectName, 0, 1) '0, 1 Default project icons.
                    End If
                Case "Directory"
                    ProjInfo(ProjectID & ".Proj").Type = ADVL_Utilities_Library_1.Project.Types.Directory
                    ProjInfo(ProjectID & ".Proj").IconNumber = 2
                    ProjInfo(ProjectID & ".Proj").OpenIconNumber = 3
                    Dim node As TreeNode()
                    'If ApplicationName = trvAppTree.TopNode.Name Then
                    If ApplicationName = trvAppTree.Nodes(0).Name Then
                        node = trvAppTree.Nodes.Find(ApplicationName, False)
                    Else
                        'node = trvAppTree.TopNode.Nodes.Find(ApplicationName, False)
                        node = trvAppTree.Nodes(0).Nodes.Find(ApplicationName, False)
                    End If
                    If node Is Nothing Then
                        'Node not found.
                        Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                    ElseIf node.Length = 0 Then
                        'Node not found.
                        Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                    Else
                        trvAppTree.SelectedNode = node(0)
                        trvAppTree.SelectedNode.Nodes.Add(ProjectID & ".Proj", ProjectName, 2, 3) '2, 3 Directory project icons.
                    End If
                Case "Archive"
                    ProjInfo(ProjectID & ".Proj").Type = ADVL_Utilities_Library_1.Project.Types.Archive
                    ProjInfo(ProjectID & ".Proj").IconNumber = 4
                    ProjInfo(ProjectID & ".Proj").OpenIconNumber = 5
                    Dim node As TreeNode()
                    'If ApplicationName = trvAppTree.TopNode.Name Then
                    If ApplicationName = trvAppTree.Nodes(0).Name Then
                        node = trvAppTree.Nodes.Find(ApplicationName, False)
                    Else
                        'node = trvAppTree.TopNode.Nodes.Find(ApplicationName, False)
                        node = trvAppTree.Nodes(0).Nodes.Find(ApplicationName, False)
                    End If
                    If node Is Nothing Then
                        'Node not found.
                        Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                    ElseIf node.Length = 0 Then
                        'Node not found.
                        Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                    Else
                        trvAppTree.SelectedNode = node(0)
                        trvAppTree.SelectedNode.Nodes.Add(ProjectID & ".Proj", ProjectName, 4, 5) '4, 5 Archive project icons.
                    End If
                Case "Hybrid"
                    ProjInfo(ProjectID & ".Proj").Type = ADVL_Utilities_Library_1.Project.Types.Hybrid
                    ProjInfo(ProjectID & ".Proj").IconNumber = 6
                    ProjInfo(ProjectID & ".Proj").OpenIconNumber = 7

                    Dim node As TreeNode()

                    'If ApplicationName = trvAppTree.TopNode.Name Then
                    If ApplicationName = trvAppTree.Nodes(0).Name Then
                        node = trvAppTree.Nodes.Find(ApplicationName, False)
                    Else
                        'node = trvAppTree.TopNode.Nodes.Find(ApplicationName, False)
                        node = trvAppTree.Nodes(0).Nodes.Find(ApplicationName, False)
                    End If

                    If node Is Nothing Then
                        'Node not found.
                        Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                    ElseIf node.Length = 0 Then
                        'Node not found.
                        Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                    Else
                        trvAppTree.SelectedNode = node(0)
                        trvAppTree.SelectedNode.Nodes.Add(ProjectID & ".Proj", ProjectName, 6, 7) '6, 7 Hybrid project icons.
                    End If
                Case Else
                    Message.AddWarning("Unknown project type: " & ProjectType & vbCrLf)
            End Select
        End If
    End Sub

    Private Sub ReadDragDropArchiveProjectInfo(ByVal ProjectPath As String)
        'Read the Project Information from an Archive Project.

        Dim ProjectInfo As System.Xml.Linq.XDocument

        Dim Zip As New ADVL_Utilities_Library_1.ZipComp
        Zip.ArchivePath = ProjectPath

        If Zip.EntryExists("Project_Info_ADVL_2.xml") Then
            ProjectInfo = XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8""?>" & Zip.GetText("Project_Info_ADVL_2.xml"))
        Else
            'Convert the ADVL_Project_Info.xml file into a Project_Info_ADVL_2.xml file:
            Dim ProjInfoConversion As New ADVL_Utilities_Library_1.FormatConvert.ProjectInfoFileConversion
            ProjInfoConversion.ProjectType = ADVL_Utilities_Library_1.FormatConvert.ProjectInfoFileConversion.ProjectTypes.Archive
            ProjInfoConversion.ProjectPath = ProjectPath
            ProjInfoConversion.InputFormatCode = ADVL_Utilities_Library_1.FormatConvert.ProjectInfoFileConversion.FormatCodes.ADVL_1
            ProjInfoConversion.OutputFormatCode = ADVL_Utilities_Library_1.FormatConvert.ProjectInfoFileConversion.FormatCodes.ADVL_2
            ProjInfoConversion.Convert()
            If Zip.EntryExists("Project_Info_ADVL_2.xml") Then
                ProjectInfo = XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8""?>" & Zip.GetText("Project_Info_ADVL_2.xml"))
            Else
                Message.AddWarning("The Project Information file could not be converted to the ADVL_2 format version." & vbCrLf)
                Exit Sub
            End If
        End If

        If ProjectInfo Is Nothing Then
            Message.AddWarning("Project Info file not found in the Archive project" & vbCrLf)
            Exit Sub
        End If

        'UPDATES 12Jan2020 -------------------------------------------------------------------------------------
        Dim ProjectNetworkName As String
        If Zip.EntryExists("Project_Params_ADVL_2.xml") Then
            Dim ParameterInfo As System.Xml.Linq.XDocument
            ParameterInfo = XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8""?>" & Zip.GetText("Project_Params_ADVL_2.xml"))

            Dim ProNetNames = From names In ParameterInfo.<ProjectParameterList>.<Parameter>
                              Where names.<Name>.Value = "ProNetName"
                              Select names

            If ProNetNames.Count = 0 Then
                ProjectNetworkName = ""
                Message.Add("The Project Parameters file did not contain an ProNetName parameter." & vbCrLf)
            ElseIf ProNetNames.Count = 1 Then
                ProjectNetworkName = ProNetNames(0).<Value>.Value
            Else
                ProjectNetworkName = ProNetNames(0).<Value>.Value
                Message.Add("The Project Parameters file contained more than one ProNetName parameter." & vbCrLf)
            End If

        Else
            ProjectNetworkName = ""
            Message.Add("The project did not contain a Project Parameters file." & vbCrLf)
        End If

        'Message.Add(vbCrLf) 'Add a blank line.
        'END UPDATES 12Jan2020 ---------------------------------------------------------------------------------

        Dim ProjectName As String
        If ProjectInfo.<Project>.<Name>.Value = Nothing Then
            ProjectName = ""
        Else
            ProjectName = ProjectInfo.<Project>.<Name>.Value
        End If
        'Message.Add("Project Name = " & ProjectName & vbCrLf)

        'SEE UPDATES 12Jan2020 Above!!! -----------------------------------------------------------------------
        'Dim ProjectNetworkName As String
        'If ProjectInfo.<Project>.<ProNetName>.Value = Nothing Then
        '    'Check if the old AppNetName is used:
        '    If ProjectInfo.<Project>.<AppNetName>.Value = Nothing Then
        '        ProjectNetworkName = ""
        '    Else
        '        ProjectNetworkName = ProjectInfo.<Project>.<AppNetName>.Value 'Read the old parameter name: AppNetName (This is now called ProNetName.)
        '    End If
        'Else
        '    ProjectNetworkName = ProjectInfo.<Project>.<ProNetName>.Value
        'End If
        'Message.Add("Project Network Name = " & ProjectNetworkName & vbCrLf)

        Dim ProjectID As String
        If ProjectInfo.<Project>.<ID>.Value = Nothing Then
            ProjectID = ""
        Else
            ProjectID = ProjectInfo.<Project>.<ID>.Value
        End If
        'Message.Add("Project ID = " & ProjectID & vbCrLf)

        Dim ProjectType As String
        If ProjectInfo.<Project>.<Type>.Value = Nothing Then
            ProjectType = ""
        Else
            ProjectType = ProjectInfo.<Project>.<Type>.Value
        End If
        'Message.Add("Project Type = " & ProjectType & vbCrLf)

        'Message.Add("Project Path= " & ProjectPath & vbCrLf)

        Dim ProjectDescription As String
        If ProjectInfo.<Project>.<Description>.Value = Nothing Then
            ProjectDescription = ""
        Else
            ProjectDescription = ProjectInfo.<Project>.<Description>.Value
        End If
        'Message.Add("Project Description = " & ProjectDescription & vbCrLf)

        Dim ApplicationName As String
        If ProjectInfo.<Project>.<Application>.<Name>.Value = Nothing Then
            ApplicationName = ""
        Else
            ApplicationName = ProjectInfo.<Project>.<Application>.<Name>.Value
        End If
        'Message.Add("Application Name = " & ApplicationName & vbCrLf)

        Dim ParentProjectName As String
        'Legacy code version:
        If ProjectInfo.<Project>.<HostProject>.<Name>.Value = Nothing Then
            ParentProjectName = ""
        Else
            ParentProjectName = ProjectInfo.<Project>.<HostProject>.<Name>.Value
        End If

        'Updated code version:
        If ProjectInfo.<Project>.<ParentProject>.<Name>.Value = Nothing Then
            'ParentProjectName = ""  'NO NEED TO CHANGE THIS - THE CODE ABOVE SHOULD HAVE SET THE CORRECT VALUE.
        Else
            ParentProjectName = ProjectInfo.<Project>.<ParentProject>.<Name>.Value
        End If

        'Message.Add("Parent Project Name = " & ParentProjectName & vbCrLf)

        Dim ParentProjectID As String
        'Legacy code version:
        If ProjectInfo.<Project>.<HostProject>.<ID>.Value = Nothing Then
            ParentProjectID = ""
        Else
            ParentProjectID = ProjectInfo.<Project>.<HostProject>.<ID>.Value
        End If

        'Updated code version:
        If ProjectInfo.<Project>.<ParentProject>.<ID>.Value = Nothing Then
            'ParentProjectID = "" 'NO NEED TO CHANGE THIS - THE CODE ABOVE SHOULD HAVE SET THE CORRECT VALUE.
        Else
            ParentProjectID = ProjectInfo.<Project>.<ParentProject>.<ID>.Value
        End If

        'Message.Add("Parent Project ID = " & ParentProjectID & vbCrLf)

        'Add project to the Project List: ---------------------------------------------------
        If ProjectIdAvailable(ProjectID) Then
            dgvProjects.Rows.Add()
            Dim CurrentRow As Integer = dgvProjects.Rows.Count - 2
            dgvProjects.Rows(CurrentRow).Cells(0).Value = ProjectName
            dgvProjects.Rows(CurrentRow).Cells(1).Value = ProjectNetworkName
            dgvProjects.Rows(CurrentRow).Cells(2).Value = ProjectType
            dgvProjects.Rows(CurrentRow).Cells(3).Value = ProjectID
            dgvProjects.Rows(CurrentRow).Cells(4).Value = ApplicationName
            dgvProjects.Rows(CurrentRow).Cells(5).Value = ProjectDescription
            dgvProjects.AutoResizeColumns()

            Dim NewProjectInfo As New ProjSummary
            NewProjectInfo.Name = ProjectName
            NewProjectInfo.ProNetName = ProjectNetworkName
            NewProjectInfo.ID = ProjectID
            Select Case ProjectType
                Case "None"
                    NewProjectInfo.Type = ADVL_Utilities_Library_1.Project.Types.None
                Case "Directory"
                    NewProjectInfo.Type = ADVL_Utilities_Library_1.Project.Types.Directory
                Case "Archive"
                    NewProjectInfo.Type = ADVL_Utilities_Library_1.Project.Types.Archive
                Case "Hybrid"
                    NewProjectInfo.Type = ADVL_Utilities_Library_1.Project.Types.Hybrid
                Case Else
                    Message.AddWarning("Unknown project type: " & ProjectType & vbCrLf)
            End Select

            NewProjectInfo.Path = ProjectPath
            NewProjectInfo.Description = ProjectDescription
            NewProjectInfo.ApplicationName = ApplicationName
            NewProjectInfo.ParentProjectName = ParentProjectName
            NewProjectInfo.ParentProjectID = ParentProjectID

            Proj.List.Add(NewProjectInfo)

        Else
            'Message.Add("The Project is already in the list." & vbCrLf)
        End If

        'Add project to the AppTree -----------------------------------------------------
        'This is displayed in the Applcation Tree tab.
        'NOTE: ProjInfo can contain ProjectID without the project added to the tree! (This can occur if the corresponding Application Name was not found in the tree after the project was added to ProjInfo.)
        ' To handle this case, the project information is first added to NewProjInfo. It is only added tp ProjInfo later if appropriate.
        'If ProjInfo.ContainsKey(ProjectID & ".Proj") Then
        '    Message.Add("Project is already in the TreeView. Project ID = " & ProjectID & vbCrLf)
        'Else
        Dim NewProjInfo As New clsProjInfo
        NewProjInfo.Name = ProjectName
        NewProjInfo.ProNetName = ProjectNetworkName
        NewProjInfo.ID = ProjectID
        NewProjInfo.Path = ProjectPath
        NewProjInfo.Description = ProjectDescription
        NewProjInfo.ApplicationName = ApplicationName
        NewProjInfo.ParentProjectName = ParentProjectName
        NewProjInfo.ParentProjectID = ParentProjectID

        Select Case ProjectType
            Case "None"
                NewProjInfo.Type = ADVL_Utilities_Library_1.Project.Types.None
                NewProjInfo.IconNumber = 0
                NewProjInfo.OpenIconNumber = 1
                Dim node As TreeNode()
                If ApplicationName = trvAppTree.Nodes(0).Name Then
                    node = trvAppTree.Nodes.Find(ApplicationName, False)
                Else
                    node = trvAppTree.Nodes(0).Nodes.Find(ApplicationName, False)
                End If
                If node Is Nothing Then
                    'Node not found.
                    Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                ElseIf node.Length = 0 Then
                    'Node not found.
                    Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                Else
                    'Add the project information to ProjInfo():
                    ProjInfo.Add(ProjectID & ".Proj", NewProjInfo)
                    'Add the project node to the Application Tree View:
                    trvAppTree.SelectedNode = node(0)
                    trvAppTree.SelectedNode.Nodes.Add(ProjectID & ".Proj", ProjectName, 0, 1) '0, 1 Default project icons.
                End If
            Case "Directory"
                NewProjInfo.Type = ADVL_Utilities_Library_1.Project.Types.Directory
                NewProjInfo.IconNumber = 2
                NewProjInfo.OpenIconNumber = 3
                Dim node As TreeNode()
                If ApplicationName = trvAppTree.Nodes(0).Name Then
                    node = trvAppTree.Nodes.Find(ApplicationName, False)
                Else
                    node = trvAppTree.Nodes(0).Nodes.Find(ApplicationName, False)
                End If
                If node Is Nothing Then
                    'Node not found.
                    Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                ElseIf node.Length = 0 Then
                    'Node not found.
                    Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                Else
                    'Add the project information to ProjInfo():
                    ProjInfo.Add(ProjectID & ".Proj", NewProjInfo)
                    'Add the project node to the Application Tree View:
                    trvAppTree.SelectedNode = node(0)
                    trvAppTree.SelectedNode.Nodes.Add(ProjectID & ".Proj", ProjectName, 2, 3) '2, 3 Directory project icons.
                End If
            Case "Archive"
                NewProjInfo.Type = ADVL_Utilities_Library_1.Project.Types.Archive
                NewProjInfo.IconNumber = 4
                NewProjInfo.OpenIconNumber = 5
                Dim node As TreeNode()
                If ApplicationName = trvAppTree.Nodes(0).Name Then
                    node = trvAppTree.Nodes.Find(ApplicationName, False)
                Else
                    node = trvAppTree.Nodes(0).Nodes.Find(ApplicationName, False)
                End If
                If node Is Nothing Then
                    'Node not found.
                    Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                ElseIf node.Length = 0 Then
                    'Node not found.
                    Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                Else
                    'Add the project information to ProjInfo():
                    ProjInfo.Add(ProjectID & ".Proj", NewProjInfo)
                    'Add the project node to the Application Tree View:
                    trvAppTree.SelectedNode = node(0)
                    trvAppTree.SelectedNode.Nodes.Add(ProjectID & ".Proj", ProjectName, 4, 5) '4, 5 Archive project icons.
                End If
            Case "Hybrid"
                NewProjInfo.Type = ADVL_Utilities_Library_1.Project.Types.Hybrid
                NewProjInfo.IconNumber = 6
                NewProjInfo.OpenIconNumber = 7
                Dim node As TreeNode()
                If ApplicationName = trvAppTree.Nodes(0).Name Then
                    node = trvAppTree.Nodes.Find(ApplicationName, False)
                Else
                    node = trvAppTree.Nodes(0).Nodes.Find(ApplicationName, False)
                End If

                If node Is Nothing Then
                    'Node not found.
                    Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                ElseIf node.Length = 0 Then
                    'Node not found.
                    Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                Else
                    'Add the project information to ProjInfo():
                    ProjInfo.Add(ProjectID & ".Proj", NewProjInfo)
                    'Add the project node to the Application Tree View:
                    trvAppTree.SelectedNode = node(0)
                    trvAppTree.SelectedNode.Nodes.Add(ProjectID & ".Proj", ProjectName, 6, 7) '6, 7 Hybrid project icons.
                End If
            Case Else
                Message.AddWarning("Unknown project type: " & ProjectType & vbCrLf)
        End Select
    End Sub

    Private Sub ReadDragDropArchiveProjectInfo_Old(ByVal ProjectPath As String)
        'Read the Project Information from an Archive Project.

        Dim ProjectInfo As System.Xml.Linq.XDocument

        Dim Zip As New ADVL_Utilities_Library_1.ZipComp
        Zip.ArchivePath = ProjectPath

        If Zip.EntryExists("Project_Info_ADVL_2.xml") Then
            ProjectInfo = XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8""?>" & Zip.GetText("Project_Info_ADVL_2.xml"))
        Else
            'Convert the ADVL_Project_Info.xml file into a Project_Info_ADVL_2.xml file:
            Dim ProjInfoConversion As New ADVL_Utilities_Library_1.FormatConvert.ProjectInfoFileConversion
            ProjInfoConversion.ProjectType = ADVL_Utilities_Library_1.FormatConvert.ProjectInfoFileConversion.ProjectTypes.Archive
            ProjInfoConversion.ProjectPath = ProjectPath
            ProjInfoConversion.InputFormatCode = ADVL_Utilities_Library_1.FormatConvert.ProjectInfoFileConversion.FormatCodes.ADVL_1
            ProjInfoConversion.OutputFormatCode = ADVL_Utilities_Library_1.FormatConvert.ProjectInfoFileConversion.FormatCodes.ADVL_2
            ProjInfoConversion.Convert()
            If Zip.EntryExists("Project_Info_ADVL_2.xml") Then
                ProjectInfo = XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8""?>" & Zip.GetText("Project_Info_ADVL_2.xml"))
            Else
                Message.AddWarning("The Project Information file could not be converted to the ADVL_2 format version." & vbCrLf)
                Exit Sub
            End If
        End If

        If ProjectInfo Is Nothing Then
            Message.AddWarning("Project Info file not found in the Archive project" & vbCrLf)
            Exit Sub
        End If

        Message.Add(vbCrLf) 'Add a blank line.

        Dim ProjectName As String
        If ProjectInfo.<Project>.<Name>.Value = Nothing Then
            ProjectName = ""
        Else
            ProjectName = ProjectInfo.<Project>.<Name>.Value
        End If
        Message.Add("Project Name = " & ProjectName & vbCrLf)

        'Dim ProjectAppNetName As String
        'If ProjectInfo.<Project>.<AppNetName>.Value = Nothing Then
        '    ProjectAppNetName = ""
        'Else
        '    ProjectAppNetName = ProjectInfo.<Project>.<AppNetName>.Value
        'End If
        'Message.Add("Project Application Network Name = " & ProjectAppNetName & vbCrLf)

        Dim ProjectNetworkName As String
        If ProjectInfo.<Project>.<ProNetName>.Value = Nothing Then
            'Check if the old AppNetName is used:
            If ProjectInfo.<Project>.<AppNetName>.Value = Nothing Then
                ProjectNetworkName = ""
            Else
                ProjectNetworkName = ProjectInfo.<Project>.<AppNetName>.Value 'Read the old parameter name: AppNetName (This is now called ProNetName.)
            End If
        Else
            ProjectNetworkName = ProjectInfo.<Project>.<ProNetName>.Value
        End If
        Message.Add("Project Network Name = " & ProjectNetworkName & vbCrLf)

        Dim ProjectID As String
        If ProjectInfo.<Project>.<ID>.Value = Nothing Then
            ProjectID = ""
        Else
            ProjectID = ProjectInfo.<Project>.<ID>.Value
        End If
        Message.Add("Project ID = " & ProjectID & vbCrLf)

        Dim ProjectType As String
        If ProjectInfo.<Project>.<Type>.Value = Nothing Then
            ProjectType = ""
        Else
            ProjectType = ProjectInfo.<Project>.<Type>.Value
        End If
        Message.Add("Project Type = " & ProjectType & vbCrLf)

        Message.Add("Project Path= " & ProjectPath & vbCrLf)

        Dim ProjectDescription As String
        If ProjectInfo.<Project>.<Description>.Value = Nothing Then
            ProjectDescription = ""
        Else
            ProjectDescription = ProjectInfo.<Project>.<Description>.Value
        End If
        Message.Add("Project Description = " & ProjectDescription & vbCrLf)

        Dim ApplicationName As String
        If ProjectInfo.<Project>.<Application>.<Name>.Value = Nothing Then
            ApplicationName = ""
        Else
            ApplicationName = ProjectInfo.<Project>.<Application>.<Name>.Value
        End If
        Message.Add("Application Name = " & ApplicationName & vbCrLf)

        Dim ParentProjectName As String
        'Legacy code version:
        If ProjectInfo.<Project>.<HostProject>.<Name>.Value = Nothing Then
            ParentProjectName = ""
        Else
            ParentProjectName = ProjectInfo.<Project>.<HostProject>.<Name>.Value
        End If

        'Updated code version:
        If ProjectInfo.<Project>.<ParentProject>.<Name>.Value = Nothing Then
            'ParentProjectName = ""  'NO NEED TO CHANGE THIS - THE CODE ABOVE SHOULD HAVE SET THE CORRECT VALUE.
        Else
            ParentProjectName = ProjectInfo.<Project>.<ParentProject>.<Name>.Value
        End If

        Message.Add("Parent Project Name = " & ParentProjectName & vbCrLf)

        Dim ParentProjectID As String
        'Legacy code version:
        If ProjectInfo.<Project>.<HostProject>.<ID>.Value = Nothing Then
            ParentProjectID = ""
        Else
            ParentProjectID = ProjectInfo.<Project>.<HostProject>.<ID>.Value
        End If

        'Updated code version:
        If ProjectInfo.<Project>.<ParentProject>.<ID>.Value = Nothing Then
            'ParentProjectID = "" 'NO NEED TO CHANGE THIS - THE CODE ABOVE SHOULD HAVE SET THE CORRECT VALUE.
        Else
            ParentProjectID = ProjectInfo.<Project>.<ParentProject>.<ID>.Value
        End If

        Message.Add("Parent Project ID = " & ParentProjectID & vbCrLf)

        'Add project to the Project List: ---------------------------------------------------
        If ProjectIdAvailable(ProjectID) Then
            dgvProjects.Rows.Add()
            Dim CurrentRow As Integer = dgvProjects.Rows.Count - 2
            dgvProjects.Rows(CurrentRow).Cells(0).Value = ProjectName
            'dgvProjects.Rows(CurrentRow).Cells(1).Value = ProjectAppNetName 'ADDED 10Feb19
            dgvProjects.Rows(CurrentRow).Cells(1).Value = ProjectNetworkName
            dgvProjects.Rows(CurrentRow).Cells(2).Value = ProjectType
            dgvProjects.Rows(CurrentRow).Cells(3).Value = ProjectID
            dgvProjects.Rows(CurrentRow).Cells(4).Value = ApplicationName
            dgvProjects.Rows(CurrentRow).Cells(5).Value = ProjectDescription
            dgvProjects.AutoResizeColumns()

            Dim NewProjectInfo As New ProjSummary
            NewProjectInfo.Name = ProjectName
            'NewProjectInfo.AppNetName = ProjectAppNetName 'ADDED 10Feb19
            NewProjectInfo.ProNetName = ProjectNetworkName
            NewProjectInfo.ID = ProjectID
            Select Case ProjectType
                Case "None"
                    NewProjectInfo.Type = ADVL_Utilities_Library_1.Project.Types.None
                Case "Directory"
                    NewProjectInfo.Type = ADVL_Utilities_Library_1.Project.Types.Directory
                Case "Archive"
                    NewProjectInfo.Type = ADVL_Utilities_Library_1.Project.Types.Archive
                Case "Hybrid"
                    NewProjectInfo.Type = ADVL_Utilities_Library_1.Project.Types.Hybrid
                Case Else
                    Message.AddWarning("Unknown project type: " & ProjectType & vbCrLf)
            End Select

            NewProjectInfo.Path = ProjectPath
            NewProjectInfo.Description = ProjectDescription
            NewProjectInfo.ApplicationName = ApplicationName
            NewProjectInfo.ParentProjectName = ParentProjectName
            NewProjectInfo.ParentProjectID = ParentProjectID

            Proj.List.Add(NewProjectInfo)

        Else
            Message.Add("The Project is already in the list." & vbCrLf)
        End If

        'Add project to the AppTree -----------------------------------------------------
        'This is displayed in the Applcation Tree tab.
        'NOTE: ProjInfo can contain ProjectID without the project added to the tree! (This can occur if the corresponding Application Name was not found in the tree after the project was added to ProjInfo.)
        ' To handle this case, the project information is first added to NewProjInfo. It is only added tp ProjInfo later if appropriate.
        'If ProjInfo.ContainsKey(ProjectID & ".Proj") Then
        '    Message.Add("Project is already in the TreeView. Project ID = " & ProjectID & vbCrLf)
        'Else
        Dim NewProjInfo As New clsProjInfo
        'ProjInfo.Add(ProjectID & ".Proj", New clsProjInfo)
        'ProjInfo(ProjectID & ".Proj").Name = ProjectName
        NewProjInfo.Name = ProjectName
        ''ProjInfo(ProjectID & ".Proj").AppNetName = ProjectAppNetName 'ADDED 10Feb19
        'ProjInfo(ProjectID & ".Proj").ProNetName = ProjectNetworkName
        NewProjInfo.ProNetName = ProjectNetworkName
        'ProjInfo(ProjectID & ".Proj").ID = ProjectID
        NewProjInfo.ID = ProjectID

        'ProjInfo(ProjectID & ".Proj").Path = ProjectPath
        NewProjInfo.Path = ProjectPath
        'ProjInfo(ProjectID & ".Proj").Description = ProjectDescription
        NewProjInfo.Description = ProjectDescription
        'ProjInfo(ProjectID & ".Proj").ApplicationName = ApplicationName
        NewProjInfo.ApplicationName = ApplicationName
        'ProjInfo(ProjectID & ".Proj").ParentProjectName = ParentProjectName
        NewProjInfo.ParentProjectName = ParentProjectName
        'ProjInfo(ProjectID & ".Proj").ParentProjectID = ParentProjectID
        NewProjInfo.ParentProjectID = ParentProjectID

        Select Case ProjectType
            Case "None"
                'ProjInfo(ProjectID & ".Proj").Type = ADVL_Utilities_Library_1.Project.Types.None
                NewProjInfo.Type = ADVL_Utilities_Library_1.Project.Types.None
                'ProjInfo(ProjectID & ".Proj").IconNumber = 0
                NewProjInfo.IconNumber = 0
                'ProjInfo(ProjectID & ".Proj").OpenIconNumber = 1
                NewProjInfo.OpenIconNumber = 1
                Dim node As TreeNode()
                'If ApplicationName = trvAppTree.TopNode.Name Then
                If ApplicationName = trvAppTree.Nodes(0).Name Then
                    node = trvAppTree.Nodes.Find(ApplicationName, False)
                Else
                    'node = trvAppTree.TopNode.Nodes.Find(ApplicationName, False)
                    node = trvAppTree.Nodes(0).Nodes.Find(ApplicationName, False)
                End If
                If node Is Nothing Then
                    'Node not found.
                    Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                ElseIf node.Length = 0 Then
                    'Node not found.
                    Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                Else
                    'Add the project information to ProjInfo():
                    ProjInfo.Add(ProjectID & ".Proj", NewProjInfo)
                    'Add the project node to the Application Tree View:
                    trvAppTree.SelectedNode = node(0)
                    trvAppTree.SelectedNode.Nodes.Add(ProjectID & ".Proj", ProjectName, 0, 1) '0, 1 Default project icons.
                End If
            Case "Directory"
                'ProjInfo(ProjectID & ".Proj").Type = ADVL_Utilities_Library_1.Project.Types.Directory
                NewProjInfo.Type = ADVL_Utilities_Library_1.Project.Types.Directory
                'ProjInfo(ProjectID & ".Proj").IconNumber = 2
                NewProjInfo.IconNumber = 2
                'ProjInfo(ProjectID & ".Proj").OpenIconNumber = 3
                NewProjInfo.OpenIconNumber = 3
                Dim node As TreeNode()
                'If ApplicationName = trvAppTree.TopNode.Name Then
                If ApplicationName = trvAppTree.Nodes(0).Name Then
                    node = trvAppTree.Nodes.Find(ApplicationName, False)
                Else
                    'node = trvAppTree.TopNode.Nodes.Find(ApplicationName, False)
                    node = trvAppTree.Nodes(0).Nodes.Find(ApplicationName, False)
                End If
                If node Is Nothing Then
                    'Node not found.
                    Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                ElseIf node.Length = 0 Then
                    'Node not found.
                    Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                Else
                    'Add the project information to ProjInfo():
                    ProjInfo.Add(ProjectID & ".Proj", NewProjInfo)
                    'Add the project node to the Application Tree View:
                    trvAppTree.SelectedNode = node(0)
                    trvAppTree.SelectedNode.Nodes.Add(ProjectID & ".Proj", ProjectName, 2, 3) '2, 3 Directory project icons.
                End If
            Case "Archive"
                'ProjInfo(ProjectID & ".Proj").Type = ADVL_Utilities_Library_1.Project.Types.Archive
                NewProjInfo.Type = ADVL_Utilities_Library_1.Project.Types.Archive
                'ProjInfo(ProjectID & ".Proj").IconNumber = 4
                NewProjInfo.IconNumber = 4
                'ProjInfo(ProjectID & ".Proj").OpenIconNumber = 5
                NewProjInfo.OpenIconNumber = 5
                Dim node As TreeNode()
                'If ApplicationName = trvAppTree.TopNode.Name Then
                If ApplicationName = trvAppTree.Nodes(0).Name Then
                    node = trvAppTree.Nodes.Find(ApplicationName, False)
                Else
                    'node = trvAppTree.TopNode.Nodes.Find(ApplicationName, False)
                    node = trvAppTree.Nodes(0).Nodes.Find(ApplicationName, False)
                End If
                If node Is Nothing Then
                    'Node not found.
                    Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                ElseIf node.Length = 0 Then
                    'Node not found.
                    Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                Else
                    'Add the project information to ProjInfo():
                    ProjInfo.Add(ProjectID & ".Proj", NewProjInfo)
                    'Add the project node to the Application Tree View:
                    trvAppTree.SelectedNode = node(0)
                    trvAppTree.SelectedNode.Nodes.Add(ProjectID & ".Proj", ProjectName, 4, 5) '4, 5 Archive project icons.
                End If
            Case "Hybrid"
                'ProjInfo(ProjectID & ".Proj").Type = ADVL_Utilities_Library_1.Project.Types.Hybrid
                NewProjInfo.Type = ADVL_Utilities_Library_1.Project.Types.Hybrid
                'ProjInfo(ProjectID & ".Proj").IconNumber = 6
                NewProjInfo.IconNumber = 6
                'ProjInfo(ProjectID & ".Proj").OpenIconNumber = 7
                NewProjInfo.OpenIconNumber = 7
                Dim node As TreeNode()
                'If ApplicationName = trvAppTree.TopNode.Name Then
                If ApplicationName = trvAppTree.Nodes(0).Name Then
                    node = trvAppTree.Nodes.Find(ApplicationName, False)
                Else
                    'node = trvAppTree.TopNode.Nodes.Find(ApplicationName, False)
                    node = trvAppTree.Nodes(0).Nodes.Find(ApplicationName, False)
                End If

                If node Is Nothing Then
                    'Node not found.
                    Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                ElseIf node.Length = 0 Then
                    'Node not found.
                    Message.AddWarning("Application node not found for " & ApplicationName & vbCrLf)
                Else
                    'Add the project information to ProjInfo():
                    ProjInfo.Add(ProjectID & ".Proj", NewProjInfo)
                    'Add the project node to the Application Tree View:
                    trvAppTree.SelectedNode = node(0)
                    trvAppTree.SelectedNode.Nodes.Add(ProjectID & ".Proj", ProjectName, 6, 7) '6, 7 Hybrid project icons.
                End If
            Case Else
                Message.AddWarning("Unknown project type: " & ProjectType & vbCrLf)
        End Select
        'End If
    End Sub


    Private Function ProjectIdAvailable(ByVal ProjectID As String) As Boolean
        'If ProjectID is not in the Project list, ProjectIdAvaialable is set to True.

        Dim IdFound As Boolean = False
        Dim I As Integer 'Loop index
        For I = 0 To dgvProjects.Rows.Count - 1
            If dgvProjects.Rows(I).Cells(3).Value = ProjectID Then
                IdFound = True
                Exit For
            End If
        Next

        If IdFound = True Then
            Return False
        Else
            Return True
        End If

    End Function

    'Private Function ProjectNameAndAppNetNameAvailable(ByVal Name As String, ByVal AppNetName As String) As Boolean
    Private Function ProjectNameAndProNetNameAvailable(ByVal Name As String, ByVal ProNetName As String) As Boolean
        'If Name and ProNetName are not in the Project list, ProjectNameAndProNetNameAvailable is set to True.

        Dim Found As Boolean = False
        Dim I As Integer 'Loop index
        For I = 0 To dgvProjects.Rows.Count - 1
            If dgvProjects.Rows(I).Cells(0).Value = Name Then
                'If dgvProjects.Rows(I).Cells(1).Value = AppNetName Then
                If dgvProjects.Rows(I).Cells(1).Value = ProNetName Then
                    Found = True
                    Exit For
                End If
            End If
        Next

        If Found = True Then
            Return False
        Else
            Return True
        End If

    End Function

    Private Sub trvAppTree_DragEnter(sender As Object, e As DragEventArgs) Handles trvAppTree.DragEnter
        'DragEnter: An object has been dragged into the trvAppTree.

        'This code is required to get the link to the item(s) being dragged into the trvAppTree:
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Link
        End If
    End Sub

    Private Sub XmlHtmDisplay1_DragEnter(sender As Object, e As DragEventArgs) Handles XmlHtmDisplay1.DragEnter
        'DragEnter: An object has been dragged into XmlHtmDisplay1.
        'This code is required to get the link to the item(s) being dragged into XmlHtmDisplay1:
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Link
        End If
    End Sub

    Private Sub XmlHtmDisplay1_DragDrop(sender As Object, e As DragEventArgs) Handles XmlHtmDisplay1.DragDrop
        'DragDrop.

        Dim Path As String()

        Path = e.Data.GetData(DataFormats.FileDrop)

        'Check if file size is too large to display:
        Dim myFileInfo As New System.IO.FileInfo(Path(0))
        If myFileInfo.Length > XmlHtmDisplay1.Settings.XmlLargeFileSizeLimit Then
            Message.AddWarning("The file size is larger than the limit of " & XmlHtmDisplay1.Settings.XmlLargeFileSizeLimit & " bytes" & vbCrLf)
            Exit Sub
        End If

        Dim I As Integer

        If Path.Count > 0 Then
            'Open the XML file:
            Dim xmlDoc As New System.Xml.XmlDocument
            Try
                xmlDoc.Load(Path(0))
                XmlHtmDisplay1.Rtf = XmlHtmDisplay1.XmlToRtf(xmlDoc, True)
                Label26.Text = "Filename: " & System.IO.Path.GetFileName(Path(0))
            Catch ex As Exception
                Message.AddWarning("Error displaying XML file: " & ex.Message & vbCrLf)
            End Try
            If Path.Count > 1 Then
                Message.AddWarning("More than one file was dragged into the XML display window. Only the first will be displayed." & vbCrLf)
            End If
        End If

    End Sub



#End Region 'Drag Drop Project Code -----------------------------------------------------------------------



#Region " Send XMessages" '==========================================================================================

    Private Sub SendMessage()
        'Code used to send a message after a timer delay.
        'The message destination is stored in MessageDest
        'The message text is stored in MessageText
        Timer1.Interval = 100 '100ms delay
        Timer1.Enabled = True 'Start the timer.
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        'Stop timer:
        Timer1.Enabled = False

        'NOTE: There is no SendMessage code in the Message Service application!
    End Sub

    Private Sub XMsg_Instruction(Data As String, Locn As String) Handles XMsg.Instruction
        'Process an XMessage instruction.
        'An XMessage is a simplified XSequence. It is used to exchange information between Andorville™ applications.
        '
        'An XSequence file is an AL-H7™ Information Sequence stored in an XML format.
        'AL-H7™ is the name of a programming system that uses sequences of data and location value pairs to store information or processing steps.
        'Any program, mathematical expression or data set can be expressed as an Information Sequence.

        'Add code here to process the XMessage instructions.
        'See other Andorville™ applications for examples.

        If IsDBNull(Data) Then
            Data = ""
        End If

        'Intercept instructions with the prefix "WebPage_"
        If Locn.StartsWith("WebPage_") Then 'Send the Data, Location data to the correct Web Page:
            'Message.Add("Web Page Location: " & Locn & vbCrLf)
            If Locn.Contains(":") Then
                Dim EndOfWebPageNoString As Integer = Locn.IndexOf(":")
                If Locn.Contains("-") Then
                    Dim HyphenLocn As Integer = Locn.IndexOf("-")
                    If HyphenLocn < EndOfWebPageNoString Then 'Web Page Location contains a sub-location in the web page - WebPage_1-SubLocn:Locn - SubLocn:Locn will be sent to Web page 1
                        EndOfWebPageNoString = HyphenLocn
                    End If
                End If
                Dim PageNoLen As Integer = EndOfWebPageNoString - 8
                Dim WebPageNoString As String = Locn.Substring(8, PageNoLen)
                Dim WebPageNo As Integer = CInt(WebPageNoString)
                Dim WebPageData As String = Data
                Dim WebPageLocn As String = Locn.Substring(EndOfWebPageNoString + 1)

                'Message.Add("WebPageData = " & WebPageData & "  WebPageLocn = " & WebPageLocn & vbCrLf)

                WebPageFormList(WebPageNo).XMsgInstruction(WebPageData, WebPageLocn)
            Else
                Message.AddWarning("XMessage instruction location is not complete: " & Locn & vbCrLf)
            End If
        Else

            Select Case Locn

                Case "ClientProNetName"
                    ClientProNetName = Data 'The name of the Client Application Network requesting service. AD

                Case "ClientName"
                    ClientAppName = Data 'The name of the Client application requesting service.

                Case "ClientConnectionName"
                    ClientConnName = Data 'The name of the client connection requesting service.

                Case "ClientLocn" 'The Location within the Client requesting service.
                    Dim statusOK As New XElement("Status", "OK") 'Add Status OK element when the Client Location is changed
                    xlocns(xlocns.Count - 1).Add(statusOK)

                    xmessage.Add(xlocns(xlocns.Count - 1)) 'Add the instructions for the last location to the reply xmessage
                    xlocns.Add(New XElement(Data)) 'Start the new location instructions

                'Case "OnCompletion" 'Specify the last instruction to be returned on completion of the XMessage processing.
                '    CompletionInstruction = Info
                    'UPDATE:
                Case "OnCompletion"
                    OnCompletionInstruction = Data

                Case "Main"
                'Blank message - do nothing.

                'Case "Main:OnCompletion"
                '    Select Case "Stop"
                '        'Stop on completion of the instruction sequence.

                '    End Select

                Case "Main:EndInstruction"
                    Select Case Data
                        Case "Stop"
                            'Stop at the end of the instruction sequence.

                            'Add other cases here:
                    End Select

                Case "Main:Status"
                    Select Case Data
                        Case "OK"
                            'Main instructions completed OK
                    End Select

                Case "AppComCheck:ClientProNetName"
                    ACC_ProNetName = Data 'The Project Network Name of the application communication check

                Case "AppComCheck:ClientName"
                'Not currently used.

                Case "AppComCheck:ClientConnectionName"
                    ACC_ConnName = Data 'The Connection Name of the application communication check

                Case "AppComCheck:OnCompletion"
                    Select Case "Stop"
                        'Stop on completion of the instruction sequence.
                    End Select

                Case "AppComCheck:Status"
                    Select Case Data
                        Case "OK"
                            AppComCheckStatus = Data
                            'Set the Status Field in dgvConnections to OK:
                            SetConnectionStatus(ACC_ProNetName, ACC_ConnName, "OK")
                    End Select

            'Case "NewConnectionInfo:ApplicationNetworkName"
            '    SelectedAppNetName = Info
                Case "NewConnectionInfo:ProjectNetworkName"
                    SelectedProNetName = Data

                Case "NewConnectionInfo:ConnectionName"

                    'If ConnectionNameAvailable(SelectedAppNetName, Info) Then
                    If ConnectionNameAvailable(SelectedProNetName, Data) Then
                        AddNewConnection = True
                        dgvConnections.Rows.Add()
                        'Dim CurrentRow As Integer = dgvConnections.Rows.Count - 2
                        Dim CurrentRow As Integer = dgvConnections.Rows.Count - 1 'Last blank row now longer showing: 2 changed to 1
                        'dgvConnections.Rows(CurrentRow).Cells(0).Value = SelectedAppNetName 'Add the AppNet Name to dgvConnections.
                        dgvConnections.Rows(CurrentRow).Cells(0).Value = SelectedProNetName 'Add the ProNet Name to dgvConnections.
                        dgvConnections.Rows(CurrentRow).Cells(2).Value = Data 'Add the ConnectionName to dgvConnections.
                        dgvConnections.AutoResizeRows()
                    Else
                        AddNewConnection = False
                    End If

                Case "NewConnectionInfo:ApplicationName"
                    If AddNewConnection = True Then
                        'Dim CurrentRow As Integer = dgvConnections.Rows.Count - 2
                        Dim CurrentRow As Integer = dgvConnections.Rows.Count - 1 'Last blank row now longer showing: 2 changed to 1
                        dgvConnections.Rows(CurrentRow).Cells(1).Value = Data 'Add the ApplicationName to dgvConnections.
                        dgvConnections.AutoResizeRows()
                    End If

                Case "NewConnectionInfo:ProjectName"
                    If AddNewConnection = True Then
                        'Dim CurrentRow As Integer = dgvConnections.Rows.Count - 2
                        Dim CurrentRow As Integer = dgvConnections.Rows.Count - 1 'Last blank row now longer showing: 2 changed to 1
                        dgvConnections.Rows(CurrentRow).Cells(3).Value = Data
                        dgvConnections.AutoResizeRows()
                    End If

                Case "NewConnectionInfo:ProjectType"
                    If AddNewConnection = True Then
                        'Dim CurrentRow As Integer = dgvConnections.Rows.Count - 2
                        Dim CurrentRow As Integer = dgvConnections.Rows.Count - 1 'Last blank row now longer showing: 2 changed to 1
                        dgvConnections.Rows(CurrentRow).Cells(4).Value = Data
                        dgvConnections.AutoResizeRows()
                    End If

                Case "NewConnectionInfo:ProjectPath"
                    If AddNewConnection = True Then
                        'Dim CurrentRow As Integer = dgvConnections.Rows.Count - 2
                        Dim CurrentRow As Integer = dgvConnections.Rows.Count - 1 'Last blank row now longer showing: 2 changed to 1
                        dgvConnections.Rows(CurrentRow).Cells(5).Value = Data
                        dgvConnections.AutoResizeRows()
                    End If

                Case "NewConnectionInfo:GetAllWarnings"
                    If AddNewConnection = True Then
                        'Dim CurrentRow As Integer = dgvConnections.Rows.Count - 2
                        Dim CurrentRow As Integer = dgvConnections.Rows.Count - 1 'Last blank row now longer showing: 2 changed to 1
                        dgvConnections.Rows(CurrentRow).Cells(6).Value = Data
                        dgvConnections.AutoResizeRows()
                    End If

                Case "NewConnectionInfo:GetAllMessages"
                    If AddNewConnection = True Then
                        'Dim CurrentRow As Integer = dgvConnections.Rows.Count - 2
                        Dim CurrentRow As Integer = dgvConnections.Rows.Count - 1 'Last blank row now longer showing: 2 changed to 1
                        dgvConnections.Rows(CurrentRow).Cells(7).Value = Data
                        dgvConnections.AutoResizeRows()
                    End If

                Case "NewConnectionInfo:CallbackHashcode"
                    If AddNewConnection = True Then
                        'Dim CurrentRow As Integer = dgvConnections.Rows.Count - 2
                        Dim CurrentRow As Integer = dgvConnections.Rows.Count - 1 'Last blank row now longer showing: 2 changed to 1
                        dgvConnections.Rows(CurrentRow).Cells(8).Value = Data
                        dgvConnections.AutoResizeRows()
                    End If

                Case "NewConnectionInfo:ConnectionStartTime"
                    If AddNewConnection = True Then
                        'Dim CurrentRow As Integer = dgvConnections.Rows.Count - 2
                        Dim CurrentRow As Integer = dgvConnections.Rows.Count - 1 'Last blank row now longer showing: 2 changed to 1
                        dgvConnections.Rows(CurrentRow).Cells(9).Value = Data
                        dgvConnections.Rows(CurrentRow).Cells(10).Value = 0 'Current duration of the connection
                        dgvConnections.AutoResizeRows()
                    End If

                    '---------------------------------------------------------------------------------------------------------------------------------------------


           'Add an Application Info entry ---------------------------------------------------------------------------------------------------------------
                Case "ApplicationInfo:Name"
                    'Code used to add application to the Application List: (TO BE REPLACED WITH THE APPLICATION DICTIONARY.)
                    If ApplicationNameAvailable(Data) Then
                        AddNewApplication = True 'Add new application to App.List
                        dgvApplications.Rows.Add()
                        Dim CurrentRow As Integer = dgvApplications.Rows.Count - 2
                        dgvApplications.Rows(CurrentRow).Cells(0).Value = Data
                        Dim NewAppInfo As New AppSummary
                        NewAppInfo.Name = Data
                        App.List.Add(NewAppInfo)

                    Else
                        AddNewApplication = False
                        'ApplicationNameAvailable() will also have set ApplicationNo to point to the existing entry in App.List
                    End If
                    AppName = Data
                    If AppInfo.ContainsKey(Data) Then 'The Application is already in the Application Tree.
                        AddNewApp = False
                        AppName = Data
                    Else 'Store the application information - this will be used to add an applcation node to the tree later.
                        AddNewApp = True 'Add new application to AppInfo()
                        AppInfo.Add(Data, New clsAppInfo) 'Text, Description, ExecutablePath, 
                        AppName = Data
                    End If

                Case "ApplicationInfo:Text"
                    If AddNewApp = True Then
                        AppText = Data 'Save the Application taxt so that it can be displayed with the application node when it is created later.
                    Else

                        'Update the node text. 
                        Dim node As TreeNode() = trvAppTree.Nodes.Find(AppName, True)
                        If node Is Nothing Then
                            'Node key not found.
                            Message.AddWarning("No node found with the name: " & AppName & vbCrLf)
                        Else
                            If node.Count > 1 Then
                                Message.AddWarning("More than one node found with the name: " & AppName & vbCrLf)
                            Else
                                node(0).Text = Data
                            End If
                        End If
                    End If

                Case "ApplicationInfo:Directory"
                    If AddNewApplication = True Then
                        Dim CurrentRow As Integer = dgvApplications.Rows.Count - 2
                        'Applications grid now shows only Name and Description
                        App.List(CurrentRow).Directory = Data
                    Else
                        If App.List(ApplicationNo).Directory = Data Then
                            'Directory has not been changed.
                        Else
                            'Directory has been changed.
                            App.List(ApplicationNo).Directory = Data
                            Message.Add("Application directory for " & App.List(ApplicationNo).Name & " has been changed to: " & vbCrLf & App.List(ApplicationNo).Directory & vbCrLf)
                        End If
                    End If

                    If AddNewApp = True Then
                        AppInfo(AppName).Directory = Data
                    Else
                        If AppInfo(AppName).Directory = Data Then
                            'The application directory is unchanged.
                        Else
                            AppInfo(AppName).Directory = Data 'The application directory has been updated.
                        End If
                    End If

                Case "ApplicationInfo:Description"
                    If AddNewApplication = True Then
                        Dim CurrentRow As Integer = dgvApplications.Rows.Count - 2
                        dgvApplications.Rows(CurrentRow).Cells(1).Value = Data
                        dgvApplications.AutoResizeRows()
                        App.List(CurrentRow).Description = Data

                    Else
                        If App.List(ApplicationNo).Description = Data Then
                            'Executable path has not been changed.
                        Else
                            'Executable path has been changed.
                            App.List(ApplicationNo).Description = Data
                            Message.Add("Application description for " & App.List(ApplicationNo).Name & " has been changed to: " & vbCrLf & App.List(ApplicationNo).Description & vbCrLf)
                        End If
                    End If

                    If AddNewApp = True Then
                        AppInfo(AppName).Description = Data
                    Else
                        If AppInfo(AppName).Description = Data Then
                            'The application description is unchanged.
                        Else
                            AppInfo(AppName).Description = Data 'The application description has been updated.
                        End If
                    End If

                Case "ApplicationInfo:ExecutablePath"
                    If AddNewApplication = True Then
                        Dim CurrentRow As Integer = dgvApplications.Rows.Count - 2
                        'Applications grid now shows only Name and Description
                        App.List(CurrentRow).ExecutablePath = Data
                    Else
                        If App.List(ApplicationNo).ExecutablePath = Data Then
                            'Executable path has not been changed.
                        Else
                            'Executable path has been changed.
                            App.List(ApplicationNo).ExecutablePath = Data
                            Message.Add("Executable path for " & App.List(ApplicationNo).Name & " has been changed to: " & vbCrLf & App.List(ApplicationNo).ExecutablePath & vbCrLf)
                        End If
                    End If

                    If AddNewApp = True Then
                        AppInfo(AppName).ExecutablePath = Data
                        'Get the application icon:
                        Dim myIcon = System.Drawing.Icon.ExtractAssociatedIcon(Data)
                        AppTreeImageList.Images.Add(AppName, myIcon)
                        AppInfo(AppName).IconNumber = AppTreeImageList.Images.IndexOfKey(AppName)
                        AppInfo(AppName).OpenIconNumber = AppTreeImageList.Images.IndexOfKey(AppName)
                    Else
                        If AppInfo(AppName).ExecutablePath = Data Then
                            'The application executable path is unchanged.
                        Else
                            AppInfo(AppName).ExecutablePath = Data 'The application executable path has been updated.
                        End If
                    End If

            '--------------------------------------------------------------------------------------------------------------------------------------------



           'Add a Project -------------------------------------------------------------------------------------------------------------------------------

                Case "ProjectInfo:Path"
                    ProcessNewProject(Data)

           '---------------------------------------------------------------------------------------------------------------------------------------------

            'Remove a connection entry ------------------------------------------------------------------------------------------------------------------

            'Case "RemovedConnectionInfo:ApplicationNetworkName"
                Case "RemovedConnectionInfo:ProjectNetworkName"
                    SelectedProNetName = Data

                Case "RemovedConnectionInfo:ApplicationName"
                    Message.AddWarning("Instruction not in use: RemovedConnectionInfo:ApplicationName" & vbCrLf)
                    Message.AddWarning("Modify your code to use this instruction: RemovedConnectionInfo:ConnectionName" & vbCrLf)

                Case "RemovedConnectionInfo:ConnectionName"
                    RemoveConnectionWithName(SelectedProNetName, Data)

          '---------------------------------------------------------------------------------------------------------------------------------------------


          'Check the Connection =====================================================

                Case "Connection"
                    Select Case Data
                        Case "Check"
                            ConnectionCheck = True 'This indicates a successful connection check.
                    End Select
          '--------------------------------------------------------------------------

          'Start a Project -----------------------------------------------------------------------------------------------------------------------------
          'Utility Methods -----------------------------------------------------------------------------------------------------------------------------

                Case "StartProject:AppName"
                    StartProject_AppName = Data
                Case "StartProject:ConnectionName"
                    StartProject_ConnName = Data
                Case "StartProject:ProNetName"
                    StartProject_ProNetName = Data
                Case "StartProject:ProjectName"
                    StartProject_ProjName = Data
                Case "StartProject:Command"
                    Select Case Data
                        Case "Apply"
                            If ConnectionNameAvailable(StartProject_ProNetName, StartProject_ConnName) Then
                                StartProject(StartProject_ProjName, StartProject_ProNetName, StartProject_AppName, StartProject_ConnName)
                            Else
                                Message.AddWarning("Start Project failed because [" & StartProject_ProNetName & "]." & StartProject_ConnName & " is already on the connections list." & vbCrLf)
                            End If

                        Case Else
                            Message.AddWarning("Unknown Start Project Command: " & Data & vbCrLf)
                    End Select

                '---------------------------------------------------------------------------------------------------------------------------------------------

                'Utility Methods -----------------------------------------------------------------------------------------------------------------------------
                Case "Command"
                    Select Case Data
                        Case "GetApplicationList"
                            'Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                            'Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                            'Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
                            Dim applicationList As New XElement("ApplicationList")

                            For Each item In App.List
                                Dim applicationInfo As New XElement("Application")
                                Dim name As New XElement("Name", item.Name)
                                applicationInfo.Add(name)
                                'Message.Add("NApplication name: " & item.Name & vbCrLf) 'For testing
                                Dim description As New XElement("Description", item.Description)
                                applicationInfo.Add(description)
                                Dim directory As New XElement("Directory", item.Directory)
                                applicationInfo.Add(directory)
                                Dim executablePath As New XElement("ExecutablePath", item.ExecutablePath)
                                applicationInfo.Add(executablePath)
                                applicationList.Add(applicationInfo)
                            Next
                            'xmessage.Add(applicationList)
                            'doc.Add(xmessage)

                            'xlocns(xlocns.Count - 1).Add(doc)
                            xlocns(xlocns.Count - 1).Add(applicationList)
                    End Select
          '---------------------------------------------------------------------------------------------------------------------------------------------


                Case "EndOfSequence"
                    If AddNewApp = True Then
                        'Add the new application node to the tree: 
                        'NOTE: If the user has scrolled the TreeView, the tree node at the top may not be the first root tree node!
                        'trvAppTree.TopNode.Nodes.Add(AppName, AppText, AppInfo(AppName).IconNumber, AppInfo(AppName).OpenIconNumber)

                        'Try this to always add a new node to the first root tree node: (Use .Nodes(0) instead of .TopNode)
                        trvAppTree.Nodes(0).Nodes.Add(AppName, AppText, AppInfo(AppName).IconNumber, AppInfo(AppName).OpenIconNumber)

                    End If
                    AddNewConnection = False
                    AddNewApplication = False
                    AddNewApp = False
                    AppName = ""
                    StartAppName = ""
                    StartAppConnName = ""
                    'StartAppProject = ""
                    StartAppProjectName = ""
                    StartAppProjectID = ""
                    StartAppProjectPath = ""

                    'SelectedAppNetName = ""
                    SelectedProNetName = ""

                    'Clear the Application Communication Check variables:
                    ACC_ProNetName = ""
                    ACC_ConnName = ""

                    'Clear Start Project variables:
                    StartProject_AppName = ""
                    StartProject_ConnName = ""
                    StartProject_ProNetName = ""
                    StartProject_ProjID = ""
                    StartProject_ProjName = ""

                    Dim statusOK As New XElement("Status", "OK")
                    xlocns(xlocns.Count - 1).Add(statusOK)

                    ''Add the final OnCompletion instruction:
                    'Dim onCompletion As New XElement("OnCompletion", CompletionInstruction) '
                    'xlocns(xlocns.Count - 1).Add(onCompletion)
                    'CompletionInstruction = "Stop" 'Reset the Completion Instruction

                    Select Case EndInstruction
                        Case "Stop"
                            'No instructions.

                            'Add any other Cases here:

                        Case Else
                            Message.AddWarning("Unknown End Instruction: " & EndInstruction & vbCrLf)
                    End Select
                    EndInstruction = "Stop"

                    ''Final Version:
                    ''Add the final EndInstruction:
                    'Dim xEndInstruction As New XElement("EndInstruction", OnCompletionInstruction)
                    'xlocns(xlocns.Count - 1).Add(xEndInstruction)
                    'OnCompletionInstruction = "Stop" 'Reset the OnCompletion Instruction

                    'Add the final EndInstruction:
                    If OnCompletionInstruction = "Stop" Then
                        'Final EndInstruction is not required.
                    Else
                        Dim xEndInstruction As New XElement("EndInstruction", OnCompletionInstruction)
                        xlocns(xlocns.Count - 1).Add(xEndInstruction)
                        OnCompletionInstruction = "Stop" 'Reset the OnCompletion Instruction
                    End If

                Case Else
                    'Message.Add("Instruction not recognised:  " & Locn & "    Property:  " & Data & vbCrLf)
                    Message.AddWarning("Unknown location: " & Locn & vbCrLf)
                    Message.AddWarning("            data: " & Data & vbCrLf)
            End Select
        End If
    End Sub

    Private Sub SetConnectionStatus(ByVal ProNetName As String, ByVal ConnName As String, ByVal Status As String)
        'Set the status field if the specified row in dgvConnections to the specified Status string

        'Find the row in dgvConnections with Project Network Name = ProNetName and Connection Name = ConnName.
        Dim I As Integer 'Loop index
        For I = 0 To dgvConnections.Rows.Count - 1
            If dgvConnections.Rows(I).Cells(2).Value = ConnName Then
                If dgvConnections.Rows(I).Cells(0).Value = ProNetName Then
                    'The Connection named ConnName has been found in the Project Network named ProNetName.
                    dgvConnections.Rows(I).DefaultCellStyle.ForeColor = Color.Black 'ADDED 9Apr20
                    dgvConnections.Rows(I).Cells(11).Value = Status
                    Exit For
                End If
            End If
        Next

    End Sub


    Private Sub chkWrapText_CheckedChanged(sender As Object, e As EventArgs) Handles chkWrapText.CheckedChanged
        If chkWrapText.Checked Then
            XmlHtmDisplay1.WordWrap = True
        Else
            XmlHtmDisplay1.WordWrap = False
        End If
    End Sub

    Private Sub Label34_Click(sender As Object, e As EventArgs) Handles Label34.Click
        XmlHtmDisplay1.Clear()
        Label26.Text = "Filename:"
    End Sub

    Private Sub ToolStripMenuItem1_OpenProject_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1_OpenProject.Click
        StartProject()
    End Sub



    Private Sub Label26_Click(sender As Object, e As EventArgs) Handles Label26.Click

    End Sub

    Private Sub Label26_MouseHover(sender As Object, e As EventArgs) Handles Label26.MouseHover
        'Update the ToolTip text:
        ToolTip1.SetToolTip(Label26, Label26.Text) 'This will allow the full filename to be read if it is cropped at the edge of the window.
    End Sub


    Private Sub XmlHtmDisplay2_DragEnter(sender As Object, e As DragEventArgs) Handles XmlHtmDisplay2.DragEnter
        'DragEnter: An object has been dragged into XmlHtmDisplay2 - View HTML tab.
        'This code is required to get the link to the item(s) being dragged into XmlHtmDisplay2:
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Link
        End If

    End Sub

    Private Sub XmlHtmDisplay2_DragDrop(sender As Object, e As DragEventArgs) Handles XmlHtmDisplay2.DragDrop
        'An item has been DragDropped into the View HTML tab.

        Dim Path As String()

        Path = e.Data.GetData(DataFormats.FileDrop)

        Dim I As Integer

        If Path.Count > 0 Then
            'Open the HTML file:

            Try
                Dim rtbData As New IO.MemoryStream
                Dim fileStream As New IO.FileStream(Path(0), System.IO.FileMode.Open)
                Dim myData(fileStream.Length) As Byte
                fileStream.Read(myData, 0, fileStream.Length) 'Additional information: Buffer cannot be null.
                Dim streamWriter As New IO.BinaryWriter(rtbData)
                streamWriter.Write(myData)
                fileStream.Close()

                XmlHtmDisplay2.Clear()
                rtbData.Position = 0

                XmlHtmDisplay2.LoadFile(rtbData, RichTextBoxStreamType.PlainText)
                Dim htmText As String = XmlHtmDisplay2.Text
                XmlHtmDisplay2.Rtf = XmlHtmDisplay2.HmlToRtf(htmText)

                Label39.Text = "Filename: " & System.IO.Path.GetFileName(Path(0))
            Catch ex As Exception
                Message.AddWarning("Error displaying HTML file: " & ex.Message & vbCrLf)
            End Try
            If Path.Count > 1 Then
                Message.AddWarning("More than one file was dragged into the HTML display window. Only the first will be displayed." & vbCrLf)
            End If
        End If

    End Sub

    Private Sub Label38_Click(sender As Object, e As EventArgs) Handles Label38.Click
        XmlHtmDisplay2.Clear()
        Label39.Text = "Filename:"

    End Sub




#End Region 'Send XMessages -----------------------------------------------------------------------------------------



    'Public Sub StartApp_ProjectName(ByVal AppName As String, ByVal AppNetName As String, ByVal ProjectName As String, ByVal ConnectionName As String)
    Public Sub StartApp_ProjectName(ByVal AppName As String, ByVal ProNetName As String, ByVal ProjectName As String, ByVal ConnectionName As String)
        'Start the application with the name AppName.

        If AppInfo.ContainsKey(AppName) Then
            'Start the application:
            If ProjectName = "" And ConnectionName = "" Then
                'No project selected and application will not be connected to the network.
                Shell(Chr(34) & AppInfo(AppName).ExecutablePath & Chr(34), AppWinStyle.NormalFocus) 'Start the application with no argument
            Else
                If ConnectionNameAvailable(ProNetName, ConnectionName) Then

                    Dim ProjectPath As String = Proj.FindNameAndAppNet(ProjectName, ProNetName).Path

                    'COMMENTED OUT 16Feb2021 - These lines return a blank ProjectPath when ADVL_Project_Network_1 application opens the Share Analysis System project:
                    ''Temp code - until Project Network projects have the ProNetName added:
                    'If AppName = "ADVL_Project_Network_1" Then
                    '    ProjectPath = Proj.FindNameAndAppNet(ProjectName, "").Path
                    'End If

                    If ProjectPath = "" Then
                        Message.AddWarning("Project Path not found for Project: " & ProjectName & " in ProNet: " & ProNetName & vbCrLf)
                        Exit Sub
                    End If

                    'Build the Application start message:
                    Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                    Dim ConnectDoc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                    Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class

                    'NOTE: The project should always be opened using a Project Path because project names may not be unique.
                    'Use the ProNetName and ProjectName to get the ProjectPath from the Project List.

                    Dim xProjectPath As New XElement("ProjectPath", ProjectPath)
                    xmessage.Add(xProjectPath)

                    'NOTE: This is redundant - only the Project Path is required.
                    ''NOTE: The application currently determines the ProNetName using other information!
                    'If ProNetName <> "" Then
                    '    Dim xProNetName As New XElement("ProNetName", ProNetName)
                    '    xmessage.Add(xProNetName)
                    'End If

                    If ConnectionName <> "" Then
                        Dim xconnection As New XElement("ConnectionName", ConnectionName)
                        xmessage.Add(xconnection)
                    End If
                    ConnectDoc.Add(xmessage)
                    If System.IO.File.Exists(AppInfo(AppName).ExecutablePath) Then
                        Shell(Chr(34) & AppInfo(AppName).ExecutablePath & Chr(34) & " " & Chr(34) & ConnectDoc.ToString & Chr(34), AppWinStyle.NormalFocus)
                    Else
                        Message.AddWarning("Application " & AppName & " executable path was not found: " & AppInfo(AppName).ExecutablePath & vbCrLf)
                    End If
                Else
                    Message.AddWarning("Connection name already in use: ConnName: " & ConnectionName & " in the Project Network: " & ProNetName & vbCrLf)
                End If
            End If
        End If
    End Sub

    Private Sub btnAppInfo_Click(sender As Object, e As EventArgs) Handles btnAppInfo.Click
        ApplicationInfo.ShowInfo()
    End Sub

    Private Sub btnAndorville_Click(sender As Object, e As EventArgs) Handles btnAndorville.Click
        ApplicationInfo.ShowInfo()
    End Sub

    'Private Sub Timer3_Tick(sender As Object, e As EventArgs)
    '    'Private Async Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
    '    'Keet the connection awake with each tick:

    '    'NOTE: client.IsAlive() does not work correctly here.
    '    '      client.IsAliveAsync() does not work correctly here.   


    '    ''Send a message to Message_Service_Form
    '    'Send a message to ADVL_Network_1
    '    'If the message is received, the connection is working OK.
    '    'This also resets the timeout period.
    '    'If the message is not received after a set wait time, the connection is assumed to be faulted.

    '    If ConnectionCheckStatus = "Waiting" Then
    '        Dim Duration As TimeSpan = Now - ConnectionCheckStart
    '        'Check if the connection check has passed:
    '        If ConnectionCheck = True Then
    '            'Connection is working.
    '            Message.Add(Format(Now, "HH:mm:ss") & " Connection OK." & vbCrLf)
    '            ConnectionCheckStatus = "Passed"
    '            'Timer3.Interval = 10000 '10 seconds - For testing only
    '            Timer3.Interval = TimeSpan.FromMinutes(55).TotalMilliseconds '55 minute interval
    '        Else
    '            Message.Add(Format(Now, "HH:mm:ss") & " Waiting for Connection check." & vbCrLf)
    '            If Duration.Seconds > 60 Then
    '                ConnectionCheckStatus = "Failed"
    '                Message.Add(Format(Now, "HH:mm:ss") & " Connection Fault." & vbCrLf)
    '                'Timer3.Interval = 20000 '10 seconds - For testing only
    '                Timer3.Interval = TimeSpan.FromMinutes(55).TotalMilliseconds '55 minute interval
    '            Else
    '                'Keep waiting
    '                Timer3.Interval = 4000 '4 seconds - Check again in 4 seconds.
    '            End If
    '        End If
    '    Else
    '        'Start a new Connection Check:

    '        ConnectionCheck = False 'Set this to False. If the connection is working, it will change this to True.

    '        ''Generate the XMessage to check the Message_Service_Form connection:
    '        'Generate the XMessage to check the ADVL_Network_1 connection:
    '        Dim decl As New XDeclaration("1.0", "utf-8", "yes")
    '        Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
    '        Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
    '        'Dim connCheck As New XElement("ConnectionCheck")
    '        'Dim connCheckCommand As New XElement("Command", "Apply")
    '        'connCheck.Add(connCheckCommand)
    '        Dim connCheck As New XElement("Connection", "Check")
    '        xmessage.Add(connCheck)
    '        doc.Add(xmessage)

    '        'Show the message sent to ComNet:
    '        'Message.XAddText("Message sent to " & "Message_Service_Form" & ":" & vbCrLf, "XmlSentNotice")
    '        Message.XAddText("Message sent to " & "ADVL_Network_1" & ":" & vbCrLf, "XmlSentNotice")
    '        Message.XAddXml(doc.ToString)
    '        Message.XAddText(vbCrLf, "Normal") 'Add extra line

    '        'client.SendMessageAsync("", "Message_Service_Form", doc.ToString)
    '        client.SendMessageAsync("", "ADVL_Network_1", doc.ToString)

    '        ConnectionCheckStart = Now
    '        ConnectionCheckStatus = "Waiting"
    '        'Message.Add(Format(Now, "HH:mm:ss") & " Waiting for Connection check." & vbCrLf)
    '        Timer3.Interval = 4000 '4 seconds - Check in 4 seconds.

    '        ''Check if the connection check has passed:
    '        'If ConnectionCheck = True Then 'VARIABLE ConnectionCheck IS NEVER UPDATED FAST ENOUGH TO GET HERE
    '        '    'Connection is working.
    '        '    Message.Add(Format(Now, "HH:mm:ss") & " Connection OK." & vbCrLf)
    '        '    Timer3.Interval = TimeSpan.FromMinutes(55).TotalMilliseconds '55 minute interval
    '        'Else
    '        '    ConnectionCheckStart = Now
    '        '    ConnectionCheckStatus = "Waiting"
    '        '    Message.Add(Format(Now, "HH:mm:ss") & " Waiting for Connection check." & vbCrLf)
    '        '    Timer3.Interval = 10000 '10 seconds - Check again in 10 seconds.
    '        'End If

    '    End If















    '    'Try

    '    '    'Dim ClientIsAlive As Boolean
    '    '    'Dim task As

    '    '    'Dim ClientIsAliveTask As Task(Of Boolean) = client.IsAliveAsync()
    '    '    'Dim ClientIsAlive As Boolean = Await ClientIsAliveTask()
    '    '    'Dim ClientIsAliveTask As Task(Of Boolean) = client.IsAliveAsync()
    '    '    ' Dim ClientIsAlive As Boolean = ClientIsAliveTask.

    '    '    'ClientIsAliveTask.Start() 'ERROR: Start may not be called on a promise-style task.
    '    '    'ClientIsAliveTask.Wait() 'One or more errors occurred.
    '    '    'One or more errors occurred.

    '    '    'Dim ClientIsActive As Boolean = client.IsAliveAsync().Result
    '    '    Dim ClientIsActive As Boolean = client.IsAliveAsync().Result


    '    '    'If client.IsAlive() Then
    '    '    'If ClientIsAliveTask.Result Then
    '    '    If ClientIsActive Then
    '    '        'If ClientIsAliveTask. Then
    '    '        Message.Add(Format(Now, "HH:mm:ss") & " Connection OK." & vbCrLf)
    '    '        Timer3.Interval = TimeSpan.FromMinutes(55).TotalMilliseconds '55 minute interval
    '    '    Else
    '    '        Message.Add(Format(Now, "HH:mm:ss") & " Connection Fault." & vbCrLf)
    '    '        Timer3.Interval = TimeSpan.FromMinutes(55).TotalMilliseconds '55 minute interval
    '    '    End If
    '    'Catch ex As Exception
    '    '    Message.AddWarning(ex.Message & vbCrLf)
    '    '    'Set interval to five minutes - try again in five minutes:
    '    '    Timer3.Interval = TimeSpan.FromMinutes(5).TotalMilliseconds '5 minute interval
    '    'End Try


    'End Sub



    Private Sub btnAlignMessageWindow_Click(sender As Object, e As EventArgs) Handles btnAlignMessageWindow.Click
        'Align the Message Window for the Selected connection.
        'The window will be aligned with the Network message window.

        Dim ConnectionName As String
        Dim ProNetName As String
        Dim SelRow As Integer
        If dgvConnections.SelectedRows.Count > 0 Then
            If dgvConnections.SelectedRows.Count = 1 Then
                SelRow = dgvConnections.SelectedRows(0).Index
                ProNetName = dgvConnections.Rows(SelRow).Cells(0).Value
                ConnectionName = dgvConnections.Rows(SelRow).Cells(2).Value

                'Check first if the ADVL_Network_1 connection has been selected:
                If ConnectionName = "ADVL_Network_1" Then 'Show the Network message window:
                    'Show the Messages form.
                    Message.ApplicationName = ApplicationInfo.Name
                    Message.SettingsLocn = Project.SettingsLocn
                    Message.Show()
                    Message.MessageForm.BringToFront()
                Else 'Align the Message window of the slected connection:
                    Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                    Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                    Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class

                    'Get the Message Window position:
                    If IsNothing(Message.MessageForm) Then
                        'Show the Messages form.
                        Message.ApplicationName = ApplicationInfo.Name
                        Message.SettingsLocn = Project.SettingsLocn
                        Message.Show()
                        Message.MessageForm.BringToFront()
                    End If
                    Dim WindowLeft As Integer = Message.MessageForm.Left
                    Dim WindowTop As Integer = Message.MessageForm.Top
                    Dim WindowWidth As Integer = Message.MessageForm.Width
                    Dim WindowHeight As Integer = Message.MessageForm.Height

                    Dim msgWindow As New XElement("MessageWindow")
                    Dim msgWindowLeft As New XElement("Left", WindowLeft)
                    msgWindow.Add(msgWindowLeft)
                    Dim msgWindowTop As New XElement("Top", WindowTop)
                    msgWindow.Add(msgWindowTop)
                    Dim msgWindowWidth As New XElement("Width", WindowWidth)
                    msgWindow.Add(msgWindowWidth)
                    Dim msgWindowHeight As New XElement("Height", WindowHeight)
                    msgWindow.Add(msgWindowHeight)
                    Dim msgWindowSaveSettings As New XElement("Command", "SaveSettings")
                    msgWindow.Add(msgWindowSaveSettings)
                    Dim msgWindowCommand As New XElement("Command", "BringToFront")
                    msgWindow.Add(msgWindowCommand)
                    xmessage.Add(msgWindow)
                    doc.Add(xmessage)

                    Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnectionName & ":" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(doc.ToString)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line

                    client.SendMessageAsync(ProNetName, ConnectionName, doc.ToString)
                End If
            End If
        Else
            Message.AddWarning("No connections have been selected." & vbCrLf)
        End If
    End Sub

    Private Sub btnAlignAll_Click(sender As Object, e As EventArgs) Handles btnAlignAll.Click
        'Align all message windows.

        Dim ConnectionName As String
        Dim ProNetName As String

        For Each item In dgvConnections.Rows
            If item.Cells(2).Value = "ADVL_Network_1" Then
                'This is the Network connection
            Else
                If item.Cells(11).Value = "Waiting" Then
                    'Connection waiting - probably failed!
                Else
                    ProNetName = item.Cells(0).Value
                    ConnectionName = item.Cells(2).Value

                    'Align the Message window:
                    Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                    Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                    'Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
                    Dim xmessage As New XElement("XSys") 'This indicates the start of the message in the XMessage class

                    'Get the Network App Message Window position:
                    'All other Message windows will be aligned with this.
                    If IsNothing(Message.MessageForm) Then
                        'Show the Messages form.
                        Message.ApplicationName = ApplicationInfo.Name
                        Message.SettingsLocn = Project.SettingsLocn
                        Message.Show()
                        Message.MessageForm.BringToFront()
                    End If
                    Dim WindowLeft As Integer = Message.MessageForm.Left
                    Dim WindowTop As Integer = Message.MessageForm.Top
                    Dim WindowWidth As Integer = Message.MessageForm.Width
                    Dim WindowHeight As Integer = Message.MessageForm.Height

                    Dim msgWindow As New XElement("MessageWindow")
                    Dim msgWindowLeft As New XElement("Left", WindowLeft)
                    msgWindow.Add(msgWindowLeft)
                    Dim msgWindowTop As New XElement("Top", WindowTop)
                    msgWindow.Add(msgWindowTop)
                    Dim msgWindowWidth As New XElement("Width", WindowWidth)
                    msgWindow.Add(msgWindowWidth)
                    Dim msgWindowHeight As New XElement("Height", WindowHeight)
                    msgWindow.Add(msgWindowHeight)
                    Dim msgWindowSaveSettings As New XElement("Command", "SaveSettings")
                    msgWindow.Add(msgWindowSaveSettings)
                    Dim msgWindowCommand As New XElement("Command", "BringToFront")
                    msgWindow.Add(msgWindowCommand)
                    xmessage.Add(msgWindow)
                    doc.Add(xmessage)

                    'Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnectionName & ":" & vbCrLf, "XmlSentNotice")
                    'Message.XAddXml(doc.ToString)
                    'Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    If ShowSysMessages Then
                        Message.XAddText("System Message sent to " & "[" & ProNetName & "]." & ConnectionName & ":" & vbCrLf, "XmlSentNotice")
                        Message.XAddXml(doc.ToString)
                        Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    End If

                    client.SendMessageAsync(ProNetName, ConnectionName, doc.ToString)

                    'Try
                    '    'Try sending an XMsg:
                    '    Dim AppComCheckMsg = <?xml version="1.0" encoding="utf-8"?>
                    '                         <XMsg>
                    '                             <ClientProNetName></ClientProNetName>
                    '                             <ClientName>ADVL_Network_1</ClientName>
                    '                             <ClientConnectionName>ADVL_Network_1</ClientConnectionName>
                    '                             <ClientLocn>AppComCheck</ClientLocn>
                    '                             <Command>AppComCheck</Command>
                    '                         </XMsg>

                    '    SendMessageParams.ProjectNetworkName = item.Cells(0).Value
                    '    SendMessageParams.ConnectionName = item.Cells(2).Value
                    '    SendMessageParams.Message = AppComCheckMsg.ToString
                    '    If bgwSendMessage.IsBusy Then
                    '        Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                    '    Else
                    '        item.DefaultCellStyle.ForeColor = Color.Red 'ADDED 9Apr20
                    '        item.Cells(11).Value = "Waiting"
                    '        bgwSendMessage.RunWorkerAsync(SendMessageParams)
                    '        System.Threading.Thread.Sleep(1000) 'Pause for 1000ms
                    '        Application.DoEvents()
                    '    End If
                    'Catch ex As Exception
                    '    Message.AddWarning("Application Connection Check error: " & ex.Message & vbCrLf)
                    'End Try


                End If
            End If
        Next




    End Sub

    Private Sub ShowMessageWindow(ByVal ProNetName As String, ByVal ConnName As String)
        'Show the Message window corresponding to the specified Application Netaork Name and Connection Name.

        If ConnName = "ADVL_Network_1" Then
            'Show the Messages form.
            Message.ApplicationName = ApplicationInfo.Name
            Message.SettingsLocn = Project.SettingsLocn
            Message.Show()
            Message.MessageForm.BringToFront()
        Else
            Dim decl As New XDeclaration("1.0", "utf-8", "yes")
            Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
            Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class

            'Show the Message window.
            Dim msgWindow As New XElement("MessageWindow")
            Dim msgWindowCommand As New XElement("Command", "BringToFront")
            msgWindow.Add(msgWindowCommand)
            xmessage.Add(msgWindow)
            doc.Add(xmessage)
            'Show the message sent:
            Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
            Message.XAddXml(doc.ToString)
            Message.XAddText(vbCrLf, "Normal") 'Add extra line
            client.SendMessageAsync(ProNetName, ConnName, doc.ToString)
        End If
    End Sub

    Private Sub ToolStripMenuItem1_EditWorkflowTabPage_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1_EditWorkflowTabPage.Click
        'Edit the Workflows Web Page:

        If WorkflowFileName = "" Then
            Message.AddWarning("No page to edit." & vbCrLf)
        Else
            Dim FormNo As Integer = OpenNewHtmlDisplayPage()
            HtmlDisplayFormList(FormNo).FileName = WorkflowFileName
            HtmlDisplayFormList(FormNo).OpenDocument
        End If
    End Sub

    Private Sub ToolStripMenuItem1_ShowStartPageInWorkflowTab_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1_ShowStartPageInWorkflowTab.Click
        'Show the Start Page in the Workflows Tab:
        OpenStartPage()
    End Sub

    Private Sub TabPage1_Enter(sender As Object, e As EventArgs) Handles TabPage1.Enter
        'Update the duration of each connection

        Dim I As Integer
        Dim StartTime As DateTime
        Dim CurrentDuration As TimeSpan

        For I = 0 To dgvConnections.RowCount - 1
            StartTime = Date.ParseExact(dgvConnections.Rows(I).Cells(9).Value, "d-MMM-yyyy H:mm:ss", System.Globalization.DateTimeFormatInfo.InvariantInfo)
            CurrentDuration = Now.Subtract(StartTime)
            dgvConnections.Rows(I).Cells(10).Value = CurrentDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                                                     CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                                                     CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                                                     CurrentDuration.Seconds.ToString.PadLeft(2, "0"c)
        Next

        Timer4.Interval = 5000 '5 seconds
        Timer4.Enabled = True
        Timer4.Start()

    End Sub

    Private Sub TabPage1_Leave(sender As Object, e As EventArgs) Handles TabPage1.Leave
        Timer4.Enabled = False
    End Sub

    Private Sub Timer4_Tick(sender As Object, e As EventArgs) Handles Timer4.Tick
        'Update the duration of each connection

        Dim I As Integer
        Dim StartTime As DateTime
        Dim CurrentDuration As TimeSpan

        For I = 0 To dgvConnections.RowCount - 1
            StartTime = Date.ParseExact(dgvConnections.Rows(I).Cells(9).Value, "d-MMM-yyyy H:mm:ss", System.Globalization.DateTimeFormatInfo.InvariantInfo)
            CurrentDuration = Now.Subtract(StartTime)
            dgvConnections.Rows(I).Cells(10).Value = CurrentDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                                                     CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                                                     CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                                                     CurrentDuration.Seconds.ToString.PadLeft(2, "0"c)
        Next

        'The timer initially ticks after 5 seconds.
        'The tick interval is progressively increased to 30 seconds.
        If Timer4.Interval < 30000 Then
            Timer4.Interval += 5000
        End If

    End Sub

    Private Sub XmlHtmDisplay1_TextChanged(sender As Object, e As EventArgs) Handles XmlHtmDisplay1.TextChanged

    End Sub

    Private Sub XmlHtmDisplay1_ErrorMessage(Msg As String) Handles XmlHtmDisplay1.ErrorMessage
        Message.AddWarning(Msg & vbCrLf)
    End Sub

    Private Sub dgvConnections_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvConnections.CellContentClick

    End Sub

    Private Sub dgvConnections_Click(sender As Object, e As EventArgs) Handles dgvConnections.Click
        'The Connections datagridview has been clicked.

        If chkShowMessages.Checked Or chkShowApp.Checked Then
            If dgvConnections.SelectedRows.Count > 0 Then
                Dim FirstSelRow As Integer = dgvConnections.SelectedRows(0).Index

                If dgvConnections.Rows(FirstSelRow).Cells(11).Value = "Waiting" Then
                    Message.AddWarning("Connection waiting - probably failed!" & vbCrLf)
                    Exit Sub
                End If

                Dim ProNetName As String = dgvConnections.Rows(FirstSelRow).Cells(0).Value
                Dim ConnName As String = dgvConnections.Rows(FirstSelRow).Cells(2).Value

                If ConnName = "ADVL_Network_1" Then
                    'Show the Messages form.
                    If IsNothing(Message.MessageForm) Then
                        Message.ApplicationName = ApplicationInfo.Name
                        Message.SettingsLocn = Project.SettingsLocn
                        Message.Show()
                    End If
                    Message.MessageForm.Activate()
                    Message.MessageForm.TopMost = True
                    Message.MessageForm.TopMost = False
                Else
                    Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                    Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                    'Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
                    Dim xmessage As New XElement("XSys") 'This indicates the start of the system message in the XMessage class

                    'Lines added to include a ComCheck:
                    Dim clientProNetName As New XElement("ClientProNetName", "")
                    xmessage.Add(clientProNetName)
                    Dim clientName As New XElement("ClientName", "ADVL_Network_1")
                    xmessage.Add(clientName)
                    Dim clientConnectionName As New XElement("ClientConnectionName", "ADVL_Network_1")
                    xmessage.Add(clientConnectionName)

                    If chkShowMessages.Checked Then
                        'Show the Message window.
                        Dim msgWindow As New XElement("MessageWindow")
                        Dim msgWindowCommand As New XElement("Command", "BringToFront")
                        msgWindow.Add(msgWindowCommand)
                        xmessage.Add(msgWindow)
                    End If

                    If chkShowApp.Checked Then
                        'Show the Application window.
                        Dim appWindow As New XElement("ApplicationWindow")
                        Dim appWindowCommand As New XElement("Command", "BringToFront")
                        appWindow.Add(appWindowCommand)
                        xmessage.Add(appWindow)
                    End If

                    'Lines added to include a ComCheck:
                    Dim clientLocn As New XElement("ClientLocn", "AppComCheck")
                    xmessage.Add(clientLocn)
                    Dim appComCheckCmd As New XElement("Command", "AppComCheck")
                    xmessage.Add(appComCheckCmd)

                    doc.Add(xmessage)
                    'Show the message sent:
                    'Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                    'Message.XAddXml(doc.ToString)
                    'Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    If ShowSysMessages Then
                        Message.XAddText("System Message sent to " & "[" & ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                        Message.XAddXml(doc.ToString)
                        Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    End If


                    If bgwSendMessage.IsBusy Then
                        Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                    Else
                        dgvConnections.Rows(FirstSelRow).DefaultCellStyle.ForeColor = Color.Red 'ADDED 9Apr20
                        dgvConnections.Rows(FirstSelRow).Cells(11).Value = "Waiting" 'Set the connection Status to Waiting - this be be changed to OK if the connection is working.
                        Dim SendMessageParams As New clsSendMessageParams
                        SendMessageParams.ProjectNetworkName = ProNetName
                        SendMessageParams.ConnectionName = ConnName
                        SendMessageParams.Message = doc.ToString
                        bgwSendMessage.WorkerReportsProgress = True
                        bgwSendMessage.WorkerSupportsCancellation = True
                        bgwSendMessage.RunWorkerAsync(SendMessageParams)
                    End If
                End If

            End If
        Else
            'Nothing to show.
        End If



    End Sub

    Private Sub btnShowAppListInfo_Click(sender As Object, e As EventArgs) Handles btnShowAppListInfo.Click
        'Show the information stored in AppInfo()

        Message.Add("List of applications in AppInfo(): -------------------------------" & vbCrLf & vbCrLf)

        For Each key In AppInfo.Keys
            Message.Add("Application name: " & key & vbCrLf)
            Message.Add("  Description: " & AppInfo(key).Description & vbCrLf)
            Message.Add("  Directory: " & AppInfo(key).Directory & vbCrLf)
            Message.Add("  ExecutablePath: " & AppInfo(key).ExecutablePath & vbCrLf)
            Message.Add("  IconNumber: " & AppInfo(key).IconNumber & vbCrLf)
            Message.Add("  OpenIconNumber: " & AppInfo(key).OpenIconNumber & vbCrLf & vbCrLf)
        Next

        Message.Add("-----------------------------------------------------------------" & vbCrLf & vbCrLf)
    End Sub

    Private Sub btnShowProjListInfo_Click(sender As Object, e As EventArgs) Handles btnShowProjListInfo.Click
        'Show the information stored in ProjInfo()

        Message.Add("List of projects in ProjInfo(): -------------------------------" & vbCrLf & vbCrLf)

        For Each key In ProjInfo.Keys
            Message.Add("Project key: " & key & vbCrLf)
            Message.Add("  Name: " & ProjInfo(key).Name & vbCrLf)
            Message.Add("  Description: " & ProjInfo(key).Description & vbCrLf)
            Message.Add("  Project Network Name: " & ProjInfo(key).ProNetName & vbCrLf)
            Message.Add("  ApplicationName: " & ProjInfo(key).ApplicationName & vbCrLf)
            Message.Add("  Project Type: " & ProjInfo(key).Type & vbCrLf)
            Message.Add("  Creation date: " & ProjInfo(key).CreationDate & vbCrLf)
            Message.Add("  ID: " & ProjInfo(key).ID & vbCrLf)
            Message.Add("  RelativePath: " & ProjInfo(key).RelativePath & vbCrLf)
            Message.Add("  Path: " & ProjInfo(key).Path & vbCrLf)
            Message.Add("  IconNumber: " & ProjInfo(key).IconNumber & vbCrLf)
            Message.Add("  OpenIconNumber: " & ProjInfo(key).OpenIconNumber & vbCrLf)

            Message.Add("  ParentProjectName: " & ProjInfo(key).ParentProjectName & vbCrLf)
            Message.Add("  ParentProjectID: " & ProjInfo(key).ParentProjectID & vbCrLf)
            Message.Add("  ParentProjectPath: " & ProjInfo(key).ParentProjectPath & vbCrLf & vbCrLf)

        Next

        Message.Add("-----------------------------------------------------------------" & vbCrLf & vbCrLf)
    End Sub

    Private Sub bgwComCheck_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgwComCheck.DoWork
        'The communications check thread.


        'While ConnectedToComNet
        While 1 = 1
            Try
                If client.IsAlive() Then
                    'Message.Add(Format(Now, "HH:mm:ss") & " Connection OK." & vbCrLf) 'This produces the error: Cross thread operation not valid.
                    bgwComCheck.ReportProgress(1, Format(Now, "HH:mm:ss") & " Connection OK." & vbCrLf)
                Else
                    'Message.Add(Format(Now, "HH:mm:ss") & " Connection Fault." & vbCrLf) 'This produces the error: Cross thread operation not valid.
                    bgwComCheck.ReportProgress(1, Format(Now, "HH:mm:ss") & " Connection Fault.")
                End If
            Catch ex As Exception
                bgwComCheck.ReportProgress(1, "Error in bgeComCheck_DoWork!" & vbCrLf)
                bgwComCheck.ReportProgress(1, ex.Message & vbCrLf)
            End Try

            'System.Threading.Thread.Sleep(60000) 'Sleep time in milliseconds (60 seconds) - For testing only.
            'System.Threading.Thread.Sleep(3600000) 'Sleep time in milliseconds (60 minutes)
            System.Threading.Thread.Sleep(1800000) 'Sleep time in milliseconds (30 minutes)
        End While

    End Sub

    Private Sub bgwComCheck_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgwComCheck.ProgressChanged
        Message.Add(e.UserState.ToString) 'Show the ComCheck message 
    End Sub

    Private Sub Project_Closing() Handles Project.Closing
        'The current project is closing.
        SaveFormSettings() 'Save the form settings - they are saved in the Project before is closes.
        SaveProjectSettings() 'Update this subroutine if project settings need to be saved.
        Project.Usage.SaveUsageInfo()   'Save the current project usage information.
        Project.UnlockProject() 'Unlock the current project before it Is closed.
        If ConnectedToComNet Then DisconnectFromComNet()
    End Sub

    Private Sub Project_Selected() Handles Project.Selected
        'A new project has been selected.
        RestoreFormSettings()
        Project.ReadProjectInfoFile()

        Project.ReadParameters()
        'Project.ReadParentParameters()
        'If Project.ParentParameterExists("ProNetName") Then
        '    Project.AddParameter("ProNetName", Project.ParentParameter("ProNetName").Value, Project.ParentParameter("ProNetName").Description) 'AddParameter will update the parameter if it already exists.
        '    ProNetName = Project.Parameter("ProNetName").Value
        'Else
        '    ProNetName = Project.GetParameter("ProNetName")
        'End If
        'If Project.ParentParameterExists("ProNetPath") Then 'Get the parent parameter value - it may have been updated.
        '    Project.AddParameter("ProNetPath", Project.ParentParameter("ProNetPath").Value, Project.ParentParameter("ProNetPath").Description) 'AddParameter will update the parameter if it already exists.
        '    ProNetPath = Project.Parameter("ProNetPath").Value
        'Else
        '    ProNetPath = Project.GetParameter("ProNetPath") 'If the parameter does not exist, the value is set to ""
        'End If
        'Project.SaveParameters() 'These should be saved now - child projects look for parent parameters in the parameter file.

        Project.LockProject() 'Lock the project while it is open in this application.

        Project.Usage.StartTime = Now

        ApplicationInfo.SettingsLocn = Project.SettingsLocn
        Message.SettingsLocn = Project.SettingsLocn
        Message.Show() 'Added 18May19

        'Restore the new project settings:
        RestoreProjectSettings() 'Update this subroutine if project settings need to be restored.

        ShowProjectInfo()

        ConnectToComNet()

    End Sub

    Private Async Sub btnCheckConnection_Click(sender As Object, e As EventArgs) Handles btnCheckConnection.Click
        'Check the selected connection.

        If dgvConnections.SelectedRows.Count > 0 Then
            If dgvConnections.SelectedRows.Count = 1 Then

                Dim SelRow As Integer = dgvConnections.SelectedRows(0).Index
                Dim ProjectNetworkName As String = dgvConnections.Rows(SelRow).Cells(0).Value
                Dim ConnectionName As String = dgvConnections.Rows(SelRow).Cells(2).Value

                If ConnectionName = "ADVL_Network_1" Then

                    Exit Sub
                End If

                If dgvConnections.Rows(SelRow).Cells(11).Value = "Waiting" Then
                    Message.AddWarning("Connection waiting - probably failed!" & vbCrLf)
                    Exit Sub
                End If

                Try
                    'Try sending an XMsg:
                    Dim AppComCheckMsg = <?xml version="1.0" encoding="utf-8"?>
                                         <XMsg>
                                             <ClientProNetName></ClientProNetName>
                                             <ClientName>ADVL_Network_1</ClientName>
                                             <ClientConnectionName>ADVL_Network_1</ClientConnectionName>
                                             <ClientLocn>AppComCheck</ClientLocn>
                                             <Command>AppComCheck</Command>
                                         </XMsg>

                    'THE FOLLOWING LINE HAS BEEN MOVED TO THE VAriable Declarations SECTION:
                    'Dim SendMessageParams As New clsSendMessageParams
                    SendMessageParams.ProjectNetworkName = ProjectNetworkName
                    SendMessageParams.ConnectionName = ConnectionName
                    SendMessageParams.Message = AppComCheckMsg.ToString
                    If bgwSendMessage.IsBusy Then
                        Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                    Else
                        dgvConnections.Rows(SelRow).DefaultCellStyle.ForeColor = Color.Red 'ADDED 9Apr20
                        'dgvConnections.Rows(SelRow).Cells(11).Value = "Failed" 'Assume the connection has failed. If the communication check succeeds, this will be changed to OK.
                        dgvConnections.Rows(SelRow).Cells(11).Value = "Waiting" 'Set the connection status to Waiting. If the communication check succeeds, this will be changed to OK.
                        'THE FOLLOWING 2 LINES HAVE BEEN MOVED TO Main.Load:
                        'bgwSendMessage.WorkerReportsProgress = True
                        'bgwSendMessage.WorkerSupportsCancellation = True
                        bgwSendMessage.RunWorkerAsync(SendMessageParams)
                    End If
                Catch ex As Exception
                    'bgwAppComCheck.ReportProgress(1, ex.Message)
                    Message.AddWarning("Application Connection Check error: " & ex.Message & vbCrLf)
                End Try

            Else
                Message.AddWarning("Select one connection only." & vbCrLf)
            End If
        Else
            Message.AddWarning("No connections have been selected." & vbCrLf)
        End If
    End Sub

    Private Sub bgwSendMessage_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgwSendMessage.DoWork
        'Send a message on a separate thread:

        Try
            If IsNothing(client) Then
                bgwSendMessage.ReportProgress(1, "No Connection available. Message not sent!")
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    bgwSendMessage.ReportProgress(1, "Connection state is faulted. Message not sent!")
                Else
                    Dim SendMessageParams As clsSendMessageParams = e.Argument
                    client.SendMessage(SendMessageParams.ProjectNetworkName, SendMessageParams.ConnectionName, SendMessageParams.Message)
                End If
            End If
        Catch ex As Exception
            bgwSendMessage.ReportProgress(1, ex.Message)
        End Try
    End Sub

    Private Sub bgwSendMessage_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgwSendMessage.ProgressChanged
        'Display an error message:
        Message.AddWarning("Send Message error: " & e.UserState.ToString & vbCrLf) 'Show the bgwSendMessage message 
    End Sub

    Private Sub bgwSendMessageAlt_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgwSendMessageAlt.DoWork
        'Alternative SendMessage background worker - used to send a message while instructions are being processed. 
        'Send a message on a separate thread
        Try
            If IsNothing(client) Then
                bgwSendMessageAlt.ReportProgress(1, "No Connection available. Message not sent!")
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    bgwSendMessageAlt.ReportProgress(1, "Connection state is faulted. Message not sent!")
                Else
                    Dim SendMessageParamsAlt As clsSendMessageParams = e.Argument
                    client.SendMessage(SendMessageParamsAlt.ProjectNetworkName, SendMessageParamsAlt.ConnectionName, SendMessageParamsAlt.Message)
                End If
            End If
        Catch ex As Exception
            bgwSendMessageAlt.ReportProgress(1, ex.Message)
        End Try
    End Sub

    Private Sub bgwSendMessageAlt_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgwSendMessageAlt.ProgressChanged
        'Display an error message:
        Message.AddWarning("Send Message error: " & e.UserState.ToString & vbCrLf) 'Show the bgwSendMessageAlt message 
    End Sub

    'Private Async Sub bgwAppComCheck_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgwAppComCheck.DoWork
    Private Sub bgwAppComCheck_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgwAppComCheck.DoWork
        'The Application Communications check thread.

        Dim GetStateTask As Task(Of String)
        Dim ConnectionState As String = ""
        Dim ProjectNetworkName As String = ""
        Dim ConnectionName As String = ""

        Dim SelRow As Integer

        SelRow = dgvConnections.SelectedRows(0).Index
        ProjectNetworkName = dgvConnections.Rows(SelRow).Cells(0).Value
        ConnectionName = dgvConnections.Rows(SelRow).Cells(2).Value
        Try
            'Try sending an XMsg:
            Dim AppComCheckMsg = <?xml version="1.0" encoding="utf-8"?>
                                 <XMsg>
                                     <ClientProNetName></ClientProNetName>
                                     <ClientName>ADVL_Network_1</ClientName>
                                     <ClientConnectionName>ADVL_Network_1</ClientConnectionName>
                                     <ClientLocn>AppComCheck</ClientLocn>
                                     <Command>AppComCheck</Command>
                                 </XMsg>
            client.SendMessage(ProjectNetworkName, ConnectionName, AppComCheckMsg.ToString)
        Catch ex As Exception
            'bgwAppComCheck.ReportProgress(1, ex.Message)
        End Try

    End Sub

    ''Private Async Sub bgwAppComCheck_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgwAppComCheck.DoWork
    'Private Sub bgwAppComCheck_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgwAppComCheck.DoWork
    '    'The Application Communications check thread.

    '    Dim GetStateTask As Task(Of String)
    '    Dim ConnectionState As String = ""
    '    Dim ProjectNetworkName As String = ""
    '    Dim ConnectionName As String = ""

    '    Dim SelRow As Integer

    '    SelRow = dgvConnections.SelectedRows(0).Index
    '    ProjectNetworkName = dgvConnections.Rows(SelRow).Cells(0).Value
    '    ConnectionName = dgvConnections.Rows(SelRow).Cells(2).Value
    '    'GetStateTask = client.CheckConnectionAsync(ProjectNetworkName, ConnectionName)
    '    'ConnectionState = client.CheckConnection(ProjectNetworkName, ConnectionName)
    '    'ConnectionState = Await GetStateTask
    '    'client.SendMessage(ProjectNetworkName, ConnectionName, "OK")
    '    Try
    '        'bgwAppComCheck.ReportProgress(1, "Project Network Name: " & ProjectNetworkName & "  Connection Name: " & ConnectionName & "  State: " & ConnectionState & vbCrLf)
    '        'client.SendMessage(ProjectNetworkName, ConnectionName, "OK")

    '        'Try sending an XMsg:
    '        Dim AppComCheckMsg = <?xml version="1.0" encoding="utf-8"?>
    '                             <XMsg>
    '                                 <ClientProNetName></ClientProNetName>
    '                                 <ClientName>ADVL_Network_1</ClientName>
    '                                 <ClientConnectionName>ADVL_Network_1</ClientConnectionName>
    '                                 <ClientLocn>AppComCheck</ClientLocn>
    '                                 <Command>AppComCheck</Command>
    '                             </XMsg>
    '        client.SendMessage(ProjectNetworkName, ConnectionName, AppComCheckMsg.ToString)
    '    Catch ex As Exception
    '        'bgwAppComCheck.ReportProgress(1, ex.Message)
    '    End Try

    'End Sub

    'Private Async Sub bgwAppComCheck_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgwAppComCheck.DoWork
    '    'The Application Communications check thread.

    '    Dim GetStateTask As Task(Of String)
    '    Dim ConnectionState As String = ""
    '    Dim ProjectNetworkName As String = ""
    '    Dim ConnectionName As String = ""

    '    Dim SelRow As Integer

    '    SelRow = dgvConnections.SelectedRows(0).Index
    '    ProjectNetworkName = dgvConnections.Rows(SelRow).Cells(0).Value
    '    ConnectionName = dgvConnections.Rows(SelRow).Cells(2).Value
    '    GetStateTask = client.CheckConnectionAsync(ProjectNetworkName, ConnectionName)
    '    ConnectionState = Await GetStateTask
    '    bgwAppComCheck.ReportProgress(1, "Project Network Name: " & ProjectNetworkName & "  Connection Name: " & ConnectionName & "  State: " & ConnectionState & vbCrLf)

    '    'Try
    '    '    ConnectionState = Await GetStateTask
    '    '    'NOTE: Cannot use Message - it was created on another thread!!!
    '    '    'Message.Add("Project Network Name: " & ProjectNetworkName & "  Connection Name: " & ConnectionName & "  State: " & ConnectionState & vbCrLf)
    '    '    bgwAppComCheck.ReportProgress(1, "Project Network Name: " & ProjectNetworkName & "  Connection Name: " & ConnectionName & "  State: " & ConnectionState & vbCrLf)
    '    'Catch ex As Exception
    '    '    'NOTE: Cannot use Message - it was created on another thread!!!
    '    '    'Message.AddWarning("AppComCheck error: " & ex.Message & vbCrLf)

    '    '    'NOTE: ReportProgress here produes the error message:
    '    '    'An exception of type 'System.InvalidOperationException' occurred in System.dll but was not handled in user code
    '    '    'Additional information: This operation has already had OperationCompleted called on it and further calls are illegal.
    '    '    'bgwAppComCheck.ReportProgress(1, "AppComCheck error: " & ex.Message & vbCrLf)
    '    'End Try

    'End Sub

    Private Sub bgwAppComCheck_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgwAppComCheck.ProgressChanged
        Message.Add(e.UserState.ToString) 'Show the AppComCheck message 
    End Sub

    Private Sub btnRemoveWaiting_Click(sender As Object, e As EventArgs) Handles btnRemoveWaiting.Click
        'Remove connections that have Waiting status - they have probably failed.

        'Find all the rows in dgvConnections with Status = Waiting:
        Dim I As Integer 'Loop index
        For I = 0 To dgvConnections.Rows.Count - 1
            If dgvConnections.Rows(I).Cells(11).Value = "Waiting" Then
                client.DisconnectAsync(dgvConnections.Rows(I).Cells(0).Value, dgvConnections.Rows(I).Cells(2).Value) 'Disconnect the Connection that is still Waiting
            End If
        Next

    End Sub

    Private Sub btnCheckAllConnections_Click(sender As Object, e As EventArgs) Handles btnCheckAllConnections.Click
        'Check a  connection in dgvConnections.

        For Each item In dgvConnections.Rows
            If item.Cells(2).Value = "ADVL_Network_1" Then
                'This is the Network connection - no check needed.
            Else
                If item.Cells(11).Value = "Waiting" Then
                    'Connection waiting - probably failed!
                Else
                    Try
                        'Try sending an XMsg:
                        Dim AppComCheckMsg = <?xml version="1.0" encoding="utf-8"?>
                                             <XMsg>
                                                 <ClientProNetName></ClientProNetName>
                                                 <ClientName>ADVL_Network_1</ClientName>
                                                 <ClientConnectionName>ADVL_Network_1</ClientConnectionName>
                                                 <ClientLocn>AppComCheck</ClientLocn>
                                                 <Command>AppComCheck</Command>
                                             </XMsg>

                        SendMessageParams.ProjectNetworkName = item.Cells(0).Value
                        SendMessageParams.ConnectionName = item.Cells(2).Value
                        SendMessageParams.Message = AppComCheckMsg.ToString
                        If bgwSendMessage.IsBusy Then
                            Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                        Else
                            item.DefaultCellStyle.ForeColor = Color.Red 'ADDED 9Apr20
                            item.Cells(11).Value = "Waiting"
                            bgwSendMessage.RunWorkerAsync(SendMessageParams)
                            System.Threading.Thread.Sleep(1000) 'Pause for 1000ms
                            Application.DoEvents()
                        End If
                    Catch ex As Exception
                        Message.AddWarning("Application Connection Check error: " & ex.Message & vbCrLf)
                    End Try
                End If
            End If
        Next

    End Sub

    Private Sub dgvProjects_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvProjects.CellContentClick

    End Sub

    Private Sub XMsg_ErrorMsg(ErrMsg As String) Handles XMsg.ErrorMsg
        Message.AddWarning(ErrMsg & vbCrLf)
    End Sub

    Private Sub XMsgLocal_Instruction(Info As String, Locn As String) Handles XMsgLocal.Instruction

    End Sub

    Private Sub Message_ShowXMessagesChanged(Show As Boolean) Handles Message.ShowXMessagesChanged
        ShowXMessages = Show
    End Sub

    Private Sub Message_ShowSysMessagesChanged(Show As Boolean) Handles Message.ShowSysMessagesChanged
        ShowSysMessages = Show
    End Sub

    Private Sub btnRedrawTree_Click(sender As Object, e As EventArgs) Handles btnRedrawTree.Click
        trvAppTree.Refresh()
    End Sub

    Private Sub btnGetIpList_Click(sender As Object, e As EventArgs) Handles btnGetIpList.Click
        'Get a list of IP addresses connected to the local network:

        Dim strBuilder As New System.Text.StringBuilder
        Dim strHostName As New String("")
        strHostName = Net.Dns.GetHostName

        strBuilder.Append("Local machine host name: " & strHostName & vbCrLf)
        strBuilder.Append(vbCrLf)

        Dim IpEntry As Net.IPHostEntry = Net.Dns.GetHostEntry(strHostName)
        Dim addr As Net.IPAddress() = IpEntry.AddressList

        For Each item In addr
            strBuilder.Append("IP Address: " & item.ToString & vbCrLf)
        Next

        strBuilder.Append("---End --" & vbCrLf & vbCrLf)

        txtIpList.Text = strBuilder.ToString

    End Sub

    Private Sub btnGetIpList2_Click(sender As Object, e As EventArgs) Handles btnGetIpList2.Click
        'Get a list of IP addresses connected to the local network: (Version 2)
        'https://social.msdn.microsoft.com/Forums/en-US/c5d470d4-665a-4a71-ae5c-1dddae3bd21a/visual-basic-form-how-to-see-connected-devices-on-network?forum=vbgeneral

        Dim adapters As Net.NetworkInformation.NetworkInterface() = Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces
        Dim adapter As Net.NetworkInformation.NetworkInterface
        Dim strBuilder As New System.Text.StringBuilder

        For Each adapter In adapters
            Dim properties As Net.NetworkInformation.IPInterfaceProperties = adapter.GetIPProperties
            strBuilder.Append("Description: " & adapter.Description & vbCrLf)
            strBuilder.Append("DNS suffix: " & properties.DnsSuffix & vbCrLf)
            strBuilder.Append("DNS enabled: " & properties.IsDnsEnabled.ToString & vbCrLf)
            strBuilder.Append("Dynamically configured DNS: " & properties.IsDynamicDnsEnabled.ToString & vbCrLf)
            strBuilder.Append(vbCrLf)
        Next

        strBuilder.Append("---End --" & vbCrLf & vbCrLf)
        txtIpList.Text = strBuilder.ToString

    End Sub

    Private Sub btnGetIpList3_Click(sender As Object, e As EventArgs) Handles btnGetIpList3.Click

        Dim adapters As Net.NetworkInformation.NetworkInterface() = Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces
        Dim adapter As Net.NetworkInformation.NetworkInterface

        For Each adapter In adapters
            Dim properties As Net.NetworkInformation.IPInterfaceProperties = adapter.GetIPProperties
            rtbIpList.AppendText("Description: " & adapter.Description & vbCrLf)
            rtbIpList.AppendText("DNS suffix: " & properties.DnsSuffix & vbCrLf)
            rtbIpList.AppendText("DNS enabled: " & properties.IsDnsEnabled.ToString & vbCrLf)
            rtbIpList.AppendText("Dynamically configured DNS: " & properties.IsDynamicDnsEnabled.ToString & vbCrLf)
            rtbIpList.AppendText(vbCrLf)

        Next

    End Sub

    Private Sub btnGetIpList4_Click(sender As Object, e As EventArgs) Handles btnGetIpList4.Click

        'Dim mc As New Management.ManagementClass("Win32_NetworkAdapterConfiguration")
        'Dim mc As New Management.Instrumentation.

        'https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-networkadapterconfiguration?redirectedfrom=MSDN


        'Dim strComputer As String = "."

        'Dim objWMIService As Object = GetObject("winmgmts:\\" & "." & "\root\CIMV2")
        'objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration",, 48)


        'https://www.codeproject.com/Questions/482204/getingplusMACplusaddressplusbyplususingplusVB-netp

        Dim nics() As Net.NetworkInformation.NetworkInterface = Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces

        Dim nic As Net.NetworkInformation.NetworkInterface

        For Each nic In nics
            rtbIpList.AppendText("Name: " & nic.Name & vbCrLf)
            rtbIpList.AppendText("Description: " & nic.Description & vbCrLf)
            rtbIpList.AppendText("Address: " & nic.GetPhysicalAddress.ToString & vbCrLf)
            rtbIpList.AppendText("OperationalStatus: " & nic.OperationalStatus & vbCrLf)
            rtbIpList.AppendText("OperationalStatus.ToString: " & nic.OperationalStatus.ToString & vbCrLf)
            rtbIpList.AppendText("Type: " & nic.GetType.ToString & vbCrLf)
            rtbIpList.AppendText("Number of DnsAddresses: " & nic.GetIPProperties.DnsAddresses.Count & vbCrLf)

            For Each item In nic.GetIPProperties.DnsAddresses
                rtbIpList.AppendText("  DnsAddresses: " & item.ToString & vbCrLf)
            Next

            'rtbIpList.AppendText("DnsAddresses: " & nic.GetIPProperties.DnsAddresses(0).ToString & vbCrLf)


            rtbIpList.AppendText(vbCrLf)

        Next


        'https://www.codeproject.com/Tips/358946/Retrieving-IP-And-MAC-addresses-for-a-LAN



    End Sub

    Private Sub btnPing_Click(sender As Object, e As EventArgs) Handles btnPing.Click
        'Ping the IP Address in txtIP

        'https://www.c-sharpcorner.com/UploadFile/moosestafa/my-network-scanner-aka-hacking-101-in-VB-Net/

        If txtIP.Text = "" Then
            'No IP Address specified
        Else
            Dim png As New Net.NetworkInformation.Ping
            Dim pr As Net.NetworkInformation.PingReply
            'Dim IPAddressStr As String

            Try
                pr = png.Send(txtIP.Text)
                If pr.Status = Net.NetworkInformation.IPStatus.Success Then
                    'rtbIpList.AppendText("Computer name: " & Net.Dns.GetHostEntry(pr.Address.ToString).ToString & " IP Address: " & txtIP.Text & " roundtrip time: " & pr.RoundtripTime.ToString & "ms" & vbCrLf)
                    rtbIpList.AppendText(vbCrLf & " IP Address: " & txtIP.Text & " roundtrip time: " & pr.RoundtripTime.ToString & "ms" & vbCrLf)
                    rtbIpList.AppendText("  pr.Address.ToString: " & pr.Address.ToString & vbCrLf)

                    'rtbIpList.AppendText("  Net.Dns.GetHostEntry(pr.Address).HostName: " & Net.Dns.GetHostEntry(pr.Address.ToString).HostName & vbCrLf) 'No such host is known

                Else
                    Beep()
                    rtbIpList.AppendText(" . ")
                End If
            Catch ex As Exception
                Message.AddWarning(ex.Message)
            End Try

        End If


    End Sub

    Private Sub btnHostName_Click(sender As Object, e As EventArgs) Handles btnHostName.Click
        Try
            rtbIpList.AppendText("  HostName: " & Net.Dns.GetHostEntry(txtIP.Text).HostName & vbCrLf)
        Catch ex As Exception
            Beep()
            rtbIpList.AppendText(ex.Message)
        End Try

    End Sub

    Private Sub btnPingNext255_Click(sender As Object, e As EventArgs) Handles btnPingNext255.Click
        'Ping the next 255 IP addresses from the start address shown in the text box:

        If txtStartIP.Text = "" Then

        Else
            Dim png As New Net.NetworkInformation.Ping
            Dim pr As Net.NetworkInformation.PingReply
            Dim StartIPAddr As String = txtStartIP.Text.Trim
            Dim IPAddrStr As String
            Dim IPAddr As Net.IPAddress
            Dim Host As Net.IPHostEntry
            Dim I As Integer


            For I = 0 To 255
                IPAddrStr = StartIPAddr & "." & I
                Message.Add(I & " ")
                Try
                    pr = png.Send(IPAddrStr, 900)
                    If pr.Status = Net.NetworkInformation.IPStatus.Success Then
                        rtbIpList.AppendText(vbCrLf & " IP Address: " & IPAddrStr & " roundtrip time: " & pr.RoundtripTime.ToString & "ms" & vbCrLf)
                        'Application.DoEvents()
                        IPAddr = Net.IPAddress.Parse(IPAddrStr)
                        Host = Net.Dns.GetHostEntry(IPAddr)
                        rtbIpList.AppendText("Host = " & Host.HostName.ToString & vbCrLf)

                    End If

                Catch ex As Exception

                End Try
            Next

        End If
    End Sub


    Private Sub PingAsync255(ByVal SubIPAddr As String)
        'Test code to ping next 256 addresses using async method:

        'http://www.vbforums.com/showthread.php?573829-Async-Ping-How-To&highlight=ping

        Dim myPing As Net.NetworkInformation.Ping
        Dim Timeout As Integer = 4000 'Timeout for pink request in ms
        Dim I As Integer
        Dim PingAddr As String
        rtbIpList.AppendText("Pinging 254 addresses from the sub address: " & SubIPAddr & vbCrLf)
        For I = 0 To 254
            Message.Add("ThreadCount: " & System.Diagnostics.Process.GetCurrentProcess().Threads.Count & vbCrLf)
            PingAddr = SubIPAddr & "." & I.ToString
            myPing = New Net.NetworkInformation.Ping
            AddHandler myPing.PingCompleted, AddressOf PingResult
            'rtbIpList.AppendText("Pinging " & PingAddr & " ")
            'rtbIpList.AppendText(PingAddr & " ")
            myPing.SendAsync(PingAddr, Timeout, PingAddr)
        Next

    End Sub

    Private Sub PingResult(ByVal sender As Object, ByVal e As System.Net.NetworkInformation.PingCompletedEventArgs)

        Dim PingResultStr As String = ""

        'rtbIpList.AppendText("Result from: " & e.UserState.ToString & vbCrLf) 'e.UserState is the UserToken passed in myPing.SendAsync()
        'PingResultStr = "Result from: " & e.UserState.ToString & vbCrLf 'e.UserState is the UserToken passed in myPing.SendAsync() (PingAddr)

        If e.Error Is Nothing Then
            If e.Reply.Status = Net.NetworkInformation.IPStatus.Success Then
                PingResultStr = "Result from: " & e.UserState.ToString & vbCrLf 'e.UserState is the UserToken passed in myPing.SendAsync() (PingAddr)
                PingResultStr &= "Status: " & e.Reply.Status.ToString & vbCrLf
                PingResultStr &= "  Round trip time: " & e.Reply.RoundtripTime.ToString & vbCrLf
                rtbIpList.AppendText(PingResultStr & vbCrLf)
            End If
        Else
                'PingResultStr &= "Error: " & e.Error.Message & vbCrLf
                'If e.Error.InnerException IsNot Nothing Then
                '    PingResultStr &= "More information: " & e.Error.InnerException.Message & vbCrLf
                'End If
            End If

        'rtbIpList.AppendText(PingResultStr & vbCrLf)

        With DirectCast(sender, Net.NetworkInformation.Ping)
            RemoveHandler .PingCompleted, AddressOf PingResult
            .Dispose()
        End With
    End Sub

    Private Sub btnPingAsyncNext255_Click(sender As Object, e As EventArgs) Handles btnPingAsyncNext255.Click
        PingAsync255(txtStartIP.Text)
    End Sub

    Private Sub btnPingAsyncNext255_255_Click(sender As Object, e As EventArgs) Handles btnPingAsyncNext255_255.Click
        PingAsyncNext255_255(txtStartIP.Text)
    End Sub

    Private Sub PingAsyncNext255_255(ByVal SubIPAddr As String)


        Dim myPing As Net.NetworkInformation.Ping
        Dim Timeout As Integer = 4000 'Timeout for pink request in ms
        Dim I As Integer
        Dim J As Integer
        Dim PingAddr As String
        rtbIpList.AppendText("Pinging 254.254 addresses from the sub address: " & SubIPAddr & vbCrLf)
        For I = 0 To 20
            'Message.Add("ThreadCount: " & vbCrLf)
            For J = 0 To 254
                If J Mod 128 = 0 Then Message.Add(vbCrLf & "ThreadCount: " & System.Diagnostics.Process.GetCurrentProcess().Threads.Count & vbCrLf)
                'Message.Add(System.Diagnostics.Process.GetCurrentProcess().Threads.Count & " ")
                Message.Add(".")
                PingAddr = SubIPAddr & "." & I.ToString & "." & J.ToString
                myPing = New Net.NetworkInformation.Ping
                AddHandler myPing.PingCompleted, AddressOf PingResult
                myPing.SendAsync(PingAddr, Timeout, PingAddr)
            Next
        Next
    End Sub









    'Private Sub SearchIPAddresses(ByVal SubNetStr As String)
    '    'Search for IP Address on a sub-net
    '    If Searching Then
    '        Message.AddWarning("Already searching." & vbCrLf)
    '    Else
    '        'Create the search thread:
    '        SearchThread = New Threading.Thread(AddressOf Search)
    '    End If
    '    xxx
    'End Sub

    'Private Sub Search()

    'End Sub


    'Private Class pingObj
    '    Property addr As Net.IPAddress
    '    Property I As Integer
    'End Class


    'Private Sub btnShowResult_Click(sender As Object, e As EventArgs) Handles btnShowResult.Click
    '    Message.Add("Application Communications Check result: " & AppComCheckStatus & vbCrLf)
    'End Sub


    Public Function MyHostStatus() As String
        'Return the status of myHost
        If myHost Is Nothing Then
            Return "No Service Host"
        Else
            Return myHost.State.ToString
        End If
    End Function

    Public Sub StopMessageService()
        'Stop the Message Service
        Try
            myHost.Abort()
        Catch ex As Exception
            Message.AddWarning("Error stopping message service:" & vbCrLf & ex.Message & vbCrLf)
        End Try

    End Sub

    Public Sub RestartMessageService()
        'Restart the message service
        SetUpHost()
    End Sub

    Private Sub DataGridView2_DragEnter(sender As Object, e As DragEventArgs) Handles DataGridView2.DragEnter
        'DragEnter: An object has been dragged into DataGridView2 - View Zip Archive tab.
        'This code is required to get the link to the item(s) being dragged into XDataGridView2:
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Link
        End If

    End Sub

    Private Sub DataGridView2_DragDrop(sender As Object, e As DragEventArgs) Handles DataGridView2.DragDrop
        'An item has been DragDropped into the View Zip Archive tab.

        Dim Path As String()

        Path = e.Data.GetData(DataFormats.FileDrop)

        Dim I As Integer

        If Path.Count > 0 Then
            'Open the archive file
            DataGridView2.Rows.Clear()
            Dim Zip As System.IO.Compression.ZipArchive

            Try
                ZipFilePath = Path(0)
                txtZipFileDir.Text = System.IO.Path.GetDirectoryName(Path(0))
                txtZipFileName.Text = System.IO.Path.GetFileName(Path(0))
                'Zip = System.IO.Compression.ZipFile.OpenRead(Path(0))
                Zip = System.IO.Compression.ZipFile.OpenRead(ZipFilePath)
                For Each entry As System.IO.Compression.ZipArchiveEntry In Zip.Entries
                    'DataGridView2.Rows.Add(entry.Name, entry.LastWriteTime, entry.Length, entry.CompressedLength, entry.CompressedLength / entry.Length * 100)
                    'DataGridView2.Rows.Add(entry.Name, entry.FullName, entry.LastWriteTime, entry.Length, entry.CompressedLength, entry.CompressedLength / entry.Length * 100)
                    'DataGridView2.Rows.Add(entry.FullName, entry.Name, entry.LastWriteTime, entry.Length, entry.CompressedLength, entry.CompressedLength / entry.Length * 100)
                    DataGridView2.Rows.Add(System.IO.Path.GetDirectoryName(entry.FullName), entry.Name, entry.LastWriteTime, entry.Length, entry.CompressedLength, entry.CompressedLength / entry.Length * 100)
                Next
                Zip.Dispose()
                DataGridView2.AutoResizeColumns()
            Catch ex As Exception
                Message.AddWarning("Error displaying contents of Zip archive: " & ex.Message & vbCrLf)
            End Try
            If Path.Count > 1 Then
                Message.AddWarning(Path.Count & " Zip files dropped. Only the first one is shown." & vbCrLf)
            End If
        Else
            Message.AddWarning("No zip file path found." & vbCrLf)
        End If
    End Sub

    Private Sub btnCopy_Click(sender As Object, e As EventArgs) Handles btnCopy.Click
        'Copy the selected file(s).

        If DataGridView2.SelectedRows.Count > 0 Then

            'Check if the Temp directory exists:
            If System.IO.Directory.Exists(Project.Path & "\Temp") Then
                'The Temp directory exists.
            Else
                'Create the Temp directory:
                System.IO.Directory.CreateDirectory(Project.Path & "\Temp")
            End If

            If System.IO.Directory.Exists(Project.Path & "\Temp") Then
                Dim FileName As String
                Dim Directory As String
                Dim ExtractDir As String = Project.Path & "\Temp"
                Dim FileList As New System.Collections.Specialized.StringCollection

                'Copy each selected file to \Temp and add the path to the file list:
                If DataGridView2.SelectedRows.Count > 0 Then
                    Dim Zip As System.IO.Compression.ZipArchive
                    Zip = System.IO.Compression.ZipFile.OpenRead(ZipFilePath)
                    For Each item As DataGridViewRow In DataGridView2.SelectedRows
                        Directory = item.Cells(0).Value.Trim
                        'FileName = item.Cells(0).Value
                        FileName = item.Cells(1).Value
                        'Dim myEntry As System.IO.Compression.ZipArchiveEntry = Zip.GetEntry(FileName)
                        Dim myEntry As System.IO.Compression.ZipArchiveEntry
                        If Directory = "" Then
                            myEntry = Zip.GetEntry(FileName)
                        Else
                            myEntry = Zip.GetEntry(Directory & "\" & FileName)
                        End If

                        If IsNothing(myEntry) Then
                            Message.AddWarning("This file was not found in the archive: " & FileName & vbCrLf)
                        Else
                            'Copy the file:
                            myEntry.ExtractToFile(ExtractDir & "\" & FileName, True)
                            FileList.Add(ExtractDir & "\" & FileName)
                        End If
                    Next
                    Zip.Dispose()
                    Clipboard.SetFileDropList(FileList)
                End If
            Else
                Message.AddWarning("The temporary directory does not exist: " & Project.Path & "\Temp" & vbCrLf)
            End If
        End If
    End Sub

    Private Sub btnPaste_Click(sender As Object, e As EventArgs) Handles btnPaste.Click

        Dim FileList As System.Collections.Specialized.StringCollection

        If Clipboard.ContainsFileDropList Then
            FileList = Clipboard.GetFileDropList()
            Dim FileName As String
            Dim ZipPath As String
            Dim Zip As System.IO.Compression.ZipArchive
            Zip = ZipFile.Open(ZipFilePath, ZipArchiveMode.Update)
            For Each item In FileList
                If System.IO.File.Exists(item) Then
                    FileName = System.IO.Path.GetFileName(item)
                    ZipPath = Path.Combine(txtZipDirectory.Text.Trim, FileName)
                    'If IsNothing(Zip.GetEntry(FileName)) Then
                    If IsNothing(Zip.GetEntry(ZipPath)) Then
                        'Zip.CreateEntryFromFile(item, FileName)
                        Zip.CreateEntryFromFile(item, ZipPath)

                        ''Check if the item is a directory:
                        'If File.GetAttributes(item).HasFlag(FileAttributes.Directory) Then
                        '    CreateEntryFromAny(Zip, item, "")
                        'Else
                        '    Zip.CreateEntryFromFile(item, FileName)
                        'End If

                        'If System.IO.Directory.Exists(item) Then 'Create an entry from a directory:

                        '    CreateEntryFromAny(Zip, item, "")
                        'Else
                        '    Zip.CreateEntryFromFile(item, FileName)
                        'End If
                    Else
                        Message.AddWarning("A file with the same name is already in the archive: " & item & vbCrLf)
                    End If
                    'UPDATE:
                    'Dim entry As ZipArchiveEntry = Zip.GetEntry(FileName)
                    'If IsNothing(entry) Then

                    'Else
                    '    entry.Delete()
                    'End If
                    'Dim newEntry As ZipArchiveEntry = Zip.CreateEntry(FileName)

                ElseIf System.IO.Directory.Exists(item) Then 'Create an entry from a directory:
                    'FileName = System.IO.Path.GetFileName(item) 'The name of the directory
                    'If IsNothing(Zip.GetEntry(FileName)) Then
                    '    'CreateEntryFromAny(Zip, item, "")
                    '    CreateEntry(Zip, item, "")
                    'Else
                    '    Message.AddWarning("A directory with the same name is already in the archive: " & item & vbCrLf)
                    'End If
                    CreateEntry(Zip, item, txtZipDirectory.Text.Trim) 'Paste the directory in the directory shown in txtZipDirectory
                Else
                    Message.AddWarning("The file or directory to paste was not found: " & item & vbCrLf)
                End If
            Next
            Zip.Dispose()
            'GetArchiveFileList()

            'Update the Archive file list:
            DataGridView2.Rows.Clear()
            Try
                Zip = System.IO.Compression.ZipFile.OpenRead(ZipFilePath)
                For Each entry As System.IO.Compression.ZipArchiveEntry In Zip.Entries
                    'DataGridView2.Rows.Add(entry.Name, entry.LastWriteTime, entry.Length, entry.CompressedLength, entry.CompressedLength / entry.Length * 100)
                    'DataGridView2.Rows.Add(entry.Name, entry.FullName, entry.LastWriteTime, entry.Length, entry.CompressedLength, entry.CompressedLength / entry.Length * 100)
                    'DataGridView2.Rows.Add(entry.FullName, entry.Name, entry.LastWriteTime, entry.Length, entry.CompressedLength, entry.CompressedLength / entry.Length * 100)
                    DataGridView2.Rows.Add(System.IO.Path.GetDirectoryName(entry.FullName), entry.Name, entry.LastWriteTime, entry.Length, entry.CompressedLength, entry.CompressedLength / entry.Length * 100)
                Next
                Zip.Dispose()
                DataGridView2.AutoResizeColumns()
            Catch ex As Exception
                Message.AddWarning("Error displaying contents of Zip archive: " & ex.Message & vbCrLf)
            End Try
        Else
            Message.AddWarning("The Clipboard does not contain a file list." & vbCrLf)
        End If
    End Sub

    'https://stackoverflow.com/questions/15133626/creating-directories-in-a-ziparchive-c-sharp-net-4-5

    'Private Sub CreateEntryFromAny(ByRef myZip As ZipArchive, ByVal sourceName As String, ByVal entryName As String)
    Private Sub CreateEntry(ByRef myZip As ZipArchive, ByVal sourceName As String, ByVal entryName As String)
        'Create an entry in myZip from a file or a directory.

        Dim FileName As String = Path.GetFileName(sourceName)
        If File.GetAttributes(sourceName).HasFlag(FileAttributes.Directory) Then
            CreateEntryFromDirectory(myZip, sourceName, Path.Combine(entryName, FileName))
        Else
            myZip.CreateEntryFromFile(sourceName, Path.Combine(entryName, FileName)) 'Example: Path.Combine("Temp", "myFile.txt") = "Temp\myFile.Text"
        End If

    End Sub

    Private Sub CreateEntryFromDirectory(ByRef myZip As ZipArchive, ByVal sourceDirName As String, ByVal entryName As String)
        'Create an entry from a directory.

        Dim FileList As String() = Directory.GetFiles(sourceDirName).Concat(Directory.GetDirectories(sourceDirName)).ToArray()
        'myZip.CreateEntry(Path.Combine(entryName, Path.GetFileName(sourceDirName)))
        For Each fileName As String In FileList
            'CreateEntryFromAny(myZip, fileName, entryName)
            CreateEntry(myZip, fileName, entryName)
        Next
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        'Delete the selected entry

        If DataGridView2.SelectedRows.Count > 0 Then
            Dim Zip As System.IO.Compression.ZipArchive
            'Zip = System.IO.Compression.ZipFile.OpenRead(ZipFilePath)
            Zip = System.IO.Compression.ZipFile.Open(ZipFilePath, ZipArchiveMode.Update)
            Dim FileName As String
            Dim Directory As String
            Dim myEntry As System.IO.Compression.ZipArchiveEntry
            For Each item As DataGridViewRow In DataGridView2.SelectedRows
                Directory = item.Cells(0).Value.Trim
                FileName = item.Cells(1).Value
                If Directory = "" Then
                    myEntry = Zip.GetEntry(FileName)
                Else
                    myEntry = Zip.GetEntry(Directory & "\" & FileName)
                End If
                Try
                    myEntry.Delete()
                Catch ex As Exception
                    Message.AddWarning("Error deleting archive entry: " & ex.Message & vbCrLf)
                End Try

            Next
            Zip.Dispose()

            'Update the Archive file list:
            DataGridView2.Rows.Clear()
            Try
                Zip = System.IO.Compression.ZipFile.OpenRead(ZipFilePath)
                For Each entry As System.IO.Compression.ZipArchiveEntry In Zip.Entries
                    DataGridView2.Rows.Add(System.IO.Path.GetDirectoryName(entry.FullName), entry.Name, entry.LastWriteTime, entry.Length, entry.CompressedLength, entry.CompressedLength / entry.Length * 100)
                Next
                Zip.Dispose()
                DataGridView2.AutoResizeColumns()
            Catch ex As Exception
                Message.AddWarning("Error displaying contents of Zip archive: " & ex.Message & vbCrLf)
            End Try
        End If
    End Sub

    'Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    '    txtZipDirectory.Text = DataGridView2.Rows(e.RowIndex).Cells(0).Value
    'End Sub



    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        txtZipDirectory.Text = DataGridView2.Rows(e.RowIndex).Cells(0).Value
    End Sub

    Private Sub btnOpen_Click(sender As Object, e As EventArgs) Handles btnOpen.Click

        Try
            OpenFileDialog1.InitialDirectory = txtZipFileDir.Text
            OpenFileDialog1.FileName = ZipFilePath
            My.Computer.Keyboard.SendKeys("{HOME}") 'To show the full selected file path on OpenFileDialog1

            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                Dim FilePath As String = OpenFileDialog1.FileName
                txtZipFileDir.Text = System.IO.Path.GetDirectoryName(FilePath)
                txtZipFileName.Text = System.IO.Path.GetFileName(FilePath)
                DataGridView2.Rows.Clear()
                Dim Zip As System.IO.Compression.ZipArchive
                Try
                    ZipFilePath = FilePath
                    Zip = System.IO.Compression.ZipFile.OpenRead(ZipFilePath)
                    For Each entry As System.IO.Compression.ZipArchiveEntry In Zip.Entries
                        DataGridView2.Rows.Add(System.IO.Path.GetDirectoryName(entry.FullName), entry.Name, entry.LastWriteTime, entry.Length, entry.CompressedLength, entry.CompressedLength / entry.Length * 100)
                    Next
                    Zip.Dispose()
                    DataGridView2.AutoResizeColumns()
                Catch ex As Exception
                    Message.AddWarning("Error displaying contents of Zip archive: " & ex.Message & vbCrLf)
                End Try
            End If
        Catch ex As Exception
            Message.AddWarning("Error opening an archive file: " & ex.Message & vbCrLf)
        End Try
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        'Close the zip file.
        txtZipDirectory.Text = ""
        DataGridView2.Rows.Clear()
    End Sub

    Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click

        FolderBrowserDialog1.SelectedPath = txtZipFileDir.Text
        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            txtZipFileDir.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        'Save the Zip archive file in the specified directory with the specified file name.

        Dim NewFilePath As String = System.IO.Path.Combine(txtZipFileDir.Text.Trim, txtZipFileName.Text.Trim)
        'Message.Add("Zip archive file path: " & NewFilePath & vbCrLf)
        If System.IO.File.Exists(NewFilePath) Then
            Message.AddWarning("A file with this name already exists: " & NewFilePath & vbCrLf)
        Else
            'Copy the file at ZipFilePath to NewFilePath:
            My.Computer.FileSystem.CopyFile(ZipFilePath, NewFilePath)
            ZipFilePath = NewFilePath 'Set the ZipFilePath to NewFilePath
        End If

    End Sub

    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        'Create a new Zip archive file in the specified directory with the specified file name.

        Dim NewFilePath As String = System.IO.Path.Combine(txtZipFileDir.Text.Trim, txtZipFileName.Text.Trim)
        'Message.Add("Zip archive file path: " & NewFilePath & vbCrLf)

        If System.IO.File.Exists(NewFilePath) Then
            Message.AddWarning("A file with this name already exists: " & NewFilePath & vbCrLf)
        Else
            Try
                txtZipDirectory.Text = ""
                DataGridView2.Rows.Clear()
                Dim Zip As System.IO.Compression.ZipArchive
                Zip = System.IO.Compression.ZipFile.Open(NewFilePath, ZipArchiveMode.Create)
                ZipFilePath = NewFilePath 'Set the ZipFilePath to NewFilePath
                Zip.Dispose()
            Catch ex As Exception
                Message.AddWarning("Error creating Zip archive file: " & ex.Message & vbCrLf)
            End Try
        End If

    End Sub



#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Class clsSendMessageParams
        'Parameters used when sending a message using the Message Service.
        Public ProjectNetworkName As String
        Public ConnectionName As String
        Public Message As String
    End Class


End Class 'Main

Public Class App
    'Class holds a list of applications.
    'This is used by the App and ProjectApp objects that contain a list of all apps and project apps.

    Public List As New List(Of AppSummary) 'A list of applications

#Region "Application Methods" '--------------------------------------------------------------------------------------

    Public Function FindName(ByVal AppName As String) As AppSummary
        'Return the AppSummary corresponding to the Application with name AppName
        Dim FoundName As AppSummary

        FoundName = List.Find(Function(item As AppSummary)
                                  If IsNothing(item) Then
                                      '
                                  Else
                                      Return item.Name = AppName
                                  End If
                              End Function)
        If IsNothing(FoundName) Then
            'Return New ApplicationInfo 'Return blank record.
            Return New AppSummary 'Return blank record.
        Else
            Return FoundName
        End If
    End Function

#End Region 'Application Methods ------------------------------------------------------------------------------------

End Class 'App

Public Class AppSummary
    'Class holds summary information about an Application.
    'This is used by the App class and displayed in the Application List tab.

    Private _name As String = ""
    Property Name As String 'The name of the Application.
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Private _description As String = ""
    Property Description As String 'A description of the Application.
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

    Private _directory As String = ""
    Property Directory As String  'The directory containing the application.
        Get
            Return _directory
        End Get
        Set(value As String)
            _directory = value
        End Set
    End Property

    Private _executablePath As String = ""
    Property ExecutablePath As String 'The path of the Application Executable File.
        Get
            Return _executablePath
        End Get
        Set(value As String)
            _executablePath = value
        End Set
    End Property

End Class 'AppSummary

Public Class Proj
    'Class holds a list of projects.
    'This is used by the Proj object that contain a list of all projects.

    Public List As New List(Of ProjSummary) 'A list of projects

#Region "Application Methods" '--------------------------------------------------------------------------------------

    'Public Function FindName(ByVal AppName As String) As ApplicationInfo
    Public Function FindID(ByVal ProjID As String) As ProjSummary
        'Return the ProjSummary corresponding to the Project with ID ProjID

        Dim FoundID As ProjSummary

        FoundID = List.Find(Function(item As ProjSummary)
                                If IsNothing(item) Then
                                    '
                                Else
                                    Return item.ID = ProjID
                                End If
                            End Function)
        If IsNothing(FoundID) Then
            Return New ProjSummary 'Return blank record.
        Else
            Return FoundID
        End If
    End Function

    'Public Function FindNameAndAppNet(ByVal Name As String, ByVal AppNetName As String) As ProjSummary
    Public Function FindNameAndAppNet(ByVal Name As String, ByVal ProNetName As String) As ProjSummary
        'Return the ProjSummary corresponding to the Project with specified Name and ProNetName.

        Dim FoundProj As ProjSummary

        FoundProj = List.Find(Function(item As ProjSummary)
                                  'Return item.Name = Name And item.AppNetName = AppNetName
                                  Return item.Name = Name And item.ProNetName = ProNetName
                              End Function)
        If IsNothing(FoundProj) Then
            Return New ProjSummary
        Else
            Return FoundProj
        End If

    End Function

#End Region 'Application Methods ------------------------------------------------------------------------------------

End Class 'Proj

Public Class ProjSummary
    'Class holds summary information about a project.
    'This is used by the Proj class and displayed in the Project List tab.

    Private _name As String = "" 'The name of the project.
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    'Private _appNetName As String = "" 'The name of the Application Network containing the project. (Added 10Feb19.)
    'Property AppNetName As String
    '    Get
    '        Return _appNetName
    '    End Get
    '    Set(value As String)
    '        _appNetName = value
    '    End Set
    'End Property

    Private _proNetName As String = "" 'The name of the Project Network containing the project. 
    Property ProNetName As String
        Get
            Return _proNetName
        End Get
        Set(value As String)
            _proNetName = value
        End Set
    End Property

    Private _iD As String = "" 'The project ID.
    Property ID As String
        Get
            Return _iD
        End Get
        Set(value As String)
            _iD = value
        End Set
    End Property

    Private _type As ADVL_Utilities_Library_1.Project.Types = ADVL_Utilities_Library_1.Project.Types.Directory 'The type of location (None, Directory, Archive, Hybrid).
    Property Type As ADVL_Utilities_Library_1.Project.Types
        Get
            Return _type
        End Get
        Set(value As ADVL_Utilities_Library_1.Project.Types)
            _type = value
        End Set
    End Property

    Private _path As String = "" 'The path to the Project directory or archive.
    Property Path As String
        Get
            Return _path
        End Get
        Set(value As String)
            _path = value
        End Set
    End Property

    Private _description As String = "" 'A description of the project.
    Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

    Private _applicationName As String = "" 'The name of the application that created the project.
    Property ApplicationName As String
        Get
            Return _applicationName
        End Get
        Set(value As String)
            _applicationName = value
        End Set
    End Property

    Private _parentProjectName As String = "" 'The Name of the Parent Project.
    Property ParentProjectName As String
        Get
            Return _parentProjectName
        End Get
        Set(value As String)
            _parentProjectName = value
        End Set
    End Property

    Private _parentProjectID As String = "" 'The parent project ID.
    Property ParentProjectID As String
        Get
            Return _parentProjectID
        End Get
        Set(value As String)
            _parentProjectID = value
        End Set
    End Property

End Class

Public Class clsAppInfo
    'Information about each Application in the AppTreeView.
    'This is stored in the AppInfo dictionary.

    'Note: The Name is the key for the AppInfo dictionary. It does not need to be repeated in this class.
    'Note: The Text label is not stored in the ProjInfo dictionary. It is displayed in the AppTreeView.     'Text            The text label shown in the AppTreeView.

    'Description     A description of the application.
    'ExecutablePath  The path to the applications executable file.
    'Directory       The application directory.
    'IconNumber      The AppTreeImageList index number of the application's icon.
    'OpenIconNumber  The AppTreeImageList index number of the application's icon for an open application.

    Private _description As String = "" 'A description of the application.
    Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

    Private _executablePath As String = "" 'The path to the applications executable file.
    Property ExecutablePath As String
        Get
            Return _executablePath
        End Get
        Set(value As String)
            _executablePath = value
        End Set
    End Property

    Private _directory As String = "" 'The application directory.
    Property Directory As String
        Get
            Return _directory
        End Get
        Set(value As String)
            _directory = value
        End Set
    End Property

    Private _iconNumber As Integer = 0 'The AppTreeImageList index number of the application's icon.
    Property IconNumber As Integer
        Get
            Return _iconNumber
        End Get
        Set(value As Integer)
            _iconNumber = value
        End Set
    End Property

    Private _openIconNumber As Integer = 0 'The AppTreeImageList index number of the application's open icon.
    Property OpenIconNumber As Integer
        Get
            Return _openIconNumber
        End Get
        Set(value As Integer)
            _openIconNumber = value
        End Set
    End Property

End Class 'clsAppInfo

Public Class clsProjInfo
    'Information about each Project in the AppTreeView.
    'This is stored in the ProjectInfo dictionary.
    'The dictionary key is the ID and ".Proj"

    'Name               The name of the project. (The name may be duplicated in other projects.)
    'REMOVED: AppNetName         The name of the Application Network containig the project. (Added 10Feb19.)
    'ProNetName         The name of the Project Network containig the project. (Added 10Feb19.)
    'CreationDate
    'Description        A description of the project.
    'Type               The type of project (Directory, Archive, Hybrid or None.) (If the type is None, the Default project will be used.)
    'Path               The path of the project directory or archive.
    'RelativePath       The path of the project directory or archive relative to the Parent Project.
    'ID                 The project ID. this is the hashcode generated from the string ProjectName & " " & CreationDate.
    'ApplicationName    The name of the application that uses the project.
    'ParentProjectName  If the project is contained within another project (the Parent), this is the name of the parent project.
    'ParentProjectID    If the project is contained within another project (the Parent), this is the ID of the parent project.
    'ParentProjectPath  The path of the parent project.
    'IconNumber         The AppTreeImageList index number of the project's icon.
    'OpenIconNumber     The AppTreeImageList index number of the project's icon for an open project.

    Private _name As String = "" 'The Name of the project.
    Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    'ADDED 10Feb19:
    'Private _appNetName As String = "" 'The name of the Application Network containing the project.
    'Property AppNetName As String
    '    Get
    '        Return _appNetName
    '    End Get
    '    Set(value As String)
    '        _appNetName = value
    '    End Set
    'End Property

    Private _proNetName As String = "" 'The name of the Project Network containing the project.
    Property ProNetName As String
        Get
            Return _proNetName
        End Get
        Set(value As String)
            _proNetName = value
        End Set
    End Property

    Private _creationDate As DateTime = "1-Jan-2000 12:00:00" 'The project creation date.
    Property CreationDate As DateTime
        Get
            Return _creationDate
        End Get
        Set(value As DateTime)
            _creationDate = value
        End Set
    End Property

    Private _description As String = "" 'A description of the project.
    Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

    Private _type As ADVL_Utilities_Library_1.Project.Types = ADVL_Utilities_Library_1.Project.Types.Directory 'The type of location (None, Directory, Archive, Hybrid).
    Property Type As ADVL_Utilities_Library_1.Project.Types
        Get
            Return _type
        End Get
        Set(value As ADVL_Utilities_Library_1.Project.Types)
            _type = value
        End Set
    End Property

    Private _path As String = "" 'The path to the Project directory or archive.
    Property Path As String
        Get
            Return _path
        End Get
        Set(value As String)
            _path = value
        End Set
    End Property

    Private _relativePath As String = "" 'The path relative to the Parent Project. (eg \Import for a directory or \Import.AdvlProject for an archive)
    Property RelativePath As String
        Get
            Return _relativePath
        End Get
        Set(value As String)
            _relativePath = value
        End Set
    End Property

    Private _iD As String = "" 'The ID code of the project. This is the hashcode generated from the ProjectName and CreationDate.
    Property ID As String
        Get
            Return _iD
        End Get
        Set(value As String)
            _iD = value
        End Set
    End Property

    Private _applicationName As String = "" 'The name of the application that created the project.
    Property ApplicationName As String
        Get
            Return _applicationName
        End Get
        Set(value As String)
            _applicationName = value
        End Set
    End Property

    Private _parentProjectName As String = "" 'The Name of the Parent Project.
    Property ParentProjectName As String
        Get
            Return _parentProjectName
        End Get
        Set(value As String)
            _parentProjectName = value
        End Set
    End Property

    Private _parentProjectID As String = "" 'The ID code of the Parent Project. This is the hashcode generated from the ParentProjectName and CreationDate.
    Property ParentProjectID As String
        Get
            Return _parentProjectID
        End Get
        Set(value As String)
            _parentProjectID = value
        End Set
    End Property

    Private _parentProjectPath As String = "" 'The path to the Parent Project directory (or archive?).
    Property ParentProjectPath As String
        Get
            Return _parentProjectPath
        End Get
        Set(value As String)
            _parentProjectPath = value
        End Set
    End Property

    Private _iconNumber As Integer = 0 'The AppTreeImageList index number of the project's icon.
    Property IconNumber As Integer
        Get
            Return _iconNumber
        End Get
        Set(value As Integer)
            _iconNumber = value
        End Set
    End Property

    Private _openIconNumber As Integer = 0 'The AppTreeImageList index number of the project's open icon.
    Property OpenIconNumber As Integer
        Get
            Return _openIconNumber
        End Get
        Set(value As Integer)
            _openIconNumber = value
        End Set
    End Property

End Class 'clsProjInfo








