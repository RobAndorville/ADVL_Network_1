'==============================================================================================================================================================================================
'
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

    Public WithEvents NewHtmlDisplay As frmHtmlDisplay
    Public HtmlDisplayFormList As New ArrayList 'Used for displaying multiple HtmlDisplay forms.

    Public WithEvents NewWebPage As frmWebPage
    Public WebPageFormList As New ArrayList 'Used for displaying multiple WebView forms.

    'Project \ Add Reference \ Assemblies \ Framework \ System.ServiceModel
    Private Shared myHost As ServiceHost
    Dim smb As ServiceMetadataBehavior


    Public WithEvents XMsg As New ADVL_Utilities_Library_1.XMessage
    Dim XDoc As New System.Xml.XmlDocument
    Public Status As New System.Collections.Specialized.StringCollection
    Dim ClientAppName As String 'The name of the client requesting coordinate operations
    Dim MessageText As String 'The text of a message sent through the MessageExchange
    Dim MessageDest As String 'The destination of a message sent through the MessageExchange.

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
    Dim SelectedAppNetName 'The selected Application Network Name - Used when starting a new application or removing a connection.
    '----------------------------------------------------------------------------------------------------------------------------------

    'Application List: ----------------------------------------------------------------------------------------------------------------
    Public App As New App 'App contains a list of all applications. App also contains methods to read, add and save the list.
    'The list is read from an xml file on startup and saved to an xml file on exit.
    'The list is displayed in the dgvApplications datagridview - in the Application List tab.

    'Project List: ----------------------------------------------------------------------------------------------------------------
    Dim Proj As New Proj 'Proj contains a list of all projects. Proj also contains methods to read, add and save the list.
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
    Private StartProject_AppNetName As String 'The Application Network name
    Private StartProject_ProjID As String   'The project ID
    Private StartProject_ProjName As String ' The project name

#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Properties - All the properties used in this form and this application" '============================================================================================================

    Private _instrReceived As String = "" 'Contains Instructions received from the Application Network message service.
    Property InstrReceived As String
        Get
            Return _instrReceived
        End Get
        Set(value As String)
            If value = Nothing Then
                Message.Add("Empty message received!")
            Else
                _instrReceived = value

                Message.XAddText("Message received: " & vbCrLf, "XmlReceivedNotice")

                Dim XDocRec As New System.Xml.XmlDocument
                XDocRec.LoadXml(_instrReceived)

                If _instrReceived.StartsWith("<XMsg>") Then 'This is an XMessage set of instructions.
                    Try
                        Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
                        XDoc.LoadXml(XmlHeader & vbCrLf & _instrReceived)

                        Message.XAddXml(XDoc)
                        Message.XAddText(vbCrLf, "Normal") 'Add extra line

                        XMsg.Run(XDoc, Status)
                    Catch ex As Exception
                        Message.Add("Error running XMsg: " & ex.Message & vbCrLf)
                    End Try

                    'XMessage has been run.
                    'Reply to this message:
                    'Add the message reply to the XMessages window:
                    If ClientAppName = "" Then
                        'No client to send a message to!
                    Else

                        Message.XAddText("Message sent to " & ClientAppName & ":" & vbCrLf, "XmlSentNotice")
                        Message.XAddXml(MessageText)
                        Message.XAddText(vbCrLf, "Normal") 'Add extra line

                        MessageDest = ClientAppName
                        'SendMessage sends the contents of MessageText to MessageDest.
                        SendMessage() 'This subroutine triggers the timer to send the message after a short delay.
                    End If
                Else
                End If
            End If
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

    Private _startPageFileName As String = "" 'The file name of the html document displayed in the Start Page tab.
    Public Property StartPageFileName As String
        Get
            Return _startPageFileName
        End Get
        Set(value As String)
            _startPageFileName = value
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
                               <!---->
                               <SelectedTabIndex><%= TabControl1.SelectedIndex %></SelectedTabIndex>
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
                           </FormSettings>

        'Add code to include other settings to save after the comment line <!---->

        Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & "_" & Me.Text & ".xml"
        Project.SaveXmlSettings(SettingsFileName, settingsData)
    End Sub

    Private Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & "_" & Me.Text & ".xml"

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

            'Add code to read other saved setting here:
            If Settings.<FormSettings>.<SelectedTabIndex>.Value <> Nothing Then TabControl1.SelectedIndex = Settings.<FormSettings>.<SelectedTabIndex>.Value

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

        End If
    End Sub

    Private Sub ReadApplicationInfo()
        'Read the Application Information.

        If ApplicationInfo.FileExists Then
            ApplicationInfo.ReadFile()
        Else
            'There is no Application_Info.xml file.
            DefaultAppProperties() 'Create a new Application Info file with default application properties:
        End If
    End Sub

    Private Sub DefaultAppProperties()
        'These properties will be saved in the Application_Info.xml file in the application directory.
        'If this file is deleted, it will be re-created using these default application properties.

        'Change this to show your application Name, Description and Creation Date.
        ApplicationInfo.Name = "ADVL_Message_Service_1"

        'ApplicationInfo.ApplicationDir is set when the application is started.
        ApplicationInfo.ExecutablePath = Application.ExecutablePath

        'The ADVL_Message_Service application hosts the message service.
        'This is used by Andorville™ software applications To exchange information.

        ApplicationInfo.Description = "The Message Service application hosts the Message Service. This is used by Andorville™ software applications to exchange information."
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
        Project.ApplicationName = ApplicationInfo.Name

        'Set up Message object:
        Message.ApplicationName = ApplicationInfo.Name

        'Set up a temporary initial settings location:
        Dim TempLocn As New ADVL_Utilities_Library_1.FileLocation
        TempLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        TempLocn.Path = ApplicationInfo.ApplicationDir
        Message.SettingsLocn = TempLocn

        Me.Show() 'Show this form before showing the Message form - This will show the App icon on top in the TaskBar.

        'Start showing messages here - Message system is set up.
        Message.AddText("------------------- Starting Application: ADVL Message Service ---------------------- " & vbCrLf, "Heading")
        Message.AddText("Application usage: Total duration = " & Format(ApplicationUsage.TotalDuration.TotalHours, "#.##") & " hours" & vbCrLf, "Normal")

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
            Message.Add("Project.Type.ToString  " & Project.Type.ToString & vbCrLf)
            Message.Add("Project.Path  " & Project.Path & vbCrLf)

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
            End If
        Else
            Project.LockProject() 'Lock the project while it is open in this application.
            ProjectSelected = False 'Reset the Project Selected flag.
        End If

        'START Initialise the form: ===============================================================

        'Set up dgvConnections
        dgvConnections.ColumnCount = 11
        dgvConnections.Columns(0).HeaderText = "Application Network Name"
        dgvConnections.Columns(1).HeaderText = "Application Name"
        dgvConnections.Columns(2).HeaderText = "Connection Name"
        dgvConnections.Columns(3).HeaderText = "Project Name"
        dgvConnections.Columns(4).HeaderText = "Project Type"
        dgvConnections.Columns(5).HeaderText = "Project Path"
        dgvConnections.Columns(6).HeaderText = "Get All Warnings"
        dgvConnections.Columns(7).HeaderText = "Get All Messages"
        dgvConnections.Columns(8).HeaderText = "Callback HashCode"
        dgvConnections.Columns(9).HeaderText = "Connection Start Time"
        dgvConnections.Columns(10).HeaderText = "Connection Duration"
        dgvConnections.Rows.Clear()
        dgvConnections.AutoResizeColumns()
        dgvConnections.AutoResizeRows()
        dgvConnections.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells

        'Set up dgvApplications:
        'Columns in the DataGridView are: Application Name, Description
        dgvApplications.ColumnCount = 2
        dgvApplications.Columns(0).HeaderText = "Name"
        dgvApplications.Columns(1).HeaderText = "Description"
        dgvApplications.Columns(1).DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgvApplications.Rows.Clear()
        dgvApplications.AutoResizeColumns()
        dgvApplications.AutoResizeRows()
        dgvApplications.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells

        'Set up dgvProjects:
        dgvProjects.ColumnCount = 6
        dgvProjects.Columns(0).HeaderText = "Name"
        dgvProjects.Columns(1).HeaderText = "Application Network" 'ADDED 10Feb19
        dgvProjects.Columns(2).HeaderText = "Type"
        dgvProjects.Columns(3).HeaderText = "ID"
        dgvProjects.Columns(4).HeaderText = "Application"
        dgvProjects.Columns(5).HeaderText = "Description"
        dgvProjects.Rows.Clear()
        dgvProjects.AutoResizeColumns()
        dgvProjects.AutoResizeRows()
        dgvProjects.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells

        Me.WebBrowser1.ObjectForScripting = Me

        InitialiseForm() 'Initialise the form for a new project.

        SetUpHost()

        'END   Initialise the form: ---------------------------------------------------------------

        RestoreFormSettings() 'Restore the form settings
        RestoreProjectSettings() 'Restore the Project settings

        ReadApplicationList() 'The list of all Applications Stored in the Application Directory.
        ReadGlobalProjectList() 'This is the list of all projects.

        ShowProjectInfo() 'Show the project information.

        Message.AddText("------------------- Started OK -------------------------------------------------------------------------- " & vbCrLf & vbCrLf, "Heading")

    End Sub

    Private Sub InitialiseForm()
        'Initialise the form for a new project.

        OpenStartPage()

        AppTreeImageList.Images.Clear()
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

        txtTotalDuration.Text = Project.Usage.TotalDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                                Project.Usage.TotalDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                                Project.Usage.TotalDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                                Project.Usage.TotalDuration.Seconds.ToString.PadLeft(2, "0"c)

        txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c)

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Application

        'Check first if there are open connections:
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
                Message.AddWarning("There are " & dgvConnections.Rows.Count - 1 & " connections still open!" & vbCrLf)
                Message.AddWarning("Close these connections before closing the Application Network." & vbCrLf)
                LastExitAttempt = Now
                Exit Sub
            End If
        End If

        SaveAppTree()

        ApplicationInfo.WriteFile() 'Update the Application Information file.

        Project.SaveLastProjectInfo() 'Save information about the last project used.

        'Project.SaveProjectInfoFile() 'Update the Project Information file. This is not required unless there is a change made to the project.

        Project.Usage.SaveUsageInfo() 'Save Project usage information.

        Project.UnlockProject() 'Unlock the project.

        ApplicationUsage.SaveUsageInfo() 'Save Application usage information.
        ApplicationInfo.UnlockApplication()

        WriteApplicationListAdvl_2() 'List of Application stored in the Application Directory.
        'WriteProjectListAdvl_2()
        WriteGlobalProjectListAdvl_2()

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

#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Methods - The main actions performed by this form." '===========================================================================================================================

    Private Sub SetUpHost()
        'Set Up Host:

        'Code Source:
        'https://msdn.microsoft.com/en-us/library/ms731758(v=vs.110).aspx?cs-save-lang=1&cs-lang=vb#code-snippet-4

        Dim baseAddress As Uri = New Uri("http://localhost:8733/ADVLService")
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

    End Sub

    Public Function ConnectionNameAvailable(ByVal AppNetName As String, ByVal ConnName As String) As Boolean
        'If AppNetName-ConnName is already on the dgvConnections list, the name is not available for a new connection and the function returns False.

        Dim NameFound As Boolean = False
        Dim I As Integer 'Loop index
        For I = 0 To dgvConnections.Rows.Count - 1
            If dgvConnections.Rows(I).Cells(2).Value = ConnName Then
                If dgvConnections.Rows(I).Cells(0).Value = AppNetName Then
                    'The Connection named ConnName has been found in the Application Network named AppNetName.
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

    Public Sub RemoveConnectionWithName(ByVal AppNetName As String, ByVal ConnName As String)
        'Remove the connection entry from dgvConnections with the Application Network Name = AppNetName and Connection Name = ConnName.

        Dim I As Integer 'Loop index
        For I = 0 To dgvConnections.Rows.Count - 1
            If dgvConnections.Rows(I).Cells(2).Value = ConnName Then
                If dgvConnections.Rows(I).Cells(0).Value = AppNetName Then
                    'The Connection named ConnName has been found in the Application Network named AppNetName.
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

        Message.Add("START: SaveAppTree()" & vbCrLf)

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

                If tnc(I).Nodes.Count > 0 Then
                    Message.Add("Node name = " & tnc(I).Name & " IsExpanded: " & tnc(I).IsExpanded & vbCrLf)
                End If

                'If NodeKey = "Application_Network" Then 'This the root node of the Application Tree.
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

                        'Dim myProjRelativePath As New XElement("RelativePath", ProjInfo(NodeKey).RelativePath)
                        'myNode.Add(myProjRelativePath)

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
        'Open the StartPage.html file and display in the Start Page tab.

        If Project.DataFileExists("StartPage.html") Then
            StartPageFileName = "StartPage.html"
            DisplayStartPage()
        Else
            CreateStartPage()
            StartPageFileName = "StartPage.html"
            DisplayStartPage()
        End If

    End Sub

    Public Sub DisplayStartPage()
        'Display the StartPage.html file in the Start Page tab.

        If Project.DataFileExists(StartPageFileName) Then
            Dim rtbData As New IO.MemoryStream
            Project.ReadData(StartPageFileName, rtbData)
            rtbData.Position = 0
            Dim sr As New IO.StreamReader(rtbData)
            WebBrowser1.DocumentText = sr.ReadToEnd()
        Else
            Message.AddWarning("Web page file not found: " & StartPageFileName & vbCrLf)
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
        sb.Append("</head>" & vbCrLf)

        sb.Append("<body style=""font-family:arial;"">" & vbCrLf & vbCrLf)

        sb.Append("<h2>" & "Andorville&trade; Message Service" & "</h2>" & vbCrLf & vbCrLf) 'Add the page title.
        sb.Append("<hr>") 'Add a horizontal divider line.
        sb.Append("<p>The Message Service is used by Andorville&trade; applications to exchange information.</p>" & vbCrLf) 'Add an application description.
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
        sb.Append("<head>" & vbCrLf)
        sb.Append("<title>" & DocumentTitle & "</title>" & vbCrLf)
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

        If System.IO.File.Exists(ApplicationInfo.ApplicationDir & "\" & "Application_List_ADVL_2.xml") Then 'Latest format version of the Application List found.
            ReadApplicationListAdvl_2()
        Else 'The Application List was found.
            Message.AddWarning("The Application List Xml document was not found." & vbCrLf)
        End If
    End Sub

    Private Sub ReadApplicationListAdvl_2()
        'Read the Application_List.xml file in the Application Directory. (ADVL_2 format version.)

        If System.IO.File.Exists(ApplicationInfo.ApplicationDir & "\Application_List_ADVL_2.xml") Then
            Dim AppListXDoc As System.Xml.Linq.XDocument
            AppListXDoc = XDocument.Load(ApplicationInfo.ApplicationDir & "\Application_List_ADVL_2.xml")

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

        ApplicationListXDoc.Save(ApplicationInfo.ApplicationDir & "\Application_List_ADVL_2.xml")

    End Sub

    Private Sub ReadGlobalProjectList()
        'Read the Project List. 

        If System.IO.File.Exists(ApplicationInfo.ApplicationDir & "\" & "Global_Project_List_ADVL_2.xml") Then 'Latest format version of the Project List found.
            'ReadProjectListAdvl_2()
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
                If item.<AppNetName>.Value Is Nothing Then
                    NewProj.AppNetName = ""
                Else
                    NewProj.AppNetName = item.<AppNetName>.Value
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
            dgvProjects.Rows(Index).Cells(1).Value = Proj.List(Index).AppNetName 'ADDED 10Feb19
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
                                              <AppNetName><%= item.AppNetName %></AppNetName>
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

    Private Sub btnProject_Click(sender As Object, e As EventArgs) Handles btnProject.Click
        Project.SelectProject()
    End Sub

    Private Sub btnOpenProject2_Click(sender As Object, e As EventArgs) Handles btnOpenProject2.Click
        If Project.Type = ADVL_Utilities_Library_1.Project.Types.Archive Then

        Else
            Process.Start(Project.Path)
        End If
    End Sub

    Private Sub btnOpenSettings_Click(sender As Object, e As EventArgs) Handles btnOpenSettings.Click
        If Project.SettingsLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory Then
            Process.Start(Project.SettingsLocn.Path)
        End If
    End Sub

    Private Sub btnOpenData_Click(sender As Object, e As EventArgs) Handles btnOpenData.Click
        If Project.DataLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory Then
            Process.Start(Project.DataLocn.Path)
        End If
    End Sub

    Private Sub btnOpenSystem_Click(sender As Object, e As EventArgs) Handles btnOpenSystem.Click
        If Project.SystemLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory Then
            Process.Start(Project.SystemLocn.Path)
        End If
    End Sub

    Private Sub btnOpenAppDir_Click(sender As Object, e As EventArgs) Handles btnOpenAppDir.Click
        Process.Start(ApplicationInfo.ApplicationDir)
    End Sub

    Private Sub btnParameters_Click(sender As Object, e As EventArgs) Handles btnParameters.Click
        Project.ShowParameters()
    End Sub


#Region " Methods Called by JavaScript" '============================================================================
    '- A collection of methods that can be called by JavaScript in a web page shown in WebBrowser1

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

    'Public Sub RestoreHtmlSettings_Old(ByVal FileName As String)
    '    'Restore the Html settings for a web page.

    '    Dim XDocSettings As New System.Xml.Linq.XDocument
    '    Project.ReadXmlData(FileName, XDocSettings)

    '    If XDocSettings Is Nothing Then

    '    Else
    '        Dim XSettings As New System.Xml.XmlDocument
    '        Try
    '            XSettings.LoadXml(XDocSettings.ToString)

    '            'Run the Settings file:
    '            XSeq.RunXSequence(XSettings, XStatus)
    '        Catch ex As Exception
    '            Message.AddWarning("Error restoring HTML settings. " & ex.Message & vbCrLf)
    '        End Try
    '    End If
    'End Sub

    Public Sub RestoreHtmlSettings()
        'Restore the Html settings for a web page.

        Dim SettingsFileName As String = StartPageFileName & "Settings"

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


    Private Sub XSeq_ErrorMsg(ErrMsg As String) Handles XSeq.ErrorMsg
        Message.AddWarning(ErrMsg & vbCrLf)
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

    Public Function GetFormNo() As String
        'Return FormNo.ToString
        Return "-1"
    End Function

    Public Sub AddText(ByVal Msg As String, ByVal TextType As String)
        Message.AddText(Msg, TextType)
    End Sub

    Public Sub AddMessage(ByVal Msg As String)
        Message.Add(Msg)
    End Sub

    Public Sub AddWarning(ByVal Msg As String)
        Message.AddWarning(Msg)
    End Sub


    Public Sub SendXMessage(ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMessage to the application with the connection name ConnName.

    End Sub

    Public Sub RunXSequence(ByVal XSequence As String)
        'Run the XMSequence
        Dim XmlSeq As New System.Xml.XmlDocument
        XmlSeq.LoadXml(XSequence)
        XSeq.RunXSequence(XmlSeq, Status)

    End Sub

    Public Sub OpenWebPage(ByVal WebPageFileName As String)
        'Open a Web Page from the WebPageFileName.
        '  Pass the ParentName Property to the new web page. The is the name of this web page that is opening the new page.
        '  Pass the ParentWebPageFormNo Property to the new web page. This is the FormNo of this web page that is opening the new page.
        '    A hash code is generated from the ParentName. This is used to define a file name to save and restore the Web Page settings.
        '    The new web page can pass instructions back to the ParentWebPage using its ParentWebPageFormNo.

        Dim NewFormNo As Integer = OpenNewWebPage()

        WebPageFormList(NewFormNo).ParentWebPageFileName = StartPageFileName 'Set the Parent Web Page property.
        WebPageFormList(NewFormNo).ParentWebPageFormNo = -1 'Set the Parent Form Number property.
        WebPageFormList(NewFormNo).Description = ""             'The web page description can be blank.
        WebPageFormList(NewFormNo).FileDirectory = ""           'Only Web files in the Project directory can be opened from another Web Page Form.
        WebPageFormList(NewFormNo).FileName = WebPageFileName  'Set the web page file name to be opened.
        WebPageFormList(NewFormNo).OpenDocument                'Open the web page file name.

    End Sub


#End Region 'Methods Called by JavaScript ---------------------------------------------------------------------------


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

    Private Sub dgvProjects_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvProjects.CellContentClick

    End Sub

    Private Sub dgvProjects_SelectionChanged(sender As Object, e As EventArgs) Handles dgvProjects.SelectionChanged
        If dgvProjects.SelectedRows.Count > 0 Then
            Dim RowNo As Integer = dgvProjects.SelectedRows(0).Index
            If Proj.List.Count > RowNo Then
                txtProjectPath.Text = Proj.List(RowNo).Path
            End If
        End If
    End Sub

    Private Sub dgvProjects_Click(sender As Object, e As EventArgs) Handles dgvProjects.Click

    End Sub

    Private Sub XMsg_Instruction(Info As String, Locn As String) Handles XMsg.Instruction
        'Process each Property Path and Property Value instruction.

        Select Case Locn

            Case "NewConnectionInfo:ApplicationNetworkName"
                SelectedAppNetName = Info

            Case "NewConnectionInfo:ConnectionName"

                If ConnectionNameAvailable(SelectedAppNetName, Info) Then
                    AddNewConnection = True
                    dgvConnections.Rows.Add()
                    Dim CurrentRow As Integer = dgvConnections.Rows.Count - 2
                    dgvConnections.Rows(CurrentRow).Cells(0).Value = SelectedAppNetName 'Add the AppNet Name to dgvConnections.
                    dgvConnections.Rows(CurrentRow).Cells(2).Value = Info 'Add the ConnectionName to dgvConnections.
                    dgvConnections.AutoResizeRows()
                Else
                    AddNewConnection = False
                End If

            Case "NewConnectionInfo:ApplicationName"
                If AddNewConnection = True Then
                    Dim CurrentRow As Integer = dgvConnections.Rows.Count - 2
                    dgvConnections.Rows(CurrentRow).Cells(1).Value = Info 'Add the ApplicationName to dgvConnections.
                    dgvConnections.AutoResizeRows()
                End If

            Case "NewConnectionInfo:ProjectName"
                If AddNewConnection = True Then
                    Dim CurrentRow As Integer = dgvConnections.Rows.Count - 2
                    dgvConnections.Rows(CurrentRow).Cells(3).Value = Info
                    dgvConnections.AutoResizeRows()
                End If

            Case "NewConnectionInfo:ProjectType"
                If AddNewConnection = True Then
                    Dim CurrentRow As Integer = dgvConnections.Rows.Count - 2
                    dgvConnections.Rows(CurrentRow).Cells(4).Value = Info
                    dgvConnections.AutoResizeRows()
                End If

            Case "NewConnectionInfo:ProjectPath"
                If AddNewConnection = True Then
                    Dim CurrentRow As Integer = dgvConnections.Rows.Count - 2
                    dgvConnections.Rows(CurrentRow).Cells(5).Value = Info
                    dgvConnections.AutoResizeRows()
                End If

            Case "NewConnectionInfo:GetAllWarnings"
                If AddNewConnection = True Then
                    Dim CurrentRow As Integer = dgvConnections.Rows.Count - 2
                    dgvConnections.Rows(CurrentRow).Cells(6).Value = Info
                    dgvConnections.AutoResizeRows()
                End If

            Case "NewConnectionInfo:GetAllMessages"
                If AddNewConnection = True Then
                    Dim CurrentRow As Integer = dgvConnections.Rows.Count - 2
                    dgvConnections.Rows(CurrentRow).Cells(7).Value = Info
                    dgvConnections.AutoResizeRows()
                End If

            Case "NewConnectionInfo:CallbackHashcode"
                If AddNewConnection = True Then
                    Dim CurrentRow As Integer = dgvConnections.Rows.Count - 2
                    dgvConnections.Rows(CurrentRow).Cells(8).Value = Info
                    dgvConnections.AutoResizeRows()
                End If

            Case "NewConnectionInfo:ConnectionStartTime"
                If AddNewConnection = True Then
                    Dim CurrentRow As Integer = dgvConnections.Rows.Count - 2
                    dgvConnections.Rows(CurrentRow).Cells(9).Value = Info
                    dgvConnections.Rows(CurrentRow).Cells(10).Value = 0 'Current duration of the connection
                    dgvConnections.AutoResizeRows()
                End If

           '---------------------------------------------------------------------------------------------------------------------------------------------


           'Add an Application Info entry ---------------------------------------------------------------------------------------------------------------
            Case "ApplicationInfo:Name"
                'Code used to add application to the Application List: (TO BE REPLACED WITH THE APPLICATION DICTIONARY.)
                If ApplicationNameAvailable(Info) Then
                    AddNewApplication = True
                    dgvApplications.Rows.Add()
                    Dim CurrentRow As Integer = dgvApplications.Rows.Count - 2
                    dgvApplications.Rows(CurrentRow).Cells(0).Value = Info
                    Dim NewAppInfo As New AppSummary
                    NewAppInfo.Name = Info
                    App.List.Add(NewAppInfo)

                Else
                    AddNewApplication = False
                End If
                AppName = Info
                If AppInfo.ContainsKey(Info) Then
                    AddNewApp = False
                    AppName = Info
                Else
                    AddNewApp = True
                    AppInfo.Add(Info, New clsAppInfo) 'Text, Description, ExecutablePath, 
                    AppName = Info
                End If

            Case "ApplicationInfo:Text"
                If AddNewApp = True Then
                    AppText = Info
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
                            node(0).Text = Info
                        End If
                    End If
                End If

            Case "ApplicationInfo:Directory"
                If AddNewApplication = True Then
                    Dim CurrentRow As Integer = dgvApplications.Rows.Count - 2
                    'Applications grid now shows only Name and Description
                    App.List(CurrentRow).Directory = Info
                End If
                If AddNewApp = True Then
                    AppInfo(AppName).Directory = Info
                Else
                    If AppInfo(AppName).Directory = Info Then
                        'The application directory is unchanged.
                    Else
                        AppInfo(AppName).Directory = Info 'The application directory has been updated.
                    End If
                End If

            Case "ApplicationInfo:Description"
                If AddNewApplication = True Then
                    Dim CurrentRow As Integer = dgvApplications.Rows.Count - 2
                    dgvApplications.Rows(CurrentRow).Cells(1).Value = Info
                    dgvApplications.AutoResizeRows()
                    App.List(CurrentRow).Description = Info

                Else
                    If App.List(ApplicationNo).Description = Info Then
                        'Executable path has not been changed.
                    Else
                        'Executable path has been changed.
                        App.List(ApplicationNo).Description = Info
                        Message.Add("Application description for " & App.List(ApplicationNo).Name & " has been changed to: " & vbCrLf & App.List(ApplicationNo).Description & vbCrLf)
                    End If
                End If

                If AddNewApp = True Then
                    AppInfo(AppName).Description = Info
                Else
                    If AppInfo(AppName).Description = Info Then
                        'The application description is unchanged.
                    Else
                        AppInfo(AppName).Description = Info 'The application description has been updated.
                    End If
                End If

            Case "ApplicationInfo:ExecutablePath"
                If AddNewApplication = True Then
                    Dim CurrentRow As Integer = dgvApplications.Rows.Count - 2
                    'Applications grid now shows only Name and Description
                    App.List(CurrentRow).ExecutablePath = Info
                Else
                    If App.List(ApplicationNo).ExecutablePath = Info Then
                        'Executable path has not been changed.
                    Else
                        'Executable path has been changed.
                        App.List(ApplicationNo).ExecutablePath = Info
                        Message.Add("Executable path for " & App.List(ApplicationNo).Name & " has been changed to: " & vbCrLf & App.List(ApplicationNo).ExecutablePath & vbCrLf)
                    End If
                End If

                If AddNewApp = True Then
                    AppInfo(AppName).ExecutablePath = Info
                    'Get the application icon:
                    Dim myIcon = System.Drawing.Icon.ExtractAssociatedIcon(Info)
                    AppTreeImageList.Images.Add(AppName, myIcon)
                    AppInfo(AppName).IconNumber = AppTreeImageList.Images.IndexOfKey(AppName)
                    AppInfo(AppName).OpenIconNumber = AppTreeImageList.Images.IndexOfKey(AppName)
                Else
                    If AppInfo(AppName).ExecutablePath = Info Then
                        'The application executable path is unchanged.
                    Else
                        AppInfo(AppName).ExecutablePath = Info 'The application executable path has been updated.
                    End If
                End If

            '----------------------------------------------------------------------------------------------------------------------------------------------


           'Add a Project -------------------------------------------------------------------------------------------------------------------------------
            Case "ProjectInfo:Path"
                ProcessNewProject(Info)

           '---------------------------------------------------------------------------------------------------------------------------------------------

            'Remove a connection entry --------------------------------------------------------------------------------------------------------------------

            Case "RemovedConnectionInfo:ApplicationNetworkName"
                SelectedAppNetName = Info

            Case "RemovedConnectionInfo:ApplicationName"
                'RemoveConnectionWithAppName(Info)
                Message.AddWarning("Instruction not in use: RemovedConnectionInfo:ApplicationName" & vbCrLf)
                Message.AddWarning("Modify your code to use this instruction: RemovedConnectionInfo:ConnectionName" & vbCrLf)

            Case "RemovedConnectionInfo:ConnectionName"
                RemoveConnectionWithName(SelectedAppNetName, Info)

          '---------------------------------------------------------------------------------------------------------------------------------------------




            Case "EndOfSequence"
                If AddNewApp = True Then
                    'Add the new application node to the tree: 
                    trvAppTree.TopNode.Nodes.Add(AppName, AppText, AppInfo(AppName).IconNumber, AppInfo(AppName).OpenIconNumber)
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

                SelectedAppNetName = ""

            Case Else
                Message.Add("Instruction not recognised:  " & Locn & "    Property:  " & Info & vbCrLf)

        End Select

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
            'Application node
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
                        UpdateAppTreeImageIndexes(trvAppTree.TopNode) 'This is needed to update the TreeView node icons.
                    End If

                End If
            End If
        End If
    End Sub

    Private Sub UpdateAppTreeImageIndexes(ByRef Node As TreeNode)
        'Update the AppTree inages indexes.

        If Node.Name.EndsWith(".Proj") Then
            'Project node - The project icon indexes do not change.
        Else
            'Application node - update the icons.
            Node.ImageIndex = AppInfo(Node.Name).IconNumber
            Node.SelectedImageIndex = AppInfo(Node.Name).OpenIconNumber
        End If

        For Each ChildNode As TreeNode In Node.Nodes
            UpdateAppTreeImageIndexes(ChildNode)
        Next

    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        'Start the selected application.

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

        Message.Add(vbCrLf & "Processing Project:" & vbCrLf)
        Message.Add("Project path: " & ProjectPath & vbCrLf)

        'Check if ProjectPath is a File or a Directory:
        Dim Attr As System.IO.FileAttributes = IO.File.GetAttributes(ProjectPath)
        If Attr.HasFlag(IO.FileAttributes.Directory) Then
            Message.Add("Project path is a Directory." & vbCrLf)
            If System.IO.File.Exists(ProjectPath & "\Project_Info_ADVL_2.xml") Then
                Message.Add("The directory is an Andorville(TM) project." & vbCrLf)
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
                Message.Add("The file is an Andorville(TM) project." & vbCrLf)
                ReadDragDropArchiveProjectInfo(ProjectPath)
            Else
                Message.Add("The file is not an Andorville(TM) project." & vbCrLf)
            End If
        End If

    End Sub

    Private Sub ReadDragDropDirectoryProjectInfo(ByVal ProjectPath As String)
        'Read the Project Information from a Directory Project.

        Dim ProjectInfo As System.Xml.Linq.XDocument = XDocument.Load(ProjectPath & "\Project_Info_ADVL_2.xml")

        Dim ProjectAppNetName As String
        If System.IO.File.Exists(ProjectPath & "\ProjectParams_ADVL2.xml") Then
            Dim ParameterInfo As System.Xml.Linq.XDocument = XDocument.Load(ProjectPath & "\ProjectParams_ADVL2.xml")
            Dim AppNetNames = From names In ParameterInfo.<ProjectParameterList>.<Parameter>
                              Where names.<Name>.Value = "AppNetName"
                              Select names

            If AppNetNames.Count = 0 Then
                ProjectAppNetName = ""
                Message.Add("The Project Parameters file did not contain an AppNetName parameter." & vbCrLf)
            ElseIf AppNetNames.Count = 1 Then
                ProjectAppNetName = AppNetNames(0).<Value>.Value
            Else
                ProjectAppNetName = AppNetNames(0).<Value>.Value
                Message.Add("The Project Parameters file contained more than one AppNetName parameter." & vbCrLf)
            End If
        Else
            ProjectAppNetName = ""
            Message.Add("The project did not contain a Project Parameters file." & vbCrLf)
        End If

        Message.Add(vbCrLf) 'Add a blank line.

        Dim ProjectName As String
        If ProjectInfo.<Project>.<Name>.Value = Nothing Then
            ProjectName = ""
        Else
            ProjectName = ProjectInfo.<Project>.<Name>.Value
        End If
        Message.Add("Project Name = " & ProjectName & vbCrLf)

        Dim ProjectID As String
        If ProjectInfo.<Project>.<ID>.Value = Nothing Then
            ProjectID = ""
            Message.AddWarning("The Project ID is blank." & vbCrLf)
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
            dgvProjects.Rows(CurrentRow).Cells(1).Value = ProjectAppNetName 'ADDED 10Feb19
            dgvProjects.Rows(CurrentRow).Cells(2).Value = ProjectType
            dgvProjects.Rows(CurrentRow).Cells(3).Value = ProjectID
            dgvProjects.Rows(CurrentRow).Cells(4).Value = ApplicationName
            dgvProjects.Rows(CurrentRow).Cells(5).Value = ProjectDescription
            dgvProjects.AutoResizeColumns()

                Dim NewProjectInfo As New ProjSummary
            NewProjectInfo.Name = ProjectName
            NewProjectInfo.AppNetName = ProjectAppNetName 'Added 10Feb19
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
            Message.Add("Project is already in the TreeView. Project ID = " & ProjectID & vbCrLf)
        Else
            ProjInfo.Add(ProjectID & ".Proj", New clsProjInfo)
            ProjInfo(ProjectID & ".Proj").Name = ProjectName
            ProjInfo(ProjectID & ".Proj").AppNetName = ProjectAppNetName
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
                        If ApplicationName = trvAppTree.TopNode.Name Then
                            node = trvAppTree.Nodes.Find(ApplicationName, False)
                        Else
                            node = trvAppTree.TopNode.Nodes.Find(ApplicationName, False)
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
                        If ApplicationName = trvAppTree.TopNode.Name Then
                            node = trvAppTree.Nodes.Find(ApplicationName, False)
                        Else
                            node = trvAppTree.TopNode.Nodes.Find(ApplicationName, False)
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
                        If ApplicationName = trvAppTree.TopNode.Name Then
                            node = trvAppTree.Nodes.Find(ApplicationName, False)
                        Else
                            node = trvAppTree.TopNode.Nodes.Find(ApplicationName, False)
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
                        If ApplicationName = trvAppTree.TopNode.Name Then
                            node = trvAppTree.Nodes.Find(ApplicationName, False)
                        Else
                            node = trvAppTree.TopNode.Nodes.Find(ApplicationName, False)
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

        Message.Add(vbCrLf) 'Add a blank line.

        Dim ProjectName As String
        If ProjectInfo.<Project>.<Name>.Value = Nothing Then
            ProjectName = ""
        Else
            ProjectName = ProjectInfo.<Project>.<Name>.Value
        End If
        Message.Add("Project Name = " & ProjectName & vbCrLf)

        Dim ProjectAppNetName As String
        If ProjectInfo.<Project>.<AppNetName>.Value = Nothing Then
            ProjectAppNetName = ""
        Else
            ProjectAppNetName = ProjectInfo.<Project>.<AppNetName>.Value
        End If
        Message.Add("Project Application Network Name = " & ProjectAppNetName & vbCrLf)

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
            dgvProjects.Rows(CurrentRow).Cells(1).Value = ProjectAppNetName 'ADDED 10Feb19
            dgvProjects.Rows(CurrentRow).Cells(2).Value = ProjectType
            dgvProjects.Rows(CurrentRow).Cells(3).Value = ProjectID
            dgvProjects.Rows(CurrentRow).Cells(4).Value = ApplicationName
            dgvProjects.Rows(CurrentRow).Cells(5).Value = ProjectDescription
            dgvProjects.AutoResizeColumns()

            Dim NewProjectInfo As New ProjSummary
            NewProjectInfo.Name = ProjectName
            NewProjectInfo.AppNetName = ProjectAppNetName 'ADDED 10Feb19
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
        If ProjInfo.ContainsKey(ProjectID & ".Proj") Then
            Message.Add("Project is already in the TreeView. Project ID = " & ProjectID & vbCrLf)
        Else
            ProjInfo.Add(ProjectID & ".Proj", New clsProjInfo)
            ProjInfo(ProjectID & ".Proj").Name = ProjectName
            ProjInfo(ProjectID & ".Proj").AppNetName = ProjectAppNetName 'ADDED 10Feb19
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
                    If ApplicationName = trvAppTree.TopNode.Name Then
                        node = trvAppTree.Nodes.Find(ApplicationName, False)
                    Else
                        node = trvAppTree.TopNode.Nodes.Find(ApplicationName, False)
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
                    If ApplicationName = trvAppTree.TopNode.Name Then
                        node = trvAppTree.Nodes.Find(ApplicationName, False)
                    Else
                        node = trvAppTree.TopNode.Nodes.Find(ApplicationName, False)
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
                    If ApplicationName = trvAppTree.TopNode.Name Then
                        node = trvAppTree.Nodes.Find(ApplicationName, False)
                    Else
                        node = trvAppTree.TopNode.Nodes.Find(ApplicationName, False)
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
                    If ApplicationName = trvAppTree.TopNode.Name Then
                        node = trvAppTree.Nodes.Find(ApplicationName, False)
                    Else
                        node = trvAppTree.TopNode.Nodes.Find(ApplicationName, False)
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

    Private Function ProjectNameAndAppNetNameAvailable(ByVal Name As String, ByVal AppNetName As String) As Boolean
        'If Name and AppNetName is not in the Project list, ProjectNameAndAppNetNameAvailable is set to True.

        Dim Found As Boolean = False
        Dim I As Integer 'Loop index
        For I = 0 To dgvProjects.Rows.Count - 1
            If dgvProjects.Rows(I).Cells(0).Value = Name Then
                If dgvProjects.Rows(I).Cells(1).Value = AppNetName Then
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

        txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c)

        Timer2.Interval = 5000 '5 seconds
        Timer2.Enabled = True
        Timer2.Start()
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        'Update the current duration:

        txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c)
    End Sub

    Private Sub TabPage2_Leave(sender As Object, e As EventArgs) Handles TabPage2.Leave
        Timer2.Enabled = False
    End Sub

    Private Sub btnSendMessage_Click(sender As Object, e As EventArgs) Handles btnSendMessage.Click
        'Test code - Try to send a message to the selected connection.

        Dim ConnectionName As String
        Dim SelRow As Integer
        If dgvConnections.SelectedRows.Count > 0 Then
            If dgvConnections.SelectedRows.Count = 1 Then
                SelRow = dgvConnections.SelectedRows(0).Index
                ConnectionName = dgvConnections.Rows(SelRow).Cells(1).Value
                Message.Add("Selected connection name: " & ConnectionName & vbCrLf)
                'SendMessage(ConnectionName, "Test")
                'SendMessage()

            Else

            End If
        Else
            Message.AddWarning("No connections have been selected." & vbCrLf)
        End If

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

    Private Sub Label26_Click(sender As Object, e As EventArgs) Handles Label26.Click

    End Sub

    Private Sub Label26_MouseHover(sender As Object, e As EventArgs) Handles Label26.MouseHover
        'Update the ToolTip text:
        ToolTip1.SetToolTip(Label26, Label26.Text) 'This will allow the full filename to be read if it is cropped at the edge of the window.
    End Sub


    Private Sub XmlHtmDisplay2_DragEnter(sender As Object, e As DragEventArgs) Handles XmlHtmDisplay2.DragEnter
        'DragEnter: An object has been dragged into XmlHtmDisplay2.
        'This code is required to get the link to the item(s) being dragged into XmlHtmDisplay2:
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Link
        End If

    End Sub

    Private Sub XmlHtmDisplay2_DragDrop(sender As Object, e As DragEventArgs) Handles XmlHtmDisplay2.DragDrop
        'DragDrop.

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

    Private Sub btnOpenProject_Click(sender As Object, e As EventArgs) Handles btnOpenProject.Click
        'Start the selected project

        If dgvProjects.SelectedRows.Count = 0 Then
            Message.AddWarning("No project has been selected." & vbCrLf)
        ElseIf dgvProjects.SelectedRows.Count = 1 Then
            Dim SelRowNo As Integer = dgvProjects.SelectedRows(0).Index
            Dim StartAppName As String = dgvProjects.Rows(SelRowNo).Cells(4).Value
            Dim StartAppConnName As String = dgvProjects.Rows(SelRowNo).Cells(4).Value 'Use the AppName as the Connection Name. (The connection name scan be duplicated al long as the AppNetNames are different.)
            Dim StartAppProjectPath As String = Proj.List(SelRowNo).Path

            Dim AppNetName As String = Proj.List(SelRowNo).AppNetName
            If StartAppName = "ADVL_Application_Network_1" Then AppNetName = Proj.List(SelRowNo).Name

            If ConnectionNameAvailable(AppNetName, StartAppConnName) Then
                StartApp_ProjectPath(StartAppName, StartAppProjectPath, StartAppConnName)
            Else
                Message.AddWarning("Connection name: " & StartAppConnName & " already used in the Application Network: " & AppNetName & vbCrLf)
            End If

        Else 'More than one project selected.
            Message.AddWarning("Two or more projects have been selected. Code to start these will be added later." & vbCrLf)
        End If

    End Sub


#End Region 'Send XMessages -----------------------------------------------------------------------------------------

    Private Sub XSeq_Instruction(Info As String, Locn As String) Handles XSeq.Instruction
        'Execute each instruction produced by running the XSeq file.

        Select Case Locn

            'Start Project commands: ----------------------------------------------------
            Case "StartProject:AppName"
                StartProject_AppName = Info

            Case "StartProject:ConnectionName"
                StartProject_ConnName = Info

            Case "StartProject:AppNetName"
                StartProject_AppNetName = Info

            Case "StartProject:ProjectID"
                StartProject_ProjID = Info

            Case "StartProject:ProjectName"
                StartProject_ProjName = Info

            Case "StartProject:Command"
                Select Case Info
                    Case "Apply"
                        If StartProject_ProjName <> "" Then
                            StartApp_ProjectName(StartProject_AppName, StartProject_AppNetName, StartProject_ProjName, StartProject_ConnName)
                        ElseIf StartProject_ProjID <> "" Then

                        Else
                            Message.AddWarning("Project not specified. Project Name and Project ID are blank." & vbCrLf)
                        End If
                    Case Else
                        Message.AddWarning("Unknown Start Project command : " & Info & vbCrLf)
                End Select

            'END Start project commands ---------------------------------------------


            Case "Settings"

            Case "EndOfSequence"
                Message.Add("End of processing sequence" & Info & vbCrLf)
                'Clear the StartProject variables:
                StartProject_AppName = ""
                StartProject_ConnName = ""
                StartProject_AppNetName = ""
                StartProject_ProjID = ""
                StartProject_ProjName = ""

            Case Else
                Message.AddWarning("Unknown location: " & Locn & "  Info: " & Info & vbCrLf)

        End Select
    End Sub

    Public Sub StartApp_ProjectName(ByVal AppName As String, ByVal AppNetName As String, ByVal ProjectName As String, ByVal ConnectionName As String)
        'Start the application with the name AppName.

        If AppInfo.ContainsKey(AppName) Then
            'Start the application:
            If ProjectName = "" And ConnectionName = "" Then
                'No project selected and application will not be connected to the network.
                Shell(Chr(34) & AppInfo(AppName).ExecutablePath & Chr(34), AppWinStyle.NormalFocus) 'Start the application with no argument
            Else
                If ConnectionNameAvailable(AppNetName, ConnectionName) Then

                    Dim ProjectPath As String = Proj.FindNameAndAppNet(ProjectName, AppNetName).Path

                    'Temp code - until Application Network projects have the AppNetName added:
                    If AppName = "ADVL_Application_Network_1" Then
                        ProjectPath = Proj.FindNameAndAppNet(ProjectName, "").Path
                    End If

                    If ProjectPath = "" Then
                        Message.AddWarning("Project Path not found for Project: " & ProjectName & " in AppNet: " & AppNetName & vbCrLf)
                        Exit Sub
                    End If

                    'Build the Application start message:
                    Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                    Dim ConnectDoc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                    Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class

                    'NOTE: The project should always be opened using a Project Path because project names may not be unique.
                    'Use the AppNetName and ProjectName to get the ProjectPath from the Project List.

                    Dim xProjectPath As New XElement("ProjectPath", ProjectPath)
                    xmessage.Add(xProjectPath)

                    'NOTE: The application currently determines the AppNetName using other information!
                    If AppNetName <> "" Then
                        Dim xAppNetName As New XElement("AppNetName", AppNetName)
                        xmessage.Add(xAppNetName)
                    End If

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
                    Message.AddWarning("Connection name already in use: ConnName: " & ConnectionName & " in the Application Network: " & AppNetName & vbCrLf)
                End If
                'End If
            End If
        End If
    End Sub

    Private Sub btnAppInfo_Click(sender As Object, e As EventArgs) Handles btnAppInfo.Click
        ApplicationInfo.ShowInfo()
    End Sub

    Private Sub btnAndorville_Click(sender As Object, e As EventArgs) Handles btnAndorville.Click
        ApplicationInfo.ShowInfo()
    End Sub

#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------


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

    Public Function FindNameAndAppNet(ByVal Name As String, ByVal AppNetName As String) As ProjSummary
        'Return the ProjSummary corresponding to the Project with specified Name and AppNetName.

        Dim FoundProj As ProjSummary

        FoundProj = List.Find(Function(item As ProjSummary)
                                  Return item.Name = Name And item.AppNetName = AppNetName
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

    Private _appNetName As String = "" 'The name of the Application Network containing the project. (Added 10Feb19.)
    Property AppNetName As String
        Get
            Return _appNetName
        End Get
        Set(value As String)
            _appNetName = value
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
    'AppNetName         The name of the Application Network containig the project. (Added 10Feb19.)
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
    Private _appNetName As String = "" 'The name of the Application Network containing the project.
    Property AppNetName As String
        Get
            Return _appNetName
        End Get
        Set(value As String)
            _appNetName = value
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






