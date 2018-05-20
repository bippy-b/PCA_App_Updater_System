Imports System.Xml
Imports System.IO
Imports System.Configuration
Imports System.Exception
Imports System.Diagnostics


Public Class frmECSUpdater
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lblProcess As System.Windows.Forms.Label
    Friend WithEvents pbDownload As System.Windows.Forms.ProgressBar
    Friend WithEvents evtlogApp As System.Diagnostics.EventLog
    Friend WithEvents lblProgrammer As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmECSUpdater))
        Me.pbDownload = New System.Windows.Forms.ProgressBar
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.lblProcess = New System.Windows.Forms.Label
        Me.evtlogApp = New System.Diagnostics.EventLog
        Me.lblProgrammer = New System.Windows.Forms.Label
        CType(Me.evtlogApp, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pbDownload
        '
        Me.pbDownload.Location = New System.Drawing.Point(16, 224)
        Me.pbDownload.Name = "pbDownload"
        Me.pbDownload.Size = New System.Drawing.Size(264, 23)
        Me.pbDownload.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(8, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(272, 48)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "ICON Application  Updater"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(8, 88)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(280, 23)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "version 1.3"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblProcess
        '
        Me.lblProcess.Location = New System.Drawing.Point(8, 192)
        Me.lblProcess.Name = "lblProcess"
        Me.lblProcess.Size = New System.Drawing.Size(272, 23)
        Me.lblProcess.TabIndex = 3
        Me.lblProcess.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'evtlogApp
        '
        Me.evtlogApp.Log = "Application"
        Me.evtlogApp.Source = "AppUpdater"
        Me.evtlogApp.SynchronizingObject = Me
        '
        'lblProgrammer
        '
        Me.lblProgrammer.AccessibleDescription = "Label that shows the name of the programmer"
        Me.lblProgrammer.AccessibleName = "Programmerlabel"
        Me.lblProgrammer.AccessibleRole = System.Windows.Forms.AccessibleRole.None
        Me.lblProgrammer.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProgrammer.ForeColor = System.Drawing.Color.White
        Me.lblProgrammer.Location = New System.Drawing.Point(8, 168)
        Me.lblProgrammer.Name = "lblProgrammer"
        Me.lblProgrammer.Size = New System.Drawing.Size(280, 23)
        Me.lblProgrammer.TabIndex = 4
        Me.lblProgrammer.Text = "written by David Benedict"
        Me.lblProgrammer.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'frmECSUpdater
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.RoyalBlue
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblProgrammer)
        Me.Controls.Add(Me.lblProcess)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.pbDownload)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmECSUpdater"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.TopMost = True
        CType(Me.evtlogApp, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Shared DownloadFrom As String
    Shared DownloadTo As String
    Shared AppToLaunch As String
    'Shared myLog As New EventLog

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Initialize the default values for the progress bar
        pbDownload.Minimum = 0
        pbDownload.Maximum = 100
        pbDownload.Value = 0
        pbDownload.ForeColor = System.Drawing.Color.Red
        pbDownload.BackColor = System.Drawing.Color.RoyalBlue
        evtlogApp.Source = "David's Application Updater"
        Dim myICON As New ICON.Maintenance.IconCommon
        'myLog.Source = "David's Application Updater"
        'If Not evtlogApp.SourceExists("David's Application Updater") Then
        '    evtlogApp.CreateEventSource("David's Application Updater", "ApplicationUpdater")
        'End If

        'Set Directories
        DownloadFrom = CType(New System.Configuration.AppSettingsReader().GetValue("DownloadDir", GetType(String)), String)
        DownloadTo = CType(New System.Configuration.AppSettingsReader().GetValue("DownloadToDir", GetType(String)), String)
        AppToLaunch = CType(New System.Configuration.AppSettingsReader().GetValue("ProcessToStart", GetType(String)), String)

        'If the direcotry does not exist, then create it
        If myICON.CheckIfFileExists(DownloadTo) = False Then
            Try
                Directory.CreateDirectory(DownloadTo)
            Catch ex As Exception
                MsgBox("Please check to see if you have permissions to the directory: " & DownloadTo & Chr(13) & Chr(10) & ex.Message, MsgBoxStyle.OKOnly, "An error has occured")
                Exit Sub
            End Try
        End If
        'Show our form
        Me.Show()
        'Let the user know what we are doing
        lblProcess.Text = "Getting Server Manifest File..."
        'Get the server manifest
        Try
            File.Copy(DownloadFrom.ToString & "ServerManifest.xml", DownloadTo.ToString & "ServerManifest.xml", True)
        Catch ex As Exception
            MessageBox.Show("An error has occured while downloading the manifest file.  Please ask David to check the event log.", "Oops!", MessageBoxButtons.OK)
            'evtlogApp.WriteEntry("David's Application Updater", "An error occured updating an application: " & ex.Message, EventLogEntryType.Error)
        Finally
            pbDownload.Value = 25
        End Try
        'Check to see if the file has been updated
        Dim NeedToCopy As Boolean = CompareDates(DownloadTo & "ServerManifest.xml", DownloadTo & "ClientManifest.xml")

        If NeedToCopy = True Then
            lblProcess.Text = "Downloading new files..."
            pbDownload.Value = 50
            Me.Refresh()
            UpdateFiles(DownloadTo & "ServerManifest.xml", DownloadFrom.ToLower, DownloadTo.ToLower)
        Else
            'No updates have been added
            lblProcess.Text = "No updates have been added."
            pbDownload.Value = 75
            Me.Refresh()
            'evtlogApp.WriteEntry("David's Application Updater", "Manifest files match.  No files updated.", EventLogEntryType.Information)
        End If
        'FileCopy(CType( & "ServerManifest.xml"), CType(System.Configuration.ConfigurationSettings.AppSettings.GetValues("DownloadToDir"),String") & "ServerManifest.xml"))
        'Old version v1.0
        ''Download the ECS database file
        'Try
        '    lblProcess.Text = "Copying ECS database..."
        '    Me.Refresh()
        '    'Check to make sure the network is available
        '    If File.Exists("\\Dc1\pro_apps\claims\ECS\everyone\ecs.mdb") Then
        '       File.Copy("\\Dc1\pro_apps\claims\ECS\everyone\ecs.mdb", "C:\ECS\", True)
        '    Else
        '        MsgBox("The database can not be found.", MsgBoxStyle.OKOnly, "ecs.mdb is not available")
        '    End If

        'Catch ex As Exception
        '    MsgBox("Unable to copy the database file" & " " & Err.Description, MsgBoxStyle.Information, "Updating ECS...")
        'End Try
        ''Increment the progress bar
        'pbDownload.Value = 50
        ''Download the ECS executable
        'Try
        '    lblProcess.Text = "Copying ECS executable..."
        '    Me.Refresh()
        '    If File.Exists("\\Dc1\pro_apps\claims\ECS\everyone\ECSSystem.exe") Then
        '      File.Copy("\\Dc1\pro_apps\claims\ECS\everyone\ECSSystem.exe", "C:\ECS\ECSSystem.exe", True)
        '    Else
        '        MsgBox("The executable file can not be found.", MsgBoxStyle.OKOnly, "ECSSystem.exe is not available.")
        '    End If

        'Catch ex As Exception
        '    MsgBox("Unable to copy the executable file" & " " & Err.Description, MsgBoxStyle.Information, "Updating ECS...")
        'End Try
        ''Increment the progress bar
        'pbDownload.Value = 100
        'lblProcess.Text = "Finished updating... Launching ECS."
        lblProcess.Text = "Finished updating your application."
        pbDownload.Value = 100
        Me.Refresh()
        Me.TopMost = False
        Process.Start(DownloadTo & AppToLaunch)
        Me.Visible = False
        Me.Refresh()
        Me.Close()
    End Sub
    Function CompareDates(ByVal server_filename As String, ByVal client_filename As String)
        Dim ClientTime As String
        Dim ServerTime As String
        Try
            If File.Exists(server_filename) Then
                ServerTime = System.IO.File.GetLastWriteTime(server_filename).ToString
            Else
                ServerTime = 0
            End If
            If File.Exists(client_filename) Then
                ClientTime = System.IO.File.GetLastWriteTime(client_filename).ToString
            Else
                'evtlogApp.WriteEntry("David's Application Updater", "No Client Manifest.  Probably the first time run.", EventLogEntryType.Information)
                ClientTime = 1
            End If
        Catch e As Exception
            MessageBox.Show("An error has occured when comparing the manifest file.  Please ask David to check the event log.", "Oops!", MessageBoxButtons.OK)
            'evtlogApp.WriteEntry("David's Application Updater", e.Message, EventLogEntryType.Error)
        End Try

        'Check to see if Server timestamp is different than client timestamp
        If ServerTime = ClientTime Then
            'If the timestamps are the same then no need to copy files, return false
            Return False
        Else
            'If the timestamps are different then we need to copy files
            Return True
        End If
    End Function
    Function UpdateFiles(ByVal server_filename As String, ByVal FromDir As String, ByVal ToDir As String)
        Dim i As Int32 = 0
        Dim j As Int32 = 0
        Dim ServerFileName As String
        Dim ServerDate As String
        Dim ClientDate As String

        Debug.WriteLine("FromDir = " & FromDir)
        Debug.WriteLine("ToDir = " & ToDir)

        'Create the XML Reader
        Dim stream As System.IO.StreamReader
        stream = New StreamReader(server_filename)
        Dim xmlRDR As New XmlTextReader(stream)
        Do While (xmlRDR.Read())


            Select Case xmlRDR.NodeType
                Case XmlNodeType.Element ' The node is an Element
                    If xmlRDR.Name = "file" Then
                        'Debug.WriteLine("Begin Element: " + xmlRDR.Name)
                        While (xmlRDR.MoveToNextAttribute()) ' Read attributes
                            'Debug.WriteLine("Name and Value: " & xmlRDR.Name + " -- " & xmlRDR.Value)
                            If xmlRDR.Name = "name" Then
                                ServerFileName = xmlRDR.Value
                            End If
                            If xmlRDR.Name = "date" Then
                                ServerDate = xmlRDR.Value
                            End If
                        End While
                        'Debug.WriteLine("End Element")
                        Try
                            Debug.WriteLine(ToDir & Replace(ServerFileName, FromDir, ""))
                            Debug.WriteLine("ServerFileName = " & ServerFileName)
                            Debug.WriteLine("FromDir = " & FromDir)
                            ClientDate = System.IO.File.GetLastWriteTime(ToDir & Replace(ServerFileName, FromDir, ""))
                        Catch e As Exception
                            ClientDate = 0
                        End Try
                        If ClientDate = ServerDate Then
                            'Increment the non copied files
                            j = j + 1
                        Else
                            Debug.WriteLine(Replace(ServerFileName, FromDir, ""))
                            If Replace(ServerFileName, FromDir, "") = "servermanifest.xml" Then
                                'Do Nothing
                            Else
                                'Change the logic here to properly get the file name
                                File.Copy(ServerFileName, ToDir & Replace(ServerFileName, FromDir, ""), True)
                                i = i + 1
                            End If

                        End If
                    End If


                Case XmlNodeType.DocumentType ' The node is a DocumentType
                    'Debug.WriteLine("Begin Doc Type" & xmlRDR.NodeType & " " & xmlRDR.Name & " " & xmlRDR.Value)
            End Select

        Loop
        'evtlogApp.WriteEntry("David's Application Updater", "Number of files updated: " & i, EventLogEntryType.Information)
        'evtlogApp.WriteEntry("David's Application Updater", "Number of files checked but not updated: " & j, EventLogEntryType.Information)
        Try
            System.IO.File.Copy(ToDir & "ServerManifest.xml", ToDir & "ClientManifest.xml", True)
        Catch e As Exception
            MessageBox.Show("An error has occured while updating the clientmanifest.xml file.  Please let David B know.", "Oops!", MessageBoxButtons.OK)
            'evtlogApp.WriteEntry("David's Application Updater", "An error occured copying the ServerManifest to the Client Manifest: " & e.Message, EventLogEntryType.Error)
        Finally
            pbDownload.Value = 75
        End Try
        Me.Refresh()
    End Function

End Class
