Imports System.Web
Imports System.Xml
Imports System.IO

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
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmECSUpdater))
        Me.pbDownload = New System.Windows.Forms.ProgressBar
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.lblProcess = New System.Windows.Forms.Label
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
        Me.Label1.Location = New System.Drawing.Point(8, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(272, 23)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "ECS Launcher"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(8, 88)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(280, 23)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "version 1.0"
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
        'frmECSUpdater
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.ControlBox = False
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
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Initialize the progress bar
        pbDownload.Minimum = 0
        pbDownload.Maximum = 100
        pbDownload.Value = 0

        Me.Show()
        'Download the ECS database file
        Try
            lblProcess.Text = "Copying ECS database..."
            Me.Refresh()
            If File.Exists("\\Dc1\pro_apps\claims\ECS\everyone\ecs.mdb") Then
                File.Copy("\\Dc1\pro_apps\claims\ECS\everyone\ecs.mdb", "C:\ECS\", True)
            Else
                MsgBox("The database can not be found.", MsgBoxStyle.OKOnly, "ecs.mdb is not available")
            End If

        Catch ex As Exception
            MsgBox("Unable to copy the database file" & " " & Err.Description, MsgBoxStyle.Information, "Updating ECS...")
        End Try
        'Increment the progress bar
        pbDownload.Value = 50
        'Download the ECS executable
        Try
            lblProcess.Text = "Copying ECS executable..."
            Me.Refresh()
            If File.Exists("\\Dc1\pro_apps\claims\ECS\everyone\ECSSystem.exe") Then
                File.Copy("\\Dc1\pro_apps\claims\ECS\everyone\ECSSystem.exe", "C:\ECS\ECSSystem.exe", True)
            Else
                MsgBox("The executable file can not be found.", MsgBoxStyle.OKOnly, "ECSSystem.exe is not available.")
            End If

        Catch ex As Exception
            MsgBox("Unable to copy the executable file" & " " & Err.Description, MsgBoxStyle.Information, "Updating ECS...")
        End Try
        'Increment the progress bar
        pbDownload.Value = 100
        lblProcess.Text = "Finished updating... Launching ECS."
        Me.TopMost = False
        Process.Start("C:\ECS\ECSSystem.exe")
        Me.Visible = False
        Me.Refresh()
        Me.Close()

        'MsgBox("Finished copy executable", MsgBoxStyle.Information, "Updating ECS...")
        'Read the ECS Version XML File
        'Dim doc As XmlDocument = New XmlDocument
        'doc.Load("ServerManifest.xml")
        'Create an XmlNamespaceManager for resolving namespaces.
        'Dim nsmgr As XmlNamespaceManager = New XmlNamespaceManager(doc.NameTable)
        'nsmgr.AddNamespace("bk", "urn:samples")
        'Select the book node with the matching attribute value.
        'Dim book As XmlNode
        'Dim root As XmlElement = doc.DocumentElement
        'book = root.SelectSingleNode("descendant::book[@bk:ISBN='1-861001-57-6']", nsmgr)
    End Sub
End Class
