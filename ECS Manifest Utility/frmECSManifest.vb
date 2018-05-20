Imports System.IO
Imports System.Xml

Public Class frmManifest
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
    Friend WithEvents fbdECSFolder As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents tbECSFolder As System.Windows.Forms.TextBox
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents btnCreateFile As System.Windows.Forms.Button
    Friend WithEvents pbCreating As System.Windows.Forms.ProgressBar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.fbdECSFolder = New System.Windows.Forms.FolderBrowserDialog
        Me.btnBrowse = New System.Windows.Forms.Button
        Me.tbECSFolder = New System.Windows.Forms.TextBox
        Me.btnCreateFile = New System.Windows.Forms.Button
        Me.pbCreating = New System.Windows.Forms.ProgressBar
        Me.SuspendLayout()
        '
        'fbdECSFolder
        '
        Me.fbdECSFolder.Description = "Shared ECS Files"
        '
        'btnBrowse
        '
        Me.btnBrowse.Location = New System.Drawing.Point(400, 16)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.TabIndex = 0
        Me.btnBrowse.Text = "&Browse"
        '
        'tbECSFolder
        '
        Me.tbECSFolder.Location = New System.Drawing.Point(8, 16)
        Me.tbECSFolder.Name = "tbECSFolder"
        Me.tbECSFolder.Size = New System.Drawing.Size(376, 20)
        Me.tbECSFolder.TabIndex = 1
        Me.tbECSFolder.Text = ""
        '
        'btnCreateFile
        '
        Me.btnCreateFile.Location = New System.Drawing.Point(200, 48)
        Me.btnCreateFile.Name = "btnCreateFile"
        Me.btnCreateFile.TabIndex = 2
        Me.btnCreateFile.Text = "Go"
        '
        'pbCreating
        '
        Me.pbCreating.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pbCreating.Location = New System.Drawing.Point(0, 87)
        Me.pbCreating.Name = "pbCreating"
        Me.pbCreating.Size = New System.Drawing.Size(488, 23)
        Me.pbCreating.TabIndex = 4
        '
        'frmManifest
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(488, 110)
        Me.Controls.Add(Me.pbCreating)
        Me.Controls.Add(Me.btnCreateFile)
        Me.Controls.Add(Me.tbECSFolder)
        Me.Controls.Add(Me.btnBrowse)
        Me.Name = "frmManifest"
        Me.Text = "AppUpdater Manifest Tool"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Show()
    End Sub
    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        Dim result As DialogResult = Me.fbdECSFolder.ShowDialog()
        If (result = DialogResult.OK) Then
            tbECSFolder.Text = fbdECSFolder.SelectedPath & "\"
        End If
        If (result = DialogResult.Cancel) Then
            'User canceled so Do Nothing
        End If
    End Sub

    Private Sub btnCreateFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateFile.Click
        If tbECSFolder.Text.ToString = "" Then
            MessageBox.Show("You did not select a directory for processing!", "No directory selected", MessageBoxButtons.OK, MessageBoxIcon.Hand)
        Else
            If tbECSFolder.Text.ToString.Substring(tbECSFolder.Text.Length - 1, 1) <> "\" Then
                MsgBox("Please end the directory with a '\'", MsgBoxStyle.OKOnly, "Missing trailing slash")
            Else
                Dim myFiles As String() = System.IO.Directory.GetFiles(tbECSFolder.Text.ToString, "*")
                Dim i As Integer = 0
                Dim xmlWTR As New XmlTextWriter(tbECSFolder.Text.ToString & "ServerManifest.xml", System.Text.Encoding.UTF8)
                xmlWTR.WriteStartDocument()
                xmlWTR.WriteStartElement("AppUpdater")
                Do
                    xmlWTR.WriteStartElement("file")
                    'xmlWTR.WriteStartElement("name")
                    xmlWTR.WriteAttributeString("name", myFiles(i).ToLower)
                    'xmlWTR.WriteEndAttribute()
                    'xmlWTR.WriteStartElement("date")
                    xmlWTR.WriteAttributeString("date", System.IO.File.GetLastWriteTime(myFiles(i).ToString).ToString)
                    'xmlWTR.WriteEndAttribute()
                    xmlWTR.WriteEndElement()
                    i = i + 1

                Loop Until i > (myFiles.Length - 1)
                xmlWTR.WriteEndDocument()
                xmlWTR.Flush()
                xmlWTR.Close()

                MsgBox(tbECSFolder.Text.ToString & "ServerManifest.xml has been created!", MsgBoxStyle.Information, "Done!")
            End If 'tbECSFolder.Text.ToString.Substring
        End If 'tbECSFolder.Text.ToString = ""

    End Sub
End Class
