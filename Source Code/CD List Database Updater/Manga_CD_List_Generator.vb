Imports System
Imports System.IO
Imports System.Collections

Public Class Main_Program
    Inherits System.Windows.Forms.Form

    Dim inputtedpath As String
    Dim recordcounter As Integer
    Dim selecteddrive As String
    Dim selectedpath As String
    Private Selected_Database As String
    Private Selected_Table As String

    Dim filenamelist As New System.Collections.ArrayList()

    Dim checklistset As Integer
    Dim maxchecklistset As Integer

    Public Structure MyFileItem
        Public FilenameString As String
        Public SizeString As String
        Public CheckedStatus As Boolean
    End Structure

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        selectedpath = ""
        selecteddrive = ""
        recordcounter = 0
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
    Friend WithEvents Select_Database_Dialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents CheckedListBox1 As System.Windows.Forms.CheckedListBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents display_foldernames As System.Windows.Forms.CheckBox
    Friend WithEvents Display_Path As System.Windows.Forms.CheckBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents refresh_button As System.Windows.Forms.Button
    Friend WithEvents Display_Extension As System.Windows.Forms.CheckBox
    Friend WithEvents Display_Size As System.Windows.Forms.CheckBox
    Friend WithEvents ParentPath As System.Windows.Forms.Label
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Button9 As System.Windows.Forms.Button
    Friend WithEvents Button8 As System.Windows.Forms.Button
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Columns As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Datasource As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents updateresultlabel As System.Windows.Forms.Label
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button10 As System.Windows.Forms.Button
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Button11 As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Main_Program))
        Me.Select_Database_Dialog = New System.Windows.Forms.OpenFileDialog
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Button5 = New System.Windows.Forms.Button
        Me.display_foldernames = New System.Windows.Forms.CheckBox
        Me.Display_Path = New System.Windows.Forms.CheckBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.refresh_button = New System.Windows.Forms.Button
        Me.Display_Extension = New System.Windows.Forms.CheckBox
        Me.Display_Size = New System.Windows.Forms.CheckBox
        Me.ParentPath = New System.Windows.Forms.Label
        Me.Button6 = New System.Windows.Forms.Button
        Me.Button9 = New System.Windows.Forms.Button
        Me.Button8 = New System.Windows.Forms.Button
        Me.Button7 = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.CheckedListBox1 = New System.Windows.Forms.CheckedListBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Button10 = New System.Windows.Forms.Button
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Button11 = New System.Windows.Forms.Button
        Me.Columns = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Datasource = New System.Windows.Forms.Label
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.updateresultlabel = New System.Windows.Forms.Label
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Select_Database_Dialog
        '
        Me.Select_Database_Dialog.DefaultExt = "mdb"
        Me.Select_Database_Dialog.Filter = "Microsoft Access files|*.mdb|All files|*.*"
        Me.Select_Database_Dialog.Title = "Select Database"
        '
        'Button5
        '
        Me.Button5.BackColor = System.Drawing.Color.Gainsboro
        Me.Button5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button5.Location = New System.Drawing.Point(456, 80)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(8, 8)
        Me.Button5.TabIndex = 15
        Me.Button5.TabStop = False
        Me.ToolTip1.SetToolTip(Me.Button5, "Select none")
        '
        'display_foldernames
        '
        Me.display_foldernames.BackColor = System.Drawing.Color.Transparent
        Me.display_foldernames.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.display_foldernames.Location = New System.Drawing.Point(240, 48)
        Me.display_foldernames.Name = "display_foldernames"
        Me.display_foldernames.Size = New System.Drawing.Size(136, 24)
        Me.display_foldernames.TabIndex = 4
        Me.display_foldernames.Text = "Include Folder Names"
        Me.ToolTip1.SetToolTip(Me.display_foldernames, "Include folder names in search")
        '
        'Display_Path
        '
        Me.Display_Path.BackColor = System.Drawing.Color.Transparent
        Me.Display_Path.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Display_Path.Location = New System.Drawing.Point(376, 48)
        Me.Display_Path.Name = "Display_Path"
        Me.Display_Path.Size = New System.Drawing.Size(88, 24)
        Me.Display_Path.TabIndex = 3
        Me.Display_Path.Text = "Display Path"
        Me.ToolTip1.SetToolTip(Me.Display_Path, "Display the full filepath in search results")
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Gainsboro
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(16, 32)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(96, 23)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Select Folder"
        Me.ToolTip1.SetToolTip(Me.Button1, "Allows you to select the folder that will be recursively searched")
        '
        'refresh_button
        '
        Me.refresh_button.BackColor = System.Drawing.Color.Gainsboro
        Me.refresh_button.Enabled = False
        Me.refresh_button.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.refresh_button.Location = New System.Drawing.Point(136, 32)
        Me.refresh_button.Name = "refresh_button"
        Me.refresh_button.TabIndex = 1
        Me.refresh_button.Text = "Refresh List"
        Me.ToolTip1.SetToolTip(Me.refresh_button, "Refreshes the search results")
        '
        'Display_Extension
        '
        Me.Display_Extension.BackColor = System.Drawing.Color.Transparent
        Me.Display_Extension.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Display_Extension.Location = New System.Drawing.Point(240, 24)
        Me.Display_Extension.Name = "Display_Extension"
        Me.Display_Extension.Size = New System.Drawing.Size(120, 24)
        Me.Display_Extension.TabIndex = 2
        Me.Display_Extension.Text = "Display Extension"
        Me.ToolTip1.SetToolTip(Me.Display_Extension, "Display the file extension in all search results")
        '
        'Display_Size
        '
        Me.Display_Size.BackColor = System.Drawing.Color.Transparent
        Me.Display_Size.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Display_Size.Location = New System.Drawing.Point(376, 24)
        Me.Display_Size.Name = "Display_Size"
        Me.Display_Size.Size = New System.Drawing.Size(112, 24)
        Me.Display_Size.TabIndex = 5
        Me.Display_Size.Text = "Display File Size"
        Me.ToolTip1.SetToolTip(Me.Display_Size, "Dislpay the file size of each file in search")
        '
        'ParentPath
        '
        Me.ParentPath.BackColor = System.Drawing.Color.Transparent
        Me.ParentPath.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ParentPath.ForeColor = System.Drawing.Color.Firebrick
        Me.ParentPath.Location = New System.Drawing.Point(16, 304)
        Me.ParentPath.Name = "ParentPath"
        Me.ParentPath.Size = New System.Drawing.Size(584, 24)
        Me.ParentPath.TabIndex = 8
        Me.ParentPath.Text = "Parent Path:"
        Me.ToolTip1.SetToolTip(Me.ParentPath, "The path selected on which to search")
        '
        'Button6
        '
        Me.Button6.BackColor = System.Drawing.Color.Gainsboro
        Me.Button6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button6.Location = New System.Drawing.Point(456, 96)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(8, 8)
        Me.Button6.TabIndex = 16
        Me.Button6.TabStop = False
        Me.ToolTip1.SetToolTip(Me.Button6, "Select all")
        '
        'Button9
        '
        Me.Button9.BackColor = System.Drawing.Color.Gainsboro
        Me.Button9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button9.Location = New System.Drawing.Point(456, 176)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(8, 8)
        Me.Button9.TabIndex = 19
        Me.Button9.TabStop = False
        Me.ToolTip1.SetToolTip(Me.Button9, "Show Next")
        '
        'Button8
        '
        Me.Button8.BackColor = System.Drawing.Color.Gainsboro
        Me.Button8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button8.Location = New System.Drawing.Point(456, 160)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(8, 8)
        Me.Button8.TabIndex = 18
        Me.Button8.TabStop = False
        Me.ToolTip1.SetToolTip(Me.Button8, "Show All")
        '
        'Button7
        '
        Me.Button7.BackColor = System.Drawing.Color.Gainsboro
        Me.Button7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button7.Location = New System.Drawing.Point(456, 144)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(8, 8)
        Me.Button7.TabIndex = 17
        Me.Button7.TabStop = False
        Me.ToolTip1.SetToolTip(Me.Button7, "Show Previous")
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Firebrick
        Me.Label1.Location = New System.Drawing.Point(16, 288)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(296, 16)
        Me.Label1.TabIndex = 20
        Me.Label1.Text = "0 records displayed"
        Me.ToolTip1.SetToolTip(Me.Label1, "The number of results resulting from the search")
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.Gainsboro
        Me.Button2.Enabled = False
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(16, 32)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(112, 23)
        Me.Button2.TabIndex = 7
        Me.Button2.Text = "Select Database"
        Me.ToolTip1.SetToolTip(Me.Button2, "Allows you to select a database to which the results will be sent to")
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.Gainsboro
        Me.Button4.Location = New System.Drawing.Point(456, 64)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(8, 8)
        Me.Button4.TabIndex = 9
        Me.Button4.TabStop = False
        Me.ToolTip1.SetToolTip(Me.Button4, "Clear Textbox")
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.Gainsboro
        Me.Button3.Enabled = False
        Me.Button3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.Location = New System.Drawing.Point(16, 32)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(112, 23)
        Me.Button3.TabIndex = 8
        Me.Button3.Text = "Update Database"
        Me.ToolTip1.SetToolTip(Me.Button3, "Allows you to update the selected database")
        '
        'TabControl1
        '
        Me.TabControl1.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Top
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.ShowToolTips = True
        Me.TabControl1.Size = New System.Drawing.Size(744, 384)
        Me.TabControl1.TabIndex = 15
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.White
        Me.TabPage1.BackgroundImage = CType(resources.GetObject("TabPage1.BackgroundImage"), System.Drawing.Image)
        Me.TabPage1.Controls.Add(Me.CheckedListBox1)
        Me.TabPage1.Controls.Add(Me.GroupBox1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(736, 355)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Step 1"
        '
        'CheckedListBox1
        '
        Me.CheckedListBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.CheckedListBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckedListBox1.HorizontalScrollbar = True
        Me.CheckedListBox1.Location = New System.Drawing.Point(32, 88)
        Me.CheckedListBox1.Name = "CheckedListBox1"
        Me.CheckedListBox1.Size = New System.Drawing.Size(432, 195)
        Me.CheckedListBox1.Sorted = True
        Me.CheckedListBox1.TabIndex = 7
        Me.CheckedListBox1.TabStop = False
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.Button10)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Button5)
        Me.GroupBox1.Controls.Add(Me.display_foldernames)
        Me.GroupBox1.Controls.Add(Me.Display_Path)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.refresh_button)
        Me.GroupBox1.Controls.Add(Me.Display_Extension)
        Me.GroupBox1.Controls.Add(Me.Display_Size)
        Me.GroupBox1.Controls.Add(Me.ParentPath)
        Me.GroupBox1.Controls.Add(Me.Button6)
        Me.GroupBox1.Controls.Add(Me.Button9)
        Me.GroupBox1.Controls.Add(Me.Button8)
        Me.GroupBox1.Controls.Add(Me.Button7)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(16, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(704, 336)
        Me.GroupBox1.TabIndex = 13
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Step 1 - Select Files"
        '
        'Button10
        '
        Me.Button10.BackColor = System.Drawing.Color.Transparent
        Me.Button10.Enabled = False
        Me.Button10.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button10.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button10.Location = New System.Drawing.Point(616, 304)
        Me.Button10.Name = "Button10"
        Me.Button10.TabIndex = 21
        Me.Button10.Text = "Next >>"
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.Color.White
        Me.TabPage2.BackgroundImage = CType(resources.GetObject("TabPage2.BackgroundImage"), System.Drawing.Image)
        Me.TabPage2.Controls.Add(Me.GroupBox2)
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(736, 355)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Step 2"
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.Button11)
        Me.GroupBox2.Controls.Add(Me.Columns)
        Me.GroupBox2.Controls.Add(Me.Panel1)
        Me.GroupBox2.Controls.Add(Me.Button2)
        Me.GroupBox2.Controls.Add(Me.Datasource)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(16, 8)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(704, 336)
        Me.GroupBox2.TabIndex = 14
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Step 2 - Select a Data Source"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Firebrick
        Me.Label2.Location = New System.Drawing.Point(160, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(328, 32)
        Me.Label2.TabIndex = 23
        Me.Label2.Text = "Now set the required values for the record fields which will be used in the SQL i" & _
        "nsert statement."
        Me.Label2.Visible = False
        '
        'Button11
        '
        Me.Button11.BackColor = System.Drawing.Color.Transparent
        Me.Button11.Enabled = False
        Me.Button11.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button11.Location = New System.Drawing.Point(616, 304)
        Me.Button11.Name = "Button11"
        Me.Button11.TabIndex = 22
        Me.Button11.Text = "Next >>"
        '
        'Columns
        '
        Me.Columns.Location = New System.Drawing.Point(584, 168)
        Me.Columns.Name = "Columns"
        Me.Columns.Size = New System.Drawing.Size(88, 16)
        Me.Columns.TabIndex = 15
        Me.Columns.Text = "0 Columns"
        Me.Columns.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Columns.Visible = False
        '
        'Panel1
        '
        Me.Panel1.AutoScroll = True
        Me.Panel1.BackColor = System.Drawing.Color.Transparent
        Me.Panel1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panel1.Location = New System.Drawing.Point(16, 72)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(448, 216)
        Me.Panel1.TabIndex = 9
        '
        'Datasource
        '
        Me.Datasource.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Datasource.ForeColor = System.Drawing.Color.Firebrick
        Me.Datasource.Location = New System.Drawing.Point(16, 296)
        Me.Datasource.Name = "Datasource"
        Me.Datasource.Size = New System.Drawing.Size(568, 24)
        Me.Datasource.TabIndex = 14
        '
        'TabPage3
        '
        Me.TabPage3.BackColor = System.Drawing.Color.White
        Me.TabPage3.BackgroundImage = CType(resources.GetObject("TabPage3.BackgroundImage"), System.Drawing.Image)
        Me.TabPage3.Controls.Add(Me.GroupBox3)
        Me.TabPage3.Location = New System.Drawing.Point(4, 25)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(736, 355)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Step 3"
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox3.Controls.Add(Me.LinkLabel1)
        Me.GroupBox3.Controls.Add(Me.updateresultlabel)
        Me.GroupBox3.Controls.Add(Me.Button4)
        Me.GroupBox3.Controls.Add(Me.RichTextBox1)
        Me.GroupBox3.Controls.Add(Me.Button3)
        Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(16, 8)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(704, 336)
        Me.GroupBox3.TabIndex = 15
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Step 3 - Update the Database"
        '
        'LinkLabel1
        '
        Me.LinkLabel1.ActiveLinkColor = System.Drawing.Color.Firebrick
        Me.LinkLabel1.ForeColor = System.Drawing.Color.Firebrick
        Me.LinkLabel1.LinkColor = System.Drawing.Color.Firebrick
        Me.LinkLabel1.Location = New System.Drawing.Point(376, 40)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(80, 16)
        Me.LinkLabel1.TabIndex = 13
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "View Report"
        Me.LinkLabel1.Visible = False
        Me.LinkLabel1.VisitedLinkColor = System.Drawing.Color.Firebrick
        '
        'updateresultlabel
        '
        Me.updateresultlabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.updateresultlabel.ForeColor = System.Drawing.Color.Firebrick
        Me.updateresultlabel.Location = New System.Drawing.Point(16, 296)
        Me.updateresultlabel.Name = "updateresultlabel"
        Me.updateresultlabel.Size = New System.Drawing.Size(664, 24)
        Me.updateresultlabel.TabIndex = 12
        '
        'RichTextBox1
        '
        Me.RichTextBox1.BackColor = System.Drawing.Color.White
        Me.RichTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RichTextBox1.Cursor = System.Windows.Forms.Cursors.Default
        Me.RichTextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RichTextBox1.Location = New System.Drawing.Point(16, 64)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ReadOnly = True
        Me.RichTextBox1.Size = New System.Drawing.Size(432, 224)
        Me.RichTextBox1.TabIndex = 11
        Me.RichTextBox1.TabStop = False
        Me.RichTextBox1.Text = ""
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem2, Me.MenuItem3})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Text = "Help"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.Text = "About"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 2
        Me.MenuItem3.Text = "Exit"
        '
        'FolderBrowserDialog1
        '
        Me.FolderBrowserDialog1.Description = "Select a root folder to be parsed. If the folder selected contains any subfolders" & _
        ", these folders will be parsed as well."
        Me.FolderBrowserDialog1.ShowNewFolderButton = False
        '
        'Main_Program
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(744, 406)
        Me.Controls.Add(Me.TabControl1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(752, 440)
        Me.Menu = Me.MainMenu1
        Me.MinimumSize = New System.Drawing.Size(752, 440)
        Me.Name = "Main_Program"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CD List Database Updater"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Sub Error_Handler(ByVal message As String)
        MsgBox("The following error has occurred: " & vbCrLf & message, MsgBoxStyle.Exclamation, "CD List Database Updater")
    End Sub


    Public Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Dim result As DialogResult = FolderBrowserDialog1.ShowDialog()
            If result = DialogResult.OK Then
                CheckedListBox1.Items.Clear()
                CheckedListBox1.Update()
                filenamelist.Clear()
                'sizelist.Clear()
                ProcessPath(FolderBrowserDialog1.SelectedPath)
                maxchecklistset = filenamelist.Count() / 200 + 1
                checklistset = 1
                displaychecklistset(checklistset)
                inputtedpath = FolderBrowserDialog1.SelectedPath
                selectedpath = FolderBrowserDialog1.SelectedPath
                selecteddrive = FolderBrowserDialog1.SelectedPath
                'CheckedListBox1.Update()
                refresh_button.Enabled = True
                Button2.Enabled = True
                Button10.BackColor = Color.IndianRed
                Button10.Enabled = True
            End If
            FolderBrowserDialog1.Dispose()



        Catch et As Exception
            Error_Handler(et.Message)
        End Try
        

    End Sub


    Public Sub ProcessPath(ByVal inputpath As String)
        Dim path As String
        ParentPath.Text = "Parent Path: "
        recordcounter = 0
        path = inputpath
        If File.Exists(path) Then
            ' This path is a file
            ProcessFile(path)
        Else
            If Directory.Exists(path) Then
                ' This path is a directory
                ProcessDirectory(path)
            Else
                MsgBox(path & " is not a valid file or directory.")
            End If
        End If
        Label1.Text = recordcounter & " records displayed"
        ParentPath.Text = ParentPath.Text & path
    End Sub 'Main


    ' Process all files in the directory passed in, and recurse on any directories 
    ' that are found to process the files they contain
    Public Sub ProcessDirectory(ByVal targetDirectory As String)
        Dim fileEntries As String() = Directory.GetFiles(targetDirectory)
        ' Process the list of files found in the directory
        Dim fileName As String
        Dim tempfilename As System.IO.DirectoryInfo
        Dim tempextension As Integer
        Dim filesize As Decimal
        Dim filesize_string As String

        If display_foldernames.Checked = True Then
            ProcessFileD(targetDirectory)
        End If

        For Each fileName In fileEntries
            tempfilename = Directory.GetParent(fileName)
            filesize = FileLen(fileName)
            If Display_Path.Checked = False Then
                fileName = Replace(fileName, tempfilename.FullName, "")
                fileName = Replace(fileName, "\", "")
            End If
            If Display_Extension.Checked = False Then
                tempextension = 0
                tempextension = fileName.LastIndexOf(".")
                If tempextension > 0 Then
                    fileName = fileName.Remove(tempextension, (fileName.Length - tempextension))
                End If
            End If
            If Display_Size.Checked = True Then

                If filesize < 1024 Then
                    filesize_string = filesize & " bytes"
                End If
                If filesize < 1048576 And filesize >= 1024 Then
                    filesize = Math.Round(filesize / 1024, 2)
                    filesize_string = filesize & " KB"
                End If
                If filesize < 1073741824 And filesize >= 1048576 Then
                    filesize = Math.Round(filesize / 1048576, 2)
                    filesize_string = filesize & " MB"
                End If
                If filesize >= 1073741824 Then
                    filesize = Math.Round(filesize / 1073741824, 2)
                    filesize_string = filesize & " GB"
                End If

            End If
            If Display_Size.Checked = True Then
                ProcessFile(fileName, filesize_string)
            Else
                ProcessFile(fileName)
            End If
        Next fileName
        Dim subdirectoryEntries As String() = Directory.GetDirectories(targetDirectory)
        ' Recurse into subdirectories of this directory
        Dim subdirectory As String
        For Each subdirectory In subdirectoryEntries
            ProcessDirectory(subdirectory)
        Next subdirectory

    End Sub 'ProcessDirectory

    ' Real logic for processing found files would go here.
    Public Sub ProcessFile(ByVal path As String)
        recordcounter = recordcounter + 1
        'CheckedListBox1.Items.Add(path, True)
        Dim itemtoadd As MyFileItem
        itemtoadd.FilenameString = path
        itemtoadd.CheckedStatus = True
        itemtoadd.SizeString = ""
        'filenamelist.Add(itemtoadd)
        SortedArraylistInsert(itemtoadd)
    End Sub 'ProcessFile

    Public Sub ProcessFileD(ByVal path As String)
        recordcounter = recordcounter + 1
        'CheckedListBox1.Items.Add(path, True)
        Dim itemtoadd As MyFileItem
        itemtoadd.FilenameString = path
        itemtoadd.CheckedStatus = True
        If Display_Size.Checked = True Then
            Dim foldersize As Long = GetFolderSize(path)
            If foldersize < 1024 Then
                itemtoadd.SizeString = foldersize & " bytes"
            End If
            If foldersize < 1048576 And foldersize >= 1024 Then
                foldersize = Math.Round(foldersize / 1024, 2)
                itemtoadd.SizeString = foldersize & " KB"
            End If
            If foldersize < 1073741824 And foldersize >= 1048576 Then
                foldersize = Math.Round(foldersize / 1048576, 2)
                itemtoadd.SizeString = foldersize & " MB"
            End If
            If foldersize >= 1073741824 Then
                foldersize = Math.Round(foldersize / 1073741824, 2)
                itemtoadd.SizeString = foldersize & " GB"
            End If

        Else
            itemtoadd.SizeString = ""
        End If
        'filenamelist.Add(itemtoadd)
        SortedArraylistInsert(itemtoadd)
    End Sub 'ProcessFile

    Public Sub ProcessFile(ByVal path As String, ByVal sizestr As String)
        recordcounter = recordcounter + 1
        'CheckedListBox1.Items.Add(path & "    [" & sizestr & "]", True)
        Dim itemtoadd As MyFileItem
        itemtoadd.FilenameString = path
        itemtoadd.CheckedStatus = True
        itemtoadd.SizeString = sizestr
        'filenamelist.Add(itemtoadd)
        SortedArraylistInsert(itemtoadd)
    End Sub 'ProcessFile

    Public Function displaychecklistset(ByVal setnumber As Integer)
        CheckedListBox1.Items.Clear()
        Dim startcount As Integer
        Dim endcount As Integer
        If Not setnumber = 0 Then
            startcount = (setnumber - 1) * 200
            If filenamelist.Count() - 1 >= startcount + 200 Then
                endcount = startcount + 200
            Else
                endcount = filenamelist.Count()
            End If

            Dim runner As Integer
            Dim itemtoadd As MyFileItem
            For runner = startcount To endcount - 1
                itemtoadd = filenamelist.Item(runner)
                If Display_Size.Checked = True Then
                    CheckedListBox1.Items.Add(itemtoadd.FilenameString & "    [" & itemtoadd.SizeString & "]", itemtoadd.CheckedStatus)
                Else
                    CheckedListBox1.Items.Add(itemtoadd.FilenameString, itemtoadd.CheckedStatus)
                End If
            Next
            CheckedListBox1.Update()
            Label1.Text = (startcount + 1) & " to " & (endcount) & " of " & filenamelist.Count() & " records displayed"
        Else
            startcount = 0
            endcount = filenamelist.Count() - 1


            Dim runner As Integer
            Dim itemtoadd As MyFileItem
            For runner = startcount To endcount
                itemtoadd = filenamelist.Item(runner)
                If Display_Size.Checked = True Then
                    CheckedListBox1.Items.Add(itemtoadd.FilenameString & "    [" & itemtoadd.SizeString & "]", itemtoadd.CheckedStatus)
                Else
                    CheckedListBox1.Items.Add(itemtoadd.FilenameString, itemtoadd.CheckedStatus)
                End If
            Next
            CheckedListBox1.Update()
            Label1.Text = (startcount + 1) & " to " & (endcount + 1) & " of " & filenamelist.Count() & " records displayed"
        End If

    End Function


    Private Sub refresh_button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles refresh_button.Click
        filenamelist.Clear()

        If Not inputtedpath.Equals("") Then
            CheckedListBox1.Items.Clear()
            CheckedListBox1.Update()
            ProcessPath(inputtedpath)
            'CheckedListBox1.Update()
            maxchecklistset = filenamelist.Count() / 200 + 1
            checklistset = 1
            displaychecklistset(checklistset)
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            DialogResult = Select_Database_Dialog.ShowDialog()
            If DialogResult = DialogResult.OK Or DialogResult = DialogResult.Yes Then

                Try

                    Selected_Database = Select_Database_Dialog.FileName


                    Dim Conn As Data.OleDb.OleDbConnection = New Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Selected_Database & ";")
                    Conn.Open()
                    Dim schemaTable As DataTable = Conn.GetOleDbSchemaTable(Data.OleDb.OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})

                    Dim frm_Select_Table_Dialog As New Select_Table_Dialog
                    frm_Select_Table_Dialog.Activate()
                    frm_Select_Table_Dialog.TableChoice = schemaTable
                    frm_Select_Table_Dialog.ShowDialog()
                    Me.Show()
                    Dim tableresult As String = frm_Select_Table_Dialog.Selected_Table.SelectedItem.ToString
                    Label2.Visible = True
                    Selected_Table = tableresult
                    frm_Select_Table_Dialog.Dispose()
                    Dim schemaTable2 As DataTable = Conn.GetOleDbSchemaTable(Data.OleDb.OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, tableresult, Nothing})

                    Dim myRow2 As DataRow
                    Dim myCol2 As DataColumn

                    Panel1.Controls.Clear()

                    RichTextBox1.Clear()

                    Dim ordinal As Integer
                    Dim columnname As String
                    Dim datatype As OleDb.OleDbType
                    For Each myRow2 In schemaTable2.Rows
                        ordinal = 0
                        columnname = ""


                        For Each myCol2 In schemaTable2.Columns
                            If myCol2.ColumnName = "DATA_TYPE" Then
                                datatype = myRow2(myCol2)
                            End If
                            If myCol2.ColumnName = "COLUMN_NAME" Then
                                columnname = myRow2(myCol2).ToString()
                            End If
                            If myCol2.ColumnName = "ORDINAL_POSITION" Then
                                ordinal = Val(myRow2(myCol2).ToString())
                            End If
                        Next
                        ordinal = ordinal - 1
                        If Not columnname.Equals("") Then
                            'MsgBox(myRow2(myCol2).ToString())
                            Dim LabelMiniMe As New System.Windows.Forms.Label
                            LabelMiniMe.Location = New System.Drawing.Point(0, (ordinal * 24))
                            LabelMiniMe.Name = "Label_" & columnname
                            LabelMiniMe.Size = New System.Drawing.Size(136, 16)
                            LabelMiniMe.Text = columnname
                            Panel1.Controls.Add(LabelMiniMe)
                            LabelMiniMe.Visible = True

                            Dim ComboBoxMiniMe As New System.Windows.Forms.ComboBox
                            ComboBoxMiniMe.Location = New System.Drawing.Point(140, (ordinal * 24))
                            ComboBoxMiniMe.Name = "ComboBox_" & columnname
                            ComboBoxMiniMe.Size = New System.Drawing.Size(136, 16)
                            If datatype.ToString().ToLower = "wchar" Then
                                ComboBoxMiniMe.Text = "'" & columnname & "'"
                                ComboBoxMiniMe.Items.Add("Selected Item String")
                                If Display_Size.Checked = True Then
                                    ComboBoxMiniMe.Items.Add("File Size String")
                                End If

                            Else
                                ComboBoxMiniMe.Text = columnname

                            End If
                            ComboBoxMiniMe.Items.Add("Ignore Column")
                            Panel1.Controls.Add(ComboBoxMiniMe)
                            ComboBoxMiniMe.Visible = True
                            If datatype.ToString().ToLower = "boolean" Then
                                ComboBoxMiniMe.Items.Add(1)
                                ComboBoxMiniMe.Items.Add(0)
                                ComboBoxMiniMe.SelectedIndex = 0
                                ComboBoxMiniMe.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
                            End If

                            Dim LabelMiniM2 As New System.Windows.Forms.Label
                            LabelMiniM2.Location = New System.Drawing.Point(280, (ordinal * 24))
                            LabelMiniM2.Name = "Label_TYPE_" & columnname
                            LabelMiniM2.Size = New System.Drawing.Size(50, 16)
                            LabelMiniM2.Text = datatype.ToString()
                            Panel1.Controls.Add(LabelMiniM2)
                            LabelMiniM2.Visible = True

                        End If
                    Next
                    Conn.Close()
                    Button3.Enabled = True
                    Button11.BackColor = Color.IndianRed
                    Button11.Enabled = True
                    Datasource.Text = "Data Source: " & Selected_Database & vbCrLf & "Table: " & Selected_Table
                    Columns.Text = (Panel1.Controls.Count / 3) & " Columns"
                Catch dberror As OleDb.OleDbException
                    Error_Handler(dberror.Message)

                Catch othererror As Exception
                    Error_Handler(othererror.Message)
                End Try
            End If
        Catch err As Exception
            Error_Handler(err.Message)
        End Try
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Try
            Button3.Text = "Updating..."
            Button3.Enabled = False
            Dim Conn As Data.OleDb.OleDbConnection = New Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Selected_Database & ";")
            Conn.Open()
            Dim counter As Integer
            Dim itemtoadd As MyFileItem
            Dim stoppedflag As Boolean = False
            RichTextBox1.Text = "[- " & Now() & " -]" & vbCrLf
            Dim success_counter, sql_counter As Long
            success_counter = 0
            sql_counter = 0
            For counter = 0 To filenamelist.Count - 1
                itemtoadd = filenamelist.Item(counter)
                If itemtoadd.CheckedStatus = True Then
                    Try
                        Dim sqlstr As String
                        sql_counter = sql_counter + 1
                        sqlstr = "insert into [" & Selected_Table & "] ("
                        Dim runne As Integer
                        For runne = 0 To Panel1.Controls.Count - 1 Step 3
                            If Not Panel1.Controls(runne + 1).Text = "Ignore Column" Then
                                sqlstr = sqlstr & "[" & Panel1.Controls(runne).Text & "],"
                            End If
                        Next
                        sqlstr = sqlstr.Remove(sqlstr.Length - 1, 1)

                        sqlstr = sqlstr & ") values ("
                        Dim runner As Integer
                        For runner = 1 To Panel1.Controls.Count - 1 Step 3
                            If Panel1.Controls(runner).Text = "Selected Item String" Then
                                'Dim itemtoadd As MyFileItem
                                'itemtoadd = filenamelist.Item(counter)
                                If itemtoadd.CheckedStatus = True Then
                                    sqlstr = sqlstr & "'" & itemtoadd.FilenameString.Replace("'", "`") & "',"
                                End If
                            Else
                                If Panel1.Controls(runner).Text = "File Size String" Then
                                    'Dim itemtoadd As MyFileItem
                                    'itemtoadd = filenamelist.Item(counter)
                                    If itemtoadd.CheckedStatus = True Then
                                        sqlstr = sqlstr & "'" & itemtoadd.SizeString & "',"
                                    End If

                                Else
                                    If Not Panel1.Controls(runner).Text = "Ignore Column" Then
                                        sqlstr = sqlstr & Panel1.Controls(runner).Text & ","
                                    End If
                                End If
                            End If
                        Next
                        sqlstr = sqlstr.Remove(sqlstr.Length - 1, 1)
                        sqlstr = sqlstr & ")"
                        RichTextBox1.Text = RichTextBox1.Text & vbCrLf & sqlstr


                        Dim recset As Data.OleDb.OleDbCommand = New Data.OleDb.OleDbCommand(sqlstr, Conn)
                        recset.ExecuteNonQuery()
                        
                        RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "- SUCCESS - " & TimeOfDay() & vbCrLf & "--------------"
                        success_counter = success_counter + 1
                    Catch sqlerror As Exception
                        Error_Handler(sqlerror.Message)
                        RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "- FAILED - " & TimeOfDay() & vbCrLf & "--------------"
                        Dim answer As Microsoft.VisualBasic.MsgBoxResult = MsgBox("Do you wish to continue processing the remaining updates?", MsgBoxStyle.OKCancel, "CD List Database Updater")
                        If answer = MsgBoxResult.Abort Or answer = MsgBoxResult.Cancel Then
                            ' MsgBox("Stopping further updates to the Database")
                            updateresultlabel.Text = "Stopping further updates to the Database. " & sql_counter & " attempted updates were made. " & success_counter & " were successful."
                            stoppedflag = True
                            Exit For
                        End If
                    End Try
                End If
            Next
          

            Button3.Text = "Update Database"
            Button3.Enabled = True
            Conn.Close()
            If stoppedflag = False Then
                'MsgBox("Database Successfully Updated")
                updateresultlabel.Text = "Database Successfully Updated. " & sql_counter & " attempted updates were made. " & success_counter & " were successful."
            End If
            If Write_HTML_Report() = True Then
                LinkLabel1.Visible = True
            End If
        Catch dberror As OleDb.OleDbException
            'MsgBox("Cannot Connect to the Datasource Specified" & vbCrLf & dberror.Message)
            updateresultlabel.Text = "Cannot Connect to the Datasource Specified" & vbCrLf & dberror.Message
        Catch othererror As Exception
            'MsgBox("Error encountered" & vbCrLf & othererror.Message)
            updateresultlabel.Text = "Error encountered" & vbCrLf & othererror.Message
        End Try

    End Sub

    Private Sub checkedlistbox1_doubleclick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckedListBox1.DoubleClick
        CheckedListBox1_SelectedIndexChanged(sender, e)
    End Sub


    Private Sub checkedlistbox1_keydown(ByVal sender As System.Object, ByVal keyselected As System.Windows.Forms.KeyEventArgs) Handles CheckedListBox1.KeyDown

        Dim selnum As Integer = CheckedListBox1.SelectedIndex()
        If selnum >= 0 Then
            If keyselected.KeyCode = Keys.F2 And keyselected.KeyData = Keys.F2 Then

                Dim frm_Change_FileString_Dialog As New Change_FileString_Dialog
                frm_Change_FileString_Dialog.Activate()
                Dim runner As Integer
                Dim stringtotest As String
                Dim resulted As Integer
                Dim itemtoadd As MyFileItem

                CheckedListBox1.Sorted = False

                For runner = 0 To filenamelist.Count() - 1
                    stringtotest = CheckedListBox1.SelectedItem()
                    If Display_Size.Checked Then
                        resulted = stringtotest.LastIndexOf("[")
                        stringtotest = stringtotest.Remove(resulted, stringtotest.Length - resulted)
                    End If
                    stringtotest = stringtotest.TrimEnd


                    itemtoadd = filenamelist.Item(runner)
                    If itemtoadd.FilenameString = stringtotest Then
                        frm_Change_FileString_Dialog.oldstring.Text = itemtoadd.FilenameString
                        frm_Change_FileString_Dialog.newstring.Text = itemtoadd.FilenameString
                        Exit For
                    End If

                Next


                Try
                    frm_Change_FileString_Dialog.ShowDialog()
                Catch et As Exception
                    Error_Handler(et.Message)
                End Try

                If Not frm_Change_FileString_Dialog.newstring.Text = "" Then
                    itemtoadd.FilenameString = frm_Change_FileString_Dialog.newstring.Text
                    filenamelist.Item(runner) = itemtoadd

                    CheckedListBox1.Items.RemoveAt(selnum)
                    'CheckedListBox1.Items.Add(itemtoadd.FilenameString, True)
                    'CheckedListBox1.SelectedIndex = CheckedListBox1.Items.Count - 1
                    'Dim checktest As Boolean = CheckedListBox1.GetItemChecked(selnum)
                    If Display_Size.Checked = True Then
                        CheckedListBox1.Items.Insert(selnum, itemtoadd.FilenameString & "    [" & itemtoadd.SizeString & "]")
                    Else
                        CheckedListBox1.Items.Insert(selnum, itemtoadd.FilenameString)
                    End If

                    If itemtoadd.CheckedStatus = True Then
                        CheckedListBox1.SetItemChecked(selnum, True)
                    Else
                        CheckedListBox1.SetItemChecked(selnum, False)
                    End If

                    'CheckedListBox1.Update()
                    'CheckedListBox1.SelectedValue = itemtoadd.FilenameString
                    ' CheckedListBox1.Items.Item(selnum) = frm_Change_FileString_Dialog.newstring.Text
                    ' MsgBox(CheckedListBox1.Items.Item(selnum))
                    'CheckedListBox1.SelectedIndex = selnum
                    'CheckedListBox1.Focus()
                End If
                frm_Change_FileString_Dialog.Dispose()

            End If
        End If

        keyselected.Handled = True
    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckedListBox1.SelectedIndexChanged
        Dim selnum As Integer = CheckedListBox1.SelectedIndex()
        Dim runner As Integer
        Dim stringtotest As String
        Dim resulted As Integer
        If selnum >= 0 Then
            For runner = 0 To filenamelist.Count() - 1
                stringtotest = CheckedListBox1.SelectedItem()
                If Display_Size.Checked Then
                    resulted = stringtotest.LastIndexOf("[")
                    stringtotest = stringtotest.Remove(resulted, stringtotest.Length - resulted)
                End If
                stringtotest = stringtotest.TrimEnd
                If CheckedListBox1.GetItemChecked(selnum) = False Then
                    Dim itemtoadd As MyFileItem
                    itemtoadd = filenamelist.Item(runner)
                    If itemtoadd.FilenameString = stringtotest Then
                        'MsgBox("Removed: " & stringtotest)
                        itemtoadd.CheckedStatus = False
                        filenamelist.Item(runner) = itemtoadd
                        'filenamelist.RemoveAt(runner)
                        'If sizelist.Count > 0 Then
                        'sizelist.RemoveAt(runner)
                        'End If
                        Exit For

                    End If
                Else
                    Dim itemtoadd As MyFileItem
                    itemtoadd = filenamelist.Item(runner)
                    If itemtoadd.FilenameString = stringtotest Then
                        ' MsgBox("Added: " & stringtotest)
                        itemtoadd.CheckedStatus = True
                        filenamelist.Item(runner) = itemtoadd
                        'filenamelist.RemoveAt(runner)
                        'If sizelist.Count > 0 Then
                        'sizelist.RemoveAt(runner)
                        'End If
                        Exit For

                    End If
                End If

            Next
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        RichTextBox1.Clear()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim runner As Integer
        For runner = 0 To filenamelist.Count() - 1
            Dim itemtoadd As MyFileItem
            itemtoadd = filenamelist.Item(runner)
            itemtoadd.CheckedStatus = False
            filenamelist.Item(runner) = itemtoadd
        Next
        For runner = 0 To CheckedListBox1.Items.Count() - 1
            CheckedListBox1.SetItemChecked(runner, False)
        Next
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim runner As Integer
        For runner = 0 To filenamelist.Count() - 1
            Dim itemtoadd As MyFileItem
            itemtoadd = filenamelist.Item(runner)
            itemtoadd.CheckedStatus = True
            filenamelist.Item(runner) = itemtoadd
        Next
        For runner = 0 To CheckedListBox1.Items.Count() - 1
            CheckedListBox1.SetItemChecked(runner, True)
        Next

        '************************************************************************************
        'ArraylistSort(filenamelist)
        '***************************************************************************************
    End Sub

    'Public Function ArraylistSort(ByVal inputarraylist As System.Collections.ArrayList)
    '   outputarraylist.Clear()
    '  Dim runner As Integer
    ' Dim itemtoadd As MyFileItem
    ''MsgBox("here")
    'For runner = 0 To inputarraylist.Count() - 1

    '   itemtoadd = inputarraylist.Item(runner)
    '  ArraylistSortSub(1, outputarraylist.Count(), itemtoadd)
    'Next
    'inputarraylist = outputarraylist
    'For runner = 0 To inputarraylist.Count() - 1
    '   itemtoadd = outputarraylist.Item(runner)
    '  RichTextBox1.Text = RichTextBox1.Text & itemtoadd.FilenameString & vbCrLf
    'Next
    'End Function

    'Protected Function ArraylistSortSub(ByVal segstart As Integer, ByVal segend As Integer, ByVal itemtoinput As MyFileItem)
    'MsgBox("now")
    'RichTextBox1.Text = RichTextBox1.Text & segstart & "  " & segend & vbCrLf
    '   If filenamelist.Count() = 0 Then
    '      filenamelist.Add(itemtoinput)
    ' Else
    '   Dim itemtocheckagainst As New MyFileItem()
    '    If segstart = segend Then
    '      itemtocheckagainst = filenamelist.Item(segstart - 1)
    '     Dim result As Integer = String.Compare(itemtoinput.FilenameString, itemtocheckagainst.FilenameString)
    '    If result = 0 Or result < 0 Then
    '       filenamelist.Insert(segstart - 1, itemtoinput)
    '  End If
    ' If result > 0 Then
    '    filenamelist.Insert(segstart, itemtoinput)
    'End If
    'Else
    'Dim midpoint As Integer
    'midpoint = (segend - segstart + 1) / 2 + 1
    'itemtocheckagainst = filenamelist.Item(midpoint - 1)
    'Dim result2 As Integer = String.Compare(itemtoinput.FilenameString, itemtocheckagainst.FilenameString)
    'If result2 = 0 Or result2 < 0 Then
    'ArraylistSortSub(segstart, midpoint - 1, itemtoinput)
    'End If
    'If result2 > 0 Then
    'ArraylistSortSub(midpoint + 1, segend, itemtoinput)
    'End If

    'End If


    'End If
    'End Function

    Protected Function SortedArraylistInsert(ByVal itemtoinput As MyFileItem)
        Dim insertflag As Boolean = False
        Dim runner = New Integer
        Dim itemtocheckagainst As MyFileItem

        If filenamelist.Count() = 0 Then
            filenamelist.Add(itemtoinput)
            'MsgBox(itemtoinput.FilenameString)
            insertflag = True
        End If
        If insertflag = False Then
            If filenamelist.Count() = 1 Then
                itemtocheckagainst = filenamelist.Item(0)
                If String.Compare(itemtoinput.FilenameString, itemtocheckagainst.FilenameString) <= 0 Then
                    filenamelist.Insert(0, itemtoinput)
                    'MsgBox("insert" & itemtoinput.FilenameString)
                Else
                    filenamelist.Add(itemtoinput)
                    'MsgBox("add" & itemtoinput.FilenameString)
                End If
                insertflag = True
            End If
        End If

        '        If insertflag = False Then
        '       If filenamelist.Count() > 1 And filenamelist.Count() < 4 Then
        '          Dim check1 As Boolean
        '         Dim check2 As Boolean
        '        For runner = 0 To filenamelist.Count() - 2
        '           check1 = False
        '          check2 = False
        '         itemtocheckagainst = filenamelist.Item(runner)
        '        If String.Compare(itemtoinput.FilenameString, itemtocheckagainst.FilenameString) >= 0 Then
        '           check1 = True
        '      End If
        '     itemtocheckagainst = filenamelist.Item(runner + 1)
        'If String.Compare(itemtoinput.FilenameString, itemtocheckagainst.FilenameString) <= 0 Then
        'check2 = True
        '       End If
        '      If check2 = True And check1 = True Then
        'filenamelist.Insert(runner, itemtoinput)
        'insertflag = True
        'Exit For
        '       End If
        '  Next

        'End If
        'End If

        If insertflag = False Then
            Dim midpoint As Integer
            midpoint = filenamelist.Count() Mod 2
            If midpoint = 0 Then
                midpoint = filenamelist.Count() \ 2
            Else
                midpoint = filenamelist.Count() \ 2 + 1
            End If
            'Dim seg1midpoint As Integer
            'seg1midpoint = midpoint Mod 2
            'If seg1midpoint = 0 Then
            'seg1midpoint = midpoint \ 2
            'Else
            '   seg1midpoint = midpoint \ 2 + 1
            'End If
            'Dim seg2midpoint As Integer
            'seg2midpoint = (filenamelist.Count() - midpoint + 1) Mod 2
            'If seg2midpoint = 0 Then
            'seg2midpoint = (filenamelist.Count() - midpoint + 1) \ 2
            'Else
            'seg2midpoint = (filenamelist.Count() - midpoint + 1) \ 2 + 1
            'End If

            Dim segment As Integer

            itemtocheckagainst = filenamelist.Item(filenamelist.Count() - 1)
            Dim result5 As Integer = String.Compare(itemtoinput.FilenameString, itemtocheckagainst.FilenameString)
            '       itemtocheckagainst = filenamelist.Item(seg2midpoint - 1)
            '      Dim result4 As Integer = String.Compare(itemtoinput.FilenameString, itemtocheckagainst.FilenameString)
            itemtocheckagainst = filenamelist.Item(midpoint - 1)
            Dim result3 As Integer = String.Compare(itemtoinput.FilenameString, itemtocheckagainst.FilenameString)
            '     itemtocheckagainst = filenamelist.Item(seg1midpoint - 1)
            '    Dim result2 As Integer = String.Compare(itemtoinput.FilenameString, itemtocheckagainst.FilenameString)
            itemtocheckagainst = filenamelist.Item(0)
            Dim result1 As Integer = String.Compare(itemtoinput.FilenameString, itemtocheckagainst.FilenameString)

            If result1 <= 0 Or result3 = 0 Or result5 >= 0 Then
                If result1 <= 0 And insertflag = False Then
                    filenamelist.Insert(0, itemtoinput)
                    insertflag = True
                End If
      
                If result5 >= 0 And insertflag = False Then
                    filenamelist.Add(itemtoinput)
                    insertflag = True
                End If
                If result3 = 0 And insertflag = False Then
                    filenamelist.Insert(midpoint - 1, itemtoinput)
                    insertflag = True
                End If
            End If
            If insertflag = False Then

                If result3 < 0 Then segment = 1 Else segment = 2
                Dim check1 As Boolean
                Dim check2 As Boolean
                If segment = 1 Then

                    For runner = 0 To midpoint - 1
                        check1 = False
                        check2 = False
                        itemtocheckagainst = filenamelist.Item(runner)
                        If String.Compare(itemtoinput.FilenameString, itemtocheckagainst.FilenameString) >= 0 Then
                            check1 = True
                        End If
                        itemtocheckagainst = filenamelist.Item(runner + 1)
                        If String.Compare(itemtoinput.FilenameString, itemtocheckagainst.FilenameString) <= 0 Then
                            check2 = True
                        End If
                        If check2 = True And check1 = True And insertflag = False Then
                            filenamelist.Insert(runner + 1, itemtoinput)
                            insertflag = True
                            Exit For
                        End If
                    Next
                End If
                If segment = 2 Then

                    For runner = midpoint - 1 To filenamelist.Count() - 2
                        check1 = False
                        check2 = False
                        itemtocheckagainst = filenamelist.Item(runner)
                        If String.Compare(itemtoinput.FilenameString, itemtocheckagainst.FilenameString) >= 0 Then
                            check1 = True
                        End If
                        itemtocheckagainst = filenamelist.Item(runner + 1)
                        If String.Compare(itemtoinput.FilenameString, itemtocheckagainst.FilenameString) <= 0 Then
                            check2 = True
                        End If
                        If check2 = True And check1 = True And insertflag = False Then
                            filenamelist.Insert(runner + 1, itemtoinput)
                            insertflag = True
                            Exit For
                        End If
                    Next
                End If
                
            End If

        End If
        If insertflag = False Then
            MsgBox("Nothing was entered into the arraylist", MsgBoxStyle.Exclamation, "CD List Database Updater 2.0")
        Else
          
        End If
    End Function


    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        checklistset = 0
        displaychecklistset(checklistset)
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        checklistset = checklistset - 1
        If checklistset <= 0 Then checklistset = 1
        displaychecklistset(checklistset)
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        checklistset = checklistset + 1
        'MsgBox("max: " & maxchecklistset)
        'MsgBox("current: " & checklistset)
        If checklistset >= maxchecklistset Then checklistset = maxchecklistset
        'MsgBox("max: " & maxchecklistset)
        'MsgBox("current: " & checklistset)
        displaychecklistset(checklistset)
    End Sub

    Function GetFolderSize(ByVal DirPath As String, _
    Optional ByVal IncludeSubFolders As Boolean = True) As Long

        Dim lngDirSize As Long
        Dim objFileInfo As FileInfo
        Dim objDir As DirectoryInfo = New DirectoryInfo(DirPath)
        Dim objSubFolder As DirectoryInfo

        Try

            'add length of each file
            For Each objFileInfo In objDir.GetFiles()
                lngDirSize += objFileInfo.Length
            Next

            'call recursively to get sub folders
            'if you don't want this set optional
            'parameter to false 
            If IncludeSubFolders Then
                For Each objSubFolder In objDir.GetDirectories()
                    lngDirSize += GetFolderSize(objSubFolder.FullName)
                Next
            End If

        Catch Ex As Exception
            Error_Handler(Ex.Message)
        End Try
        Return lngDirSize
    End Function






    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        TabControl1.SelectedTab = TabPage2

    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Try
            Me.Close()
            Me.Dispose()
        Catch err As Exception
            Error_Handler(err.Message)
        End Try
    End Sub

   
    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        TabControl1.SelectedTab = TabPage3
    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        Try
            Dim AboutDialog1 As New About_Dialog
            Dim result As DialogResult = AboutDialog1.ShowDialog()
        Catch err As Exception
            Error_Handler(err.Message)
        End Try
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        Try
            Dim helpscreen As New Help
            'Dim result As DialogResult
            'result = helpscreen.ShowDialog()
            helpscreen.Show()
            'helpscreen.Dispose()
        Catch et As Exception
            Error_Handler(et.Message)
        End Try
    End Sub


    Private Function Write_HTML_Report() As Boolean
        Try
            Dim f As FileInfo
            f = New FileInfo(Application.ExecutablePath())
            Dim outputpath, outputdirectory As String
            outputpath = f.DirectoryName() & "\Reports\Report.htm"
            outputdirectory = f.DirectoryName() & "\Reports"

            Dim di As DirectoryInfo = New DirectoryInfo(outputdirectory)
            If di.Exists = False Then
                di.Create()
            End If
            Dim objStreamWriter = New System.IO.StreamWriter(outputpath, False, System.Text.Encoding.Unicode)

            objStreamWriter.WriteLine("<html><head><Title>CD List Database Updater :: Report (" & TimeOfDay() & ")</Title></head>")
            objStreamWriter.WriteLine("<body bgproperties=""fixed"" background=""" & f.DirectoryName() & "\Reports\Background_Image.jpg"">")
            objStreamWriter.writeline("<table width=""75%"" border=""0""><tr><td>")
            objStreamWriter.writeline("<h1>CD List Database Updater Report</h1>")
            objStreamWriter.writeline("<p>" & ParentPath.Text & "<br>")
            objStreamWriter.writeline("<p>" & Replace(Datasource.Text, "Table:", "<br>Table:") & "</p>")
            objStreamWriter.writeline("<p>")
            Dim texttowrite As String
            texttowrite = RichTextBox1.Text
            texttowrite = Replace(texttowrite, "--------------", "</li>")
            texttowrite = Replace(texttowrite, "insert into ", "</i><li>insert into ")
            texttowrite = Replace(texttowrite, "- SUCCESS ", "<br><i>- SUCCESS ")
            texttowrite = Replace(texttowrite, "- FAILED ", "<br><i>- FAILED ")
            texttowrite = Replace(texttowrite, " -]", "</b></p><ul><i>")
            texttowrite = Replace(texttowrite, "[- ", "<b>Report generated at ")

            objStreamWriter.WriteLine(texttowrite)
            objStreamWriter.writeline("</i></li></ul>")
            objStreamWriter.writeline("<p><i>" & updateresultlabel.Text & "</i></p>")
            objStreamWriter.writeline("</td></tr></table>")
            objStreamWriter.WriteLine("</body></html>")
            objStreamWriter.Close()
            Return (True)
        Catch et As Exception
                Error_Handler(et.Message)
                Return (False)
            End Try
    End Function

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            Dim f As FileInfo
            f = New FileInfo(Application.ExecutablePath())
            Dim outputpath As String
            outputpath = f.DirectoryName() & "/Reports/Report.htm"
            System.Diagnostics.Process.Start(outputpath)
        Catch et As Exception
            Error_Handler(et.Message)
        End Try
    End Sub
End Class