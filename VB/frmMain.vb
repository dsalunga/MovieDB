Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class frmMain
    Inherits System.Windows.Forms.Form
    Dim NewArrivals As Long = 0
    Dim SortedColumn As Integer = 0
    Dim ColumnSortedAsc As Boolean = True
    'Dim SelectedItemIndex As Integer
    Dim lastSELECTString As String, lastORDERString As String

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
    Friend WithEvents menuItemViewsList As System.Windows.Forms.MenuItem
    Friend WithEvents ToolBarButton5 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton4 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton3 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton2 As System.Windows.Forms.ToolBarButton
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents toolBarMain As System.Windows.Forms.ToolBar
    Friend WithEvents statusBarMain As System.Windows.Forms.StatusBar
    Friend WithEvents panelProperties As System.Windows.Forms.Panel
    Friend WithEvents panelTop As System.Windows.Forms.Panel
    Friend WithEvents listViewRecords As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader10 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader11 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader12 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader13 As System.Windows.Forms.ColumnHeader
    Friend WithEvents imageListLarge As System.Windows.Forms.ImageList
    Friend WithEvents contextMenuViews As System.Windows.Forms.ContextMenu
    Friend WithEvents menuItemViewsLargeIcon As System.Windows.Forms.MenuItem
    Friend WithEvents menuItemViewsDetails As System.Windows.Forms.MenuItem
    Friend WithEvents contextMenuRecords As System.Windows.Forms.ContextMenu
    Friend WithEvents ToolBarButton6 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton7 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton8 As System.Windows.Forms.ToolBarButton
    Friend WithEvents toolBarButtonViews As System.Windows.Forms.ToolBarButton
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents tabPageProperties As System.Windows.Forms.TabPage
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents contextMenuProperties As System.Windows.Forms.ContextMenu
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel19 As System.Windows.Forms.Panel
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Panel12 As System.Windows.Forms.Panel
    Friend WithEvents Panel9 As System.Windows.Forms.Panel
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Panel7 As System.Windows.Forms.Panel
    Friend WithEvents Panel8 As System.Windows.Forms.Panel
    Friend WithEvents txtYear As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Panel10 As System.Windows.Forms.Panel
    Friend WithEvents txtRating As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Panel11 As System.Windows.Forms.Panel
    Friend WithEvents txtCategory As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents txtProducer As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Panel22 As System.Windows.Forms.Panel
    Friend WithEvents txtDirector As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents panelStarring As System.Windows.Forms.Panel
    Friend WithEvents txtStarring As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents Panel23 As System.Windows.Forms.Panel
    Friend WithEvents pictureBoxThumb As System.Windows.Forms.PictureBox
    Friend WithEvents txtCondition As System.Windows.Forms.TextBox
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents Panel24 As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtType As System.Windows.Forms.TextBox
    Friend WithEvents Panel21 As System.Windows.Forms.Panel
    Friend WithEvents txtDateReturned As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Panel16 As System.Windows.Forms.Panel
    Friend WithEvents txtDateBorrowed As System.Windows.Forms.TextBox
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Panel17 As System.Windows.Forms.Panel
    Friend WithEvents txtBorrower As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Panel20 As System.Windows.Forms.Panel
    Friend WithEvents txtStatus As System.Windows.Forms.TextBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Panel15 As System.Windows.Forms.Panel
    Friend WithEvents Panel18 As System.Windows.Forms.Panel
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents Panel14 As System.Windows.Forms.Panel
    Friend WithEvents Panel13 As System.Windows.Forms.Panel
    Friend WithEvents txtRecordDate As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents cmdFilter As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents cboFilterBy As System.Windows.Forms.ComboBox
    Friend WithEvents cboFilter As System.Windows.Forms.ComboBox
    Friend WithEvents imageListSmall As System.Windows.Forms.ImageList
    Friend WithEvents imageListSmall2 As System.Windows.Forms.ImageList
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.toolBarMain = New System.Windows.Forms.ToolBar()
        Me.ToolBarButton2 = New System.Windows.Forms.ToolBarButton()
        Me.ToolBarButton3 = New System.Windows.Forms.ToolBarButton()
        Me.ToolBarButton4 = New System.Windows.Forms.ToolBarButton()
        Me.ToolBarButton5 = New System.Windows.Forms.ToolBarButton()
        Me.ToolBarButton7 = New System.Windows.Forms.ToolBarButton()
        Me.ToolBarButton8 = New System.Windows.Forms.ToolBarButton()
        Me.ToolBarButton6 = New System.Windows.Forms.ToolBarButton()
        Me.toolBarButtonViews = New System.Windows.Forms.ToolBarButton()
        Me.contextMenuViews = New System.Windows.Forms.ContextMenu()
        Me.menuItemViewsLargeIcon = New System.Windows.Forms.MenuItem()
        Me.menuItemViewsList = New System.Windows.Forms.MenuItem()
        Me.menuItemViewsDetails = New System.Windows.Forms.MenuItem()
        Me.imageListLarge = New System.Windows.Forms.ImageList(Me.components)
        Me.panelTop = New System.Windows.Forms.Panel()
        Me.cmdFilter = New System.Windows.Forms.Button()
        Me.cboFilter = New System.Windows.Forms.ComboBox()
        Me.cboFilterBy = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.statusBarMain = New System.Windows.Forms.StatusBar()
        Me.panelProperties = New System.Windows.Forms.Panel()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.contextMenuProperties = New System.Windows.Forms.ContextMenu()
        Me.tabPageProperties = New System.Windows.Forms.TabPage()
        Me.Panel21 = New System.Windows.Forms.Panel()
        Me.txtDateReturned = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Panel16 = New System.Windows.Forms.Panel()
        Me.txtDateBorrowed = New System.Windows.Forms.TextBox()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.Panel17 = New System.Windows.Forms.Panel()
        Me.txtBorrower = New System.Windows.Forms.TextBox()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Panel20 = New System.Windows.Forms.Panel()
        Me.txtStatus = New System.Windows.Forms.TextBox()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.Panel15 = New System.Windows.Forms.Panel()
        Me.Panel18 = New System.Windows.Forms.Panel()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.Panel14 = New System.Windows.Forms.Panel()
        Me.Panel13 = New System.Windows.Forms.Panel()
        Me.txtRecordDate = New System.Windows.Forms.TextBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Panel24 = New System.Windows.Forms.Panel()
        Me.txtType = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Panel19 = New System.Windows.Forms.Panel()
        Me.txtCondition = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Panel12 = New System.Windows.Forms.Panel()
        Me.Panel9 = New System.Windows.Forms.Panel()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Panel7 = New System.Windows.Forms.Panel()
        Me.Panel8 = New System.Windows.Forms.Panel()
        Me.txtYear = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Panel10 = New System.Windows.Forms.Panel()
        Me.txtRating = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Panel11 = New System.Windows.Forms.Panel()
        Me.txtCategory = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.txtProducer = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Panel22 = New System.Windows.Forms.Panel()
        Me.txtDirector = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.panelStarring = New System.Windows.Forms.Panel()
        Me.txtStarring = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.Panel23 = New System.Windows.Forms.Panel()
        Me.pictureBoxThumb = New System.Windows.Forms.PictureBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.listViewRecords = New System.Windows.Forms.ListView()
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader6 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader8 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader7 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader12 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader10 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader9 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader11 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader13 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.contextMenuRecords = New System.Windows.Forms.ContextMenu()
        Me.imageListSmall = New System.Windows.Forms.ImageList(Me.components)
        Me.Splitter1 = New System.Windows.Forms.Splitter()
        Me.imageListSmall2 = New System.Windows.Forms.ImageList(Me.components)
        Me.panelTop.SuspendLayout()
        Me.panelProperties.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.tabPageProperties.SuspendLayout()
        Me.Panel21.SuspendLayout()
        Me.Panel16.SuspendLayout()
        Me.Panel17.SuspendLayout()
        Me.Panel20.SuspendLayout()
        Me.Panel13.SuspendLayout()
        Me.Panel24.SuspendLayout()
        Me.Panel19.SuspendLayout()
        Me.Panel8.SuspendLayout()
        Me.Panel10.SuspendLayout()
        Me.Panel11.SuspendLayout()
        Me.Panel6.SuspendLayout()
        Me.Panel22.SuspendLayout()
        Me.panelStarring.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel23.SuspendLayout()
        CType(Me.pictureBoxThumb, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'toolBarMain
        '
        Me.toolBarMain.Appearance = System.Windows.Forms.ToolBarAppearance.Flat
        Me.toolBarMain.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.ToolBarButton2, Me.ToolBarButton3, Me.ToolBarButton4, Me.ToolBarButton5, Me.ToolBarButton7, Me.ToolBarButton8, Me.ToolBarButton6, Me.toolBarButtonViews})
        Me.toolBarMain.DropDownArrows = True
        Me.toolBarMain.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.toolBarMain.ImageList = Me.imageListLarge
        Me.toolBarMain.Location = New System.Drawing.Point(0, 0)
        Me.toolBarMain.Name = "toolBarMain"
        Me.toolBarMain.ShowToolTips = True
        Me.toolBarMain.Size = New System.Drawing.Size(692, 52)
        Me.toolBarMain.TabIndex = 0
        Me.toolBarMain.TextAlign = System.Windows.Forms.ToolBarTextAlign.Right
        '
        'ToolBarButton2
        '
        Me.ToolBarButton2.ImageIndex = 0
        Me.ToolBarButton2.Name = "ToolBarButton2"
        Me.ToolBarButton2.Text = "New"
        Me.ToolBarButton2.ToolTipText = "New"
        '
        'ToolBarButton3
        '
        Me.ToolBarButton3.ImageIndex = 3
        Me.ToolBarButton3.Name = "ToolBarButton3"
        Me.ToolBarButton3.Text = "Edit"
        Me.ToolBarButton3.ToolTipText = "Edit"
        '
        'ToolBarButton4
        '
        Me.ToolBarButton4.ImageIndex = 5
        Me.ToolBarButton4.Name = "ToolBarButton4"
        Me.ToolBarButton4.Text = "Delete"
        Me.ToolBarButton4.ToolTipText = "Delete"
        '
        'ToolBarButton5
        '
        Me.ToolBarButton5.Name = "ToolBarButton5"
        Me.ToolBarButton5.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'ToolBarButton7
        '
        Me.ToolBarButton7.ImageIndex = 6
        Me.ToolBarButton7.Name = "ToolBarButton7"
        Me.ToolBarButton7.Text = "Refresh"
        '
        'ToolBarButton8
        '
        Me.ToolBarButton8.Name = "ToolBarButton8"
        Me.ToolBarButton8.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'ToolBarButton6
        '
        Me.ToolBarButton6.ImageIndex = 0
        Me.ToolBarButton6.Name = "ToolBarButton6"
        Me.ToolBarButton6.Text = "Options"
        Me.ToolBarButton6.Visible = False
        '
        'toolBarButtonViews
        '
        Me.toolBarButtonViews.DropDownMenu = Me.contextMenuViews
        Me.toolBarButtonViews.ImageIndex = 7
        Me.toolBarButtonViews.Name = "toolBarButtonViews"
        Me.toolBarButtonViews.Style = System.Windows.Forms.ToolBarButtonStyle.DropDownButton
        Me.toolBarButtonViews.Text = "Views"
        Me.toolBarButtonViews.ToolTipText = "Views"
        '
        'contextMenuViews
        '
        Me.contextMenuViews.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.menuItemViewsLargeIcon, Me.menuItemViewsList, Me.menuItemViewsDetails})
        '
        'menuItemViewsLargeIcon
        '
        Me.menuItemViewsLargeIcon.Index = 0
        Me.menuItemViewsLargeIcon.RadioCheck = True
        Me.menuItemViewsLargeIcon.Text = "Large Icon"
        '
        'menuItemViewsList
        '
        Me.menuItemViewsList.Index = 1
        Me.menuItemViewsList.RadioCheck = True
        Me.menuItemViewsList.Text = "List"
        '
        'menuItemViewsDetails
        '
        Me.menuItemViewsDetails.Index = 2
        Me.menuItemViewsDetails.RadioCheck = True
        Me.menuItemViewsDetails.Text = "Details"
        '
        'imageListLarge
        '
        Me.imageListLarge.ImageStream = CType(resources.GetObject("imageListLarge.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imageListLarge.TransparentColor = System.Drawing.Color.Transparent
        Me.imageListLarge.Images.SetKeyName(0, "")
        Me.imageListLarge.Images.SetKeyName(1, "")
        Me.imageListLarge.Images.SetKeyName(2, "")
        Me.imageListLarge.Images.SetKeyName(3, "")
        Me.imageListLarge.Images.SetKeyName(4, "")
        Me.imageListLarge.Images.SetKeyName(5, "")
        Me.imageListLarge.Images.SetKeyName(6, "")
        Me.imageListLarge.Images.SetKeyName(7, "")
        Me.imageListLarge.Images.SetKeyName(8, "")
        Me.imageListLarge.Images.SetKeyName(9, "")
        '
        'panelTop
        '
        Me.panelTop.Controls.Add(Me.cmdFilter)
        Me.panelTop.Controls.Add(Me.cboFilter)
        Me.panelTop.Controls.Add(Me.cboFilterBy)
        Me.panelTop.Controls.Add(Me.Label2)
        Me.panelTop.Controls.Add(Me.Label1)
        Me.panelTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.panelTop.Location = New System.Drawing.Point(0, 52)
        Me.panelTop.Name = "panelTop"
        Me.panelTop.Size = New System.Drawing.Size(692, 39)
        Me.panelTop.TabIndex = 1
        '
        'cmdFilter
        '
        Me.cmdFilter.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdFilter.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdFilter.Location = New System.Drawing.Point(384, 8)
        Me.cmdFilter.Name = "cmdFilter"
        Me.cmdFilter.Size = New System.Drawing.Size(75, 23)
        Me.cmdFilter.TabIndex = 5
        Me.cmdFilter.Text = "Filter Now"
        '
        'cboFilter
        '
        Me.cboFilter.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboFilter.Location = New System.Drawing.Point(256, 8)
        Me.cboFilter.Name = "cboFilter"
        Me.cboFilter.Size = New System.Drawing.Size(112, 21)
        Me.cboFilter.Sorted = True
        Me.cboFilter.TabIndex = 4
        '
        'cboFilterBy
        '
        Me.cboFilterBy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFilterBy.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboFilterBy.Items.AddRange(New Object() {"Title", "Cast", "Director", "Condition", "Category", "Status", "Borrower", "New Arrivals"})
        Me.cboFilterBy.Location = New System.Drawing.Point(72, 8)
        Me.cboFilterBy.Name = "cboFilterBy"
        Me.cboFilterBy.Size = New System.Drawing.Size(121, 21)
        Me.cboFilterBy.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(16, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 23)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Filter by:"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(216, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 23)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Filter:"
        '
        'statusBarMain
        '
        Me.statusBarMain.Location = New System.Drawing.Point(0, 592)
        Me.statusBarMain.Name = "statusBarMain"
        Me.statusBarMain.Size = New System.Drawing.Size(692, 22)
        Me.statusBarMain.TabIndex = 5
        Me.statusBarMain.Text = "StatusBar"
        '
        'panelProperties
        '
        Me.panelProperties.BackColor = System.Drawing.SystemColors.ControlLight
        Me.panelProperties.Controls.Add(Me.TabControl1)
        Me.panelProperties.Dock = System.Windows.Forms.DockStyle.Left
        Me.panelProperties.Location = New System.Drawing.Point(0, 91)
        Me.panelProperties.Name = "panelProperties"
        Me.panelProperties.Size = New System.Drawing.Size(288, 501)
        Me.panelProperties.TabIndex = 6
        '
        'TabControl1
        '
        Me.TabControl1.ContextMenu = Me.contextMenuProperties
        Me.TabControl1.Controls.Add(Me.tabPageProperties)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabControl1.ItemSize = New System.Drawing.Size(42, 18)
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(288, 501)
        Me.TabControl1.TabIndex = 0
        '
        'tabPageProperties
        '
        Me.tabPageProperties.BackColor = System.Drawing.SystemColors.ControlLight
        Me.tabPageProperties.Controls.Add(Me.Panel21)
        Me.tabPageProperties.Controls.Add(Me.Panel16)
        Me.tabPageProperties.Controls.Add(Me.Panel17)
        Me.tabPageProperties.Controls.Add(Me.Panel20)
        Me.tabPageProperties.Controls.Add(Me.Panel15)
        Me.tabPageProperties.Controls.Add(Me.Panel18)
        Me.tabPageProperties.Controls.Add(Me.Label25)
        Me.tabPageProperties.Controls.Add(Me.Panel14)
        Me.tabPageProperties.Controls.Add(Me.Panel13)
        Me.tabPageProperties.Controls.Add(Me.Panel24)
        Me.tabPageProperties.Controls.Add(Me.Panel19)
        Me.tabPageProperties.Controls.Add(Me.Panel12)
        Me.tabPageProperties.Controls.Add(Me.Panel9)
        Me.tabPageProperties.Controls.Add(Me.Label16)
        Me.tabPageProperties.Controls.Add(Me.Panel7)
        Me.tabPageProperties.Controls.Add(Me.Panel8)
        Me.tabPageProperties.Controls.Add(Me.Panel10)
        Me.tabPageProperties.Controls.Add(Me.Panel11)
        Me.tabPageProperties.Controls.Add(Me.Panel6)
        Me.tabPageProperties.Controls.Add(Me.Panel22)
        Me.tabPageProperties.Controls.Add(Me.Panel5)
        Me.tabPageProperties.Controls.Add(Me.panelStarring)
        Me.tabPageProperties.Controls.Add(Me.Panel4)
        Me.tabPageProperties.Controls.Add(Me.Panel3)
        Me.tabPageProperties.Controls.Add(Me.Panel2)
        Me.tabPageProperties.Controls.Add(Me.Panel1)
        Me.tabPageProperties.Location = New System.Drawing.Point(4, 22)
        Me.tabPageProperties.Name = "tabPageProperties"
        Me.tabPageProperties.Size = New System.Drawing.Size(280, 475)
        Me.tabPageProperties.TabIndex = 0
        Me.tabPageProperties.Text = "Properties"
        '
        'Panel21
        '
        Me.Panel21.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Panel21.Controls.Add(Me.txtDateReturned)
        Me.Panel21.Controls.Add(Me.Label5)
        Me.Panel21.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel21.Location = New System.Drawing.Point(0, 446)
        Me.Panel21.Name = "Panel21"
        Me.Panel21.Size = New System.Drawing.Size(276, 24)
        Me.Panel21.TabIndex = 89
        '
        'txtDateReturned
        '
        Me.txtDateReturned.BackColor = System.Drawing.SystemColors.ControlLight
        Me.txtDateReturned.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtDateReturned.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDateReturned.Location = New System.Drawing.Point(72, 0)
        Me.txtDateReturned.Name = "txtDateReturned"
        Me.txtDateReturned.ReadOnly = True
        Me.txtDateReturned.Size = New System.Drawing.Size(204, 21)
        Me.txtDateReturned.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(0, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(72, 24)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Returned:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel16
        '
        Me.Panel16.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Panel16.Controls.Add(Me.txtDateBorrowed)
        Me.Panel16.Controls.Add(Me.Label27)
        Me.Panel16.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel16.Location = New System.Drawing.Point(0, 422)
        Me.Panel16.Name = "Panel16"
        Me.Panel16.Size = New System.Drawing.Size(276, 24)
        Me.Panel16.TabIndex = 88
        '
        'txtDateBorrowed
        '
        Me.txtDateBorrowed.BackColor = System.Drawing.SystemColors.ControlLight
        Me.txtDateBorrowed.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtDateBorrowed.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDateBorrowed.Location = New System.Drawing.Point(72, 0)
        Me.txtDateBorrowed.Name = "txtDateBorrowed"
        Me.txtDateBorrowed.ReadOnly = True
        Me.txtDateBorrowed.Size = New System.Drawing.Size(204, 21)
        Me.txtDateBorrowed.TabIndex = 1
        '
        'Label27
        '
        Me.Label27.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label27.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label27.Location = New System.Drawing.Point(0, 0)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(72, 24)
        Me.Label27.TabIndex = 0
        Me.Label27.Text = "Borrowed:"
        Me.Label27.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel17
        '
        Me.Panel17.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Panel17.Controls.Add(Me.txtBorrower)
        Me.Panel17.Controls.Add(Me.Label22)
        Me.Panel17.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel17.Location = New System.Drawing.Point(0, 398)
        Me.Panel17.Name = "Panel17"
        Me.Panel17.Size = New System.Drawing.Size(276, 24)
        Me.Panel17.TabIndex = 87
        '
        'txtBorrower
        '
        Me.txtBorrower.BackColor = System.Drawing.SystemColors.ControlLight
        Me.txtBorrower.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtBorrower.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBorrower.Location = New System.Drawing.Point(72, 0)
        Me.txtBorrower.Name = "txtBorrower"
        Me.txtBorrower.ReadOnly = True
        Me.txtBorrower.Size = New System.Drawing.Size(204, 21)
        Me.txtBorrower.TabIndex = 1
        '
        'Label22
        '
        Me.Label22.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label22.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.Location = New System.Drawing.Point(0, 0)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(72, 24)
        Me.Label22.TabIndex = 0
        Me.Label22.Text = "Borrower:"
        Me.Label22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel20
        '
        Me.Panel20.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Panel20.Controls.Add(Me.txtStatus)
        Me.Panel20.Controls.Add(Me.Label24)
        Me.Panel20.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel20.Location = New System.Drawing.Point(0, 374)
        Me.Panel20.Name = "Panel20"
        Me.Panel20.Size = New System.Drawing.Size(276, 24)
        Me.Panel20.TabIndex = 86
        '
        'txtStatus
        '
        Me.txtStatus.BackColor = System.Drawing.SystemColors.ControlLight
        Me.txtStatus.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtStatus.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStatus.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtStatus.Location = New System.Drawing.Point(72, 0)
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.Size = New System.Drawing.Size(204, 21)
        Me.txtStatus.TabIndex = 1
        '
        'Label24
        '
        Me.Label24.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label24.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label24.Location = New System.Drawing.Point(0, 0)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(72, 24)
        Me.Label24.TabIndex = 0
        Me.Label24.Text = "Status:"
        Me.Label24.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel15
        '
        Me.Panel15.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Panel15.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel15.Location = New System.Drawing.Point(0, 366)
        Me.Panel15.Name = "Panel15"
        Me.Panel15.Size = New System.Drawing.Size(276, 8)
        Me.Panel15.TabIndex = 85
        '
        'Panel18
        '
        Me.Panel18.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Panel18.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel18.Location = New System.Drawing.Point(0, 364)
        Me.Panel18.Name = "Panel18"
        Me.Panel18.Size = New System.Drawing.Size(276, 2)
        Me.Panel18.TabIndex = 84
        '
        'Label25
        '
        Me.Label25.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Label25.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label25.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label25.Location = New System.Drawing.Point(0, 348)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(276, 16)
        Me.Label25.TabIndex = 83
        Me.Label25.Text = "Availability"
        '
        'Panel14
        '
        Me.Panel14.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Panel14.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel14.Location = New System.Drawing.Point(0, 334)
        Me.Panel14.Name = "Panel14"
        Me.Panel14.Size = New System.Drawing.Size(276, 14)
        Me.Panel14.TabIndex = 82
        '
        'Panel13
        '
        Me.Panel13.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Panel13.Controls.Add(Me.txtRecordDate)
        Me.Panel13.Controls.Add(Me.Label20)
        Me.Panel13.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel13.Location = New System.Drawing.Point(0, 310)
        Me.Panel13.Name = "Panel13"
        Me.Panel13.Size = New System.Drawing.Size(276, 24)
        Me.Panel13.TabIndex = 81
        '
        'txtRecordDate
        '
        Me.txtRecordDate.BackColor = System.Drawing.SystemColors.ControlLight
        Me.txtRecordDate.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtRecordDate.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRecordDate.Location = New System.Drawing.Point(72, 0)
        Me.txtRecordDate.Name = "txtRecordDate"
        Me.txtRecordDate.ReadOnly = True
        Me.txtRecordDate.Size = New System.Drawing.Size(204, 21)
        Me.txtRecordDate.TabIndex = 1
        '
        'Label20
        '
        Me.Label20.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label20.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.Location = New System.Drawing.Point(0, 0)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(72, 24)
        Me.Label20.TabIndex = 0
        Me.Label20.Text = "Record Date:"
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel24
        '
        Me.Panel24.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Panel24.Controls.Add(Me.txtType)
        Me.Panel24.Controls.Add(Me.Label3)
        Me.Panel24.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel24.Location = New System.Drawing.Point(0, 286)
        Me.Panel24.Name = "Panel24"
        Me.Panel24.Size = New System.Drawing.Size(276, 24)
        Me.Panel24.TabIndex = 80
        '
        'txtType
        '
        Me.txtType.BackColor = System.Drawing.SystemColors.ControlLight
        Me.txtType.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtType.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtType.Location = New System.Drawing.Point(72, 0)
        Me.txtType.Name = "txtType"
        Me.txtType.ReadOnly = True
        Me.txtType.Size = New System.Drawing.Size(204, 21)
        Me.txtType.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(0, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 24)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Type:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel19
        '
        Me.Panel19.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Panel19.Controls.Add(Me.txtCondition)
        Me.Panel19.Controls.Add(Me.Label18)
        Me.Panel19.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel19.Location = New System.Drawing.Point(0, 262)
        Me.Panel19.Name = "Panel19"
        Me.Panel19.Size = New System.Drawing.Size(276, 24)
        Me.Panel19.TabIndex = 70
        '
        'txtCondition
        '
        Me.txtCondition.BackColor = System.Drawing.SystemColors.ControlLight
        Me.txtCondition.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtCondition.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCondition.Location = New System.Drawing.Point(72, 0)
        Me.txtCondition.Name = "txtCondition"
        Me.txtCondition.ReadOnly = True
        Me.txtCondition.Size = New System.Drawing.Size(204, 21)
        Me.txtCondition.TabIndex = 1
        '
        'Label18
        '
        Me.Label18.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label18.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.Location = New System.Drawing.Point(0, 0)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(72, 24)
        Me.Label18.TabIndex = 0
        Me.Label18.Text = "Condition:"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel12
        '
        Me.Panel12.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Panel12.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel12.Location = New System.Drawing.Point(0, 254)
        Me.Panel12.Name = "Panel12"
        Me.Panel12.Size = New System.Drawing.Size(276, 8)
        Me.Panel12.TabIndex = 69
        '
        'Panel9
        '
        Me.Panel9.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Panel9.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel9.Location = New System.Drawing.Point(0, 252)
        Me.Panel9.Name = "Panel9"
        Me.Panel9.Size = New System.Drawing.Size(276, 2)
        Me.Panel9.TabIndex = 68
        '
        'Label16
        '
        Me.Label16.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Label16.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label16.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(0, 236)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(276, 16)
        Me.Label16.TabIndex = 67
        Me.Label16.Text = "Record Media"
        '
        'Panel7
        '
        Me.Panel7.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Panel7.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel7.Location = New System.Drawing.Point(0, 220)
        Me.Panel7.Name = "Panel7"
        Me.Panel7.Size = New System.Drawing.Size(276, 16)
        Me.Panel7.TabIndex = 66
        '
        'Panel8
        '
        Me.Panel8.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Panel8.Controls.Add(Me.txtYear)
        Me.Panel8.Controls.Add(Me.Label11)
        Me.Panel8.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel8.Location = New System.Drawing.Point(0, 196)
        Me.Panel8.Name = "Panel8"
        Me.Panel8.Size = New System.Drawing.Size(276, 24)
        Me.Panel8.TabIndex = 65
        '
        'txtYear
        '
        Me.txtYear.BackColor = System.Drawing.SystemColors.ControlLight
        Me.txtYear.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtYear.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtYear.Location = New System.Drawing.Point(72, 0)
        Me.txtYear.Name = "txtYear"
        Me.txtYear.ReadOnly = True
        Me.txtYear.Size = New System.Drawing.Size(204, 21)
        Me.txtYear.TabIndex = 1
        '
        'Label11
        '
        Me.Label11.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label11.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(0, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(72, 24)
        Me.Label11.TabIndex = 0
        Me.Label11.Text = "Year:"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel10
        '
        Me.Panel10.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Panel10.Controls.Add(Me.txtRating)
        Me.Panel10.Controls.Add(Me.Label15)
        Me.Panel10.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel10.Location = New System.Drawing.Point(0, 172)
        Me.Panel10.Name = "Panel10"
        Me.Panel10.Size = New System.Drawing.Size(276, 24)
        Me.Panel10.TabIndex = 64
        '
        'txtRating
        '
        Me.txtRating.BackColor = System.Drawing.SystemColors.ControlLight
        Me.txtRating.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtRating.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRating.Location = New System.Drawing.Point(72, 0)
        Me.txtRating.Name = "txtRating"
        Me.txtRating.ReadOnly = True
        Me.txtRating.Size = New System.Drawing.Size(204, 21)
        Me.txtRating.TabIndex = 1
        '
        'Label15
        '
        Me.Label15.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label15.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(0, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(72, 24)
        Me.Label15.TabIndex = 0
        Me.Label15.Text = "Rating:"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel11
        '
        Me.Panel11.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Panel11.Controls.Add(Me.txtCategory)
        Me.Panel11.Controls.Add(Me.Label13)
        Me.Panel11.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel11.Location = New System.Drawing.Point(0, 148)
        Me.Panel11.Name = "Panel11"
        Me.Panel11.Size = New System.Drawing.Size(276, 24)
        Me.Panel11.TabIndex = 63
        '
        'txtCategory
        '
        Me.txtCategory.BackColor = System.Drawing.SystemColors.ControlLight
        Me.txtCategory.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtCategory.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCategory.Location = New System.Drawing.Point(72, 0)
        Me.txtCategory.Name = "txtCategory"
        Me.txtCategory.ReadOnly = True
        Me.txtCategory.Size = New System.Drawing.Size(204, 21)
        Me.txtCategory.TabIndex = 1
        '
        'Label13
        '
        Me.Label13.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label13.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(0, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(72, 24)
        Me.Label13.TabIndex = 0
        Me.Label13.Text = "Category:"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel6
        '
        Me.Panel6.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Panel6.Controls.Add(Me.txtProducer)
        Me.Panel6.Controls.Add(Me.Label9)
        Me.Panel6.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel6.Location = New System.Drawing.Point(0, 124)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(276, 24)
        Me.Panel6.TabIndex = 62
        '
        'txtProducer
        '
        Me.txtProducer.BackColor = System.Drawing.SystemColors.ControlLight
        Me.txtProducer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtProducer.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtProducer.Location = New System.Drawing.Point(72, 0)
        Me.txtProducer.Name = "txtProducer"
        Me.txtProducer.ReadOnly = True
        Me.txtProducer.Size = New System.Drawing.Size(204, 21)
        Me.txtProducer.TabIndex = 1
        '
        'Label9
        '
        Me.Label9.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label9.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(0, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(72, 24)
        Me.Label9.TabIndex = 0
        Me.Label9.Text = "Producer:"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel22
        '
        Me.Panel22.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Panel22.Controls.Add(Me.txtDirector)
        Me.Panel22.Controls.Add(Me.Label7)
        Me.Panel22.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel22.Location = New System.Drawing.Point(0, 100)
        Me.Panel22.Name = "Panel22"
        Me.Panel22.Size = New System.Drawing.Size(276, 24)
        Me.Panel22.TabIndex = 61
        '
        'txtDirector
        '
        Me.txtDirector.BackColor = System.Drawing.SystemColors.ControlLight
        Me.txtDirector.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtDirector.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDirector.Location = New System.Drawing.Point(72, 0)
        Me.txtDirector.Name = "txtDirector"
        Me.txtDirector.ReadOnly = True
        Me.txtDirector.Size = New System.Drawing.Size(204, 21)
        Me.txtDirector.TabIndex = 1
        '
        'Label7
        '
        Me.Label7.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label7.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(0, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(72, 24)
        Me.Label7.TabIndex = 0
        Me.Label7.Text = "Director:"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel5
        '
        Me.Panel5.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel5.Location = New System.Drawing.Point(0, 96)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(276, 4)
        Me.Panel5.TabIndex = 60
        '
        'panelStarring
        '
        Me.panelStarring.Controls.Add(Me.txtStarring)
        Me.panelStarring.Controls.Add(Me.Label4)
        Me.panelStarring.Dock = System.Windows.Forms.DockStyle.Top
        Me.panelStarring.Location = New System.Drawing.Point(0, 58)
        Me.panelStarring.Name = "panelStarring"
        Me.panelStarring.Size = New System.Drawing.Size(276, 38)
        Me.panelStarring.TabIndex = 59
        '
        'txtStarring
        '
        Me.txtStarring.BackColor = System.Drawing.SystemColors.ControlLight
        Me.txtStarring.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtStarring.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStarring.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtStarring.Location = New System.Drawing.Point(72, 0)
        Me.txtStarring.Multiline = True
        Me.txtStarring.Name = "txtStarring"
        Me.txtStarring.ReadOnly = True
        Me.txtStarring.Size = New System.Drawing.Size(204, 38)
        Me.txtStarring.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(0, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 38)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Cast:"
        '
        'Panel4
        '
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel4.Location = New System.Drawing.Point(0, 50)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(276, 8)
        Me.Panel4.TabIndex = 58
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.SystemColors.ControlDark
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Location = New System.Drawing.Point(0, 48)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(276, 2)
        Me.Panel3.TabIndex = 57
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.lblTitle)
        Me.Panel2.Controls.Add(Me.Panel23)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(276, 48)
        Me.Panel2.TabIndex = 56
        '
        'lblTitle
        '
        Me.lblTitle.BackColor = System.Drawing.SystemColors.ControlLight
        Me.lblTitle.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblTitle.Font = New System.Drawing.Font("Trebuchet MS", 12.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.ForeColor = System.Drawing.Color.Red
        Me.lblTitle.Location = New System.Drawing.Point(64, 0)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(212, 48)
        Me.lblTitle.TabIndex = 1
        Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel23
        '
        Me.Panel23.Controls.Add(Me.pictureBoxThumb)
        Me.Panel23.Dock = System.Windows.Forms.DockStyle.Left
        Me.Panel23.Location = New System.Drawing.Point(0, 0)
        Me.Panel23.Name = "Panel23"
        Me.Panel23.Size = New System.Drawing.Size(64, 48)
        Me.Panel23.TabIndex = 0
        '
        'pictureBoxThumb
        '
        Me.pictureBoxThumb.Location = New System.Drawing.Point(16, 8)
        Me.pictureBoxThumb.Name = "pictureBoxThumb"
        Me.pictureBoxThumb.Size = New System.Drawing.Size(32, 32)
        Me.pictureBoxThumb.TabIndex = 0
        Me.pictureBoxThumb.TabStop = False
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel1.Location = New System.Drawing.Point(276, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(4, 475)
        Me.Panel1.TabIndex = 0
        '
        'listViewRecords
        '
        Me.listViewRecords.AllowColumnReorder = True
        Me.listViewRecords.BackColor = System.Drawing.SystemColors.Window
        Me.listViewRecords.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6, Me.ColumnHeader8, Me.ColumnHeader7, Me.ColumnHeader12, Me.ColumnHeader10, Me.ColumnHeader9, Me.ColumnHeader11, Me.ColumnHeader1, Me.ColumnHeader13})
        Me.listViewRecords.ContextMenu = Me.contextMenuRecords
        Me.listViewRecords.Dock = System.Windows.Forms.DockStyle.Fill
        Me.listViewRecords.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.listViewRecords.FullRowSelect = True
        Me.listViewRecords.LargeImageList = Me.imageListLarge
        Me.listViewRecords.Location = New System.Drawing.Point(288, 91)
        Me.listViewRecords.MultiSelect = False
        Me.listViewRecords.Name = "listViewRecords"
        Me.listViewRecords.Size = New System.Drawing.Size(404, 501)
        Me.listViewRecords.SmallImageList = Me.imageListSmall
        Me.listViewRecords.TabIndex = 8
        Me.listViewRecords.UseCompatibleStateImageBehavior = False
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Title"
        Me.ColumnHeader2.Width = 175
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Starring"
        Me.ColumnHeader3.Width = 400
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Director"
        Me.ColumnHeader4.Width = 150
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Producer"
        Me.ColumnHeader5.Width = 200
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Category"
        Me.ColumnHeader6.Width = 150
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "Rating"
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "Year"
        '
        'ColumnHeader12
        '
        Me.ColumnHeader12.Text = "Record Date"
        Me.ColumnHeader12.Width = 100
        '
        'ColumnHeader10
        '
        Me.ColumnHeader10.Text = "Type"
        Me.ColumnHeader10.Width = 100
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "Condition"
        '
        'ColumnHeader11
        '
        Me.ColumnHeader11.Text = "Status"
        Me.ColumnHeader11.Width = 100
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Borrower"
        Me.ColumnHeader1.Width = 150
        '
        'ColumnHeader13
        '
        Me.ColumnHeader13.Text = "Record #"
        Me.ColumnHeader13.Width = 0
        '
        'contextMenuRecords
        '
        '
        'imageListSmall
        '
        Me.imageListSmall.ImageStream = CType(resources.GetObject("imageListSmall.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imageListSmall.TransparentColor = System.Drawing.Color.Transparent
        Me.imageListSmall.Images.SetKeyName(0, "")
        Me.imageListSmall.Images.SetKeyName(1, "")
        Me.imageListSmall.Images.SetKeyName(2, "")
        Me.imageListSmall.Images.SetKeyName(3, "")
        '
        'Splitter1
        '
        Me.Splitter1.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Splitter1.Location = New System.Drawing.Point(288, 91)
        Me.Splitter1.MinExtra = 0
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(2, 501)
        Me.Splitter1.TabIndex = 9
        Me.Splitter1.TabStop = False
        '
        'imageListSmall2
        '
        Me.imageListSmall2.ImageStream = CType(resources.GetObject("imageListSmall2.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imageListSmall2.TransparentColor = System.Drawing.Color.Transparent
        Me.imageListSmall2.Images.SetKeyName(0, "")
        Me.imageListSmall2.Images.SetKeyName(1, "")
        Me.imageListSmall2.Images.SetKeyName(2, "")
        Me.imageListSmall2.Images.SetKeyName(3, "")
        Me.imageListSmall2.Images.SetKeyName(4, "")
        Me.imageListSmall2.Images.SetKeyName(5, "")
        Me.imageListSmall2.Images.SetKeyName(6, "")
        Me.imageListSmall2.Images.SetKeyName(7, "")
        Me.imageListSmall2.Images.SetKeyName(8, "")
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.ControlLight
        Me.ClientSize = New System.Drawing.Size(692, 614)
        Me.Controls.Add(Me.Splitter1)
        Me.Controls.Add(Me.listViewRecords)
        Me.Controls.Add(Me.panelProperties)
        Me.Controls.Add(Me.statusBarMain)
        Me.Controls.Add(Me.panelTop)
        Me.Controls.Add(Me.toolBarMain)
        Me.MinimumSize = New System.Drawing.Size(700, 600)
        Me.Name = "frmMain"
        Me.Text = "Movie Records 2.0"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.panelTop.ResumeLayout(False)
        Me.panelProperties.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.tabPageProperties.ResumeLayout(False)
        Me.Panel21.ResumeLayout(False)
        Me.Panel21.PerformLayout()
        Me.Panel16.ResumeLayout(False)
        Me.Panel16.PerformLayout()
        Me.Panel17.ResumeLayout(False)
        Me.Panel17.PerformLayout()
        Me.Panel20.ResumeLayout(False)
        Me.Panel20.PerformLayout()
        Me.Panel13.ResumeLayout(False)
        Me.Panel13.PerformLayout()
        Me.Panel24.ResumeLayout(False)
        Me.Panel24.PerformLayout()
        Me.Panel19.ResumeLayout(False)
        Me.Panel19.PerformLayout()
        Me.Panel8.ResumeLayout(False)
        Me.Panel8.PerformLayout()
        Me.Panel10.ResumeLayout(False)
        Me.Panel10.PerformLayout()
        Me.Panel11.ResumeLayout(False)
        Me.Panel11.PerformLayout()
        Me.Panel6.ResumeLayout(False)
        Me.Panel6.PerformLayout()
        Me.Panel22.ResumeLayout(False)
        Me.Panel22.PerformLayout()
        Me.panelStarring.ResumeLayout(False)
        Me.panelStarring.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel23.ResumeLayout(False)
        CType(Me.pictureBoxThumb, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub formMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        SetViewType(True)
        'LoadData("SELECT * FROM MovieRecords", " ORDER BY Title")

        'pictureBoxThumb.Image = imageListLarge.Images(0)
        pictureBoxThumb.Visible = False

        'RefreshRecords()
        'RefreshRecords()

        cboFilterBy.SelectedIndex = 0
    End Sub

    Sub SetViewType(Optional ByVal listViewLargeIcon As Boolean = False, Optional ByVal listViewList As Boolean = False, Optional ByVal listViewDetails As Boolean = False)
        menuItemViewsDetails.Checked = listViewDetails
        menuItemViewsList.Checked = listViewList
        menuItemViewsLargeIcon.Checked = listViewLargeIcon

        If listViewLargeIcon Then listViewRecords.View = View.LargeIcon
        If listViewList Then listViewRecords.View = View.List
        If listViewDetails Then listViewRecords.View = View.Details
    End Sub

    Public Sub LoadData(ByVal strSELECTString As String, ByVal strORDERString As String)
        Dim movieReader As OleDbDataReader
        Dim movieConnection As OleDbConnection
        Dim movieCommand As OleDbCommand
        Dim tmpItem As ListViewItem

        Dim currCursor As Cursor = Cursor.Current
        Cursor.Current = Cursors.WaitCursor

        ' Create columns for the items and subitems.
        'listView1.Columns.Add("Item Column", -2, HorizontalAlignment.Left)
        lastSELECTString = strSELECTString
        lastORDERString = strORDERString

        'MessageBox.Show(strSELECTString & strORDERString)

        movieConnection = New OleDbConnection(conStr)
        movieCommand = New OleDbCommand(strSELECTString & strORDERString, movieConnection)

        movieConnection.Open()
        movieReader = movieCommand.ExecuteReader()

        'listViewRecords.Visible = False
        listViewRecords.Items.Clear()

        'movieReader.Read()
        'MessageBox.Show(movieReader.Item(5))
        'listViewRecords.Columns.Clear()

        'listviewrecords.Columns.Add(
        Do While (movieReader.Read())

            tmpItem = New ListViewItem(movieReader.GetString(1)) 'Title
            Try
                tmpItem.SubItems.Add(movieReader.GetString(2))          'Starring
            Catch ex As Exception
                tmpItem.SubItems.Add("")
            End Try

            Try
                tmpItem.SubItems.Add(movieReader.GetString(3))          'Director
            Catch ex As Exception
                tmpItem.SubItems.Add("")
            End Try

            Try
                tmpItem.SubItems.Add(movieReader.GetString(4))          'Producer
            Catch ex As Exception
                tmpItem.SubItems.Add("")
            End Try

            Try
                tmpItem.SubItems.Add(movieReader.GetString(5))          'Category
            Catch ex As Exception
                tmpItem.SubItems.Add("")
            End Try

            Try
                tmpItem.SubItems.Add(movieReader.GetString(6))          'Rating
            Catch ex As Exception
                tmpItem.SubItems.Add("")
            End Try

            Try
                tmpItem.SubItems.Add(movieReader.GetString(7))          'Year
            Catch ex As Exception
                tmpItem.SubItems.Add("")
            End Try

            Try
                tmpItem.SubItems.Add(FormatDate(movieReader.GetDateTime(8)))          'Record Date
            Catch ex As Exception
                tmpItem.SubItems.Add("")
            End Try

            Try
                Dim strType As String
                strType = movieReader.GetString(9)
                tmpItem.SubItems.Add(strType)
                If strType.StartsWith("VCD") Then
                    If InStr(strType, "Recopy") <> 0 Then
                        tmpItem.ImageIndex = 2
                    ElseIf InStr(strType, "Original") <> 0 Then
                        tmpItem.ImageIndex = 0
                    Else
                        tmpItem.ImageIndex = 3
                    End If
                ElseIf strType.StartsWith("DVD") Then
                    tmpItem.ImageIndex = 1
                Else
                    tmpItem.ImageIndex = 3
                End If
            Catch ex As Exception
                tmpItem.SubItems.Add("")
            End Try

            If movieReader.IsDBNull(10) Then
                tmpItem.SubItems.Add("")
            Else
                tmpItem.SubItems.Add(movieReader.GetString(10))
            End If

            Try
                tmpItem.SubItems.Add(movieReader.GetString(11))
            Catch ex As Exception
                tmpItem.SubItems.Add("")
            End Try

            Try
                tmpItem.SubItems.Add(movieReader.GetString(12))
            Catch ex As Exception
                tmpItem.SubItems.Add("")
            End Try

            tmpItem.SubItems.Add(movieReader.GetInt32(0))
            listViewRecords.Items.Add(tmpItem)

            If NewArrivals <> 0 Then
                If listViewRecords.Items.Count = NewArrivals Then Exit Do
            End If
        Loop

        '********* Automatically set the first record's property
        'If listViewRecords.Items.Count > 0 Then
        'currentRecord = CLng(listViewRecords.Items(0).SubItems(12).Text)
        'SetProperties()
        'End If
        statusBarMain.Text = listViewRecords.Items.Count.ToString & " Record(s)"
        'listViewRecords.Visible = True
        'MessageBox.Show(movieReader.FieldCount.ToString)

        ' Always call Close when done reading.
        If Not (movieReader Is Nothing) Then
            movieReader.Close()
        End If

        ' Close the connection when done with it.
        If (movieConnection.State = ConnectionState.Open) Then
            movieConnection.Close()
        End If

        Cursor.Current = currCursor
    End Sub

    Private Sub menuItemViewsLargeIcon_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles menuItemViewsLargeIcon.Click
        SetViewType(True)
    End Sub

    Private Sub menuItemViewsList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles menuItemViewsList.Click
        SetViewType(, True)
    End Sub

    Private Sub menuItemViewsDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles menuItemViewsDetails.Click
        SetViewType(, , True)
    End Sub

    Private Sub toolBarMain_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles toolBarMain.ButtonClick
        Select Case toolBarMain.Buttons.IndexOf(e.Button)
            Case 0
                NewRecord()
            Case 1
                If ARecordSelected Then
                    EditRecord()
                End If
            Case 2
                If ARecordSelected Then
                    DeleteRecord()
                End If
            Case 4
                RefreshRecords()
            Case 7
                If listViewRecords.View = View.LargeIcon Then
                    SetViewType(, True)
                ElseIf listViewRecords.View = View.List Then
                    SetViewType(, , True)
                Else
                    SetViewType(True)
                End If
        End Select
    End Sub

    Private Sub listViewRecords_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles listViewRecords.SelectedIndexChanged
        If listViewRecords.SelectedItems.Count <> 0 Then
            SetProperties()
        Else
            ClearProperties()
        End If
    End Sub

    Sub SetProperties()
        Dim movieReader As OleDbDataReader
        Dim movieConnection As OleDbConnection
        Dim movieCommand As OleDbCommand
        Dim tmpUnreturned As Boolean = False

        Try
            currentRecord = CLng(listViewRecords.SelectedItems(0).SubItems(12).Text)
            'statusBarMain.Text = listViewRecords.SelectedItems.Count.ToString & " Item(s) selected."
        Catch ex As Exception
        End Try
        ' Create columns for the items and subitems.
        'listView1.Columns.Add("Item Column", -2, HorizontalAlignment.Left)

        movieConnection = New OleDbConnection(conStr)
        movieCommand = New OleDbCommand("SELECT * FROM MovieRecords WHERE RecordNo=" & currentRecord.ToString, movieConnection)

        movieConnection.Open()
        movieReader = movieCommand.ExecuteReader()

        movieReader.Read()

        lblTitle.Text = movieReader.GetString(1)

        Try
            txtStarring.Text = movieReader.GetString(2)
        Catch ex As Exception
            txtStarring.Text = ""
        End Try

        Try
            txtDirector.Text = movieReader.GetString(3)
        Catch ex As Exception
            txtDirector.Text = ""
        End Try

        Try
            txtProducer.Text = movieReader.GetString(4)
        Catch ex As Exception
            txtProducer.Text = ""
        End Try

        Try
            txtCategory.Text = movieReader.GetString(5)
        Catch ex As Exception
            txtCategory.Text = ""
        End Try

        Try
            txtRating.Text = movieReader.GetString(6)
        Catch ex As Exception
            txtRating.Text = ""
        End Try

        Try
            txtYear.Text = movieReader.GetString(7)
        Catch ex As Exception
            txtYear.Text = ""
        End Try

        Try
            txtCondition.Text = movieReader.GetString(10)
        Catch ex As Exception
            txtCondition.Text = ""
        End Try

        Try
            'txtType.Text = movieReader.GetString(9)
            Dim strType As String
            strType = movieReader.GetString(9)
            txtType.Text = strType
            If strType.StartsWith("VCD") Then
                If InStr(strType, "Recopy") <> 0 Then
                    'tmpItem.ImageIndex = 2
                    pictureBoxThumb.Image = imageListLarge.Images(2)
                ElseIf InStr(strType, "Original") <> 0 Then
                    'tmpItem.ImageIndex = 0
                    pictureBoxThumb.Image = imageListLarge.Images(0)
                Else
                    'tmpItem.ImageIndex = 3
                    pictureBoxThumb.Image = imageListLarge.Images(3)
                End If
            ElseIf strType.StartsWith("DVD") Then
                pictureBoxThumb.Image = imageListLarge.Images(1)
            Else
                pictureBoxThumb.Image = imageListLarge.Images(3)
            End If
            'Else
            'tmpItem.ImageIndex = 1
            'pictureBoxThumb.Image = imageListLarge.Images(1)
            'End If
        Catch ex As Exception
            txtType.Text = ""
        End Try

        Try
            txtRecordDate.Text = FormatDate(movieReader.GetDateTime(8))
        Catch ex As Exception
            txtRecordDate.Text = ""
        End Try

        If movieReader.IsDBNull(11) Then
            txtStatus.Text = ""
        Else
            Dim tmpStrStatus As String = movieReader.GetString(11)
            txtStatus.Text = tmpStrStatus
            If Trim(tmpStrStatus) = "Unreturned" Then
                tmpUnreturned = True
                txtStatus.ForeColor = Color.Red
            Else
                tmpUnreturned = False
                txtStatus.ForeColor = Color.Black
            End If
        End If

        Try
            txtBorrower.Text = movieReader.GetString(12)
        Catch ex As Exception
            txtBorrower.Text = ""
        End Try

        If movieReader.IsDBNull(13) Then
            txtDateBorrowed.Text = ""
        Else
            txtDateBorrowed.Text = FormatDate(movieReader.GetDateTime(13))
        End If

        If movieReader.IsDBNull(14) Then
            txtDateReturned.Text = ""
        Else
            If tmpUnreturned Then
                txtDateReturned.Text = ""
            Else
                txtDateReturned.Text = FormatDate(movieReader.GetDateTime(14))
            End If
        End If

        pictureBoxThumb.Visible = True
        ARecordSelected = True
        ' Always call Close when done reading.
        If Not (movieReader Is Nothing) Then
            movieReader.Close()
        End If

        ' Close the connection when done with it.
        If (movieConnection.State = ConnectionState.Open) Then
            movieConnection.Close()
        End If
    End Sub

    Private Sub cboFilterBy_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboFilterBy.SelectedValueChanged
        Dim movieReader As OleDbDataReader
        Dim movieConnection As OleDbConnection
        Dim movieCommand As OleDbCommand

        Select Case cboFilterBy.Text
            Case "Title"
                'cboFilter.Text = ""
            Case "Starring"
                'cboFilter.Text = ""
            Case "Director"
                'cboFilter.Text = ""
            Case "Condition"
                movieConnection = New OleDbConnection(conStr)
                movieCommand = New OleDbCommand("SELECT DISTINCT Condition FROM MovieRecords", movieConnection)

                movieConnection.Open()
                movieReader = movieCommand.ExecuteReader()

                cboFilter.Items.Clear()
                Do While (movieReader.Read())
                    If movieReader.GetString(0) <> "" Then
                        cboFilter.Items.Add(movieReader.GetString(0))
                    End If
                Loop
                cboFilter.Text = ""

                If Not (movieReader Is Nothing) Then
                    movieReader.Close()
                End If

                movieConnection.Close()

            Case "Category"
                movieConnection = New OleDbConnection(conStr)
                movieCommand = New OleDbCommand("SELECT DISTINCT Category FROM MovieRecords", movieConnection)

                movieConnection.Open()
                movieReader = movieCommand.ExecuteReader()

                cboFilter.Items.Clear()
                Do While (movieReader.Read())
                    If movieReader.GetString(0) <> "" Then
                        cboFilter.Items.Add(movieReader.GetString(0))
                    End If
                Loop
                cboFilter.Text = ""

                If Not (movieReader Is Nothing) Then
                    movieReader.Close()
                End If

                movieConnection.Close()

            Case "Status"
                movieConnection = New OleDbConnection(conStr)
                movieCommand = New OleDbCommand("SELECT DISTINCT Status FROM MovieRecords", movieConnection)

                movieConnection.Open()
                movieReader = movieCommand.ExecuteReader()

                cboFilter.Items.Clear()
                Do While (movieReader.Read())
                    'If Not movieReader.IsDBNull(0) Then
                    'If movieReader.GetString(0) <> "" Then
                    Try
                        cboFilter.Items.Add(movieReader.GetString(0))
                    Catch ex As Exception
                        cboFilter.Items.Add("")
                    End Try

                    'End If
                    'End If
                Loop
                cboFilter.Text = ""

                If Not (movieReader Is Nothing) Then
                    movieReader.Close()
                End If

                movieConnection.Close()

            Case "Borrower"
                movieConnection = New OleDbConnection(conStr)
                movieCommand = New OleDbCommand("SELECT DISTINCT Borrower FROM MovieRecords", movieConnection)

                movieConnection.Open()
                movieReader = movieCommand.ExecuteReader()

                cboFilter.Items.Clear()
                Do While (movieReader.Read())
                    If Not movieReader.IsDBNull(0) Then
                        If movieReader.GetString(0) <> "" Then
                            cboFilter.Items.Add(movieReader.GetString(0))
                        End If
                    End If
                Loop
                cboFilter.Text = ""

                If Not (movieReader Is Nothing) Then
                    movieReader.Close()
                End If

                movieConnection.Close()
            Case "New Arrivals"
                cboFilter.Items.Clear()
                cboFilter.Items.Add("10")
                cboFilter.Items.Add("20")
                cboFilter.Items.Add("50")
                cboFilter.SelectedIndex = 0
        End Select

        cboFilter.Select()
    End Sub

    Private Sub cmdFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFilter.Click

        ClearProperties()
        FilterRecords()

        'If Not String.IsNullOrEmpty(cboFilterBy.Text) And Not String.IsNullOrEmpty(cboFilter.Text) Then

        'End If
    End Sub

    Sub FilterRecords()

        If Not String.IsNullOrEmpty(cboFilterBy.Text) And Not String.IsNullOrEmpty(cboFilter.Text) Then
            Dim tmpSELECTString As String = String.Empty, tmpORDERString As String = " ORDER BY Title"
            If cboFilterBy.Text = "New Arrivals" Then
                If IsNumeric(cboFilter.Text) Then
                    NewArrivals = CLng(cboFilter.Text)
                    'LoadData("SELECT * FROM MovieRecords ORDER BY RecordDate DESC")
                    tmpSELECTString = "SELECT * FROM MovieRecords"
                    tmpORDERString = " ORDER BY RecordDate DESC"
                    'NewArrivals = 0
                End If
            Else
                NewArrivals = 0
                Select Case cboFilterBy.Text
                    Case "Title"
                        'LoadData("SELECT * FROM MovieRecords WHERE Title LIKE '%" & cboFilter.Text & "%' ORDER BY Title")
                        tmpSELECTString = "SELECT * FROM MovieRecords WHERE Title LIKE '%" & cboFilter.Text & "%'"
                    Case "Cast"
                        'LoadData("SELECT * FROM MovieRecords WHERE Starring LIKE '%" & cboFilter.Text & "%' ORDER BY Title")
                        tmpSELECTString = "SELECT * FROM MovieRecords WHERE Starring LIKE '%" & cboFilter.Text & "%'"
                    Case "Director"
                        'LoadData("SELECT * FROM MovieRecords WHERE Director LIKE '%" & cboFilter.Text & "%' ORDER BY Title")
                        tmpSELECTString = "SELECT * FROM MovieRecords WHERE Director LIKE '%" & cboFilter.Text & "%'"
                    Case "Condition"
                        'LoadData("SELECT * FROM MovieRecords WHERE Condition LIKE '%" & cboFilter.Text & "%' ORDER BY Title")
                        tmpSELECTString = "SELECT * FROM MovieRecords WHERE Condition LIKE '%" & cboFilter.Text & "%'"
                    Case "Category"
                        'LoadData("SELECT * FROM MovieRecords WHERE Category LIKE '%" & cboFilter.Text & "%' ORDER BY Title")
                        tmpSELECTString = "SELECT * FROM MovieRecords WHERE Category LIKE '%" & cboFilter.Text & "%'"
                    Case "Status"
                        'LoadData("SELECT * FROM MovieRecords WHERE Status='" & cboFilter.Text & "' ORDER BY Borrower")
                        tmpSELECTString = "SELECT * FROM MovieRecords WHERE Status='" & cboFilter.Text & "'"
                        tmpORDERString = " ORDER BY Borrower"
                    Case "Borrower"
                        'LoadData("SELECT * FROM MovieRecords WHERE Borrower LIKE '%" & cboFilter.Text & "%' ORDER BY Title")
                        tmpSELECTString = "SELECT * FROM MovieRecords WHERE Borrower LIKE '%" & cboFilter.Text & "%'"
                End Select
            End If

            LoadData(tmpSELECTString, tmpORDERString)

        End If
    End Sub

    Private Sub cboFilter_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cboFilter.KeyPress
        Select Case e.KeyChar
            Case vbCr
                FilterRecords()
        End Select
    End Sub

    Private Sub contextMenuRecords_Popup(ByVal sender As Object, ByVal e As System.EventArgs) Handles contextMenuRecords.Popup
        Dim tmpMenuItem As MenuItem

        contextMenuRecords.MenuItems.Clear()

        tmpMenuItem = New MenuItem("&New Record", New EventHandler(AddressOf mnuItemNewRecord_Click))
        contextMenuRecords.MenuItems.Add(tmpMenuItem)


        If listViewRecords.SelectedItems.Count <> 0 Then
            tmpMenuItem = New MenuItem("-")
            contextMenuRecords.MenuItems.Add(tmpMenuItem)

            tmpMenuItem = New MenuItem("&Edit", New EventHandler(AddressOf mnuItemEditRecord_Click))
            tmpMenuItem.DefaultItem = True
            contextMenuRecords.MenuItems.Add(tmpMenuItem)

            tmpMenuItem = New MenuItem("&Delete", New EventHandler(AddressOf mnuItemDeleteRecord_Click))
            contextMenuRecords.MenuItems.Add(tmpMenuItem)

            tmpMenuItem = New MenuItem("-")
            contextMenuRecords.MenuItems.Add(tmpMenuItem)

            If listViewRecords.SelectedItems(0).SubItems(10).Text = "Unreturned" Then
                tmpMenuItem = New MenuItem("&Return", New EventHandler(AddressOf mnuItemAvailabilityIn_Click))
                contextMenuRecords.MenuItems.Add(tmpMenuItem)
            Else
                tmpMenuItem = New MenuItem("&Borrow", New EventHandler(AddressOf mnuItemAvailabilityOut_Click))
                contextMenuRecords.MenuItems.Add(tmpMenuItem)
            End If

            'tmpMenuItem = New MenuItem("&Set Status...", New EventHandler(AddressOf mnuItemAvailability_Click))
            'contextMenuRecords.MenuItems.Add(tmpMenuItem)
        End If

        tmpMenuItem = New MenuItem("-")
        contextMenuRecords.MenuItems.Add(tmpMenuItem)

        tmpMenuItem = New MenuItem("&Refresh", New EventHandler(AddressOf mnuItemRefresh_Click))
        contextMenuRecords.MenuItems.Add(tmpMenuItem)

        'tmpMenuItem = New MenuItem("-")
        'contextMenuRecords.MenuItems.Add(tmpMenuItem)

        tmpMenuItem = New MenuItem("&View")
        tmpMenuItem.MergeMenu(contextMenuViews)
        contextMenuRecords.MenuItems.Add(tmpMenuItem)

    End Sub

    Private Sub mnuItemEditRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        EditRecord()
    End Sub

    Private Sub mnuItemDeleteRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        DeleteRecord()
    End Sub

    Private Sub mnuItemNewRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        NewRecord()
    End Sub

    'Private Sub mnuItemAvailability_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    'SetAvailabilityOptions()
    'End Sub

    Private Sub mnuItemAvailabilityOut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        SetAvailabilityOut()
    End Sub

    Private Sub mnuItemAvailabilityIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        SetAvailabilityIn()
    End Sub

    Private Sub mnuItemRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        RefreshRecords()
    End Sub

    'Sub SetAvailabilityOptions()
    'MessageBox.Show(listViewRecords.SelectedItems(0).SubItems(12).Text)
    'End Sub

    Sub SetAvailabilityIn()
        ExecuteCommand("UPDATE MovieRecords SET Status='Returned', DateReturned=#" & FormatDateTime(Now.Date) & "# WHERE RecordNo=" & currentRecord.ToString)
        listViewRecords.SelectedItems(0).SubItems(10).Text = "Returned"
        SetProperties()
    End Sub

    Sub SetAvailabilityOut()
        Dim fBorrower As New frmBorrower()
        'MessageBox.Show(listViewRecords.SelectedItems(0).SubItems(12).Text)
        If fBorrower.ShowDialog(Me) = DialogResult.OK Then
            listViewRecords.SelectedItems(0).SubItems(10).Text = "Unreturned"
            listViewRecords.SelectedItems(0).SubItems(11).Text = lastBorrower
            SetProperties()
        End If
    End Sub

    Private Sub EditRecord()
        Dim fEdit As New frmAddEdit()
        'fEdit.ShowDialog(Me)
        If fEdit.ShowDialog(Me) = DialogResult.OK Then
            Dim movieReader As OleDbDataReader
            Dim movieConnection As OleDbConnection
            Dim movieCommand As OleDbCommand
            Dim tmpItem As ListViewItem

            movieConnection = New OleDbConnection(conStr)
            movieCommand = New OleDbCommand("SELECT * FROM MovieRecords WHERE RecordNo=" & currentRecord.ToString, movieConnection)

            movieConnection.Open()
            movieReader = movieCommand.ExecuteReader()

            movieReader.Read()

            With listViewRecords.SelectedItems(0) 'listViewRecords.Items(SelectedItemIndex)
                .Text = movieReader.GetString(1)
                'listviewrecords.ListViewItemCollectio
                '.Text = "Daniel Pogi"
                'tmpItem = New ListViewItem(movieReader.GetString(1), 0) 'Title
                Try
                    'tmpItem.SubItems.Add(movieReader.GetString(2))          'Starring
                    .SubItems(1).Text = movieReader.GetString(2)
                Catch ex As Exception
                    .SubItems(1).Text = ""
                End Try

                Try
                    'tmpItem.SubItems.Add(movieReader.GetString(2))          'Starring
                    .SubItems(2).Text = movieReader.GetString(3)
                Catch ex As Exception
                    .SubItems(2).Text = ""
                End Try

                Try
                    'tmpItem.SubItems.Add(movieReader.GetString(2))          'Starring
                    .SubItems(3).Text = movieReader.GetString(4)
                Catch ex As Exception
                    .SubItems(3).Text = ""
                End Try

                Try
                    'tmpItem.SubItems.Add(movieReader.GetString(2))          'Starring
                    .SubItems(4).Text = movieReader.GetString(5)
                Catch ex As Exception
                    .SubItems(4).Text = ""
                End Try

                Try
                    'tmpItem.SubItems.Add(movieReader.GetString(2))          'Starring
                    .SubItems(5).Text = movieReader.GetString(6)
                Catch ex As Exception
                    .SubItems(5).Text = ""
                End Try

                Try
                    'tmpItem.SubItems.Add(movieReader.GetString(2))          'Starring
                    .SubItems(6).Text = movieReader.GetString(7)
                Catch ex As Exception
                    .SubItems(6).Text = ""
                End Try

                Try
                    'tmpItem.SubItems.Add(FormatDate(movieReader.GetDateTime(8)))          'Record Date
                    .SubItems(7).Text = FormatDate(movieReader.GetDateTime(8))
                Catch ex As Exception
                    'tmpItem.SubItems.Add("")
                    .SubItems(7).Text = ""
                End Try

                Try
                    'tmpItem.SubItems.Add(movieReader.GetString(9))
                    .SubItems(8).Text = movieReader.GetString(9)
                Catch ex As Exception
                    'tmpItem.SubItems.Add("")
                    .SubItems(8).Text = ""
                End Try

                Try
                    'tmpItem.SubItems.Add(movieReader.GetString(2))          'Starring
                    .SubItems(9).Text = movieReader.GetString(10)
                Catch ex As Exception
                    .SubItems(9).Text = ""
                End Try

                Try
                    'tmpItem.SubItems.Add(movieReader.GetString(2))          'Starring
                    .SubItems(10).Text = movieReader.GetString(11)
                Catch ex As Exception
                    .SubItems(10).Text = ""
                End Try

                Try
                    'tmpItem.SubItems.Add(movieReader.GetString(2))          'Starring
                    .SubItems(11).Text = movieReader.GetString(12)
                Catch ex As Exception
                    .SubItems(11).Text = ""
                End Try


                'tmpItem.SubItems.Add(movieReader.GetInt32(0).ToString)
                .SubItems(12).Text = movieReader.GetInt32(0).ToString
                'listViewRecords.Items(SelectedItemIndex) = tmpItem
            End With
            If Not (movieReader Is Nothing) Then
                movieReader.Close()
            End If

            If (movieConnection.State = ConnectionState.Open) Then
                movieConnection.Close()
            End If

            SetProperties()
        End If
        fEdit.Dispose()
    End Sub

    Sub NewRecord()
        Dim fAdd As New frmAddEdit()
        AddNewRecord = True
        fAdd.ShowDialog(Me)
        fAdd.Dispose()
        AddNewRecord = False
    End Sub

    Sub DeleteRecord()
        Dim tmpMessageResult As DialogResult

        tmpMessageResult = MessageBox.Show("Are you sure you want to delete " & Chr(34) & listViewRecords.SelectedItems(0).SubItems(0).Text & _
            Chr(34) & "?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If tmpMessageResult = DialogResult.Yes Then
            ExecuteCommand("DELETE FROM MovieRecords WHERE RecordNo=" & currentRecord.ToString)
            'If cboFilterBy.Text = "" Then
            'RefreshRecords()
            'Else
            'FilterRecords()
            'End If
            listViewRecords.FocusedItem.Remove()
        End If
    End Sub

    Sub RefreshRecords()
        NewArrivals = 0
        LoadData("SELECT * FROM MovieRecords", " ORDER BY Title")
        'LoadData(lastSELECTString, lastORDERString)
    End Sub

    Sub ClearProperties()
        lblTitle.Text = ""
        txtStarring.Text = ""
        txtDirector.Text = ""
        txtProducer.Text = ""
        txtCategory.Text = ""
        txtRating.Text = ""
        txtYear.Text = ""
        txtCondition.Text = ""
        txtType.Text = ""
        txtRecordDate.Text = ""
        txtStatus.Text = ""
        txtBorrower.Text = ""
        txtDateBorrowed.Text = ""
        txtDateReturned.Text = ""

        pictureBoxThumb.Visible = False
        ARecordSelected = False
    End Sub

    Private Sub listViewRecords_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles listViewRecords.ColumnClick
        Dim sortString As String
        If NewArrivals <> 0 Then Exit Sub
        Select Case e.Column
            Case RecordColumns.Title
                sortString = " ORDER BY Title"
            Case RecordColumns.Starring
                sortString = " ORDER BY Starring"
            Case RecordColumns.Director
                sortString = " ORDER BY Director"
            Case RecordColumns.Producer
                sortString = " ORDER BY Producer"
            Case RecordColumns.Category
                sortString = " ORDER BY Category"
            Case RecordColumns.Rating
                sortString = " ORDER BY Rating"
            Case RecordColumns.Year
                sortString = " ORDER BY [Year]"
            Case RecordColumns.RecordDate
                sortString = " ORDER BY RecordDate"
            Case RecordColumns.Type
                sortString = " ORDER BY Type"
            Case RecordColumns.Condition
                sortString = " ORDER BY Condition"
            Case RecordColumns.Status
                sortString = " ORDER BY Status"
            Case RecordColumns.Borrower
                sortString = " ORDER BY Borrower"
        End Select
        If SortedColumn = e.Column Then
            SortedColumn = -1
            ColumnSortedAsc = False
        Else
            SortedColumn = e.Column
            ColumnSortedAsc = True
        End If

        If ColumnSortedAsc = False Then
            sortString = sortString & " DESC"
        End If

        'sortString = "SELECT * FROM MovieRecords" & sortString
        'LoadData("SELECT * FROM MovieRecords", sortString)
        LoadData(lastSELECTString, sortString)
    End Sub

    Private Sub listViewRecords_ItemActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles listViewRecords.ItemActivate
        'SelectedItemIndex = listViewRecords.SelectedItems(0).Index
        EditRecord()
    End Sub

    Private Sub frmMain_Validated(sender As System.Object, e As System.EventArgs) Handles MyBase.Validated
        MessageBox.Show("validated")
    End Sub

End Class
