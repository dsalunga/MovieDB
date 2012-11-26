Imports System.Data.OleDb

Public Class frmAddEdit
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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents txtTitle As System.Windows.Forms.TextBox
    Friend WithEvents lblStarring As System.Windows.Forms.Label
    Friend WithEvents txtStarring As System.Windows.Forms.TextBox
    Friend WithEvents lblDirector As System.Windows.Forms.Label
    Friend WithEvents txtDirector As System.Windows.Forms.TextBox
    Friend WithEvents lblProducer As System.Windows.Forms.Label
    Friend WithEvents txtProducer As System.Windows.Forms.TextBox
    Friend WithEvents lblCategory As System.Windows.Forms.Label
    Friend WithEvents lblRating As System.Windows.Forms.Label
    Friend WithEvents lblYear As System.Windows.Forms.Label
    Friend WithEvents lblRecordDate As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents lblCondition As System.Windows.Forms.Label
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents cboCategory As System.Windows.Forms.ComboBox
    Friend WithEvents dtpRecordDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents cboRating As System.Windows.Forms.ComboBox
    Friend WithEvents cboCondition As System.Windows.Forms.ComboBox
    Friend WithEvents cboType As System.Windows.Forms.ComboBox
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents cboFileQuality As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtDiscNo As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cboSubtitles As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboFileType As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboFileFormat As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboMyRating As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtYear As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cboCategory = New System.Windows.Forms.ComboBox()
        Me.dtpRecordDate = New System.Windows.Forms.DateTimePicker()
        Me.cboRating = New System.Windows.Forms.ComboBox()
        Me.lblRecordDate = New System.Windows.Forms.Label()
        Me.lblYear = New System.Windows.Forms.Label()
        Me.txtYear = New System.Windows.Forms.TextBox()
        Me.lblRating = New System.Windows.Forms.Label()
        Me.lblCategory = New System.Windows.Forms.Label()
        Me.lblProducer = New System.Windows.Forms.Label()
        Me.txtProducer = New System.Windows.Forms.TextBox()
        Me.lblDirector = New System.Windows.Forms.Label()
        Me.txtDirector = New System.Windows.Forms.TextBox()
        Me.lblStarring = New System.Windows.Forms.Label()
        Me.txtStarring = New System.Windows.Forms.TextBox()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.txtTitle = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.txtDiscNo = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cboCondition = New System.Windows.Forms.ComboBox()
        Me.cboType = New System.Windows.Forms.ComboBox()
        Me.lblCondition = New System.Windows.Forms.Label()
        Me.lblType = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.cboSubtitles = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cboFileType = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cboFileQuality = New System.Windows.Forms.ComboBox()
        Me.cboFileFormat = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cboMyRating = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdOK
        '
        Me.cmdOK.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdOK.Location = New System.Drawing.Point(322, 488)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(75, 23)
        Me.cmdOK.TabIndex = 10
        Me.cmdOK.Text = "&Save"
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdCancel.Location = New System.Drawing.Point(410, 488)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 11
        Me.cmdCancel.Text = "&Cancel"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cboMyRating)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.cboCategory)
        Me.GroupBox1.Controls.Add(Me.dtpRecordDate)
        Me.GroupBox1.Controls.Add(Me.cboRating)
        Me.GroupBox1.Controls.Add(Me.lblRecordDate)
        Me.GroupBox1.Controls.Add(Me.lblYear)
        Me.GroupBox1.Controls.Add(Me.txtYear)
        Me.GroupBox1.Controls.Add(Me.lblRating)
        Me.GroupBox1.Controls.Add(Me.lblCategory)
        Me.GroupBox1.Controls.Add(Me.lblProducer)
        Me.GroupBox1.Controls.Add(Me.txtProducer)
        Me.GroupBox1.Controls.Add(Me.lblDirector)
        Me.GroupBox1.Controls.Add(Me.txtDirector)
        Me.GroupBox1.Controls.Add(Me.lblStarring)
        Me.GroupBox1.Controls.Add(Me.txtStarring)
        Me.GroupBox1.Controls.Add(Me.lblTitle)
        Me.GroupBox1.Controls.Add(Me.txtTitle)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(473, 239)
        Me.GroupBox1.TabIndex = 86
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Movie Record"
        '
        'cboCategory
        '
        Me.cboCategory.DropDownWidth = 180
        Me.cboCategory.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboCategory.Items.AddRange(New Object() {"2D Animation", "3D Animation", "Action", "Adventure", "Comedy", "Crime", "Documentary", "Drama", "Horror", "Musical", "Romance", "Suspense", "Tagalog", "Thriller"})
        Me.cboCategory.Location = New System.Drawing.Point(322, 113)
        Me.cboCategory.Name = "cboCategory"
        Me.cboCategory.Size = New System.Drawing.Size(128, 21)
        Me.cboCategory.TabIndex = 5
        '
        'dtpRecordDate
        '
        Me.dtpRecordDate.CustomFormat = "MMMM dd, yyyy"
        Me.dtpRecordDate.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpRecordDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpRecordDate.Location = New System.Drawing.Point(322, 177)
        Me.dtpRecordDate.Name = "dtpRecordDate"
        Me.dtpRecordDate.Size = New System.Drawing.Size(128, 21)
        Me.dtpRecordDate.TabIndex = 7
        '
        'cboRating
        '
        Me.cboRating.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboRating.Items.AddRange(New Object() {"PG", "PG-7", "PG-13", "PG-15", "R-17", "R-18"})
        Me.cboRating.Location = New System.Drawing.Point(322, 145)
        Me.cboRating.Name = "cboRating"
        Me.cboRating.Size = New System.Drawing.Size(128, 21)
        Me.cboRating.TabIndex = 6
        '
        'lblRecordDate
        '
        Me.lblRecordDate.AutoSize = True
        Me.lblRecordDate.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRecordDate.Location = New System.Drawing.Point(250, 181)
        Me.lblRecordDate.Name = "lblRecordDate"
        Me.lblRecordDate.Size = New System.Drawing.Size(67, 13)
        Me.lblRecordDate.TabIndex = 90
        Me.lblRecordDate.Text = "Record Date"
        '
        'lblYear
        '
        Me.lblYear.AutoSize = True
        Me.lblYear.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblYear.Location = New System.Drawing.Point(16, 180)
        Me.lblYear.Name = "lblYear"
        Me.lblYear.Size = New System.Drawing.Size(83, 13)
        Me.lblYear.TabIndex = 88
        Me.lblYear.Text = "Year of Release"
        '
        'txtYear
        '
        Me.txtYear.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtYear.Location = New System.Drawing.Point(106, 177)
        Me.txtYear.Name = "txtYear"
        Me.txtYear.Size = New System.Drawing.Size(79, 21)
        Me.txtYear.TabIndex = 4
        '
        'lblRating
        '
        Me.lblRating.AutoSize = True
        Me.lblRating.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRating.Location = New System.Drawing.Point(250, 149)
        Me.lblRating.Name = "lblRating"
        Me.lblRating.Size = New System.Drawing.Size(38, 13)
        Me.lblRating.TabIndex = 86
        Me.lblRating.Text = "Rating"
        '
        'lblCategory
        '
        Me.lblCategory.AutoSize = True
        Me.lblCategory.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCategory.Location = New System.Drawing.Point(250, 117)
        Me.lblCategory.Name = "lblCategory"
        Me.lblCategory.Size = New System.Drawing.Size(52, 13)
        Me.lblCategory.TabIndex = 84
        Me.lblCategory.Text = "Category"
        '
        'lblProducer
        '
        Me.lblProducer.AutoSize = True
        Me.lblProducer.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProducer.Location = New System.Drawing.Point(16, 148)
        Me.lblProducer.Name = "lblProducer"
        Me.lblProducer.Size = New System.Drawing.Size(50, 13)
        Me.lblProducer.TabIndex = 82
        Me.lblProducer.Text = "Producer"
        '
        'txtProducer
        '
        Me.txtProducer.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtProducer.Location = New System.Drawing.Point(106, 145)
        Me.txtProducer.Name = "txtProducer"
        Me.txtProducer.Size = New System.Drawing.Size(128, 21)
        Me.txtProducer.TabIndex = 3
        '
        'lblDirector
        '
        Me.lblDirector.AutoSize = True
        Me.lblDirector.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDirector.Location = New System.Drawing.Point(16, 116)
        Me.lblDirector.Name = "lblDirector"
        Me.lblDirector.Size = New System.Drawing.Size(45, 13)
        Me.lblDirector.TabIndex = 80
        Me.lblDirector.Text = "Director"
        '
        'txtDirector
        '
        Me.txtDirector.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDirector.Location = New System.Drawing.Point(106, 113)
        Me.txtDirector.Name = "txtDirector"
        Me.txtDirector.Size = New System.Drawing.Size(128, 21)
        Me.txtDirector.TabIndex = 2
        '
        'lblStarring
        '
        Me.lblStarring.AutoSize = True
        Me.lblStarring.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStarring.Location = New System.Drawing.Point(16, 60)
        Me.lblStarring.Name = "lblStarring"
        Me.lblStarring.Size = New System.Drawing.Size(29, 13)
        Me.lblStarring.TabIndex = 78
        Me.lblStarring.Text = "Cast"
        '
        'txtStarring
        '
        Me.txtStarring.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStarring.Location = New System.Drawing.Point(106, 57)
        Me.txtStarring.Multiline = True
        Me.txtStarring.Name = "txtStarring"
        Me.txtStarring.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtStarring.Size = New System.Drawing.Size(344, 48)
        Me.txtStarring.TabIndex = 1
        '
        'lblTitle
        '
        Me.lblTitle.AutoSize = True
        Me.lblTitle.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.Location = New System.Drawing.Point(16, 28)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(58, 13)
        Me.lblTitle.TabIndex = 76
        Me.lblTitle.Text = "Movie Title"
        '
        'txtTitle
        '
        Me.txtTitle.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTitle.Location = New System.Drawing.Point(106, 25)
        Me.txtTitle.Name = "txtTitle"
        Me.txtTitle.Size = New System.Drawing.Size(344, 21)
        Me.txtTitle.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtDiscNo)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.cboCondition)
        Me.GroupBox2.Controls.Add(Me.cboType)
        Me.GroupBox2.Controls.Add(Me.lblCondition)
        Me.GroupBox2.Controls.Add(Me.lblType)
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(12, 377)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(473, 92)
        Me.GroupBox2.TabIndex = 87
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Media"
        '
        'txtDiscNo
        '
        Me.txtDiscNo.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDiscNo.Location = New System.Drawing.Point(106, 51)
        Me.txtDiscNo.Name = "txtDiscNo"
        Me.txtDiscNo.Size = New System.Drawing.Size(79, 21)
        Me.txtDiscNo.TabIndex = 82
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(16, 51)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(37, 13)
        Me.Label5.TabIndex = 81
        Me.Label5.Text = "Disc #"
        '
        'cboCondition
        '
        Me.cboCondition.DropDownWidth = 200
        Me.cboCondition.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboCondition.Location = New System.Drawing.Point(322, 24)
        Me.cboCondition.Name = "cboCondition"
        Me.cboCondition.Size = New System.Drawing.Size(128, 21)
        Me.cboCondition.TabIndex = 9
        '
        'cboType
        '
        Me.cboType.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboType.Location = New System.Drawing.Point(106, 24)
        Me.cboType.Name = "cboType"
        Me.cboType.Size = New System.Drawing.Size(128, 21)
        Me.cboType.TabIndex = 8
        '
        'lblCondition
        '
        Me.lblCondition.AutoSize = True
        Me.lblCondition.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCondition.Location = New System.Drawing.Point(250, 29)
        Me.lblCondition.Name = "lblCondition"
        Me.lblCondition.Size = New System.Drawing.Size(52, 13)
        Me.lblCondition.TabIndex = 77
        Me.lblCondition.Text = "Condition"
        '
        'lblType
        '
        Me.lblType.AutoSize = True
        Me.lblType.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblType.Location = New System.Drawing.Point(16, 24)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(31, 13)
        Me.lblType.TabIndex = 79
        Me.lblType.Text = "Type"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.cboSubtitles)
        Me.GroupBox3.Controls.Add(Me.Label4)
        Me.GroupBox3.Controls.Add(Me.cboFileType)
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.cboFileQuality)
        Me.GroupBox3.Controls.Add(Me.cboFileFormat)
        Me.GroupBox3.Controls.Add(Me.Label1)
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(12, 267)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(473, 95)
        Me.GroupBox3.TabIndex = 88
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Movie File"
        '
        'cboSubtitles
        '
        Me.cboSubtitles.DropDownWidth = 200
        Me.cboSubtitles.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboSubtitles.Items.AddRange(New Object() {"English", "French", "Japanese", "German", "Italian", "Portuguese", "Spanish", "Tagalog", "Other", "None"})
        Me.cboSubtitles.Location = New System.Drawing.Point(322, 57)
        Me.cboSubtitles.Name = "cboSubtitles"
        Me.cboSubtitles.Size = New System.Drawing.Size(128, 21)
        Me.cboSubtitles.TabIndex = 82
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(250, 62)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(48, 13)
        Me.Label4.TabIndex = 83
        Me.Label4.Text = "Subtitles"
        '
        'cboFileType
        '
        Me.cboFileType.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboFileType.Items.AddRange(New Object() {"BDRip", "DVDRip", "R5", "Screener", "WorkPrint"})
        Me.cboFileType.Location = New System.Drawing.Point(106, 57)
        Me.cboFileType.Name = "cboFileType"
        Me.cboFileType.Size = New System.Drawing.Size(128, 21)
        Me.cboFileType.TabIndex = 80
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(16, 57)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(50, 13)
        Me.Label3.TabIndex = 81
        Me.Label3.Text = "File Type"
        '
        'cboFileQuality
        '
        Me.cboFileQuality.DropDownWidth = 200
        Me.cboFileQuality.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboFileQuality.Items.AddRange(New Object() {"Excellent", "Good", "Watchable"})
        Me.cboFileQuality.Location = New System.Drawing.Point(322, 24)
        Me.cboFileQuality.Name = "cboFileQuality"
        Me.cboFileQuality.Size = New System.Drawing.Size(128, 21)
        Me.cboFileQuality.TabIndex = 9
        '
        'cboFileFormat
        '
        Me.cboFileFormat.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboFileFormat.Items.AddRange(New Object() {"AVI", "MP4", "MKV", "WMV"})
        Me.cboFileFormat.Location = New System.Drawing.Point(106, 24)
        Me.cboFileFormat.Name = "cboFileFormat"
        Me.cboFileFormat.Size = New System.Drawing.Size(128, 21)
        Me.cboFileFormat.TabIndex = 8
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(250, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(60, 13)
        Me.Label1.TabIndex = 77
        Me.Label1.Text = "File Quality"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(16, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 13)
        Me.Label2.TabIndex = 79
        Me.Label2.Text = "File Format"
        '
        'cboMyRating
        '
        Me.cboMyRating.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboMyRating.Items.AddRange(New Object() {"1", "2", "3", "4", "4.5", "5", "5.5", "6", "6.5", "7", "7.5", "8", "8.5", "9", "9.5", "10"})
        Me.cboMyRating.Location = New System.Drawing.Point(322, 207)
        Me.cboMyRating.Name = "cboMyRating"
        Me.cboMyRating.Size = New System.Drawing.Size(77, 21)
        Me.cboMyRating.TabIndex = 91
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(250, 211)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(55, 13)
        Me.Label6.TabIndex = 92
        Me.Label6.Text = "My Rating"
        '
        'frmAddEdit
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(504, 523)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdCancel)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAddEdit"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "frmAddEdit"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Sub SetData()
        Dim movieReader As OleDbDataReader
        Dim movieConnection As OleDbConnection
        Dim movieCommand As OleDbCommand

        movieConnection = New OleDbConnection(conStr)
        movieCommand = New OleDbCommand("SELECT * FROM MovieRecords WHERE RecordNo=" & currentRecord.ToString, movieConnection)

        movieConnection.Open()
        movieReader = movieCommand.ExecuteReader()

        movieReader.Read()
        Try
            'MessageBox.Show(movieReader.GetString(1) & ", " & strSelected)
            txtTitle.Text = movieReader.GetString(1)
            txtStarring.Text = movieReader.GetString(2)
            txtDirector.Text = movieReader.GetString(3)
            txtProducer.Text = movieReader.GetString(4)
            cboCategory.Text = movieReader.GetString(5)
            cboRating.Text = movieReader.GetString(6)
            txtYear.Text = movieReader.GetString(7)
            cboCondition.Text = movieReader.GetString(10)
            cboType.Text = movieReader.GetString(9)
            dtpRecordDate.Text = FormatDate(movieReader.GetDateTime(8))

            cboMyRating.Text = DataHelper.DbObjectToString(movieReader("MyRating"))
            cboFileFormat.Text = DataHelper.DbObjectToString(movieReader("FileFormat"))
            cboFileQuality.Text = DataHelper.DbObjectToString(movieReader("FileQuality"))
            cboFileType.Text = DataHelper.DbObjectToString(movieReader("FileType"))
            cboSubtitles.Text = DataHelper.DbObjectToString(movieReader("Subtitles"))
            txtDiscNo.Text = DataHelper.DbObjectToString(movieReader("DiscNumber"))

        Catch ex As Exception
            MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Always call Close when done reading.
        If Not (movieReader Is Nothing) Then
            movieReader.Close()
        End If

        ' Close the connection when done with it.
        If (movieConnection.State = ConnectionState.Open) Then
            movieConnection.Close()
        End If
    End Sub

    Private Sub frmAddEdit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Form_Init()
        If AddNewRecord Then
            Me.Text = "Add Record"
            cmdOK.Text = "&Add"
            dtpRecordDate.Value = Date.Now
            txtTitle.Select()
        Else
            Me.Text = "Edit Record"
            cmdOK.DialogResult = DialogResult.OK
            SetData()
        End If
    End Sub


    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        If Trim(txtTitle.Text) = "" Then
            MessageBox.Show("Record must contain a title.", "No title", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            If AddNewRecord Then
                ExecuteCommand("INSERT INTO MovieRecords(Title,Starring,Director,Producer,[Year]," & _
                    "Category,Rating,RecordDate, Type, Condition," & _
                    "MyRating, FileFormat, FileQuality, FileType, Subtitles, DiscNumber)" & _
                    " VALUES(" & FormatString(txtTitle.Text) & ", '" & txtStarring.Text & _
                    "', '" & txtDirector.Text & "', '" & txtProducer.Text & _
                    "', '" & txtYear.Text & "', '" & cboCategory.Text & "', '" & cboRating.Text & _
                    "', #" & FormatDateTime(dtpRecordDate.Value) & _
                    "#, '" & cboType.Text & "', '" & cboCondition.Text & _
                    "', " & cboMyRating.Text & ", '" & cboFileFormat.Text & "','" & cboFileQuality.Text & _
                    "', '" & cboFileType.Text & "', '" & cboSubtitles.Text & "', '" & txtDiscNo.Text & _
                    "')")

                MessageBox.Show("Record successfully added!", "Record Added", MessageBoxButtons.OK, MessageBoxIcon.Information)

                txtTitle.Text = ""
                txtStarring.Text = ""
                txtDirector.Text = ""
                txtProducer.Text = ""
                cboCategory.Text = ""
                cboRating.Text = ""
                txtYear.Text = ""
                cboCondition.Text = ""
                cboType.Text = ""
                dtpRecordDate.Text = Date.Now

                cboMyRating.Text = String.Empty
                cboFileFormat.Text = String.Empty
                cboFileQuality.Text = String.Empty
                cboFileType.Text = String.Empty
                cboSubtitles.Text = String.Empty
                txtDiscNo.Text = String.Empty

                txtTitle.Select()
            Else
                ExecuteCommand("UPDATE MovieRecords SET Title=" & FormatString(txtTitle.Text) & _
                    ", Starring='" & txtStarring.Text & "', Director='" & txtDirector.Text & _
                    "', Producer='" & txtProducer.Text & "', [Year]='" & txtYear.Text & _
                    "', Category='" & cboCategory.Text & "', Rating='" & cboRating.Text & _
                    "', RecordDate=#" & FormatDateTime(dtpRecordDate.Value) & _
                    "#, Type='" & cboType.Text & "', Condition='" & cboCondition.Text & _
                    "', MyRating=" & cboMyRating.Text & ", FileFormat='" & cboFileFormat.Text & _
                    "', FileQuality='" & cboFileQuality.Text & "', FileType='" & cboFileType.Text & _
                    "', Subtitles='" & cboSubtitles.Text & "', DiscNumber='" & txtDiscNo.Text & _
                    "' WHERE RecordNo=" & currentRecord.ToString)
                Me.Close()
            End If
        End If
    End Sub

    Sub Form_Init()
        Dim movieReader As OleDbDataReader
        Dim movieConnection As OleDbConnection
        Dim movieCommand As OleDbCommand

        'Case "Condition"
        movieConnection = New OleDbConnection(conStr)
        movieCommand = New OleDbCommand("SELECT DISTINCT Condition FROM MovieRecords", movieConnection)

        movieConnection.Open()
        movieReader = movieCommand.ExecuteReader()

        cboCondition.Items.Clear()
        Do While (movieReader.Read())
            If movieReader.GetString(0) <> "" Then
                cboCondition.Items.Add(movieReader.GetString(0))
            End If
        Loop

        If Not (movieReader Is Nothing) Then
            movieReader.Close()
        End If
        movieConnection.Close()


        'Case "Category"
        movieConnection = New OleDbConnection(conStr)
        movieCommand = New OleDbCommand("SELECT DISTINCT Category FROM MovieRecords", movieConnection)

        movieConnection.Open()
        movieReader = movieCommand.ExecuteReader()

        cboCategory.Items.Clear()
        Do While (movieReader.Read())
            If movieReader.GetString(0) <> "" Then
                cboCategory.Items.Add(movieReader.GetString(0))
            End If
        Loop

        If Not (movieReader Is Nothing) Then
            movieReader.Close()
        End If

        movieConnection.Close()
        'Case Rating
        movieConnection = New OleDbConnection(conStr)
        movieCommand = New OleDbCommand("SELECT DISTINCT Rating FROM MovieRecords", movieConnection)

        movieConnection.Open()
        movieReader = movieCommand.ExecuteReader()

        cboRating.Items.Clear()
        Do While (movieReader.Read())
            If movieReader.GetString(0) <> "" Then
                cboRating.Items.Add(movieReader.GetString(0))
            End If
        Loop

        If Not (movieReader Is Nothing) Then
            movieReader.Close()
        End If

        movieConnection.Close()

        'Case Type
        movieConnection = New OleDbConnection(conStr)
        movieCommand = New OleDbCommand("SELECT DISTINCT Type FROM MovieRecords", movieConnection)

        movieConnection.Open()
        movieReader = movieCommand.ExecuteReader()

        cboType.Items.Clear()
        Do While (movieReader.Read())
            If movieReader.GetString(0) <> "" Then
                cboType.Items.Add(movieReader.GetString(0))
            End If
        Loop

        If Not (movieReader Is Nothing) Then
            movieReader.Close()
        End If

        movieConnection.Close()
    End Sub

    Private Function FormatString(ByVal StringToFormat As String) As String
        FormatString = Chr(34) & Trim(StringToFormat) & Chr(34)
    End Function

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFileFormat.SelectedIndexChanged

    End Sub
End Class
