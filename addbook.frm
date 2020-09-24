VERSION 5.00
Begin VB.Form fromAddressBook 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Address Book"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   Icon            =   "addbook.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   8610
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   7680
      Top             =   5040
   End
   Begin VB.TextBox txtRowid 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DataField       =   "pincode"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   345
      Left            =   5370
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   5130
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DataField       =   "byear"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   660
      Left            =   1170
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   4845
      Width           =   3675
   End
   Begin VB.ComboBox CboGender 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DataField       =   "stream"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   315
      ItemData        =   "addbook.frx":0ECA
      Left            =   1170
      List            =   "addbook.frx":0ED4
      Sorted          =   -1  'True
      TabIndex        =   2
      Text            =   "CboGender"
      Top             =   930
      Width           =   2220
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7140
      Top             =   5010
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   4605
      Left            =   4950
      TabIndex        =   37
      Top             =   510
      Visible         =   0   'False
      Width           =   3555
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2430
         TabIndex        =   47
         Top             =   180
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   300
         TabIndex        =   46
         Top             =   210
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   225
         Left            =   1350
         TabIndex        =   45
         Top             =   180
         Width           =   855
      End
      Begin VB.ListBox lstSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   2565
         ItemData        =   "addbook.frx":0EE6
         Left            =   150
         List            =   "addbook.frx":0EE8
         TabIndex        =   27
         ToolTipText     =   "Double click on Selected Item"
         Top             =   1830
         Width           =   3255
      End
      Begin VB.TextBox txtcontaining 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   315
         Left            =   1080
         TabIndex        =   28
         Top             =   990
         Width           =   2325
      End
      Begin VB.CommandButton cmdCancel1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Ca&ncel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2430
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1380
         Width           =   975
      End
      Begin VB.ComboBox cboSearch 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   315
         ItemData        =   "addbook.frx":0EEA
         Left            =   1080
         List            =   "addbook.frx":0F0F
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   600
         Width           =   2325
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Value"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Left            =   330
         TabIndex        =   39
         Top             =   990
         Width           =   675
      End
      Begin VB.Label lblfield 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Left            =   300
         TabIndex        =   38
         Top             =   600
         Width           =   705
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2790
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5910
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1935
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5910
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Searc&h"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1515
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6270
      Width           =   2520
   End
   Begin VB.CommandButton cmdexit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3630
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5910
      Width           =   855
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2775
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5910
      Width           =   855
   End
   Begin VB.CommandButton cmdedit 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5910
      Width           =   855
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1065
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5910
      Width           =   855
   End
   Begin VB.CommandButton cmdlast 
      BackColor       =   &H00C0E0FF&
      Height          =   315
      Left            =   3450
      MaskColor       =   &H0000C0C0&
      Picture         =   "addbook.frx":0F76
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Last Record"
      Top             =   5580
      Width           =   645
   End
   Begin VB.CommandButton cmdnext 
      BackColor       =   &H00C0E0FF&
      Height          =   315
      Left            =   2805
      Picture         =   "addbook.frx":10E8
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Next Record"
      Top             =   5580
      Width           =   645
   End
   Begin VB.CommandButton cmdprev 
      BackColor       =   &H00C0E0FF&
      Height          =   315
      Left            =   2160
      Picture         =   "addbook.frx":125A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Previous Record"
      Top             =   5580
      Width           =   645
   End
   Begin VB.CommandButton cmdfirst 
      BackColor       =   &H00C0E0FF&
      Height          =   315
      Left            =   1515
      Picture         =   "addbook.frx":13CC
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "First Record"
      Top             =   5580
      Width           =   645
   End
   Begin VB.ComboBox cbostream 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DataField       =   "stream"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   315
      ItemData        =   "addbook.frx":153E
      Left            =   1170
      List            =   "addbook.frx":157B
      Locked          =   -1  'True
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1350
      Width           =   2220
   End
   Begin VB.ComboBox cboyear 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DataField       =   "byear"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   315
      Left            =   3570
      Locked          =   -1  'True
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   4080
      Width           =   1275
   End
   Begin VB.ComboBox cbomonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DataField       =   "bmonth"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   315
      ItemData        =   "addbook.frx":1603
      Left            =   2010
      List            =   "addbook.frx":162B
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   4080
      Width           =   1485
   End
   Begin VB.ComboBox cbodate 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DataField       =   "bdate"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   315
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   4080
      Width           =   765
   End
   Begin VB.TextBox txtemail 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DataField       =   "byear"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   345
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   4440
      Width           =   3675
   End
   Begin VB.TextBox txtphone 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DataField       =   "phoneno"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   345
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3690
      Width           =   3675
   End
   Begin VB.TextBox txtpincode 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DataField       =   "pincode"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   345
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3300
      Width           =   2235
   End
   Begin VB.TextBox txtdistrict 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DataField       =   "district"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   345
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2910
      Width           =   2235
   End
   Begin VB.TextBox txttaluka 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DataField       =   "taluka"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   345
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2520
      Width           =   2235
   End
   Begin VB.TextBox txtaddress 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DataField       =   "address"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   735
      Left            =   1170
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1740
      Width           =   3675
   End
   Begin VB.TextBox txtname 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DataField       =   "name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   345
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   510
      Width           =   3675
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "Developed By : Chetan Patel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   8310
      TabIndex        =   59
      Top             =   7440
      Width           =   3090
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
      Caption         =   "-: Birth Day :-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   195
      Left            =   6090
      TabIndex        =   57
      Top             =   5250
      Width           =   1410
   End
   Begin VB.Label Label12 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Still to Come"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   195
      Left            =   6360
      TabIndex        =   56
      Top             =   6240
      Width           =   1260
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6030
      TabIndex        =   55
      ToolTipText     =   "It Displays Birthday which is within previous 7 Days and Next 15 Days"
      Top             =   6240
      Width           =   285
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Today"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   195
      Left            =   6360
      TabIndex        =   54
      Top             =   5910
      Width           =   630
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6030
      TabIndex        =   53
      ToolTipText     =   "It Displays Birthday which is within previous 7 Days and Next 15 Days"
      Top             =   5910
      Width           =   285
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Gone"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   195
      Left            =   6360
      TabIndex        =   52
      Top             =   5550
      Width           =   540
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6030
      TabIndex        =   51
      ToolTipText     =   "It Displays Birthday which is within previous 7 Days and Next 15 Days"
      Top             =   5550
      Width           =   285
   End
   Begin VB.Label lblBDMOnth 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2220
      TabIndex        =   50
      ToolTipText     =   "It Displays Birthday which is within previous 7 Days and Next 15 Days from today"
      Top             =   7080
      Width           =   1395
   End
   Begin VB.Label lblBDDate 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1740
      TabIndex        =   49
      ToolTipText     =   "It Displays Birthday which is within previous 7 Days and Next 15 Days from today"
      Top             =   7080
      Width           =   435
   End
   Begin VB.Label lblBDName 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   48
      ToolTipText     =   "It Displays Birthday which is within previous 7 Days and Next 15 Days from today"
      Top             =   6750
      Width           =   3915
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   225
      Left            =   105
      TabIndex        =   43
      Top             =   4905
      Width           =   930
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   90
      TabIndex        =   42
      Top             =   930
      Width           =   1035
   End
   Begin VB.Label Label3 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   0
      TabIndex        =   41
      Top             =   7440
      Width           =   8625
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "                       Address Book"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   8655
   End
   Begin VB.Label lblBirthday 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Birth Day"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   90
      TabIndex        =   36
      Top             =   4110
      Width           =   1035
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E- Mail"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   225
      Left            =   90
      TabIndex        =   35
      Top             =   4500
      Width           =   1035
   End
   Begin VB.Label lblPhone 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Phone No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   90
      TabIndex        =   34
      Top             =   3720
      Width           =   1035
   End
   Begin VB.Label lblpincode 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pin Code"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   225
      Left            =   90
      TabIndex        =   33
      Top             =   3300
      Width           =   1035
   End
   Begin VB.Label lbldist 
      BackColor       =   &H00FFFFFF&
      Caption         =   "District"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   225
      Left            =   90
      TabIndex        =   32
      Top             =   2940
      Width           =   1035
   End
   Begin VB.Label lbltaluka 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Taluka"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   225
      Left            =   90
      TabIndex        =   31
      Top             =   2550
      Width           =   1035
   End
   Begin VB.Label lblAddress 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   90
      TabIndex        =   30
      Top             =   1770
      Width           =   1035
   End
   Begin VB.Label lblStream 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stream"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   90
      TabIndex        =   29
      Top             =   1350
      Width           =   1035
   End
   Begin VB.Label lblname 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   90
      TabIndex        =   24
      Top             =   540
      Width           =   1035
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   5880
      TabIndex        =   58
      ToolTipText     =   "It Displays Birthday which is within previous 7 Days and Next 15 Days"
      Top             =   5190
      Width           =   1845
   End
End
Attribute VB_Name = "fromAddressBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rsSearch As New ADODB.Recordset
Dim flag, i, recno As Integer
Dim str1, s1, s2, msg1, passwd As String
Dim RsError  As New ADODB.Recordset
Dim rsBD As New ADODB.Recordset
Dim CurDate, intDate, intYear As Integer
Dim strCurDate, strMonth As String
Dim intDate1, intDate2, intYear1, intYear2, intIncMonth, intDecMonth, intMonthdigit As Integer
Dim strMonth1, strMonth2 As String
Dim strDate1, strDate2, strDate As Date
Private Sub cboSearch_Click()
    str1 = ""
    txtcontaining.Text = ""
End Sub

Private Sub cmdadd_Click()
    txtname.SetFocus
    txtname.Text = ""
    cbostream.Text = ""
    txtaddress.Text = ""
    txttaluka.Text = ""
    txtdistrict.Text = ""
    txtpincode.Text = ""
    cbodate.Text = ""
    cbomonth.Text = ""
    cboyear.Text = ""
    txtemail.Text = ""
    txtphone.Text = ""
    txtRemarks.Text = ""
    Call visibleON
    flag = 1
    Call lockFALSE
    
    
    
End Sub

Private Sub cmdcancel_Click()
    Call cmdfirst_Click
    Call visibleOFF
    'Call show_rec
    Call lockTRUE
End Sub

Private Sub cmdCancel1_Click()
    Frame1.Visible = False
    Me.Width = 5145
  '  Me.StartUpPosition = 2
    
End Sub

Private Sub cmdDel_Click()
On Error GoTo Err1
    If txtRowid.Text = Val(1) Then
    MsgBox "            Na. Na. Na. " & Chr(13) & "You can't Delete ( Chetan Patel )", vbOKOnly + vbCritical, "Address Book"
    Else
        msg1 = MsgBox("Are you Sure ..?", vbYesNo + vbQuestion, "Address Book")
        If msg1 = vbYes Then
                Set rs = Nothing
                rs.Open " Delete from AddressBook where rowid = " & Val(txtRowid.Text), cn, adOpenDynamic, adLockOptimistic, adCmdText
                Set rs = Nothing
        End If
            cmdfirst_Click
'            Call show_rec
    End If

Exit Sub
Err1:
       Call ErrorHandlerCheck(Me.Name & " : CmdAction_Click ")
    
End Sub

Private Sub cmdedit_Click()
    txtname.SetFocus
    Call visibleON
    Call lockFALSE
    flag = 2
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub cmdfirst_Click()
    Set rs = Nothing
    rs.Open "select top 1 * from addressbook", cn, adOpenDynamic, adLockOptimistic, adCmdText
    Call show_rec
    Set rs = Nothing
End Sub

Private Sub cmdlast_Click()
    Set rs = Nothing
    rs.Open "select top 1 * from addressbook order by rowid desc", cn, adOpenDynamic, adLockOptimistic, adCmdText
    Call show_rec
    Set rs = Nothing
End Sub

Private Sub cmdnext_Click()
    Set rs = Nothing
    rs.Open "select top 1 * from addressbook where rowid > " & txtRowid.Text, cn, adOpenDynamic, adLockOptimistic, adCmdText
    Call show_rec
    Set rs = Nothing
    
End Sub

Private Sub cmdprev_Click()
    Set rs = Nothing
    rs.Open "select top 1 * from addressbook where rowid < " & txtRowid.Text & " order by rowid desc ", cn, adOpenDynamic, adLockOptimistic, adCmdText
    Call show_rec
    Set rs = Nothing
End Sub

Private Sub cmdsave_Click()
    On Error GoTo Err1
    
    If txtname.Text = "" Then
        MsgBox "Please Enter Name", vbOKOnly + vbInformation, "AddressBook"
        txtname.SetFocus
        Exit Sub
    End If
    If cbostream.Text = "" Then
        MsgBox "Please Enter stream", vbOKOnly + vbInformation, "AddressBook"
        cbostream.SetFocus
        Exit Sub
    End If
    If txtaddress.Text = "" Then
        MsgBox "Please Enter Address", vbOKOnly + vbInformation, "AddressBook"
        txtaddress.SetFocus
        Exit Sub
    End If
    If txttaluka.Text = "" Then
        MsgBox "Please Enter Taluka Name", vbOKOnly + vbInformation, "AddressBook"
        txttaluka.SetFocus
        Exit Sub
    End If
    If txtdistrict.Text = "" Then
        MsgBox "Please Enter District Name", vbOKOnly + vbInformation, "AddressBook"
        txtdistrict.SetFocus
        Exit Sub
    End If
    If txtpincode.Text = "" Then
          txtpincode.Text = 0
    End If
    If cbodate.Text = "" Then
        MsgBox "Please Select BirthDate", vbOKOnly + vbInformation, "AddressBook"
        cbodate.SetFocus
        Exit Sub
    End If
    If cbomonth.Text = "" Then
        MsgBox "Please Select BirthMonth", vbOKOnly + vbInformation, "AddressBook"
        cbomonth.SetFocus
        Exit Sub
    End If
    If cboyear.Text = "" Then
        MsgBox "Please Enter BirthYear", vbOKOnly + vbInformation, "AddressBook"
        cboyear.SetFocus
        Exit Sub
    End If
    
    If txtemail.Text = "" Then
        txtemail.Text = ""
    End If
    If txtphone.Text = "" Then
        txtphone.Text = 0
    Else
      
    Dim rsAdd As New ADODB.Recordset
    rsAdd.CursorLocation = adUseClient

    If flag = 1 Then
    '    rs.AddNew
    
        rsAdd.Open " Insert into addressbook (Name, Gender, Stream, Address, Taluka, District, " & _
                "   BirthDay, BirthMonth,  BirthYear , EMail, PhoneNo, Pincode, Remarks) values (" & _
                " '" & txtname.Text & "' , '" & CboGender.Text & "', '" & cbostream.Text & "',  '" & txtaddress.Text & "', " & _
                " '" & txttaluka.Text & "' , '" & txtdistrict.Text & "', '" & Val(cbodate.Text) & "', '" & cbomonth.Text & "' , " & Val(cboyear.Text) & " ," & _
                " '" & txtemail.Text & "', '" & txtphone.Text & "', " & Val(txtpincode.Text) & " , '" & txtRemarks.Text & "')   ", cn, adOpenDynamic, adLockOptimistic, adCmdText
        'MsgBox "Record  Saved...", vbOKOnly + vbInformation, "Birth Day Alarm"
    ElseIf flag = 2 Then
        'Call new_rec
        If Val(txtRowid.Text) <> 1 Then
        rsAdd.Open " Update addressbook set Name = '" & txtname.Text & "' , Gender = '" & CboGender.Text & "' , " & _
                " Stream = '" & cbostream.Text & "', Address = '" & txtaddress.Text & "' , Taluka = '" & txttaluka.Text & "' , " & _
                " District  = '" & txtdistrict.Text & "', BirthDay = " & cbodate.Text & ", BirthMonth = '" & cbomonth.Text & "' , " & _
                " BirthYear = " & cboyear.Text & ", EMail = '" & txtemail.Text & "', PhoneNo = '" & txtphone.Text & "' ," & _
                " Pincode = " & Val(txtpincode.Text) & ", Remarks = '" & txtRemarks.Text & "' where rowid = " & Val(txtRowid.Text), cn, adOpenDynamic, adLockOptimistic, adCmdText
                'MsgBox "Record  Updated...", vbOKOnly + vbInformation, "Birth Day Alarm"
               
        Else
                MsgBox "        Sorry No. Cheating.." & Chr(13) & " You can't Edit Chetan Patel", vbOKOnly + vbInformation, "Birth Day Alarm"
                Call cmdfirst_Click
                Exit Sub
        End If
                MsgBox "Record Saved..", vbOKOnly + vbInformation, "Birth Day Alarm"
    End If
    flag = 0
    Call visibleOFF
    Call lockTRUE
    
    
    
    
                Set rsBD = Nothing
        ' rsBD.CursorLocation = adUseClient
    
            'strDate = Date
            strDate1 = DateAdd("d", -7, Date)
            strDate2 = DateAdd("d", 15, Date)
        
          
            intDate1 = Val(Format(strDate1, "dd"))
            intDate2 = Val(Format(strDate2, "dd"))
            
            strMonth1 = Format(strDate1, "MMMM")
            strMonth2 = Format(strDate2, "MMMM")
            
            intYear1 = Val(Format(strDate1, "yyyy"))
            intYear2 = Val(Format(strDate2, "yyyy"))
                            
            
            rsBD.Open " Select * from addressbook where (BirthDay > " & Val(intDate1) & " and BirthMonth = '" & Trim(strMonth1) & "' ) or ( BirthDay < " & Val(intDate2) & "  and  BirthMonth = '" & Trim(strMonth2) & "' ) ", cn, adOpenDynamic, adLockOptimistic, adCmdText
              
    
    End If
    
    Exit Sub
    
Err1:
       Call ErrorHandlerCheck(Me.Name & " : CmdAction_Click ")

End Sub

Private Sub cmdSearch_Click()
Frame1.Visible = True
cboSearch.Text = "Name"
txtcontaining.SetFocus
Me.Width = 8700
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
    End If
End Sub

Private Sub Form_Load()
    
On Error GoTo Err1

    cn.CursorLocation = adUseClient
    Dim dbpath As String
    
     cn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\AddressBook.mdb;Jet OLEDB:Database Password=CHEtan;Persist "
    rs.Open "select * from addressbook", cn, adOpenDynamic, adLockOptimistic

    Call show_rec
    
    Set rs = Nothing
    flag = 0
    For i = 1 To 31
        cbodate.AddItem (i)
    Next
    
    i = 0
    
    Dim j As Integer
    For j = 1950 To Year(Date)
        cboyear.AddItem (j)
    Next
    
    Me.Width = 5145
      
      
        Set rsBD = Nothing
        ' rsBD.CursorLocation = adUseClient
    
            'strDate = Date
            strDate1 = DateAdd("d", -7, Date)
            strDate2 = DateAdd("d", 15, Date)
        
          
            intDate1 = Val(Format(strDate1, "dd"))
            intDate2 = Val(Format(strDate2, "dd"))
            
            strMonth1 = Format(strDate1, "MMMM")
            strMonth2 = Format(strDate2, "MMMM")
            
            intYear1 = Val(Format(strDate1, "yyyy"))
            intYear2 = Val(Format(strDate2, "yyyy"))
                            
            
            rsBD.Open " Select * from addressbook where (BirthDay > " & Val(intDate1) & " and BirthMonth = '" & Trim(strMonth1) & "' ) or ( BirthDay < " & Val(intDate2) & "  and  BirthMonth = '" & Trim(strMonth2) & "' ) ", cn, adOpenDynamic, adLockOptimistic, adCmdText
                
            Call lockTRUE
    Exit Sub
    
    
Err1:
       Call ErrorHandlerCheck(Me.Name & " : CmdAction_Click ")
    
End Sub

Sub show_rec()
    On Error GoTo Err1
    
        If rs.RecordCount > 0 Then
            txtRowid.Text = rs!Rowid
            txtname.Text = IIf(IsNull(rs!Name) = True, "", rs!Name)
            CboGender.Text = IIf(IsNull(rs!Gender) = True, "", rs!Gender)
            cbostream.Text = IIf(IsNull(rs!Stream) = True, "", rs!Stream)
            txtaddress.Text = IIf(IsNull(rs!Address) = True, "", rs!Address)
            txttaluka.Text = IIf(IsNull(rs!Taluka) = True, "", rs!Taluka)
            txtdistrict.Text = IIf(IsNull(rs!District) = True, "", rs!District)
            txtpincode.Text = IIf(IsNull(rs!PhoneNo) = True, "", rs!PhoneNo)
            cbodate.Text = IIf(IsNull(rs!BirthDay) = True, "", rs!BirthDay)
            cbomonth.Text = IIf(IsNull(rs!BirthMonth) = True, "", rs!BirthMonth)
            cboyear.Text = IIf(IsNull(rs!BirthYear) = True, "", rs!BirthYear)
            txtemail.Text = IIf(IsNull(rs!EMail) = True, "", rs!EMail)
            txtphone.Text = IIf(IsNull(rs!PhoneNo) = True, "", rs!PhoneNo)
            txtRemarks.Text = IIf(IsNull(rs!Remarks) = True, "", rs!Remarks)
        End If
    Exit Sub
Err1:
       Call ErrorHandlerCheck(Me.Name & " : CmdAction_Click ")

End Sub

Sub visibleON()
    cmdsave.Visible = True
    cmdcancel.Visible = True
    cmdadd.Visible = False
    cmdedit.Visible = False
    cmdDel.Visible = False
    cmdexit.Visible = False
End Sub
Sub visibleOFF()
    cmdsave.Visible = False
    cmdcancel.Visible = False
    cmdadd.Visible = True
    cmdedit.Visible = True
    cmdDel.Visible = True
    cmdexit.Visible = True
End Sub

Sub lockTRUE()
    txtname.Locked = True
    cbostream.Locked = True
    txtaddress.Locked = True
    txttaluka.Locked = True
    txtdistrict.Locked = True
    txtpincode.Locked = True
    cbodate.Locked = True
    cbomonth.Locked = True
    cboyear.Locked = True
    txtemail.Locked = True
    txtphone.Locked = True
    CboGender.Locked = True
    txtRemarks.Locked = True
    cmdfirst.Enabled = True
    cmdlast.Enabled = True
    cmdprev.Enabled = True
    cmdnext.Enabled = True
    cmdSearch.Enabled = True
End Sub
Sub lockFALSE()
    cmdSearch.Enabled = False
    txtname.Locked = False
    cbostream.Locked = False
    txtaddress.Locked = False
    txttaluka.Locked = False
    txtdistrict.Locked = False
    txtpincode.Locked = False
    cbodate.Locked = False
    cbomonth.Locked = False
    cboyear.Locked = False
    txtemail.Locked = False
    txtphone.Locked = False
    CboGender.Locked = False
    txtRemarks.Locked = False
    cmdfirst.Enabled = False
    cmdlast.Enabled = False
    cmdprev.Enabled = False
    cmdnext.Enabled = False
    
    Call cmdCancel1_Click
    
End Sub

Private Sub lstSearch_DblClick()
    
On Error GoTo Err1

        Set rs = Nothing
        rs.Open " select * from addressbook where name = '" & Trim(lstSearch.Text) & "'", cn, adOpenDynamic, adLockOptimistic, adCmdText
        
        If rs.RecordCount > 0 Then
                Call show_rec
        End If
    
Exit Sub

Err1:
       Call ErrorHandlerCheck(Me.Name & " : CmdAction_Click ")


End Sub

Private Sub Timer1_Timer()
                
          
        If rsBD.EOF = True And rsBD.RecordCount > 0 Then
                  rsBD.MoveFirst
        End If
            
        If rsBD.RecordCount > 0 Then
        
              
                
        
                If Val(Format(Date, "dd")) = rsBD!BirthDay And strMonth1 = Trim(rsBD!BirthMonth) Then
                        lblBDName.ForeColor = 192
                        lblBDDate.ForeColor = 192
                        lblBDMOnth.ForeColor = 192
                ElseIf Val(Format(Date, "dd")) = rsBD!BirthDay And strMonth2 = Trim(rsBD!BirthMonth) Then
                        lblBDName.ForeColor = 192
                        lblBDDate.ForeColor = 192
                        lblBDMOnth.ForeColor = 192
                ElseIf intDate1 < rsBD!BirthDay And strMonth1 = Trim(rsBD!BirthMonth) Then
                        lblBDName.ForeColor = 12648447
                        lblBDDate.ForeColor = 12648447
                        lblBDMOnth.ForeColor = 12648447
                ElseIf intDate1 < rsBD!BirthDay And strMonth2 = Trim(rsBD!BirthMonth) Then
                        lblBDName.ForeColor = 12648447
                        lblBDDate.ForeColor = 12648447
                        lblBDMOnth.ForeColor = 12648447
                ElseIf intDate1 >= rsBD!BirthDay And strMonth1 = Trim(rsBD!BirthMonth) Then
                        lblBDName.ForeColor = 16744576
                        lblBDDate.ForeColor = 16744576
                        lblBDMOnth.ForeColor = 16744576
                ElseIf intDate1 >= rsBD!BirthDay And strMonth2 = Trim(rsBD!BirthMonth) Then
                        lblBDName.ForeColor = 16744576
                        lblBDDate.ForeColor = 16744576
                        lblBDMOnth.ForeColor = 16744576
                End If

                lblBDName.Caption = rsBD!Name
                lblBDDate.Caption = rsBD!BirthDay
                lblBDMOnth.Caption = rsBD!BirthMonth
                rsBD.MoveNext
                
        Else
                lblBDName.Visible = False
                lblBDDate.Visible = False
                lblBDMOnth.Visible = False
              
        End If
        
       
        
        
End Sub

Private Sub Timer2_Timer()
 If Label15.Left + Label15.Width > 0 Then
                Label15.Left = Val(Label15.Left) - 10
        
        Else
                Label15.Left = 8600
        
        End If
End Sub

Private Sub txtcontaining_KeyPress(KeyAscii As Integer)
        On Error GoTo Err1

        str1 = str1 & Chr(KeyAscii)
        If KeyAscii = 8 Then
            If txtcontaining.Text = "" Then
                str1 = ""
            Else
                str1 = Left(str1, Len(str1) - 2)
            
                If Len(str1) <= 0 Then
                    lstSearch.Clear
                End If
            End If
        End If
    
        Dim strQry As String
        
        If Option2.Value = True Then
                    rsSearch.Open "select * from ADDRESSBOOK where " + cboSearch.Text + " like '%" + Trim(str1) + "%'", cn, adOpenDynamic, adLockOptimistic
        ElseIf Option1.Value = True Then
                    rsSearch.Open "select * from ADDRESSBOOK where " + cboSearch.Text + " like '%" + Trim(str1) + "%' and Gender='Male'", cn, adOpenDynamic, adLockOptimistic
        ElseIf Option3.Value = True Then
                    rsSearch.Open "select * from ADDRESSBOOK where " + cboSearch.Text + " like '%" + Trim(str1) + "%' and Gender ='Female'", cn, adOpenDynamic, adLockOptimistic
        End If
        
        lstSearch.Clear
        While Not rsSearch.EOF
            lstSearch.AddItem (rsSearch.Fields(1).Value)
            rsSearch.MoveNext
        Wend
        rsSearch.Close
    
        Exit Sub
Err1:
       Call ErrorHandlerCheck(Me.Name & " : CmdAction_Click ")
    
End Sub


Public Sub ErrorHandlerCheck(Optional StrFormName As String)
    StrFormName = "Address Book"
    MsgBox Err.Description, vbCritical, "ERROR"
RsError.CursorLocation = adUseClient
    Set RsError = Nothing
        RsError.Open " Insert Into ErrorOccur(ErrorDesc ,ErrorNumber ,ErrorDate)Values ('" & Replace(Err.Description, "'", "''") & "  (" & StrFormName & ")'," & Err.Number & ",'" & Str(Now) & "')", StrConnectionString, adOpenKeyset, adLockPessimistic, adCmdText
    Set RsError = Nothing

End Sub
