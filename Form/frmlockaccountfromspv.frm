VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmlockaccountfromspv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Schedulling Blok Data"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9015
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox CheckEntry 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tampilkan data entry:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   21
      Top             =   5160
      Width           =   2565
   End
   Begin VB.Frame FrameEntry 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   2025
      Left            =   60
      TabIndex        =   20
      Top             =   5220
      Width           =   8895
      Begin VB.CheckBox chknewentry 
         BackColor       =   &H00E0E0E0&
         Caption         =   "New Entry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   180
         TabIndex        =   48
         Top             =   300
         Width           =   1725
      End
      Begin VB.CheckBox chkreguler 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reguler"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   47
         Top             =   690
         Width           =   1485
      End
      Begin VB.CheckBox chkswap 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Swap"
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
         Left            =   180
         TabIndex        =   46
         Top             =   1140
         Width           =   1755
      End
      Begin VB.CheckBox chkcurrent 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Current Entry"
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
         Left            =   180
         TabIndex        =   45
         Top             =   1530
         Width           =   1755
      End
      Begin VB.OptionButton OptSwap 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Swap"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5355
         TabIndex        =   30
         Top             =   1260
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox CmbSwap 
         Height          =   315
         ItemData        =   "frmlockaccountfromspv.frx":0000
         Left            =   3660
         List            =   "frmlockaccountfromspv.frx":0007
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1260
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.ComboBox CmbReguler 
         Height          =   315
         ItemData        =   "frmlockaccountfromspv.frx":0010
         Left            =   3660
         List            =   "frmlockaccountfromspv.frx":0017
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.OptionButton OptReguler 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reguler "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5355
         TabIndex        =   25
         Top             =   840
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.ComboBox CmbNewEntry 
         Height          =   315
         ItemData        =   "frmlockaccountfromspv.frx":0020
         Left            =   3660
         List            =   "frmlockaccountfromspv.frx":0027
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   450
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.OptionButton OptNewEntry 
         BackColor       =   &H00E0E0E0&
         Caption         =   "New entry "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5355
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bulan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   29
         Top             =   1320
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bulan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   27
         Top             =   900
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bulan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   24
         Top             =   510
         Visible         =   0   'False
         Width           =   585
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   1095
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   1931
      _Version        =   196610
      BackColor       =   14737632
      Begin VB.CheckBox CHKACCOUNT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "LUNAS COMPLETE"
         Height          =   285
         Left            =   5970
         TabIndex        =   50
         Top             =   570
         Width           =   1845
      End
      Begin VB.CheckBox CHKLUNASPENDING 
         BackColor       =   &H00E0E0E0&
         Caption         =   "LUNAS PENDING"
         Height          =   465
         Left            =   5970
         TabIndex        =   49
         Top             =   60
         Width           =   1725
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   3120
         TabIndex        =   2
         Top             =   600
         Width           =   2655
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1800
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
      Begin Threed.SSOption SSOption1 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   196610
         BackColor       =   14737632
         Caption         =   "All TeleCollection"
      End
      Begin Threed.SSOption SSOption1 
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   4
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   196610
         BackColor       =   14737632
         Caption         =   "Pilih TeleCollection"
      End
      Begin Threed.SSOption SSOption1 
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   5
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   196610
         BackColor       =   14737632
         Caption         =   "Pilih SPV Name"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "TeleCollection Name :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1575
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3825
      Left            =   60
      TabIndex        =   7
      Top             =   1290
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   6747
      _Version        =   196610
      BackColor       =   14737632
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "SK-Skip"
         Height          =   255
         Index           =   11
         Left            =   5670
         TabIndex        =   44
         Top             =   780
         Width           =   1245
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON-On Nego"
         Height          =   255
         Index           =   10
         Left            =   5700
         TabIndex        =   43
         Top             =   450
         Width           =   1245
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "PR-Prospect"
         Height          =   255
         Index           =   9
         Left            =   5700
         TabIndex        =   42
         Top             =   150
         Width           =   1245
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   2325
         Left            =   90
         TabIndex        =   31
         Top             =   1410
         Width           =   8685
         Begin VB.CommandButton Command1 
            Caption         =   "Hapus"
            Height          =   315
            Left            =   3360
            TabIndex        =   51
            Top             =   720
            Width           =   585
         End
         Begin VB.CheckBox chkmultiple 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Multiple"
            Height          =   345
            Left            =   2430
            TabIndex        =   41
            Top             =   270
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.CheckBox chksingle 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Single"
            Height          =   345
            Left            =   1620
            TabIndex        =   40
            Top             =   240
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.CommandButton cmd 
            Caption         =   ">"
            Height          =   375
            Index           =   0
            Left            =   3960
            TabIndex        =   35
            Top             =   690
            Width           =   675
         End
         Begin VB.CommandButton cmd 
            Caption         =   "<"
            Height          =   375
            Index           =   1
            Left            =   3960
            TabIndex        =   34
            Top             =   1080
            Width           =   675
         End
         Begin VB.CommandButton cmd 
            Caption         =   ">>"
            Height          =   375
            Index           =   2
            Left            =   3960
            TabIndex        =   33
            Top             =   1470
            Width           =   675
         End
         Begin VB.CommandButton cmd 
            Caption         =   "<<"
            Height          =   375
            Index           =   3
            Left            =   3960
            TabIndex        =   32
            Top             =   1860
            Width           =   675
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1575
            Left            =   90
            TabIndex        =   36
            Top             =   690
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   2778
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            TextBackground  =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   1575
            Left            =   4710
            TabIndex        =   38
            Top             =   660
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   2778
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            TextBackground  =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Destination Lead Markup"
            Height          =   255
            Left            =   4800
            TabIndex        =   39
            Top             =   390
            Width           =   2685
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Source Mark Up"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   330
            Width           =   1185
         End
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Blank Data"
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   19
         Top             =   1170
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "P O P - Progress Of Payment"
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   14
         Top             =   480
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B P - Broken Promise"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   13
         Top             =   120
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "P T P - Promise To Pay"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "OS - On Process"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "VL-Valid"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   1245
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "RP - Refuse Payment"
         Height          =   255
         Index           =   7
         Left            =   135
         TabIndex        =   9
         Top             =   480
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "S P - Settled Payment"
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   8
         Top             =   810
         Width           =   2535
      End
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Index           =   0
      Left            =   105
      TabIndex        =   15
      Top             =   8610
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   196610
      MousePointer    =   16
      Caption         =   "&Add Schedule"
      ButtonStyle     =   3
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Index           =   1
      Left            =   5880
      TabIndex        =   16
      Top             =   7665
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   196610
      MousePointer    =   16
      Caption         =   "&Release"
      ButtonStyle     =   3
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Index           =   2
      Left            =   2100
      TabIndex        =   17
      Top             =   8610
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   196610
      MousePointer    =   16
      Caption         =   "E&xit"
      ButtonStyle     =   3
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Index           =   3
      Left            =   7035
      TabIndex        =   18
      Top             =   7665
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   196610
      MousePointer    =   16
      Caption         =   "&SHUT"
      ButtonStyle     =   3
   End
   Begin TDBDate6Ctl.TDBDate StartDate 
      Height          =   315
      Left            =   1575
      TabIndex        =   54
      Top             =   7455
      Width           =   1560
      _Version        =   65536
      _ExtentX        =   2752
      _ExtentY        =   556
      Calendar        =   "frmlockaccountfromspv.frx":0030
      Caption         =   "frmlockaccountfromspv.frx":0148
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmlockaccountfromspv.frx":01B4
      Keys            =   "frmlockaccountfromspv.frx":01D2
      Spin            =   "frmlockaccountfromspv.frx":0230
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd/mm/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__/__/____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   1.12794198814265E-317
      CenturyMode     =   0
   End
   Begin TDBTime6Ctl.TDBTime StartTime 
      Height          =   315
      Left            =   3150
      TabIndex        =   55
      Top             =   7455
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   556
      Caption         =   "frmlockaccountfromspv.frx":0258
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "frmlockaccountfromspv.frx":02C4
      Spin            =   "frmlockaccountfromspv.frx":0314
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "hh:nn"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "hh:nn"
      HighlightText   =   0
      Hour12Mode      =   1
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxTime         =   0.99999
      MidnightMode    =   0
      MinTime         =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__:__"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   1.02960316199441E-317
   End
   Begin TDBDate6Ctl.TDBDate EndDate 
      Height          =   315
      Left            =   1575
      TabIndex        =   56
      Top             =   7875
      Width           =   1560
      _Version        =   65536
      _ExtentX        =   2752
      _ExtentY        =   556
      Calendar        =   "frmlockaccountfromspv.frx":033C
      Caption         =   "frmlockaccountfromspv.frx":0454
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmlockaccountfromspv.frx":04C0
      Keys            =   "frmlockaccountfromspv.frx":04DE
      Spin            =   "frmlockaccountfromspv.frx":053C
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd/mm/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__/__/____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   1.12794198814265E-317
      CenturyMode     =   0
   End
   Begin TDBTime6Ctl.TDBTime EndTime 
      Height          =   315
      Left            =   3150
      TabIndex        =   57
      Top             =   7875
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   556
      Caption         =   "frmlockaccountfromspv.frx":0564
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "frmlockaccountfromspv.frx":05D0
      Spin            =   "frmlockaccountfromspv.frx":0620
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "hh:nn"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "hh:nn"
      HighlightText   =   0
      Hour12Mode      =   1
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxTime         =   0.99999
      MidnightMode    =   0
      MinTime         =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__:__"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   1.02960316199441E-317
   End
   Begin VB.Label Label8 
      Caption         =   "End Lock:"
      Height          =   225
      Left            =   210
      TabIndex        =   53
      Top             =   7875
      Width           =   1170
   End
   Begin VB.Label Label7 
      Caption         =   "Start Lock:"
      Height          =   225
      Left            =   210
      TabIndex        =   52
      Top             =   7455
      Width           =   1170
   End
End
Attribute VB_Name = "frmlockaccountfromspv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CMDSQL As String
Dim StsVl As String
Dim StsOS As String
Dim StsON As String
Dim StsSK As String
Dim StsPR As String
Dim StsPTP As String
Dim StsBP As String
Dim StsPOP As String
Dim StsSP As String
Dim StsRP As String
Dim StsOP As String
Dim StsFresh As String
Dim Stsblank As String
Dim Stsuncontact As String
Dim spv As Boolean
'@@ 140710 Tambahan buat blok entry yang diambil dari field entry_date dan pay_dt di mgm
Dim BlokEntry As String
'@@ 061110 Blok data automatic dengan waktu
Dim StringBlokTimer As String
Dim StatusLocked As String

Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "Batch", 10 * 120
    ListView2.ColumnHeaders.ADD 1, , "Batch", 20 * 120
End Sub
Private Sub CheckEntry_Click()
    If FrameEntry.Enabled = False Then
        FrameEntry.Enabled = True
        OptNewEntry.Value = True
    Else
        FrameEntry.Enabled = False
        OptNewEntry.Value = False
        OptReguler.Value = False
        OptSwap.Value = False
    End If
End Sub


Private Sub chkmultiple_Click()
    If chkmultiple.Value = vbChecked Then
        chksingle.Value = vbUnchecked
    End If
End Sub

Private Sub chksingle_Click()
    If chksingle.Value = vbChecked Then
        chkmultiple.Value = vbUnchecked
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
Select Case Index
Case 1
    If ListView2.ListItems.Count <> 0 Then
            Set lList = ListView1.ListItems.ADD(, , ListView2.SelectedItem.Text)
            ListView2.ListItems.Remove ListView2.SelectedItem.Index
    End If
Case 3
    For i = 1 To ListView2.ListItems.Count
                Set lList = ListView1.ListItems.ADD(, , ListView2.SelectedItem.Text)
                ListView2.ListItems.Remove ListView2.SelectedItem.Index
    Next
Case 0
    If ListView1.ListItems.Count <> 0 Then
        Set lList = ListView2.ListItems.ADD(, , ListView1.SelectedItem.Text)
        ListView1.ListItems.Remove ListView1.SelectedItem.Index
    End If

Case 2
    For i = 1 To ListView1.ListItems.Count
            Set lList = ListView2.ListItems.ADD(, , ListView1.SelectedItem.Text)
                   
            ListView1.ListItems.Remove ListView1.SelectedItem.Index
    Next
End Select
End Sub

Private Sub Combo1_Click(Index As Integer)
Dim m_data As New CLS_FRMSEARCH
Dim M_objrs As ADODB.Recordset
Select Case Index
Case 0
    If spv = False Then
        Set M_objrs = m_data.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + Combo1(Index).Text + "'")
        If M_objrs.RecordCount <> 0 Then
            Combo1(0).Text = M_objrs("USERID")
            Combo1(1).Text = M_objrs("AGENT")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Else
        Set M_objrs = New ADODB.Recordset
        M_objrs.Open "select * from mandiri.SPVTBL where SPVCODE='" + Combo1(0) + "'", M_OBJCONN, adOpenDynamic, adLockBatchOptimistic
            While Not M_objrs.EOF
                Combo1(0).Text = M_objrs("SPVCODE")
                Combo1(1).Text = M_objrs("SPVNAME")
                M_objrs.MoveNext
            Wend
        Set M_objrs = Nothing
        spv = True
    End If
Case 1
    If spv = False Then
        Set M_objrs = m_data.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + Combo1(Index).Text + "'")
        If M_objrs.RecordCount <> 0 Then
            Combo1(0).Text = M_objrs("USERID")
            Combo1(1).Text = M_objrs("AGENT")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Else
        Set M_objrs = New ADODB.Recordset
        M_objrs.Open "select * from mandiri.SPVTBL where SPVNAME='" + Combo1(1) + "'", M_OBJCONN, adOpenDynamic, adLockBatchOptimistic
            While Not M_objrs.EOF
                Combo1(0).Text = M_objrs("SPVCODE")
                Combo1(1).Text = M_objrs("SPVNAME")
                M_objrs.MoveNext
            Wend
        Set M_objrs = Nothing
        spv = True
    End If
    
 End Select
 Set m_data = Nothing
 Set M_objrs = Nothing
End Sub

Private Sub Command1_Click()
    M_OBJCONN.Execute "UPDATE MGM SET EXCLUDE =NULL WHERE EXCLUDE='" + ListView1.SelectedItem.Text + "'"
    getMarkup
End Sub

Private Sub Form_Load()
    Dim M_objrs As ADODB.Recordset
    Dim m_data As New CLS_FRMSEARCH
    
    Set M_objrs = m_data.QUERY_AGENT_JADWAL(M_OBJCONN, "")
        While Not M_objrs.EOF
            Combo1(0).AddItem M_objrs("USERID")
            Combo1(1).AddItem M_objrs("AGENT")
            M_objrs.MoveNext
        Wend
    Set M_objrs = Nothing
    SSOption1(0).Value = True
    spv = False
    header
    CmbNewEntry.Text = "< 2"
    CmbReguler.Text = "< 2"
    CmbSwap.Text = "> 2"
    getMarkup
End Sub


Private Sub SSCommand1_Click(Index As Integer)
Dim M_objrs As New ADODB.Recordset
Dim sStrsql As String
Dim mwhere As String
Select Case Index
Case 0
    '@@ 061110 - awal blok dengan timer -
        StringBlokTimer = "AWAL | "
    '@@ 061110 - akhir blok dengan timer -
        StatusLocked = ""

    
If CHKLUNASPENDING.Value = vbChecked And CHKACCOUNT.Value = vbChecked Then
    If Combo1(1).Text <> Empty And SSOption1(2).Value = True Then
        sStrsql = " agent in (@LUNAS PENDING@,@LUNAS COMPLETE@) AND AGENTPREV IN (SELECT USERID FROM mandiri.USERTBL WHERE SPVCODE LIKE @%SPV%@ ) "
       'mwhere = " WHERE SPVCODE LIKE '%SPV%'"
        sStrsql = sStrsql + " AND AGENTPREV IN (SELECT USERID FROM mandiri.USERTBL WHERE SPVCODE=@" + Combo1(0).Text + "@)"
        mwhere = "Where spvcode='" + Combo1(0).Text + "'"
        STRSQL = "UPDATE mandiri.usertbl SET dilockoleh='"
        STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
        STRSQL = STRSQL + "lockdarispvbuattl ='" + sStrsql + "'" + mwhere
        'M_OBJCONN.Execute STRSQL
        
        '@@ 061110 - awal blok dengan timer -
          StringBlokTimer = StringBlokTimer + STRSQL + " | "
          StatusLocked = StatusLocked + " LUNAS PENDING- LUNAS COMPLETE"
        '@@ 061110 - akhir blok dengan timer -

        
    ElseIf Combo1(1).Text = Empty And SSOption1(2).Value = True Then
            Set M_objrs = New ADODB.Recordset
            M_objrs.CursorLocation = adUseClient
            M_objrs.Open "SELECT * FROM mandiri.USERTBL WHERE USERTYPE='6'", M_OBJCONN, adOpenDynamic, adLockOptimistic
            While Not M_objrs.EOF
                sStrsql = " agent in (@LUNAS PENDING@,@LUNAS COMPLETE@) "
                sStrsql = sStrsql + " AND AGENTPREV IN (SELECT USERID FROM mandiri.USERTBL WHERE SPVCODE=@" + M_objrs("SPVCODE") + "@)"
                mwhere = "Where spvcode='" + M_objrs("SPVCODE") + "'"
                STRSQL = "UPDATE mandiri.usertbl SET dilockoleh='"
                STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                STRSQL = STRSQL + "lockdarispvbuattl ='" + sStrsql + "'" + mwhere
                'M_OBJCONN.Execute STRSQL
                
                '@@ 061110 - awal blok dengan timer -
                StringBlokTimer = StringBlokTimer + STRSQL + " | "
                '@@ 061110 - akhir blok dengan timer -
                
                M_objrs.MoveNext
            Wend
            '@@ 061110
            StatusLocked = StatusLocked + " LUNAS PENDING- LUNAS COMPLETE"
            '@@ 061110
            Set M_objrs = Nothing
    Else
    Exit Sub
    End If
    
    'MsgBox "Data Berhasil di Blok", vbOKOnly + vbInformation, "Pesan"
    Exit Sub
  ElseIf CHKLUNASPENDING.Value = vbChecked Then
        If Combo1(1).Text <> Empty And SSOption1(2).Value = True Then
          sStrsql = " agent in (@LUNAS PENDING@) AND AGENTPREV IN (SELECT USERID FROM mandiri.USERTBL WHERE SPVCODE LIKE @%SPV%@ ) "
            mwhere = " WHERE SPVCODE LIKE '%SPV%'"
            sStrsql = sStrsql + " AND AGENTPREV IN (SELECT USERID FROM mandiri.USERTBL WHERE SPVCODE=@" + Combo1(0).Text + "@)"
            mwhere = "Where spvcode='" + Combo1(0).Text + "'"
             STRSQL = "UPDATE mandiri.usertbl SET dilockoleh='"
             STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
             STRSQL = STRSQL + " lockdarispvbuattl ='" + sStrsql + "'" + mwhere
            'M_OBJCONN.Execute STRSQL
            
         '@@ 061110 - awal blok dengan timer -
          StringBlokTimer = StringBlokTimer + STRSQL + " | "
         '@@ 061110 - akhir blok dengan timer -
         
         '@@ 061110
            StatusLocked = StatusLocked + " LUNAS PENDING- "
          '@@ 061110
            
        ElseIf Combo1(1).Text = Empty And SSOption1(2).Value = True Then
            Set M_objrs = New ADODB.Recordset
            M_objrs.CursorLocation = adUseClient
            M_objrs.Open "SELECT * FROM mandiri.USERTBL WHERE USERTYPE='6'", M_OBJCONN, adOpenDynamic, adLockOptimistic
            While Not M_objrs.EOF
                sStrsql = " agent in (@LUNAS PENDING@) "
                sStrsql = sStrsql + " AND AGENTPREV IN (SELECT USERID FROM mandiri.USERTBL WHERE SPVCODE=@" + M_objrs("SPVCODE") + "@)"
                mwhere = "Where spvcode='" + M_objrs("SPVCODE") + "'"
                STRSQL = "UPDATE mandiri.usertbl SET dilockoleh='"
                STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                STRSQL = STRSQL + " lockdarispvbuattl ='" + sStrsql + "'" + mwhere
                'M_OBJCONN.Execute STRSQL
                
                '@@ 061110 - awal blok dengan timer -
                StringBlokTimer = StringBlokTimer + STRSQL + " | "
                '@@ 061110 - akhir blok dengan timer -
                
                M_objrs.MoveNext
            Wend
            '@@ 061110
            StatusLocked = StatusLocked + " LUNAS PENDING-"
            '@@ 061110
            Set M_objrs = Nothing
    Else
        Exit Sub
        
        End If
        
        'MsgBox "Data Berhasil di Blok", vbOKOnly + vbInformation, "Pesan"
        Exit Sub
ElseIf CHKACCOUNT.Value = vbChecked Then
           sStrsql = " agent in (@LUNAS COMPLETE@) AND AGENTPREV IN (SELECT USERID FROM mandiri.USERTBL WHERE SPVCODE LIKE @%SPV%@ ) "
            mwhere = " WHERE SPVCODE LIKE '%SPV%'"
        If Combo1(1).Text <> Empty And SSOption1(2).Value = True Then
            sStrsql = sStrsql + " AND AGENTPREV IN (SELECT USERID FROM mandiri.USERTBL WHERE SPVCODE=@" + Combo1(0).Text + "@)"
            mwhere = "Where spvcode='" + Combo1(0).Text + "'"
            STRSQL = "UPDATE mandiri.usertbl SET dilockoleh='"
            STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
            STRSQL = STRSQL + "lockdarispvbuattl ='" + sStrsql + "'" + mwhere
            'M_OBJCONN.Execute STRSQL
            
            '@@ 061110 - awal blok dengan timer -
            StringBlokTimer = StringBlokTimer + STRSQL + " | "
            '@@ 061110 - akhir blok dengan timer -
            
            '@@ 061110
            StatusLocked = StatusLocked + " LUNAS COMPLETE -"
            '@@ 061110
            
        ElseIf Combo1(1).Text = Empty And SSOption1(2).Value = True Then
            Set M_objrs = New ADODB.Recordset
            M_objrs.CursorLocation = adUseClient
            M_objrs.Open "SELECT * FROM mandiri.USERTBL WHERE USERTYPE='6'", M_OBJCONN, adOpenDynamic, adLockOptimistic
            While Not M_objrs.EOF
                sStrsql = " agent in (@LUNAS COMPLETE@) "
                sStrsql = sStrsql + " AND AGENTPREV IN (SELECT USERID FROM mandiri.USERTBL WHERE SPVCODE=@" + M_objrs("SPVCODE") + "@)"
                mwhere = "Where spvcode='" + M_objrs("SPVCODE") + "'"
                STRSQL = "UPDATE mandiri.usertbl SET dilockoleh='"
                STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                STRSQL = STRSQL + "lockdarispvbuattl ='" + sStrsql + "'" + mwhere
                'M_OBJCONN.Execute STRSQL
                
                '@@ 061110 - awal blok dengan timer -
                StringBlokTimer = StringBlokTimer + STRSQL + " | "
                '@@ 061110 - akhir blok dengan timer -
                
                M_objrs.MoveNext
            Wend
            Set M_objrs = Nothing
            
            '@@ 061110
            StatusLocked = StatusLocked + " LUNAS COMPLETE -"
            '@@ 061110
            
    Else
        Exit Sub
        
        End If
        
            'MsgBox "Data Berhasil di Blok", vbOKOnly + vbInformation, "Pesan"
        Exit Sub
End If


        
        If SSOption1(0).Value = False And SSOption1(1).Value = False And SSOption1(2).Value = False Then
            MsgBox "Select DCR Name To Proccess OR All"
         Else
                If SSOption1(0).Value Then
                    Call ceksts
                    STRSQL = "UPDATE mandiri.usertbl SET F_NA=NULL,F_PR=NULL,F_VL=NULL,F_ON=NULL,F_OS=NULL,F_SK=NULL, F_OP=NULL, F_PTP=NULL, F_BP=NULL, F_POP=NULL, F_SP=NULL, F_UC=NULL, F_RP=NULL "
                    STRSQL = STRSQL + ", F_WO_DATE=NULL, F_WO_2009=NULL, F_WO_2008=NULL, F_WO_2007=NULL, F_WO_2006=NULL, F_WO_2005=NULL "
                    STRSQL = STRSQL + ", F_WO_2004=NULL, F_WO_2003=NULL, F_WO_2002=NULL, F_WO_2001=NULL, F_WO_2000=NULL, F_WO_1999=NULL,lock_entry_lpd=NULL, lockmarkup=NULL,lockdarispv=NULL Where usertype='1'"
                    M_OBJCONN.Execute (STRSQL)
                    
                    
                    STRSQL = "UPDATE mandiri.usertbl SET f_flagrender=1, lockdarispv ='"
                    STRSQL = STRSQL + getblock + "',lock_entry_lpd='"
                    STRSQL = STRSQL + BlokEntry + "',dilockoleh='"
                    STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "' "
                    STRSQL = STRSQL + " Where usertype='1'"
                    'M_OBJCONN.Execute (STRSQL)
                    
                    '@@ 061110 - awal blok dengan timer -
                    StringBlokTimer = StringBlokTimer + STRSQL + " | "
                    '@@ 061110 - akhir blok dengan timer -
                    
                    
                        If ListView2.ListItems.Count <> 0 Then
                            HLSMARKUP = Replace(GETSELECTMARKUP, "'", "@")
                            sStrsql = " UPDATE mandiri.USERTBL SET dilockoleh='"
                            sStrsql = sStrsql + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                            sStrsql = sStrsql + "lockmarkup='" + HLSMARKUP + "' FROM ( "
                            sStrsql = sStrsql + " select distinct(agent) AS AGENT FROM mandiri.MGM where exclude in(" + GETSELECTMARKUP + ")) AS C "
                            sStrsql = sStrsql + " Where USERTBL.USERID = C.agent "
                            'M_OBJCONN.Execute (sStrsql)
                            
                            '@@ 061110 - awal blok dengan timer -
                            StringBlokTimer = StringBlokTimer + sStrsql + " | "
                            '@@ 061110 - akhir blok dengan timer -
                            '@@ 061110
                            StatusLocked = StatusLocked + HLSMARKUP + "-"
                            '@@ 061110
                        End If
                    'MsgBox "Proccess to All DCR Name Done.....!"
                  End If
        
                If SSOption1(1).Value Then
                    If Combo1(0).Text = "" Then
                        MsgBox "Select DCR Name To Proccess..!"
                        Combo1(0).SetFocus
                    Else
                        Call ceksts
                        STRSQL = "UPDATE mandiri.usertbl SET F_NA=NULL,F_PR=NULL,F_VL=NULL,F_ON=NULL,F_OS=NULL,F_SK=NULL, F_OP=NULL, F_PTP=NULL, F_BP=NULL, F_POP=NULL, F_SP=NULL, F_UC=NULL, F_RP=NULL "
                        STRSQL = STRSQL + ", F_WO_DATE=NULL, F_WO_2009=NULL, F_WO_2008=NULL, F_WO_2007=NULL, F_WO_2006=NULL, F_WO_2005=NULL "
                        STRSQL = STRSQL + ", F_WO_2004=NULL, F_WO_2003=NULL, F_WO_2002=NULL, F_WO_2001=NULL, F_WO_2000=NULL, F_WO_1999=NULL,lock_entry_lpd=NULL, lockmarkup=NULL,lockdarispv=NULL Where userid='" + Combo1(0).Text + "'"
                        M_OBJCONN.Execute (STRSQL)
                        
                        STRSQL = "UPDATE mandiri.usertbl SET  f_flagrender=1,lockdarispv ='" + getblock + "',lock_entry_lpd='"
                        STRSQL = STRSQL + BlokEntry + "',dilockoleh='"
                        STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "' "
                        STRSQL = STRSQL + "Where userid='" + Combo1(0).Text + "'"
                        'M_OBJCONN.Execute STRSQL
                        
                        '@@ 061110 - awal blok dengan timer -
                        StringBlokTimer = StringBlokTimer + STRSQL + " | "
                        '@@ 061110 - akhir blok dengan timer -
                        
                        If ListView2.ListItems.Count <> 0 Then
                            HLSMARKUP = Replace(GETSELECTMARKUP, "'", "@")
                            sStrsql = " UPDATE mandiri.USERTBL SET dilockoleh='"
                            sStrsql = sStrsql + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                            sStrsql = sStrsql + "lockmarkup='" + HLSMARKUP + "' FROM ( "
                            sStrsql = sStrsql + " select distinct(agent) AS AGENT FROM mandiri.MGM where exclude in(" + GETSELECTMARKUP + ") AND AGENT='" + Combo1(0).Text + "') AS C "
                            sStrsql = sStrsql + " Where USERTBL.USERID = C.agent "
                            'M_OBJCONN.Execute (sStrsql)
                            
                            '@@ 061110 - awal blok dengan timer -
                            StringBlokTimer = StringBlokTimer + sStrsql + " | "
                            '@@ 061110 - akhir blok dengan timer -
                            
                            '@@ 061110
                            StatusLocked = StatusLocked + HLSMARKUP + "-"
                            '@@ 061110
                            
                        End If
                        
                    
                        'MsgBox "Proccess To  " + Combo1(0).Text + "  " + Combo1(1).Text + " Done.....!"
                    End If
                Else
                    If SSOption1(2).Value = True Then
                        If Combo1(0).Text = "" Then
                            MsgBox "Select SPV Name To Proccess..!"
                        Else
                            Call ceksts
                        STRSQL = "UPDATE mandiri.usertbl SET F_NA=NULL,F_PR=NULL,F_VL=NULL,F_ON=NULL,F_OS=NULL,F_SK=NULL,F_OP=NULL, F_PTP=NULL, F_BP=NULL, F_POP=NULL, F_SP=NULL, F_UC=NULL, F_RP=NULL "
                        STRSQL = STRSQL + ", F_WO_DATE=NULL, F_WO_2009=NULL, F_WO_2008=NULL, F_WO_2007=NULL, F_WO_2006=NULL, F_WO_2005=NULL "
                        STRSQL = STRSQL + ", F_WO_2004=NULL, F_WO_2003=NULL, F_WO_2002=NULL, F_WO_2001=NULL, F_WO_2000=NULL, F_WO_1999=NULL,lockmarkup=NULL,lockdarispv=null,lock_entry_lpd=null Where spvcode='" + Combo1(0).Text + "'"
                        M_OBJCONN.Execute (STRSQL)
                            
                       'If CHKLUNASPENDING.Value = vbChecked Then
                        '    STRSQL = "UPDATE usertbl SET f_flagrender=1,lockdarispvbuattl ='" + getblock + "',lock_entry_lpd='"
                         '   STRSQL = STRSQL + BlokEntry + "',fromaccount ='" + cboaccount.Text + "' Where spvcode='" + Combo1(0).Text + "'"
                          '  M_OBJCONN.Execute STRSQL
                       'End If
                       
                       'If CHKACCOUNT.Value = vbChecked Then
                            STRSQL = "UPDATE mandiri.usertbl SET f_flagrender=1,lockdarispv ='" + getblock + "',lock_entry_lpd='"
                            STRSQL = STRSQL + BlokEntry + "',dilockoleh='"
                            STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "' "
                            STRSQL = STRSQL + "Where spvcode='" + Combo1(0).Text + "'"
                            'M_OBJCONN.Execute STRSQL
                       ' End If
                       
                            '@@ 061110 - awal blok dengan timer -
                            StringBlokTimer = StringBlokTimer + STRSQL + " | "
                            '@@ 061110 - akhir blok dengan timer -
                        
                            
                        If ListView2.ListItems.Count <> 0 Then
                            HLSMARKUP = Replace(GETSELECTMARKUP, "'", "@")
                            sStrsql = " UPDATE mandiri.USERTBL SET dilockoleh='"
                            sStrsql = sStrsql + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                            sStrsql = sStrsql + "lockmarkup='" + HLSMARKUP + "' FROM ( "
                            sStrsql = sStrsql + " select distinct(agent) AS AGENT FROM mandiri.MGM where exclude in(" + GETSELECTMARKUP + ") AND AGENT IN (SELECT USERID FROM USERTBL WHERE SPVCODE='" + Combo1(0).Text + "')) AS C "
                            sStrsql = sStrsql + " Where USERTBL.USERID = C.agent "
                            'M_OBJCONN.Execute (sStrsql)
                            
                            '@@ 061110 - awal blok dengan timer -
                            StringBlokTimer = StringBlokTimer + sStrsql + " | "
                            '@@ 061110 - akhir blok dengan timer -
                            
                            '@@ 061110
                            StatusLocked = StatusLocked + HLSMARKUP + "-"
                            '@@ 061110
                        End If
                        
                            
                            'MsgBox "Proccess To  " + Combo1(0).Text + "  " + Combo1(1).Text + " Done.....!"
                        End If
                    End If
             End If
        End If
        StsVl = ""
        StsPR = ""
        StsOS = ""
        StsON = ""
        StsSK = ""
        StsOP = ""
       StsPTP = ""
       StsBP = ""
       StsPOP = ""
       StsSP = ""
       StsUC = ""
       StsRP = ""
       StsWO_Date = ""
       StsWO_2009 = ""
       StsWO_2008 = ""
       StsWO_2007 = ""
       StsWO_2006 = ""
       StsWO_2005 = ""
       StsWO_2004 = ""
       StsWO_2003 = ""
       StsWO_2002 = ""
       StsWO_2001 = ""
       StsWO_2000 = ""
       StsWO_1999 = ""
       STRSQL = ""
       
            '@@ 061110 - awal blok dengan timer -
            'cek dulu tanggalnya udah diisi apa belum??
            If IsNull(StartDate.Value) Or IsNull(EndDate.Value) Or IsNull(StartTime.Value) Or IsNull(EndTime.Value) Then
                MsgBox "Start date dan End date, tidak boleh kosong!", vbOKOnly + vbExclamation, "Peringatan"
                StatusLocked = ""
                Exit Sub
            End If
            
            StringBlokTimer = Replace(StringBlokTimer + " AKHIR ", "'", "$")
            
            Dim CmdTimer As String
            Dim AccLock As String
            
            If SSOption1(0).Value = True Then
                AccLock = "ALL"
            Else
               AccLock = Trim(Combo1(0).Text)
            End If
            
            Dim WaktuAwal As Date
            Dim WaktuAkhir As Date
            
            WaktuAwal = Format(StartDate.Value, "mm-dd-yyyy") & " " & StartTime.Value
            WaktuAkhir = Format(EndDate.Value, "mm-dd-yyyy") & " " & EndTime.Value
            
            If WaktuAwal > WaktuAkhir Then
                MsgBox "Waktu awal tidak boleh lebih besar dari waktu akhir!", vbOKOnly + vbExclamation
                StatusLocked = ""
                Exit Sub
            End If

            
            CmdTimer = "insert into mandiri.tbltemplockacc (date_lock,start_lock,end_lock,"
            CmdTimer = CmdTimer + "account_lock,lock_by,status_lock,script_lock) values ('"
            CmdTimer = CmdTimer + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + "','"
            CmdTimer = CmdTimer + Format(StartDate.Value, "yyyy-mm-dd") + " " + Format(StartTime.Value, "hh:mm:00") + "','"
            CmdTimer = CmdTimer + Format(EndDate.Value, "yyyy-mm-dd") + " " + Format(EndTime.Value, "hh:mm:00") + "','"
            CmdTimer = CmdTimer + AccLock + "','"
            CmdTimer = CmdTimer + Trim(MDIForm1.TxtUsername.Text) + "','"
            CmdTimer = CmdTimer + StatusLocked + "','"
            CmdTimer = CmdTimer + StringBlokTimer + "')"
            
            Dim m_objrsTimer As ADODB.Recordset
            Set m_objrsTimer = New ADODB.Recordset
            m_objrsTimer.CursorLocation = adUseClient
            m_objrsTimer.Open CmdTimer, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            Set m_objrsTimer = Nothing
            
            MsgBox "Lock data berhasil ditambahkan dalam schedulle!", vbOKOnly + vbInformation, "Informasi"
            Unload Me
            '@@ 061110 - akhir blok dengan timer -
 
Case 1
     If SSOption1(2).Value = True Then
               
            If CHKLUNASPENDING.Value = vbChecked Or CHKACCOUNT.Value = vbChecked Then
                    
                   Set M_objrs = New ADODB.Recordset
            M_objrs.CursorLocation = adUseClient
            
            If Combo1(0).Text = "" Then
                    M_objrs.Open "SELECT * FROM mandiri.USERTBL WHERE usertype='6'", M_OBJCONN, adOpenDynamic, adLockOptimistic
            Else
                    M_objrs.Open "SELECT * FROM mandiri.USERTBL WHERE SPVCODE='" + Combo1(0).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
            End If
            
            While Not M_objrs.EOF
                STRSQL = "UPDATE usertbl SET dilockoleh='Clear by:"
                STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                STRSQL = STRSQL + " lockdarispvbuattl=NULL WHERE SPVCODE='" + M_objrs("SPVCODE") + "'"
                M_OBJCONN.Execute STRSQL
                M_objrs.MoveNext
            Wend
            Set M_objrs = Nothing
            MsgBox "Data telah direlease"
           Exit Sub
           End If
            
                If Combo1(0).Text = "" Then
                MsgBox "CLIK DULU COMBO SPV", vbInformation + vbOKOnly, "PESAN"
                Exit Sub
               End If
                    STRSQL = "UPDATE mandiri.usertbl SET dilockoleh='Clear by:"
                    STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                    STRSQL = STRSQL + " lockdarispv=NULL,lock_entry_lpd=NULL,fromaccount=NULL,lockmarkup=NULL,lockdarispvbuattl=NULL WHERE SPVCODE='" + Combo1(0).Text + "'"
                
     Else
            If SSOption1(1).Value = True Then
                If Combo1(0).Text = "" Then
                    MsgBox "CLIK DULU COMBO NYA", vbInformation + vbOKOnly, "PESAN"
                Exit Sub
                End If
                STRSQL = "UPDATE mandiri.usertbl SET dilockoleh='Clear by:"
                STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                STRSQL = STRSQL + " lockdarispv=NULL,lock_entry_lpd=NULL,lockmarkup=NULL ,fromaccount=NULL,lockdarispvbuattl=NULL WHERE userid='" + Combo1(0).Text + "'"
            Else
                    STRSQL = "UPDATE mandiri.usertbl SET lockdarispv=NULL,lock_entry_lpd=NULL,lockmarkup=NULL,dilockoleh='Clear by:"
                    STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "'"
            End If
            
    End If
    If STRSQL <> "" Then
        M_OBJCONN.Execute STRSQL
        MsgBox "Reset Done.....!"
        End If
        
Case 2
        Unload Me

Case 3
        STRSQL1 = "UPDATE  mandiri.tblshut SET nshut=1 "
        M_OBJCONN.Execute STRSQL1
End Select

End Sub
Sub ceksts()
If Check1(0).Value Then
    StsVl = "VL-"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsVl
    
End If

If Check1(1).Value Then
    StsOP = "OS-"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsOP
End If

If Check1(2).Value Then
    StsPTP = "PTP"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsPTP + "-"
End If

If Check1(3).Value Then
    StsBP = "BP-"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsBP
End If

If Check1(4).Value Then
    StsPOP = "POP"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsPOP
End If
 
If Check1(5).Value Then
    StsSP = "SP-"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsSP
End If

If Check1(7).Value Then
    StsRP = "RP-"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsRP
End If

If Check1(6).Value Then
    Stsblank = "anto"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + "BlankData -"
End If

If Check1(9).Value Then
    StsPR = "PR-"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsPR
End If

If Check1(10).Value Then
    StsON = "ON-"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsON
End If

If Check1(11).Value Then
    StsSK = "SK-"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsSK
End If

BlokEntry = ""
bCheckNewentry = False
bCheckReguler = False
bCheckSwap = False
bCheckCurrent = False

If chknewentry.Value = vbChecked Then
    bCheckNewentry = True
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + "NewEntry -"
End If


If chkreguler.Value = vbChecked Then
   bCheckReguler = True
   '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + "Reguler -"
End If

If chkswap.Value = vbChecked Then
   bCheckSwap = True
   '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + "SWAP -"
End If


If chkcurrent.Value = vbChecked Then
   bCheckCurrent = True
   '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + "Current -"
End If






If bCheckSwap = True And bCheckNewentry = True And bCheckReguler = True And bCheckCurrent = True Then
    BlokEntry = " (date_part(''month'',entry_date) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
    BlokEntry = BlokEntry + " and date_part(''year'',entry_date)=date_part(''year'',now) or "
    BlokEntry = BlokEntry + " (date_part(''month'',pay_dt_update) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
    BlokEntry = BlokEntry + " and date_part(''year'',pay_dt_update)=date_part(''year'',now)) or "
    BlokEntry = BlokEntry + " (((date_part(''month'',pay_dt_update)< (date_part(''month'',now)- 2 ) "
    BlokEntry = BlokEntry + " and date_part(''year'',pay_dt_update)<=date_part(''year'',now)) or "
    BlokEntry = BlokEntry + " pay_dt_update isnull) and "
    BlokEntry = BlokEntry + " date_part(''month'',entry_date) < (date_part(''month'',now)-2) "
    BlokEntry = BlokEntry + " and date_part(''year'',entry_date) <= date_part(''year'',now)) or "
    BlokEntry = BlokEntry + " (date_part(''month'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "mm") + " and date_part(''year'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "yyyy") + ")"
    BlokEntry = BlokEntry + " )"
    Exit Sub
ElseIf bCheckNewentry = True And bCheckReguler = True And bCheckCurrent = True Then
    BlokEntry = " (date_part(''month'',entry_date) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
    BlokEntry = BlokEntry + " and date_part(''year'',entry_date)=date_part(''year'',now) or "
    BlokEntry = BlokEntry + " (date_part(''month'',pay_dt_update) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
    BlokEntry = BlokEntry + " and date_part(''year'',pay_dt_update)=date_part(''year'',now)) or "
    BlokEntry = BlokEntry + " (date_part(''month'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "mm") + " and date_part(''year'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "yyyy") + ")"
    BlokEntry = BlokEntry + " )"
    Exit Sub
ElseIf bCheckNewentry = True And bCheckSwap = True And bCheckCurrent = True Then
    BlokEntry = " (date_part(''month'',entry_date) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
    BlokEntry = BlokEntry + " and date_part(''year'',entry_date)=date_part(''year'',now) or "
    BlokEntry = BlokEntry + " (((date_part(''month'',pay_dt_update)< (date_part(''month'',now)- 2 ) "
    BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)<=date_part(''year'',now)) or "
    BlokEntry = BlokEntry + "pay_dt_update isnull) and "
    BlokEntry = BlokEntry + "date_part(''month'',entry_date) < (date_part(''month'',now)-2) "
    BlokEntry = BlokEntry + " and date_part(''year'',entry_date) <= date_part(''year'',now)) or "
    BlokEntry = BlokEntry + " (date_part(''month'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "mm") + " and date_part(''year'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "yyyy") + ")"
    BlokEntry = BlokEntry + " )"
    Exit Sub
ElseIf bCheckReguler = True And bCheckSwap = True And bCheckCurrent = True Then
   BlokEntry = " (date_part(''month'',pay_dt_update) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
   BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)=date_part(''year'',now) or "
   BlokEntry = BlokEntry + " (((date_part(''month'',pay_dt_update)< (date_part(''month'',now)- 2 ) "
   BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)<=date_part(''year'',now)) or "
   BlokEntry = BlokEntry + "pay_dt_update isnull) and "
   BlokEntry = BlokEntry + "date_part(''month'',entry_date) < (date_part(''month'',now)-2) "
   BlokEntry = BlokEntry + " and date_part(''year'',entry_date) <= date_part(''year'',now)) or "
   BlokEntry = BlokEntry + " (date_part(''month'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "mm") + " and date_part(''year'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "yyyy") + ")"
   BlokEntry = BlokEntry + " )"
   Exit Sub
End If




If bCheckNewentry = True Then
    BlokEntry = " date_part(''month'',entry_date) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
    BlokEntry = BlokEntry + "and date_part(''year'',entry_date)=date_part(''year'',now)"
    Exit Sub
End If


If bCheckReguler = True Then
    BlokEntry = " date_part(''month'',pay_dt_update) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
    BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)=date_part(''year'',now)"
    Exit Sub
End If


If bCheckSwap = True Then
    BlokEntry = " ((date_part(''month'',pay_dt_update)< (date_part(''month'',now)- 2 ) "
    BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)<=date_part(''year'',now)) or "
    BlokEntry = BlokEntry + "pay_dt_update isnull) and "
    BlokEntry = BlokEntry + "date_part(''month'',entry_date) < (date_part(''month'',now)-2) "
    BlokEntry = BlokEntry + " and date_part(''year'',entry_date) <= date_part(''year'',now)"
    Exit Sub
End If

If bCheckCurrent = True Then
   BlokEntry = " (date_part(''month'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "mm") + " and date_part(''year'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "yyyy") + ")"
    Exit Sub
End If

'@@ 140710 Tambahan buat blok entry

'If OptNewEntry.Value = True Then
'    BlokEntry = " datediff(''month'',entry_date,now) "
'    BlokEntry = BlokEntry + CmbNewEntry.Text
'End If
'
'If OptReguler.Value = True Then
'    BlokEntry = " datediff(''month'',pay_dt,now) "
'    BlokEntry = BlokEntry + CmbReguler.Text
'End If
'
'If OptSwap.Value = True Then
'    BlokEntry = " datediff(''month'',pay_dt,now) "
'    BlokEntry = BlokEntry + CmbSwap.Text
'End If

'@@ 150710 Ubah blok entry
'If OptNewEntry.Value = True Then
'    BlokEntry = " date_part(''month'',entry_date) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
'    BlokEntry = BlokEntry + "and date_part(''year'',entry_date)=date_part(''year'',now)"
'End If
'
'If OptReguler.Value = True Then
'    BlokEntry = " date_part(''month'',pay_dt_update) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
'    BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)=date_part(''year'',now)"
'End If
'
'If OptSwap.Value = True Then
'    BlokEntry = " ((date_part(''month'',pay_dt_update)< (date_part(''month'',now)- 2 ) "
'    BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)<=date_part(''year'',now)) or "
'    BlokEntry = BlokEntry + "pay_dt_update isnull) and "
'    BlokEntry = BlokEntry + "date_part(''month'',entry_date) < (date_part(''month'',now)-2) "
'    BlokEntry = BlokEntry + " and date_part(''year'',entry_date) <= date_part(''year'',now)"
'End If

End Sub
Public Function getblock() As String


                    STRINGBLOK = ""
                    
                    If StsVl <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                            STRINGBLOK = " substring(F_cek_new,1,3) in (@" + StsVl + "@"
                        Else
                            STRINGBLOK = STRINGBLOK + ",@" + StsVl + "@"
                        End If
                    End If
                    
                    If StsPR <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsPR + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsPR + "@"
                        End If
                    End If
                    
                    If StsPTP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsPTP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsPTP + "@"
                        End If
                    End If
                    
                    If StsPOP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsPOP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsPOP + "@"
                        End If
                    End If
                    
                    If StsBP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsBP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsBP + "@"
                        End If
                    End If
                    
                    If StsSP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)  in (@" + StsSP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsSP + "@"
                        End If
                    End If
                    
                    If StsRP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsRP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsRP + "@"
                        End If
                    End If
                    
                    If StsSK <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsSK + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsSK + "@"
                        End If
                    End If
                    
                     If StsON <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsON + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsON + "@"
                        End If
                    End If
                    
                     If StsOP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsOP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsOP + "@"
                        End If
                    End If
                    
                    
                     If Stsblank <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_NEW,1,3)   in (@@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@@"
                        End If
                    End If
                    
                    
                    
                
                    If Len(STRINGBLOK) > 0 Then
                            STRINGBLOK = STRINGBLOK + ")"
                    End If
                    getblock = STRINGBLOK
End Function
Private Sub SSOption1_Click(Index As Integer, Value As Integer)
Dim M_objrs As ADODB.Recordset
Select Case Index
Case 0
        Combo1(0).Enabled = False
        Combo1(1).Enabled = False
Case 1
        Combo1(0).Enabled = True
        Combo1(1).Enabled = True
        Combo1(0).CLEAR
        Combo1(1).CLEAR
        
        '@@221010
'        Dim M_DATA As New CLS_FRMSEARCH
'        Set M_OBJRS = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "")
'            While Not M_OBJRS.EOF
'                Combo1(0).AddItem M_OBJRS("USERID")
'                Combo1(1).AddItem M_OBJRS("AGENT")
'                M_OBJRS.MoveNext
'            Wend
'        Set M_OBJRS = Nothing
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open "select userid,agent from mandiri.usertbl where usertype='1' order by userid asc", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        While Not M_objrs.EOF
            Combo1(0).AddItem M_objrs("userid")
            Combo1(1).AddItem M_objrs("agent")
            M_objrs.MoveNext
        Wend
        
        'SSOption1(0).Value = True
        spv = False
Case 2
        Combo1(0).Enabled = True
        Combo1(1).Enabled = True
        Combo1(0).CLEAR
        Combo1(1).CLEAR
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        If UCase(MDIForm1.txtlevel.Text) = "SUPERVISOR" Then
            M_objrs.Open "select * from mandiri.SPVTBL ", M_OBJCONN, adOpenDynamic, adLockOptimistic
        ElseIf UCase(MDIForm1.txtlevel.Text) = "ADMINISTRATOR" Or UCase(MDIForm1.txtlevel.Text) = "ADMIN" Then
            M_objrs.Open "select * from mandiri.SPVTBL", M_OBJCONN, adOpenDynamic, adLockOptimistic
        ElseIf UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Then
        M_objrs.Open "select * from mandiri.SPVTBL where team='" + MDIForm1.TxtUsername.Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
        End If
            While Not M_objrs.EOF
                Combo1(0).AddItem M_objrs("SPVCODE")
                Combo1(1).AddItem M_objrs("SPVNAME")
                M_objrs.MoveNext
            Wend
        Set M_objrs = Nothing
        spv = True
'        SSOption1(0).Value = True
'        SSOption1(1).Value = True
        
End Select

End Sub
Public Sub getMarkup()
Dim list As ListItem
Dim RSNEW As New ADODB.Recordset
Set rs = New ADODB.Recordset
RSNEW.CursorLocation = adUseClient
RSNEW.Open "select distinct(exclude) from mgm WHERE (exclude <>'')", M_OBJCONN, adOpenDynamic, adLockOptimistic
ListView1.ListItems.CLEAR
While Not RSNEW.EOF
Set list = ListView1.ListItems.ADD(, , IIf(IsNull(RSNEW!exclude), "", RSNEW!exclude))
    RSNEW.MoveNext
Wend

End Sub

Public Function GETSELECTMARKUP() As String
Dim j As Integer
Dim TMPSELECTMARKUP As String
GETSELECTMARKUP = ""
For j = 1 To ListView2.ListItems.Count
        If j = 1 Then
            TMPSELECTMARKUP = TMPSELECTMARKUP + Chr(39) + ListView2.ListItems(j).Text + Chr(39)
        Else
            TMPSELECTMARKUP = TMPSELECTMARKUP + "," + Chr(39) + ListView2.ListItems(j).Text + Chr(39)
        End If
    
        
Next j
GETSELECTMARKUP = TMPSELECTMARKUP
End Function
