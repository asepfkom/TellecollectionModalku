VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form_AutoDialer 
   Caption         =   "Setting Auto Dialler"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   ScaleHeight     =   9165
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Height          =   225
      Left            =   3150
      TabIndex        =   37
      Top             =   3195
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   9180
      Left            =   0
      ScaleHeight     =   9120
      ScaleWidth      =   11685
      TabIndex        =   0
      Top             =   0
      Width           =   11745
      Begin TabDlg.SSTab SSTab1 
         Height          =   8955
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   8940
         _ExtentX        =   15769
         _ExtentY        =   15796
         _Version        =   393216
         Tab             =   2
         TabHeight       =   520
         BackColor       =   -2147483643
         TabCaption(0)   =   "Create Auto Dialer"
         TabPicture(0)   =   "Form_AutoDialer.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(1)=   "Frame1"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Auto Dialer On Running"
         TabPicture(1)   =   "Form_AutoDialer.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame5"
         Tab(1).Control(1)=   "Frame4"
         Tab(1).Control(2)=   "Frame3"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "Auto Dial Schedule"
         TabPicture(2)   =   "Form_AutoDialer.frx":0038
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Frame6"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Frame7"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).ControlCount=   2
         Begin VB.Frame Frame7 
            BackColor       =   &H80000005&
            Caption         =   "Detail Data"
            Height          =   7410
            Left            =   30
            TabIndex        =   30
            Top             =   1500
            Width           =   8865
            Begin MSComctlLib.ListView ListView2 
               Height          =   7095
               Left            =   75
               TabIndex        =   31
               Top             =   240
               Width           =   8700
               _ExtentX        =   15346
               _ExtentY        =   12515
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   0
               BackColor       =   16777215
               BorderStyle     =   1
               Appearance      =   0
               MousePointer    =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   0
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   1155
            Left            =   45
            TabIndex        =   28
            Top             =   345
            Width           =   8850
            Begin VB.CommandButton cmd_schedule 
               Caption         =   "Check Auto Dial Schedule"
               Height          =   840
               Left            =   120
               TabIndex        =   29
               Top             =   165
               Width           =   1620
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   1155
            Left            =   -74955
            TabIndex        =   25
            Top             =   345
            Width           =   8850
            Begin VB.CommandButton cmd_stop 
               Caption         =   "Stop Auto Dial"
               Height          =   840
               Left            =   1965
               TabIndex        =   27
               Top             =   165
               Width           =   1620
            End
            Begin VB.CommandButton cmd_check_running 
               Caption         =   "Check Auto Dial Running"
               Height          =   840
               Left            =   105
               TabIndex        =   26
               Top             =   165
               Width           =   1620
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H80000005&
            Caption         =   "Detail Data"
            Height          =   7410
            Left            =   -74970
            TabIndex        =   23
            Top             =   1500
            Width           =   8865
            Begin MSComctlLib.ListView ListView1 
               Height          =   7095
               Left            =   75
               TabIndex        =   24
               Top             =   240
               Width           =   8700
               _ExtentX        =   15346
               _ExtentY        =   12515
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               Checkboxes      =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   0
               BackColor       =   16777215
               BorderStyle     =   1
               Appearance      =   0
               MousePointer    =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   0
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Sampah"
            Height          =   255
            Left            =   -74955
            TabIndex        =   20
            Top             =   8640
            Visible         =   0   'False
            Width           =   750
            Begin MSComctlLib.ListView LVAgent 
               Height          =   2175
               Left            =   915
               TabIndex        =   21
               Top             =   0
               Visible         =   0   'False
               Width           =   4005
               _ExtentX        =   7064
               _ExtentY        =   3836
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               Checkboxes      =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   12582912
               BackColor       =   16777215
               BorderStyle     =   1
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   0
            End
            Begin VB.Label Label4 
               BackColor       =   &H80000005&
               Caption         =   "Agent"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   120
               TabIndex        =   22
               Top             =   330
               Visible         =   0   'False
               Width           =   960
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H80000005&
            Caption         =   "Detail Data"
            Height          =   4635
            Left            =   -74955
            TabIndex        =   16
            Top             =   4455
            Width           =   8865
            Begin MSComctlLib.ListView lv3 
               Height          =   4245
               Left            =   75
               TabIndex        =   19
               Top             =   255
               Width           =   8700
               _ExtentX        =   15346
               _ExtentY        =   7488
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               Checkboxes      =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   0
               BackColor       =   16777215
               BorderStyle     =   1
               Appearance      =   0
               MousePointer    =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   0
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H80000005&
            Caption         =   "Filter Data"
            Height          =   4140
            Left            =   -74955
            TabIndex        =   2
            Top             =   300
            Width           =   8865
            Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
               Height          =   315
               Left            =   1245
               TabIndex        =   45
               Top             =   3390
               Width           =   690
               _Version        =   65536
               _ExtentX        =   1217
               _ExtentY        =   556
               Calculator      =   "Form_AutoDialer.frx":0054
               Caption         =   "Form_AutoDialer.frx":0074
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "Form_AutoDialer.frx":00E0
               Keys            =   "Form_AutoDialer.frx":00FE
               Spin            =   "Form_AutoDialer.frx":0148
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "####0;;Null"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "####0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   100
               MinValue        =   -99999
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   -1
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin VB.ComboBox CmbRetryCall 
               Height          =   315
               Left            =   3915
               TabIndex        =   44
               Text            =   "Combo4"
               Top             =   2745
               Visible         =   0   'False
               Width           =   825
            End
            Begin VB.ComboBox Combo3 
               Height          =   315
               Index           =   1
               Left            =   1410
               TabIndex        =   42
               Top             =   2715
               Width           =   1290
            End
            Begin VB.ComboBox Combo3 
               Height          =   315
               Index           =   0
               ItemData        =   "Form_AutoDialer.frx":0170
               Left            =   1425
               List            =   "Form_AutoDialer.frx":017A
               TabIndex        =   41
               Top             =   2325
               Width           =   1275
            End
            Begin VB.CheckBox Check1 
               Height          =   225
               Left            =   1425
               TabIndex        =   36
               Top             =   4140
               Visible         =   0   'False
               Width           =   225
            End
            Begin VB.ComboBox Combo2 
               Height          =   315
               Left            =   1455
               TabIndex        =   32
               Top             =   1965
               Width           =   3330
            End
            Begin VB.CommandButton runbtn 
               BackColor       =   &H00FF8080&
               Caption         =   "Run"
               Height          =   480
               Left            =   3705
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   3420
               Width           =   1485
            End
            Begin VB.CommandButton check 
               BackColor       =   &H0080FF80&
               Caption         =   "Check Before Run"
               Height          =   480
               Left            =   2160
               MaskColor       =   &H0080FF80&
               Style           =   1  'Graphical
               TabIndex        =   17
               Top             =   3435
               Width           =   1485
            End
            Begin VB.ComboBox cbospvname 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2790
               TabIndex        =   8
               Top             =   750
               Width           =   2310
            End
            Begin VB.ComboBox cbospv 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1455
               TabIndex        =   7
               Top             =   750
               Width           =   1275
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               Left            =   1065
               TabIndex        =   6
               Top             =   330
               Width           =   4050
            End
            Begin MSComctlLib.ListView LVStatusCall 
               Height          =   3315
               Left            =   5265
               TabIndex        =   9
               Top             =   600
               Width           =   3510
               _ExtentX        =   6191
               _ExtentY        =   5847
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               Checkboxes      =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   16711680
               BackColor       =   16777215
               BorderStyle     =   1
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   0
            End
            Begin TDBDate6Ctl.TDBDate tgl1 
               Height          =   315
               Left            =   1440
               TabIndex        =   10
               Top             =   1155
               Width           =   1365
               _Version        =   65536
               _ExtentX        =   2408
               _ExtentY        =   556
               Calendar        =   "Form_AutoDialer.frx":0190
               Caption         =   "Form_AutoDialer.frx":02A8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "Form_AutoDialer.frx":0314
               Keys            =   "Form_AutoDialer.frx":0332
               Spin            =   "Form_AutoDialer.frx":0390
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd-mmm-yyyy"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               FirstMonth      =   4
               ForeColor       =   0
               Format          =   "dd-mm-yyyy"
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
               Text            =   "__-__-____"
               ValidateMode    =   0
               ValueVT         =   1
               Value           =   37468
               CenturyMode     =   0
            End
            Begin TDBDate6Ctl.TDBDate tgl2 
               Height          =   315
               Left            =   1440
               TabIndex        =   11
               Top             =   1560
               Width           =   1365
               _Version        =   65536
               _ExtentX        =   2408
               _ExtentY        =   556
               Calendar        =   "Form_AutoDialer.frx":03B8
               Caption         =   "Form_AutoDialer.frx":04D0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "Form_AutoDialer.frx":053C
               Keys            =   "Form_AutoDialer.frx":055A
               Spin            =   "Form_AutoDialer.frx":05B8
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd-mmm-yyyy"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               FirstMonth      =   4
               ForeColor       =   0
               Format          =   "dd-mm-yyyy"
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
               Text            =   "__-__-____"
               ValidateMode    =   0
               ValueVT         =   1
               Value           =   37468
               CenturyMode     =   0
            End
            Begin TDBTime6Ctl.TDBTime jam1 
               Height          =   315
               Left            =   2820
               TabIndex        =   12
               Top             =   1155
               Width           =   900
               _Version        =   65536
               _ExtentX        =   1587
               _ExtentY        =   556
               Caption         =   "Form_AutoDialer.frx":05E0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "Form_AutoDialer.frx":064C
               Spin            =   "Form_AutoDialer.frx":069C
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
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
               MaxTime         =   0.999988425925926
               MidnightMode    =   0
               MinTime         =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
               PromptChar      =   "_"
               ReadOnly        =   0
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "__:__"
               ValidateMode    =   0
               ValueVT         =   1
               Value           =   0.507210648148148
            End
            Begin TDBTime6Ctl.TDBTime jam2 
               Height          =   315
               Left            =   2835
               TabIndex        =   13
               Top             =   1560
               Width           =   900
               _Version        =   65536
               _ExtentX        =   1587
               _ExtentY        =   556
               Caption         =   "Form_AutoDialer.frx":06C4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "Form_AutoDialer.frx":0730
               Spin            =   "Form_AutoDialer.frx":0780
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
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
               MaxTime         =   0.999988425925926
               MidnightMode    =   0
               MinTime         =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
               PromptChar      =   "_"
               ReadOnly        =   0
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "__:__"
               ValidateMode    =   0
               ValueVT         =   1
               Value           =   0.507210648148148
            End
            Begin VB.Label Label13 
               Caption         =   "Retry Call status UTC"
               Height          =   390
               Left            =   210
               TabIndex        =   43
               Top             =   3360
               Width           =   930
            End
            Begin VB.Label Label12 
               BackColor       =   &H80000005&
               Caption         =   "Order By Ke 2"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   165
               TabIndex        =   40
               Top             =   2790
               Width           =   1350
            End
            Begin VB.Label Label10 
               BackColor       =   &H80000005&
               Caption         =   "Order By Ke 1"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   165
               TabIndex        =   38
               Top             =   2340
               Width           =   1200
            End
            Begin VB.Label Label9 
               BackColor       =   &H80000005&
               Caption         =   "DPD"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   1830
               TabIndex        =   35
               Top             =   4140
               Visible         =   0   'False
               Width           =   825
            End
            Begin VB.Label Label8 
               BackColor       =   &H80000005&
               Caption         =   "Outstanding"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   150
               TabIndex        =   34
               Top             =   4125
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.Label Label7 
               BackColor       =   &H80000005&
               Caption         =   "Phone To Call"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   150
               TabIndex        =   33
               Top             =   1905
               Width           =   960
            End
            Begin VB.Label Label6 
               BackColor       =   &H80000005&
               Caption         =   "To"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   150
               TabIndex        =   15
               Top             =   1575
               Width           =   315
            End
            Begin VB.Label Label5 
               BackColor       =   &H80000005&
               Caption         =   "Running From"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   150
               TabIndex        =   14
               Top             =   1170
               Width           =   1275
            End
            Begin VB.Label Label1 
               BackColor       =   &H80000005&
               Caption         =   "Campaign"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   150
               TabIndex        =   5
               Top             =   360
               Width           =   960
            End
            Begin VB.Label Label2 
               BackColor       =   &H80000005&
               Caption         =   "STATUS CALL"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   5280
               TabIndex        =   4
               Top             =   300
               Width           =   1455
            End
            Begin VB.Label Label3 
               BackColor       =   &H80000005&
               Caption         =   "SPV Name"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   165
               TabIndex        =   3
               Top             =   780
               Width           =   960
            End
         End
      End
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   495
      Left            =   3885
      TabIndex        =   39
      Top             =   4245
      Width           =   1215
   End
End
Attribute VB_Name = "Form_AutoDialer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public qpub As String
Public qpub2 As String

Private Sub cbospv_Click()
    cbospvname.ListIndex = cbospv.ListIndex
End Sub

Private Sub cbospvname_Click()
    cbospv.ListIndex = cbospvname.ListIndex
End Sub

Private Sub check_Click()
    Call check_data
    runbtn.Enabled = True
End Sub

Private Sub Check1_Click()
    Call check_data
End Sub

Private Sub Check2_Click()
    Call check_data
End Sub

Private Sub Combo1_Click()
    Call isistatuscall
End Sub

Private Sub Combo1_DropDown()
    sStrsql = "select distinct(recsource) from mgm"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    Combo1.clear
        While Not M_objrs.EOF
            Combo1.AddItem IIf(IsNull(M_objrs!recsource), "", M_objrs!recsource)
            M_objrs.MoveNext
        Wend
    Set M_objrs = Nothing
End Sub

Private Sub Combo2_DropDown()
    Combo2.clear
    Combo2.AddItem "Handphone"
    Combo2.AddItem "Home Phone"
    Combo2.AddItem "Office Phone"
End Sub

Private Sub Combo3_Click(Index As Integer)
  
  Combo3(1).clear
  Select Case Combo3(0).text
  
  Case "DPD"
        Combo3(1).AddItem "OUTSTANDING"
        Combo3(1).text = "OUTSTANDING"
  Case "OUTSTANDING"
        Combo3(1).AddItem "DPD"
        Combo3(1).text = "DPD"
  End Select
  
End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
    Call HeaderLvAgent
    'Call isistatuscall
    runbtn.Enabled = False
    
    If MDIForm1.txtlevel.text = "Supervisor" Then
        cbospv.text = MDIForm1.TxtUsername.text
        cbospvname.text = MDIForm1.txtnama.text
        cbospv.Enabled = False
        cbospvname.Enabled = False
        Call ISIAGENT
        'remark asep20200617'
        'Call isistatuscall
    ElseIf MDIForm1.txtlevel.text = "Manager" Then
        Call GETSPV
        'Call isistatuscall
    End If
'    CmbRetryCall.clear
'    CmbRetryCall.AddItem " "
'    CmbRetryCall.AddItem "1x"
'    CmbRetryCall.AddItem "2x"
'    CmbRetryCall.AddItem "3x"
    tgl1.Value = Format(Now(), "yyyy-mm-dd")
    tgl2.Value = Format(Now(), "yyyy-mm-dd")
    jam1.Value = Format(Now(), "hh:mm")
    jam2.Value = Format(DateAdd("h", 2, Now), "hh:mm")
End Sub

Private Sub GETSPV()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
        
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        
        CMDSQL = "select tbluser_userid,tbluser_name from tbluser where tbluser_kdlevel='2' and tbluser_kdstatus='1' and tbluser_mgrcode ='" + MDIForm1.TxtUsername.text + "'"
        CMDSQL = CMDSQL + " order by tbluser_userid "
        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        cbospv.clear
        While Not M_objrs.EOF
             cbospv.AddItem cnull(M_objrs!tbluser_userid)
             cbospvname.AddItem cnull(M_objrs!tbluser_name)
            M_objrs.MoveNext
        Wend
        
        Set M_objrs = Nothing
End Sub

Private Sub ISIAGENT()
    Dim sQuery As String
    Dim Rs_Agent As ADODB.Recordset
    Dim Nomor As Double
    Dim list As ListItem
    If MDIForm1.txtlevel.text = "Supervisor" Then
        sQuery = "SELECT userid,agent FROM usertbl WHERE spvcode = '" + cbospv.text + "' AND aktif = '1' AND usertype = '1' order by userid "
    ElseIf MDIForm1.txtlevel.text = "Manager" Then
        sQuery = "SELECT userid,agent FROM usertbl WHERE tbluser_mgrcode = '" + MDIForm1.TxtUsername.text + "' AND tbluser_kdstatus = '1' AND tbluser_kdlevel = '2' order by tbluser_userid "
        'sQuery = "SELECT tbluser_userid,tbluser_name,tbluser_empid FROM tbluser WHERE tbluser_mgrcode in (select tbluser_userid from tbluser where tbluser_kdlevel='5' and f_cp='0' and tbluser_kdstatus='1') AND tbluser_kdstatus = '1' AND tbluser_kdlevel = '2' order by tbluser_userid "
    End If
    Set Rs_Agent = New ADODB.Recordset
    Rs_Agent.CursorLocation = adUseClient
    Rs_Agent.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    LVAgent.ListItems.clear
    
    If Rs_Agent.RecordCount > 0 Then
        While Not Rs_Agent.EOF
            Nomor = Nomor + 1
            Set list = LVAgent.ListItems.ADD(, , Nomor)
                list.SubItems(1) = Trim(Rs_Agent("userid"))
                list.SubItems(2) = Trim(Rs_Agent("agent"))
            Rs_Agent.MoveNext
        Wend
    End If
End Sub

Private Sub HeaderLvAgent()
'    LVAgent.ColumnHeaders.Clear
    LVAgent.ColumnHeaders.ADD 1, , "No", 600
    LVAgent.ColumnHeaders.ADD 2, , "AGENT", 5000
    LVAgent.ColumnHeaders.ADD 3, , "NAMA AGENT", 9000
    
    LVStatusCall.ColumnHeaders.clear
    LVStatusCall.ColumnHeaders.ADD 1, , "STATUS CALL", 5000
    
    lv3.ColumnHeaders.clear
    lv3.ColumnHeaders.ADD 1, , "No", 10 * 120
    lv3.ColumnHeaders.ADD 2, , "AGENT", 10 * 120
    lv3.ColumnHeaders.ADD 3, , "DATA", 10 * 120
    
    ListView1.ColumnHeaders.clear
    ListView1.ColumnHeaders.ADD 1, , "No", 10 * 120
    ListView1.ColumnHeaders.ADD 2, , "Autodialer Name", 1500
    ListView1.ColumnHeaders.ADD 3, , "Execute By", 1500
    ListView1.ColumnHeaders.ADD 4, , "Tanggal Start", 1500
    ListView1.ColumnHeaders.ADD 5, , "Tanggal End", 1500
    ListView1.ColumnHeaders.ADD 6, , "RetryCall", 1500
    ListView1.ColumnHeaders.ADD 7, , "RetryCall Ongoing", 1500
    ListView1.ColumnHeaders.ADD 8, , "Jumlah Data", 1500
    ListView1.ColumnHeaders.ADD 9, , "Agent", 1500
    
    ListView2.ColumnHeaders.clear
    ListView2.ColumnHeaders.ADD 1, , "No", 10 * 120
    ListView2.ColumnHeaders.ADD 2, , "Autodialer Name", 5000
    ListView2.ColumnHeaders.ADD 3, , "Execute By", 5000
    ListView2.ColumnHeaders.ADD 4, , "Tanggal Start", 5000
    ListView2.ColumnHeaders.ADD 5, , "Tanggal End", 5000
    ListView2.ColumnHeaders.ADD 6, , "RetryCall Ongoing", 5000
    ListView2.ColumnHeaders.ADD 7, , "Jumlah Data", 5000
    
End Sub

Private Sub isistatuscall()
    Dim sQuery As String
    Dim Rs_call As ADODB.Recordset
    Dim list2 As ListItem
        
    'sQuery = "select distinct(statuscall) as status_call from mgm where statuscall is not null and agent in (select userid from usertbl where spvcode = '" & cbospv.text & "')"
    sQuery = "select distinct(statuscall) as status_call from mgm where recsource='" + Combo1.text + "' and statuscall is not null and agent in (select userid from usertbl where spvcode = '" & cbospv.text & "')"
    
    Set Rs_call = New ADODB.Recordset
    Rs_call.CursorLocation = adUseClient
    Rs_call.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    LVStatusCall.ListItems.clear
    Set list2 = LVStatusCall.ListItems.ADD(, , "NEW DATA")
    If Rs_call.RecordCount > 0 Then
        While Not Rs_call.EOF
            Set list2 = LVStatusCall.ListItems.ADD(, , Rs_call("status_call"))
            Rs_call.MoveNext
        Wend
    End If
End Sub

Private Sub check_data()
    Dim stscall, campaign, custidx, outs, hit1, hit2 As String
    
    stscall = ""
    campaign = ""
    custidx = ""
    
    If Combo1.text = "" Then
        MsgBox "Harap Pilih Campaign", vbInformation + vbOKOnly, "Informasi"
        Exit Sub
    End If
    
'    a = 0
'    For i = 1 To LVStatusCall.ListItems.Count
'        If LVStatusCall.ListItems(i).Checked = True Then
'            a = a + 1
'           stscall = stscall & "'" & LVStatusCall.ListItems(i).text & "'"
'        End If
'    Next i
    a = 0
    i = 1
    For i = 1 To LVStatusCall.ListItems.Count
        If LVStatusCall.ListItems(i).Checked = True Then
            a = a + 1
            If LVStatusCall.ListItems(i).text = "NEW DATA" Then
            stscall = stscall & "'',"
            End If
           stscall = stscall & "'" & LVStatusCall.ListItems(i).text & "',"
        End If
    Next i
    
'    If a = 0 Then
'        MsgBox "Harap Pilih Status"
'        Exit Sub
'    End If
    
    If a = 0 Then
    Else
        stscall = Left(stscall, Len(stscall) - 1)
    'Else
    End If
        
    campaign = Combo1.text
    
    'qpub = "select v_cif from mgm where campaign_code ilike '%" & campaign & "%' and coalesce(agent,'') <> '' "
    '==================='
    qpub2 = "where recsource ='" & campaign & "'"
    '============================='
    qpub = "where recsource ='" & campaign & "'"
    
    'q = "select agent, count(id) from mgm where recsource = '" & campaign & "' and agent in (select distinct userid from usertbl where spvcode='" + MDIForm1.TxtUsername.text + "' AND usertype='1' and aktif='1') and coalesce(agent,'') <> '' "
    '20200514'
    q = "select agent, count(distinct(custid)) from mgm where recsource = '" & campaign & "' and agent in (select distinct userid from usertbl where spvcode='" + MDIForm1.TxtUsername.text + "' AND usertype='1' and aktif='1') and coalesce(agent,'') <> '' "
    
    If stscall <> "" Then
        q = q + " and coalesce(statuscall,'') in (" & stscall & ")"
        qpub = qpub + " and coalesce(statuscall,'')in (" & stscall & ")"
    End If
    
        
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open q & " group by agent", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    lv3.ListItems.clear
    
    While Not M_objrs.EOF
        Set ListItem = lv3.ListItems.ADD(, , M_objrs.Bookmark)
        ListItem.SubItems(1) = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
        ListItem.SubItems(2) = IIf(IsNull(M_objrs("count")), "", M_objrs("count"))
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    
End Sub

Private Sub runbtn_Click()
    Dim autodlnm As String
    Dim tgl1x As String
    Dim tgl2x As String
    Dim mida As String
    Dim CstrRetryCall As String
    Dim c As Integer
    
   CstrRetryCall = ""
'    Select Case CmbRetryCall.text
'        Case "1x"
'          CstrRetryCall = "1"
'        Case "2x"
'           CstrRetryCall = "2"
'        Case "3x"
'           CstrRetryCall = "3"
'    End Select
    
    CstrRetryCall = TDBNumber1.Value
    
    If Combo2.text <> "" Then
    Else
        MsgBox "Harap Pilih No Telepon"
        Exit Sub
    End If
    
    If lv3.SelectedItem.Checked = False Then
        MsgBox "harap pilih Agent"
        Exit Sub
    End If
    
    
    'agentt = Left(agentt, Len(agentt) - 1)
    
    If agentt <> "" Then
        qpub = qpub + " and agent in (" & agentt & ")"
    End If
    
    '=============='
    If agentt <> "" Then
        qpub2 = qpub2 + " and agent in (" & agentt & ")"
    End If
    '============='
    
    '====asep20200511===='
    If Combo3(0).text = "OUTSTANDING" Then
        qpub = qpub + " order by oustanding desc "
    End If
    
    If Combo3(0).text = "DPD" Then
        qpub = qpub + " order by delq_amt_by_x desc "
    End If
    
    If Combo3(1).text = "OUTSTANDING" Then
        qpub = qpub + " , oustanding desc "
    End If
    
    If Combo3(1).text = "DPD" Then
        qpub = qpub + " ,delq_amt_by_x desc "
    End If
    '==============='
    
    If tgl1.text = "__-__-____" Or tgl2.text = "__-__-____" Or jam1.text = "__:__" Or jam2.text = "__:__" Then
        MsgBox "Harap Pilih Waktu Auto Dial"
        Exit Sub
    End If
    
    tgl1x = Format(tgl1.Value, "yyyy-mm-dd") & " " & jam1.Value
    tgl2x = Format(tgl2.Value, "yyyy-mm-dd") & " " & jam2.Value
    
    autodlnm = cbospv.text + "_autodial_" + Format(FungsiWaktuServer, "DD-MM-YYYY HH:MM:SS")
    
    Dim asd As String
    
    For i = 1 To lv3.ListItems.Count
        If lv3.ListItems(i).Checked = True Then
      c = c + 1
            agentt = lv3.ListItems(i).SubItems(1)
            If CstrRetryCall = "" Then
                qins = "insert into tbl_autodialer_header (autodialer_name, executor, tgl_start, tgl_end, flag_running,agent) values "
                qins = qins & "('" & autodlnm & "', '" & MDIForm1.TxtUsername.text & "' , '" & tgl1x & "', '" & tgl2x & "','RUN','" & agentt & "'); "
                M_OBJCONN.Execute qins
            Else
                qins = "insert into tbl_autodialer_header (autodialer_name, executor, tgl_start, tgl_end, flag_running, retrycall,agent) values "
                qins = qins & "('" & autodlnm & "', '" & MDIForm1.TxtUsername.text & "' , '" & tgl1x & "', '" & tgl2x & "','RUN', '" + CstrRetryCall + "', '" + agentt + "' );"
                M_OBJCONN.Execute qins
            End If
                
           asd = "select max(id_header) as mid from tbl_autodialer_header"
            
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open asd, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    mida = IIf(IsNull(M_objrs("mid")), "", M_objrs("mid"))

    '======asep20200502====='
    Dim STRSQL As String
    If Combo2.text = "Handphone" Then
        STRSQL = "insert into tbl_autodialer (id_header,autodialer_name,id_cust,phone,tgl_start,tgl_end, agent)"
        STRSQL = STRSQL + " select id_header,autodialer_name,custid,mobileno,tgl_start,tgl_end,agent from("
        STRSQL = STRSQL + " select distinct '" + mida + "'::integer as id_header,'" + autodlnm + "' as autodialer_name, '" + tgl1x + "'::timestamp as tgl_start,  '" + tgl2x + "'::timestamp as tgl_end,"
        STRSQL = STRSQL + " custid::integer, mobileno,agent from mgm " & qpub2 & " and agent= '" + agentt + "' and "
        STRSQL = STRSQL + " custid in(select custid from mgm " & qpub & "))a"
        M_OBJCONN.Execute STRSQL
    ElseIf Combo2.text = "Home Phone" Then
        STRSQL = "insert into tbl_autodialer (id_header,autodialer_name,id_cust,phone,tgl_start,tgl_end, agent)"
        STRSQL = STRSQL + " select id_header,autodialer_name,custid,homenoadd1,tgl_start,tgl_end,agent from("
        STRSQL = STRSQL + " select distinct '" + mida + "'::integer as id_header,'" + autodlnm + "' as autodialer_name, '" + tgl1x + "'::timestamp as tgl_start,  '" + tgl2x + "'::timestamp as tgl_end,"
        STRSQL = STRSQL + " custid::integer, homenoadd1,agent from mgm " & qpub2 & "  and agent= '" + agentt + "' and"
        STRSQL = STRSQL + " custid in(select custid from mgm " & qpub & "))a"
        M_OBJCONN.Execute STRSQL
    ElseIf Combo2.text = "Office Phone" Then
        STRSQL = "insert into tbl_autodialer (id_header,autodialer_name,id_cust,phone,tgl_start,tgl_end, agent)"
        STRSQL = STRSQL + " select id_header,autodialer_name,custid,officenoadd1,tgl_start,tgl_end,agent from("
        STRSQL = STRSQL + " select distinct '" + mida + "'::integer as id_header,'" + autodlnm + "' as autodialer_name, '" + tgl1x + "'::timestamp as tgl_start,  '" + tgl2x + "'::timestamp as tgl_end,"
        STRSQL = STRSQL + " custid::integer, officenoadd1,agent from mgm " & qpub2 & "  and agent= '" + agentt + "' and"
        STRSQL = STRSQL + " custid in(select custid from mgm " & qpub & "))a"
        M_OBJCONN.Execute STRSQL
    End If
                
        End If
    Next i

    MsgBox "Auto Dial Sudah Di Process"
    runbtn.Enabled = False
End Sub

Private Sub cmd_check_running_Click()
    Set M_objrs = New ADODB.Recordset
    'wer = " select distinct(autodialer_name),executor, tgl_start, tgl_end,retrycall,retrycall_ongoing from tbl_autodialer_header  where tgl_end > now() and flag_running='RUN' group by autodialer_name,executor,tgl_start,tgl_end "
    wer = "select a.autodialer_name,a.id_header,executor, a.tgl_start, a.tgl_end,retrycall,a.retrycall_ongoing,b.jumlahdata,agent from"
    wer = wer + "(select distinct(autodialer_name),id_header,executor, tgl_start, tgl_end,retrycall,retrycall_ongoing,agent from tbl_autodialer_header  where tgl_end > now() and flag_running='RUN' group by autodialer_name,executor,tgl_start,tgl_end,id_header,retrycall,retrycall_ongoing,jumlah_data,agent )a"
    wer = wer + " left join ("
    wer = wer + " select id_header,count(id_header) as jumlahdata from tbl_autodialer group by id_header)b on a.id_header = b.id_header"
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open wer, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    ListView1.ListItems.clear
    
    While Not M_objrs.EOF
        Set ListItem = ListView1.ListItems.ADD(, , M_objrs.Bookmark)
        ListItem.SubItems(1) = IIf(IsNull(M_objrs("autodialer_name")), "", M_objrs("autodialer_name"))
        ListItem.SubItems(2) = IIf(IsNull(M_objrs("executor")), "", M_objrs("executor"))
        ListItem.SubItems(3) = IIf(IsNull(M_objrs("tgl_start")), "", M_objrs("tgl_start"))
        ListItem.SubItems(4) = IIf(IsNull(M_objrs("tgl_end")), "", M_objrs("tgl_end"))
        ListItem.SubItems(5) = IIf(IsNull(M_objrs("retrycall")), "", M_objrs("retrycall"))
        ListItem.SubItems(6) = IIf(IsNull(M_objrs("retrycall_ongoing")), "", M_objrs("retrycall_ongoing"))
        ListItem.SubItems(7) = Format(IIf(IsNull(M_objrs("jumlahdata")), "0", M_objrs("jumlahdata")))
        ListItem.SubItems(8) = Format(IIf(IsNull(M_objrs("agent")), "0", M_objrs("agent")))
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
End Sub

Private Sub cmd_stop_Click()
    Dim aSql As String
    ac = 0
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked = True Then
            ac = ac + 1
           auto_name = auto_name & "'" & ListView1.ListItems(i).SubItems(1) & "',"
        End If
    Next i
    
    If ac = 0 Then
        MsgBox "Harap Pilih Auto Dial Yang Akan Diakhiri"
        Exit Sub
    End If
    auto_name = Left(auto_name, Len(auto_name) - 1)
    
    If auto_name <> "" Then
        aSql = " where autodialer_name in (" & auto_name & ")"
    End If
    
    If MsgBox("Apakah anda yakin akan mengakhiri auto dial yang dipilih?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        er1 = "delete from tbl_autodialer " & aSql & " ;"
        er1 = er1 + "update tbl_autodialer_header set flag_running = 'STOP' " & aSql & " ;"
        M_OBJCONN.Execute er1
        MsgBox "Auto Dial Berhasil Dihapus", vbOKOnly + vbInformation, "Informasi"
    Else
        Exit Sub
    End If
End Sub
Private Sub cmd_schedule_Click()
    Dim openn As String
    Set M_objrs = New ADODB.Recordset
    'wer = " select autodialer_name, executor, tgl_start, tgl_end,flag_running from tbl_autodialer_header where tgl_start >= date(now())-5"
    wer = " select autodialer_name, executor, tgl_start, tgl_end,flag_running,retrycall , retrycall_ongoing from tbl_autodialer_header where tgl_start >= date(now())-5"
    'group by autodialer_name, executor, tgl_start,tgl_end "
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open wer, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    ListView2.ListItems.clear
    
    While Not M_objrs.EOF
        openn = IIf(IsNull(M_objrs("flag_running")), "", M_objrs("flag_running"))
        If openn = "RUN" Then
            Set ListItem = ListView2.ListItems.ADD(, , M_objrs.Bookmark)
            ListItem.ForeColor = vbBlue
            ListItem.SubItems(1) = IIf(IsNull(M_objrs("autodialer_name")), "", M_objrs("autodialer_name"))
            ListItem.ListSubItems(1).ForeColor = vbBlue
            ListItem.SubItems(2) = IIf(IsNull(M_objrs("executor")), "", M_objrs("executor"))
            ListItem.ListSubItems(2).ForeColor = vbBlue
            ListItem.SubItems(3) = IIf(IsNull(M_objrs("tgl_start")), "", M_objrs("tgl_start"))
            ListItem.ListSubItems(3).ForeColor = vbBlue
            ListItem.SubItems(4) = IIf(IsNull(M_objrs("tgl_end")), "", M_objrs("tgl_end"))
            ListItem.ListSubItems(4).ForeColor = vbBlue
            ListItem.SubItems(5) = IIf(IsNull(M_objrs("retrycall")), "", M_objrs("retrycall"))
            ListItem.ListSubItems(5).ForeColor = vbBlue
            ListItem.SubItems(6) = IIf(IsNull(M_objrs("retrycall_ongoing")), "", M_objrs("retrycall_ongoing"))
            ListItem.ListSubItems(6).ForeColor = vbBlue
        Else
            Set ListItem = ListView2.ListItems.ADD(, , M_objrs.Bookmark)
            ListItem.SubItems(1) = IIf(IsNull(M_objrs("autodialer_name")), "", M_objrs("autodialer_name"))
            ListItem.SubItems(2) = IIf(IsNull(M_objrs("executor")), "", M_objrs("executor"))
            ListItem.SubItems(3) = IIf(IsNull(M_objrs("tgl_start")), "", M_objrs("tgl_start"))
            ListItem.SubItems(4) = IIf(IsNull(M_objrs("tgl_end")), "", M_objrs("tgl_end"))
            ListItem.SubItems(5) = IIf(IsNull(M_objrs("retrycall")), "", M_objrs("retrycall"))
            ListItem.SubItems(6) = IIf(IsNull(M_objrs("retrycall_ongoing")), "", M_objrs("retrycall_ongoing"))
        End If
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
End Sub
