VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form VIEW_MGMDATA 
   BackColor       =   &H00E6E6E6&
   ClientHeight    =   10020
   ClientLeft      =   495
   ClientTop       =   825
   ClientWidth     =   19755
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   Icon            =   "VIEW_MGMDATAold.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10020
   ScaleWidth      =   19755
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FCFCFC&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   10725
      Left            =   0
      TabIndex        =   21
      Top             =   15
      Width           =   19755
      Begin VB.Frame Frame9 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Caption         =   "Frame9"
         Height          =   10695
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Width           =   5235
         Begin MSComctlLib.ProgressBar ProgressBar2 
            Height          =   375
            Left            =   960
            TabIndex        =   105
            Top             =   7680
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
         End
         Begin VB.ComboBox Combo4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Index           =   1
            ItemData        =   "VIEW_MGMDATAold.frx":000C
            Left            =   1560
            List            =   "VIEW_MGMDATAold.frx":000E
            TabIndex        =   103
            Top             =   4005
            Width           =   3405
         End
         Begin VB.ComboBox Combo4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Index           =   0
            ItemData        =   "VIEW_MGMDATAold.frx":0010
            Left            =   1560
            List            =   "VIEW_MGMDATAold.frx":001A
            TabIndex        =   101
            Top             =   3630
            Width           =   3405
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   530
            Left            =   1080
            Picture         =   "VIEW_MGMDATAold.frx":0030
            ScaleHeight     =   525
            ScaleWidth      =   525
            TabIndex        =   99
            Top             =   240
            Width           =   530
         End
         Begin VB.ComboBox Combo99 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            ItemData        =   "VIEW_MGMDATAold.frx":2827
            Left            =   1560
            List            =   "VIEW_MGMDATAold.frx":2843
            TabIndex        =   97
            Top             =   6930
            Visible         =   0   'False
            Width           =   3420
         End
         Begin VB.CommandButton Command2 
            Caption         =   "EXPORT"
            Height          =   495
            Left            =   3240
            TabIndex        =   96
            Top             =   6330
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txtregion 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   1560
            TabIndex        =   94
            Top             =   3270
            Width           =   3390
         End
         Begin VB.ComboBox cmb_nmagent 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "VIEW_MGMDATAold.frx":2880
            Left            =   2970
            List            =   "VIEW_MGMDATAold.frx":2882
            TabIndex        =   93
            Top             =   1920
            Width           =   2010
         End
         Begin VB.ComboBox cmb_kdagent 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "VIEW_MGMDATAold.frx":2884
            Left            =   1560
            List            =   "VIEW_MGMDATAold.frx":2886
            TabIndex        =   92
            Top             =   1920
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   0
            Left            =   1560
            TabIndex        =   51
            Top             =   1485
            Width           =   3390
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Index           =   0
            Left            =   1905
            TabIndex        =   50
            Top             =   10275
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Index           =   1
            Left            =   3180
            TabIndex        =   49
            Top             =   10275
            Visible         =   0   'False
            Width           =   2130
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Index           =   2
            Left            =   1560
            TabIndex        =   48
            Top             =   2820
            Width           =   3420
         End
         Begin VB.TextBox txtnocard 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   47
            Top             =   1050
            Width           =   3390
         End
         Begin VB.ComboBox Combo2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "VIEW_MGMDATAold.frx":2888
            Left            =   1560
            List            =   "VIEW_MGMDATAold.frx":288A
            TabIndex        =   46
            Top             =   2370
            Width           =   3420
         End
         Begin VB.TextBox txtamount 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   4395
            Visible         =   0   'False
            Width           =   3390
         End
         Begin VB.TextBox txtcurbalance 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   44
            Top             =   4815
            Visible         =   0   'False
            Width           =   3390
         End
         Begin MSComDlg.CommonDialog CD_save 
            Left            =   0
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComctlLib.ListView lstautodial 
            Height          =   270
            Left            =   60
            TabIndex        =   100
            Top             =   180
            Visible         =   0   'False
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   476
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   33023
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   0
            MousePointer    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Order By Ke 2"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00333333&
            Height          =   255
            Index           =   19
            Left            =   0
            TabIndex        =   104
            Top             =   4020
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Order By Ke 1"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00333333&
            Height          =   255
            Index           =   18
            Left            =   0
            TabIndex        =   102
            Top             =   3645
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Bank/Fintech"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00333333&
            Height          =   255
            Index           =   17
            Left            =   0
            TabIndex        =   98
            Top             =   6930
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Region"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00333333&
            Height          =   300
            Index           =   16
            Left            =   0
            TabIndex        =   95
            Top             =   3300
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "SEARCH FILTER"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00777777&
            Height          =   375
            Left            =   1800
            TabIndex        =   91
            Top             =   285
            Width           =   3375
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C1C1C1&
            X1              =   240
            X2              =   4920
            Y1              =   840
            Y2              =   840
         End
         Begin Threed.SSCommand Command1 
            Height          =   675
            Index           =   0
            Left            =   1800
            TabIndex        =   53
            Top             =   5505
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1191
            _Version        =   196610
            ForeColor       =   4210752
            PictureFrames   =   1
            Windowless      =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "VIEW_MGMDATAold.frx":288C
            Caption         =   "&"
            ButtonStyle     =   2
            BevelWidth      =   0
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cust No"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00333333&
            Height          =   300
            Index           =   5
            Left            =   0
            TabIndex        =   57
            Top             =   1065
            Width           =   1335
         End
         Begin Threed.SSCommand Command1 
            Height          =   675
            Index           =   2
            Left            =   2880
            TabIndex        =   90
            Top             =   5520
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1191
            _Version        =   196610
            ForeColor       =   4210752
            PictureFrames   =   1
            Windowless      =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "VIEW_MGMDATAold.frx":5FF9
            Caption         =   "&"
            ButtonStyle     =   2
            BevelWidth      =   0
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00333333&
            Height          =   300
            Index           =   0
            Left            =   0
            TabIndex        =   60
            Top             =   1515
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Agent"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00333333&
            Height          =   330
            Index           =   1
            Left            =   0
            TabIndex        =   59
            Top             =   1935
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Campaign"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00333333&
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   58
            Top             =   2820
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Statuscall"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00333333&
            Height          =   345
            Index           =   8
            Left            =   0
            TabIndex        =   56
            Top             =   2445
            Width           =   1335
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00333333&
            Height          =   225
            Left            =   0
            TabIndex        =   55
            Top             =   4425
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Curr Balance"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00333333&
            Height          =   225
            Left            =   0
            TabIndex        =   54
            Top             =   4860
            Visible         =   0   'False
            Width           =   1335
         End
         Begin Threed.SSCommand Command1 
            Height          =   675
            Index           =   1
            Left            =   3960
            TabIndex        =   52
            Top             =   5520
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1191
            _Version        =   196610
            ForeColor       =   4210752
            PictureFrames   =   1
            Windowless      =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "VIEW_MGMDATAold.frx":91B7
            Caption         =   "&"
            ButtonStyle     =   2
            BevelWidth      =   0
         End
         Begin VB.Image Image1 
            Height          =   25380
            Left            =   0
            Picture         =   "VIEW_MGMDATAold.frx":C1DD
            Top             =   -735
            Width           =   5730
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Sampah"
         Height          =   1095
         Left            =   120
         TabIndex        =   23
         Top             =   7920
         Visible         =   0   'False
         Width           =   2655
         Begin VB.Frame Frame3 
            Caption         =   "Proses....!!"
            Height          =   615
            Left            =   0
            TabIndex        =   88
            Top             =   0
            Visible         =   0   'False
            Width           =   2025
            Begin MSComctlLib.ProgressBar ProgressBar1 
               Height          =   270
               Left            =   15
               TabIndex        =   89
               Top             =   0
               Width           =   1980
               _ExtentX        =   3493
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
            End
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Index           =   3
            Left            =   180
            TabIndex        =   81
            Top             =   1500
            Visible         =   0   'False
            Width           =   2580
         End
         Begin VB.ComboBox CmbStatusCek 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3225
            TabIndex        =   80
            Top             =   2685
            Width           =   1800
         End
         Begin VB.ComboBox cmbStsLastCall 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   480
            TabIndex        =   79
            Top             =   2700
            Visible         =   0   'False
            Width           =   3180
         End
         Begin VB.TextBox TxtJmlDtMgm 
            Height          =   375
            Left            =   1080
            TabIndex        =   78
            Text            =   "Text4"
            Top             =   4020
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox TxtJmlVolMgm 
            Height          =   375
            Left            =   2280
            TabIndex        =   77
            Text            =   "Text4"
            Top             =   5220
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Frame Frame5 
            Caption         =   "Status CPA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   120
            TabIndex        =   32
            Top             =   1920
            Width           =   2445
            Begin VB.ComboBox Combo3 
               Height          =   315
               ItemData        =   "VIEW_MGMDATAold.frx":11797
               Left            =   45
               List            =   "VIEW_MGMDATAold.frx":117A1
               Style           =   2  'Dropdown List
               TabIndex        =   33
               Top             =   180
               Width           =   2310
            End
         End
         Begin VB.CommandButton Cmd_listrequestdecease 
            BackColor       =   &H0080FFFF&
            Caption         =   "&List Request Acc Decease"
            Height          =   375
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   1680
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox txtalamat 
            Height          =   285
            Left            =   1005
            MaxLength       =   200
            TabIndex        =   30
            Top             =   1125
            Visible         =   0   'False
            Width           =   2325
         End
         Begin VB.TextBox TDBMask1 
            Height          =   285
            Left            =   1005
            TabIndex        =   29
            Top             =   840
            Visible         =   0   'False
            Width           =   2325
         End
         Begin VB.CommandButton CmdSearchPTP 
            Caption         =   "Search Tgl.Tagih"
            Height          =   615
            Left            =   0
            TabIndex        =   28
            Top             =   465
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton CmdReleaseUnRealesePTP 
            Caption         =   "Realese/ UnRealese PTP"
            Height          =   615
            Left            =   1020
            TabIndex        =   27
            Top             =   465
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton CmdListHotProsPect 
            Caption         =   "&List Hot Prospect"
            Height          =   375
            Left            =   0
            TabIndex        =   26
            Top             =   960
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.CommandButton cmd_review 
            Caption         =   "LIST REVIEW"
            Height          =   375
            Left            =   0
            TabIndex        =   25
            Top             =   1080
            Visible         =   0   'False
            Width           =   2470
         End
         Begin VB.CommandButton cmd_claimback_acc 
            Caption         =   "BATAL CLAIM ACCOUNT"
            Height          =   375
            Left            =   0
            TabIndex        =   24
            Top             =   1485
            Visible         =   0   'False
            Width           =   2470
         End
         Begin TDBDate6Ctl.TDBDate TdDob 
            Height          =   315
            Left            =   1080
            TabIndex        =   82
            Top             =   2100
            Visible         =   0   'False
            Width           =   1170
            _Version        =   65536
            _ExtentX        =   2064
            _ExtentY        =   556
            Calendar        =   "VIEW_MGMDATAold.frx":117B0
            Caption         =   "VIEW_MGMDATAold.frx":118C8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "VIEW_MGMDATAold.frx":11934
            Keys            =   "VIEW_MGMDATAold.frx":11952
            Spin            =   "VIEW_MGMDATAold.frx":119B0
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   16777215
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
            Format          =   "dd-mmm-yyyy"
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
            Text            =   "__-___-____"
            ValidateMode    =   0
            ValueVT         =   1
            Value           =   37475
            CenturyMode     =   0
         End
         Begin TDBMask6Ctl.TDBMask TDBMask2 
            Height          =   315
            Left            =   0
            TabIndex        =   83
            Top             =   0
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   556
            Caption         =   "VIEW_MGMDATAold.frx":119D8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "VIEW_MGMDATAold.frx":11A44
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   16777215
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   0
            Format          =   "9999-999999999999999999"
            HighlightText   =   0
            IMEMode         =   0
            IMEStatus       =   0
            LookupMode      =   0
            LookupTable     =   ""
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MousePointer    =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            PromptChar      =   "_"
            ReadOnly        =   0
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "____-__________________"
            Value           =   ""
         End
         Begin TDBTime6Ctl.TDBTime DTimeLastCall 
            Height          =   300
            Index           =   0
            Left            =   780
            TabIndex        =   84
            Top             =   1920
            Visible         =   0   'False
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1455
            _ExtentY        =   529
            Caption         =   "VIEW_MGMDATAold.frx":11A86
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "VIEW_MGMDATAold.frx":11AF2
            Spin            =   "VIEW_MGMDATAold.frx":11B42
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   16777215
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
            Text            =   "00:00"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   6.13425925925926E-04
         End
         Begin TDBDate6Ctl.TDBDate DtLastCall 
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   85
            Top             =   2160
            Visible         =   0   'False
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   556
            Calendar        =   "VIEW_MGMDATAold.frx":11B6A
            Caption         =   "VIEW_MGMDATAold.frx":11C82
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "VIEW_MGMDATAold.frx":11CEE
            Keys            =   "VIEW_MGMDATAold.frx":11D0C
            Spin            =   "VIEW_MGMDATAold.frx":11D6A
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   16777215
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
            Format          =   "dd-mmm-yyyy"
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
            Text            =   "__-___-____"
            ValidateMode    =   0
            ValueVT         =   1
            Value           =   37468
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate DtLastCall 
            Height          =   315
            Index           =   1
            Left            =   1500
            TabIndex        =   86
            Top             =   1860
            Visible         =   0   'False
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   556
            Calendar        =   "VIEW_MGMDATAold.frx":11D92
            Caption         =   "VIEW_MGMDATAold.frx":11EAA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "VIEW_MGMDATAold.frx":11F16
            Keys            =   "VIEW_MGMDATAold.frx":11F34
            Spin            =   "VIEW_MGMDATAold.frx":11F92
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   16777215
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
            Format          =   "dd-mmm-yyyy"
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
            Text            =   "__-___-____"
            ValidateMode    =   0
            ValueVT         =   1
            Value           =   37468
            CenturyMode     =   0
         End
         Begin TDBTime6Ctl.TDBTime DTimeLastCall 
            Height          =   300
            Index           =   1
            Left            =   2280
            TabIndex        =   87
            Top             =   1920
            Visible         =   0   'False
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1455
            _ExtentY        =   529
            Caption         =   "VIEW_MGMDATAold.frx":11FBA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "VIEW_MGMDATAold.frx":12026
            Spin            =   "VIEW_MGMDATAold.frx":12076
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   16777215
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
            Text            =   "20:53"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   0.870289351851852
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   15
            Left            =   -240
            TabIndex        =   35
            Top             =   1125
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Telp "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   4
            Left            =   -240
            TabIndex        =   34
            Top             =   840
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin VB.Timer Timer1 
         Left            =   12600
         Top             =   0
      End
      Begin VB.CheckBox CekDtDistribute 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Searching Data Belum Distribute"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   150
         MaskColor       =   &H000000FF&
         TabIndex        =   22
         Top             =   -345
         Visible         =   0   'False
         Width           =   3225
      End
      Begin MSComctlLib.ListView LstVwSearchMgm 
         Height          =   8925
         Left            =   5280
         TabIndex        =   61
         Top             =   0
         Width           =   13920
         _ExtentX        =   24553
         _ExtentY        =   15743
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
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
      Begin VB.Frame Frame8 
         BackColor       =   &H00FCFCFC&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   5280
         TabIndex        =   62
         Top             =   9015
         Width           =   13935
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   10920
            TabIndex        =   71
            Top             =   30
            Width           =   855
         End
         Begin VB.TextBox txttotal 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   12780
            TabIndex        =   70
            Top             =   0
            Width           =   975
         End
         Begin VB.TextBox txtjmllimit 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   9030
            TabIndex        =   69
            Top             =   0
            Width           =   885
         End
         Begin VB.TextBox txtcountpage 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2850
            Locked          =   -1  'True
            TabIndex        =   68
            Text            =   "0"
            Top             =   45
            Width           =   540
         End
         Begin VB.TextBox txtpage 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2010
            TabIndex        =   67
            Text            =   "1"
            Top             =   45
            Width           =   510
         End
         Begin VB.CommandButton cmd 
            Caption         =   "<<"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   3
            Left            =   120
            TabIndex        =   66
            Top             =   15
            Width           =   720
         End
         Begin VB.CommandButton cmd 
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   2
            Left            =   4410
            TabIndex        =   65
            Top             =   15
            Width           =   720
         End
         Begin VB.CommandButton cmd 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   1
            Left            =   840
            TabIndex        =   64
            Top             =   15
            Width           =   675
         End
         Begin VB.CommandButton cmd 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   0
            Left            =   3630
            TabIndex        =   63
            Top             =   15
            Width           =   720
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Row :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   11865
            TabIndex        =   76
            Top             =   45
            Width           =   885
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Jml Row :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   10080
            TabIndex        =   75
            Top             =   15
            Width           =   705
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Of"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2610
            TabIndex        =   74
            Top             =   75
            Width           =   315
         End
         Begin VB.Label Label16 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Page"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   1635
            TabIndex        =   73
            Top             =   75
            Width           =   435
         End
         Begin VB.Label Label31 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Row Per Page :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Left            =   8115
            TabIndex        =   72
            Top             =   0
            Width           =   1110
         End
      End
      Begin VB.Label LBLCOUNT 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   12300
         TabIndex        =   42
         Top             =   990
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status Account"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   14
         Left            =   7440
         TabIndex        =   41
         Top             =   3240
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Terakhir Telp."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   9
         Left            =   5640
         TabIndex        =   40
         Top             =   2640
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status Check"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   7320
         TabIndex        =   39
         Top             =   3480
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   11
         Left            =   6720
         TabIndex        =   38
         Top             =   2820
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "HP :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   6
         Left            =   5430
         TabIndex        =   37
         Top             =   210
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Lahir :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   2
         Left            =   5520
         TabIndex        =   36
         Top             =   2760
         Visible         =   0   'False
         Width           =   1050
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   3210
      Left            =   120
      TabIndex        =   5
      Top             =   7200
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5662
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   706
      BackColor       =   16183524
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Find"
      TabPicture(0)   =   "VIEW_MGMDATAold.frx":1209E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Schedulle"
      TabPicture(1)   =   "VIEW_MGMDATAold.frx":1455C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H00B1FDD5&
         ForeColor       =   &H80000008&
         Height          =   1425
         Left            =   -74985
         TabIndex        =   6
         Top             =   600
         Width           =   18930
         Begin VB.CheckBox Check2 
            Caption         =   "MGM Data"
            BeginProperty DataFormat 
               Type            =   4
               Format          =   "M/d/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   8
            EndProperty
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   4065
            TabIndex        =   1
            Top             =   1470
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   5
            Left            =   3360
            TabIndex        =   0
            Top             =   150
            Width           =   3480
         End
         Begin Threed.SSFrame SSFrame1 
            Height          =   510
            Left            =   1500
            TabIndex        =   7
            Top             =   540
            Width           =   3390
            _ExtentX        =   5980
            _ExtentY        =   900
            _Version        =   196610
            BackColor       =   -2147483644
            BackStyle       =   1
            Begin TDBDate6Ctl.TDBDate TDBDate1 
               Height          =   315
               Index           =   0
               Left            =   75
               TabIndex        =   2
               Top             =   60
               Width           =   1125
               _Version        =   65536
               _ExtentX        =   1984
               _ExtentY        =   556
               Calendar        =   "VIEW_MGMDATAold.frx":16B29
               Caption         =   "VIEW_MGMDATAold.frx":16C41
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "VIEW_MGMDATAold.frx":16CAD
               Keys            =   "VIEW_MGMDATAold.frx":16CCB
               Spin            =   "VIEW_MGMDATAold.frx":16D29
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
               ForeColor       =   -2147483640
               Format          =   "dd-mmm-yyyy"
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
               Text            =   "__-___-____"
               ValidateMode    =   0
               ValueVT         =   1
               Value           =   37609
               CenturyMode     =   0
            End
            Begin TDBDate6Ctl.TDBDate TDBDate1 
               Height          =   315
               Index           =   1
               Left            =   1830
               TabIndex        =   3
               Top             =   90
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               Calendar        =   "VIEW_MGMDATAold.frx":16D51
               Caption         =   "VIEW_MGMDATAold.frx":16E69
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "VIEW_MGMDATAold.frx":16ED5
               Keys            =   "VIEW_MGMDATAold.frx":16EF3
               Spin            =   "VIEW_MGMDATAold.frx":16F51
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
               ForeColor       =   -2147483640
               Format          =   "dd-mmm-yyyy"
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
               Text            =   "__-___-____"
               ValidateMode    =   0
               ValueVT         =   1
               Value           =   37609
               CenturyMode     =   0
            End
            Begin VB.Label Label1 
               BackColor       =   &H80000004&
               BackStyle       =   0  'Transparent
               Caption         =   "S/d"
               ForeColor       =   &H00000000&
               Height          =   300
               Index           =   7
               Left            =   1365
               TabIndex        =   8
               Top             =   150
               Width           =   285
            End
         End
         Begin Threed.SSCommand CmdScheduleoK 
            Height          =   690
            Left            =   5160
            TabIndex        =   4
            Top             =   540
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   1217
            _Version        =   196610
            Font3D          =   5
            MousePointer    =   16
            BackColor       =   -2147483644
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Search"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand CmdToday 
            Height          =   480
            Left            =   120
            TabIndex        =   11
            Top             =   1920
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   847
            _Version        =   196610
            Font3D          =   5
            MousePointer    =   16
            BackColor       =   -2147483644
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Today"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand CmdMissed 
            Height          =   450
            Left            =   150
            TabIndex        =   12
            Top             =   2475
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   794
            _Version        =   196610
            Font3D          =   5
            MousePointer    =   16
            BackColor       =   -2147483644
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Miss"
            ButtonStyle     =   2
         End
         Begin VB.ComboBox Combo1 
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   4
            Left            =   1485
            TabIndex        =   9
            Top             =   165
            Width           =   1905
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl Schedule :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   12
            Left            =   240
            TabIndex        =   13
            Top             =   600
            Width           =   1170
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Agent :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   13
            Left            =   180
            TabIndex        =   10
            Top             =   195
            Width           =   1170
         End
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3555
      Left            =   120
      TabIndex        =   14
      Top             =   7200
      Visible         =   0   'False
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   6271
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Card Holder Data"
      TabPicture(0)   =   "VIEW_MGMDATAold.frx":16F79
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LblTarget(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Check1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Referall Data"
      TabPicture(1)   =   "VIEW_MGMDATAold.frx":16F95
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "LblTarget(1)"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame4 
         BackColor       =   &H80000004&
         Height          =   8685
         Left            =   -74940
         TabIndex        =   16
         Top             =   345
         Width           =   15075
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   11910
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   7980
            Width           =   3045
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   7845
            Left            =   0
            TabIndex        =   18
            Top             =   120
            Width           =   15015
            _ExtentX        =   26485
            _ExtentY        =   13838
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   12582912
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
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MGM Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   3090
         MaskColor       =   &H000000FF&
         TabIndex        =   15
         Top             =   1950
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   3225
      End
      Begin VB.Label LblTarget 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Index           =   1
         Left            =   -71820
         TabIndex        =   20
         Top             =   -15
         Visible         =   0   'False
         Width           =   9465
      End
      Begin VB.Label LblTarget 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Index           =   0
         Left            =   3315
         TabIndex        =   19
         Top             =   180
         Visible         =   0   'False
         Width           =   4605
      End
   End
   Begin VB.Menu MnFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnclaim 
         Caption         =   "Claim"
      End
   End
End
Attribute VB_Name = "VIEW_MGMDATA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_cari2 As ADODB.Recordset
Dim m_cari As ADODB.Recordset
    Dim mrs_cek As ADODB.Recordset
Dim StsVl  As String
Dim StsPR As String
Dim StsSK As String
Dim StsON As String
Dim StsOS As String
Dim StsPTP As String
Dim StsBP As String
Dim StsPOP As String
Dim StsSP As String
Dim StsUC As String
Dim StsAll As String
Dim Stsblank As String
Dim Stsf_fresh As String
Dim StsRP As String
Dim StsWO_Date As String
Dim StsWO_2009 As String
Dim StsWO_2008 As String
Dim StsWO_2007 As String
Dim StsWO_2006 As String
Dim StsWO_2005 As String
Dim StsWO_2004 As String
Dim StsWO_2003 As String
Dim StsWO_2002 As String
Dim StsWO_2001 As String
Dim StsWO_2000 As String
Dim StsWO_1999 As String
Dim StsWO_2010 As String
Dim CMDSQL As String
Dim Bloked As String
Dim LUserType As String
Dim F_CEK As String
Dim WO_DATE As String
Dim f_Pending As String
Dim datajml As Double
'@@ 14072010 Blok entry list
Dim BlokedEntry As String
Dim jmlpage As String
Dim totalrows As New ADODB.Recordset
Dim IndexColumnHEader As Integer
Dim opt_hide_header() As Integer
Public qpub  As String

Private Sub HEADER_VIEW_Refferall()
    ListView1.ColumnHeaders.ADD 1, , "No", 3 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Cust Id", 5 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Priority", 1
    ListView1.ColumnHeaders.ADD 4, , "Ref Id", 10 * TXT
    ListView1.ColumnHeaders.ADD 5, , "Ref Name", 10 * TXT
    ListView1.ColumnHeaders.ADD 6, , "Nama Customer", 25 * TXT
    'ListView1.ColumnHeaders.ADD 7, , "Batch Expire Date", 25 * TXT
    ListView1.ColumnHeaders.ADD 7, , "Tgl Schedule", 10 * TXT
    ListView1.ColumnHeaders.ADD 8, , "Next Action", 12 * TXT
    ListView1.ColumnHeaders.ADD 9, , "Remarks", 17 * TXT
    ListView1.ColumnHeaders.ADD 10, , "Sts LastCall", 17 * TXT
    ListView1.ColumnHeaders.ADD 11, , "SalesCode", 8 * TXT
    ListView1.ColumnHeaders.ADD 12, , "Agent", 8 * TXT
    ListView1.ColumnHeaders.ADD 13, , "DataBase", 10 * TXT
    ListView1.ColumnHeaders.ADD 14, , "LastCall Date", 10 * TXT
    ListView1.ColumnHeaders.ADD 15, , "Code", 5 * TXT
    ListView1.ColumnHeaders.ADD 16, , "Complaint Note", 15 * TXT
    ListView1.ColumnHeaders.ADD 17, , "Check", 10 * TXT
    ListView1.ColumnHeaders.ADD 18, , "ID", 10 * TXT
End Sub
Private Sub isi_dataClaimKeGrid(gCUSTID As String, gNama As String, gnextact As String, gremarks As String, gagent As String, gnamaagent As String, grecsource As String)
    ' insert ke grid list view
Dim ListItem As ListItem
Set ListItem = LstVwSearchMgm.ListItems.ADD(, , "9999")
    ListItem.SubItems(1) = gCUSTID
    ListItem.SubItems(2) = ""
    ListItem.SubItems(3) = gNama
    ListItem.SubItems(4) = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd hh:nn")
    ListItem.SubItems(5) = gnextact
    ListItem.SubItems(6) = gremarks
    ListItem.SubItems(7) = gagent
    ListItem.SubItems(8) = gnamaagent
    ListItem.SubItems(9) = grecsource
    ListItem.SubItems(10) = ""
    ListItem.SubItems(11) = "1A"
    ListItem.SubItems(12) = ""
    ListItem.SubItems(13) = ""
    ListItem.SubItems(14) = ""
End Sub
Private Sub cmb_kdagent_Click()
    cmb_nmagent.ListIndex = cmb_kdagent.ListIndex
End Sub

Private Sub cmb_kdagent_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmb_nmagent_Click()
    cmb_kdagent.ListIndex = cmb_nmagent.ListIndex
End Sub

Private Sub cmb_nmagent_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmd_claimback_acc_Click()
    'Form_return_claim.Show 1
End Sub

Private Sub Cmd_listrequestdecease_Click()
    'Form_listreq_decease.Show 1
End Sub

Private Sub cmd_review_Click()
    FrmCustIdReview.Show
End Sub

Private Sub CmdListHotProsPect_Click()
    'FrmListHotProspect.Show vbModal
End Sub

Private Sub CmdReleaseUnRealesePTP_Click()
    If UCase(MDIForm1.txtlevel.text) = "AGENT" Then
        MsgBox "Anda tidak mendapatkan akses!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    'FrmReleaseUnRealesePTP.Show vbModal
End Sub

Private Sub CmdScheduleoK_Click()
If Combo1(4).text = Empty Then
    MsgBox "Agent Harus Diisi", vbCritical + vbOKOnly, "Informasi"
    Exit Sub
End If
If TDBDate1(0).ValueIsNull Or TDBDate1(1).ValueIsNull Then
    MsgBox "Tanggal Tidak Boleh Kosong", vbInformation + vbOKOnly, "Informasi"
    Exit Sub
End If
If TDBDate1(0).Value > TDBDate1(1).Value Then
    MsgBox "Tanggal Periode Awal harus Lebih Kecil Dari Tanggal Periode Akhir", vbInformation + vbOKOnly, "Informasi"
    Exit Sub
End If
Call cari_Schedule
End Sub

Private Sub cari_Schedule()
Dim m_data As New CLS_FRMSEARCH
Dim ListItem As ListItem
Dim VOLUMEAMOUNT As Long
'@@ 19 Juli 2010 tambahan u/ blok data
Dim Blokedsearch As String
Dim BlokedEntrysearch As String
Dim STRSQL, StrsqlBlok, strinject As String
Dim M_objrs As ADODB.Recordset
Dim blokeddatamarkup As String
Dim STSLOCKTL As String
Dim STSfromaccount As String
Dim NMAGETPREV As String
If Check2.Value = 1 Then

    LstVwSearchMgm.ListItems.clear
    SSTab1.Tab = 0
    ' searching schedule mgm
  Call CEK_STATUS_F_CEK
  
  '--------- @@Start 19 Juli 2010 tambahan bloked
   STRSQL = "select * from usertbl where userid='"
   STRSQL = STRSQL + Trim(MDIForm1.TxtUsername.text) + "'"
   Set M_objrs = New ADODB.Recordset
   M_objrs.CursorLocation = adUseClient
   M_objrs.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
   If M_objrs.RecordCount <> 0 Then
   
     STSLOCKTL = CStr(Trim(IIf(IsNull(M_objrs!lockdarispvbuattl), "", M_objrs!lockdarispvbuattl)))
                
                STSfromaccount = CStr(Trim(IIf(IsNull(M_objrs!fromaccount), "", M_objrs!fromaccount)))
                
        If UCase(MDIForm1.txtlevel.text) = "TEAMLEADER" Then
                        NMAGETPREV = STSLOCKTL
         End If
         
        
        If M_objrs("usertype") = "1" Then
            strinject = IIf(IsNull(M_objrs!lockdarispv), "", M_objrs!lockdarispv)
           If strinject = "" Then
              Blokedsearch = ""
           Else
             Blokedsearch = IIf(IsNull(M_objrs!lockdarispv), "", Replace(M_objrs!lockdarispv, "@", "'"))
           End If
           BlokedEntrysearch = ""
           BlokedEntrysearch = IIf(IsNull(M_objrs!lock_entry_lpd), "", M_objrs!lock_entry_lpd)
           blokeddatamarkup = IIf(IsNull(M_objrs!lockmarkup), "", M_objrs!lockmarkup)
           If blokeddatamarkup <> "" Then
                    F_CEK = ""
                    Blokedsearch = ""

                     BlokedEntrysearch = ""
                End If
        End If
   End If
   
   If StsWO_Date = "1" Then
            If LUserType = "1" Then
                WO_DATE = "substring(B_D,1,4) in ('" + StsWO_2009 + "','" + StsWO_2008 + "','" + StsWO_2007 + "','" + StsWO_2006 + "','" + StsWO_2005 + "', "
                WO_DATE = WO_DATE + "'" + StsWO_2004 + "', '" + StsWO_2003 + "', '" + StsWO_2002 + "', '" + StsWO_2001 + "','" + StsWO_2000 + "','" + StsWO_1999 + "','" + StsWO_2010 + "')"
            End If
      End If
  
   Set m_cari = m_data.QUERY_SEARCH_mgm(M_OBJCONN, "AGENT = '" + Combo1(4).text + "' AND (NEXTACTDATE BETWEEN '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " 00:00" + "' AND '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " 23:59" + "') ", MDIForm1.Text3.text, F_CEK, f_Pending, Blokedsearch, BlokedEntrysearch, blokeddatamarkup, WO_DATE, NMAGETPREV)
    ProgressBar1.Max = m_cari.RecordCount + 1
    If Check2.Value = 1 Then
        TxtJmlDtMgm.text = m_cari.RecordCount & " Data"
    Else
        Text2.text = m_cari.RecordCount & " Data"
    End If
    
    While Not m_cari.EOF
    ProgressBar1.Value = m_cari.Bookmark
    
    Set ListItem = LstVwSearchMgm.ListItems.ADD(, , m_cari.Bookmark)
    
        ListItem.SubItems(1) = IIf(IsNull(m_cari("CUSTID")), "", m_cari("CUSTID"))
        ListItem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
        ListItem.SubItems(3) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
        'listitem.SubItems(4) = IIf(IsNull(m_cari("TGLSOURCE")), "", DateAdd("d", 90, Format(m_cari("TGLSOURCE"), "dd-mm-yyyy")))
        ListItem.SubItems(4) = IIf(IsNull(m_cari("recsource")), "", m_cari("recsource"))
        ListItem.SubItems(5) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "yyyy/mm/dd hh:nn"))
        ListItem.SubItems(6) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
        ListItem.SubItems(7) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS")) & "-" & IIf(IsNull(m_cari("f_pending")), "", m_cari("f_pending"))
        ListItem.SubItems(8) = IIf(IsNull(m_cari("KETHSLKERJA_NEW")), "", m_cari("KETHSLKERJA_NEW"))
        ListItem.SubItems(9) = IIf(IsNull(m_cari("StatusCall")), "", m_cari("StatusCall"))
        ListItem.SubItems(11) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
        'listitem.SubItems(10) = IIf(IsNull(m_cari("NamaAGENT")), "", m_cari("NamaAGENT"))
        ListItem.SubItems(12) = Format(IIf(IsNull(m_cari("Principal")), 0, m_cari("Principal")), "##,###")
        ListItem.SubItems(13) = Format(IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo")), "##,###")

       VOLUMEAMOUNT = VOLUMEAMOUNT + IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo"))
'        'LISTITEM.SubItems(15) = IIf(IsNull(m_cari("[NO]")), "", m_cari("[NO]"))
'
'        '@@16032011 Tambahan DOB dan No KTP
'        listitem.SubItems(26) = IIf(IsNull(m_cari("dob")), "", Format(m_cari("dob"), "yyyy-mm-dd"))
'        listitem.SubItems(27) = IIf(IsNull(m_cari("ktpno")), "", m_cari("ktpno"))

         ListItem.SubItems(14) = Format(IIf(IsNull(m_cari("OpenDate")), "", m_cari("OpenDate")), "yyyy/mm/dd")
        ListItem.SubItems(15) = Format(IIf(IsNull(m_cari("B_D")), 0, m_cari("B_D")))
        ListItem.SubItems(16) = Format(IIf(IsNull(m_cari("Pay_Dt")), 0, m_cari("Pay_Dt")), "yyyy/mm/dd")
         ListItem.SubItems(17) = Format(IIf(IsNull(m_cari("lastpay")), 0, m_cari("lastpay")), "##,###")
        
        ListItem.SubItems(18) = IIf(IsNull(m_cari("TGLSTATUS")), "", Format(m_cari("TGLSTATUS"), "YYYY/MM/DD"))
        ListItem.SubItems(19) = IIf(IsNull(m_cari("TGLCALL")), "", Format(m_cari("TGLCALL"), "YYYY/MM/DD"))
        ListItem.SubItems(20) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
        ListItem.SubItems(21) = Format(IIf(IsNull(m_cari("TGLINCOMING")), "", m_cari("TGLINCOMING")), "YYYY/MM/DD")
        ListItem.SubItems(23) = IIf(IsNull(m_cari("resultcpa")), "", m_cari("resultcpa"))
        ListItem.SubItems(24) = IIf(IsNull(m_cari("tglinsertfrmcpa")), "", m_cari("tglinsertfrmcpa"))
        ListItem.SubItems(25) = Format(IIf(IsNull(m_cari("curbal")), "", m_cari("curbal")), "##,###")
        'TOTCURBALANCE = TOTCURBALANCE + IIf(IsNull(m_cari("curbal")), 0, m_cari("curbal"))
       
        '@@16032011 Tambahan DOB dan No KTP
        ListItem.SubItems(26) = IIf(IsNull(m_cari("dob")), "", Format(m_cari("dob"), "yyyy-mm-dd"))
        ListItem.SubItems(27) = IIf(IsNull(m_cari("ktpno")), "", m_cari("ktpno"))

        m_cari.MoveNext
    Wend
        If LstVwSearchMgm.ListItems.Count = 0 Then
            TxtJmlDtMgm.text = "Tidak Ada Data"
            TxtJmlVolMgm.text = "0"
        Else
            TxtJmlDtMgm.text = "Total " + CStr(m_cari.RecordCount) + " Records"
            'TxtJmlVolMgm.Text = "Total " + CStr(m_cari.RecordCount)
            TxtJmlVolMgm.text = Format(VOLUMEAMOUNT, "##,###")
        End If
        
Else
    ' searching schedule leads
    Set m_cari = m_data.QUERY_SEARCH(M_OBJCONN, "AGENT = '" + Combo1(4).text + "' AND (NEXTACTDATE BETWEEN '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " 00:00" + "' AND '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " 23:59" + "') ", MDIForm1.Text3.text)
        ListView1.ListItems.clear
        SSTab1.Tab = 1
        ' searching schedule mgm
        ProgressBar1.Max = m_cari.RecordCount + 1
        Text2.text = m_cari.RecordCount & " Data"
        While Not m_cari.EOF
        ProgressBar1.Value = m_cari.Bookmark
        Set ListItem = ListView1.ListItems.ADD(, , m_cari.Bookmark)
        ListItem.SubItems(1) = IIf(IsNull(m_cari("custid")), "", JADI_QUOTE(m_cari("custid")))
        Select Case m_cari("RECSTATUS")
        Case "1A"
            ListItem.SubItems(2) = "Available"
        Case ""
            ListItem.SubItems(2) = "Available"
        Case Else
            ListItem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
        End Select
        ListItem.SubItems(3) = IIf(IsNull(m_cari("CUSTIDREF")), "", m_cari("CUSTIDREF"))
        ListItem.SubItems(4) = IIf(IsNull(m_cari("NAMAREF")), "", m_cari("NAMAREF"))
        ListItem.SubItems(5) = IIf(IsNull(m_cari("NAME")), "", JADI_QUOTE(m_cari("NAME")))
        ListItem.SubItems(6) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "yyyy/mm/dd hh:mm"))
        ListItem.SubItems(7) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
        ListItem.SubItems(8) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS"))
        ListItem.SubItems(9) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
        ListItem.SubItems(10) = IIf(IsNull(m_cari("NamaAGENT")), "", m_cari("NamaAGENT"))
        ListItem.SubItems(11) = IIf(IsNull(m_cari("RECSOURCEREF")), "", m_cari("RECSOURCEREF"))
        ListItem.SubItems(12) = Format(IIf(IsNull(m_cari("TGLSTATUS")), "", m_cari("TGLSTATUS")), "yyyy/mm/dd")
        ListItem.SubItems(13) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
        ListItem.SubItems(14) = IIf(IsNull(m_cari("KdComplaint")), "", m_cari("KdComplaint"))
        ListItem.SubItems(15) = IIf(IsNull(m_cari("RemarkComplaint")), "", m_cari("RemarkComplaint"))
        ListItem.SubItems(16) = IIf(IsNull(m_cari("KETHSLKERJA_NEW")), "", m_cari("KETHSLKERJA_NEW"))
        m_cari.MoveNext
    Wend
End If
Set m_data = Nothing

End Sub

Private Sub CEK_STATUS_F_CEK()
Dim CMDSQL As String
Dim M_objrs As New ADODB.Recordset

F_CEK = Empty
Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        CMDSQL = "SELECT * FROM usertbl WHERE USERID = '" + MDIForm1.TxtUsername.text + "'"
         M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        
        While Not M_objrs.EOF
                StsVl = CStr(Trim(IIf(IsNull(M_objrs!f_VL), "", M_objrs!f_VL)))
                StsPR = CStr(Trim(IIf(IsNull(M_objrs!f_PR), "", M_objrs!f_PR)))
                StsPTP = CStr(Trim(IIf(IsNull(M_objrs!f_ptp), "", M_objrs!f_ptp)))
                StsBP = CStr(Trim(IIf(IsNull(M_objrs!f_bp), Empty, M_objrs!f_bp)))
                StsPOP = CStr(Trim(IIf(IsNull(M_objrs!f_pop), "", M_objrs!f_pop)))
                StsSP = CStr(Trim(IIf(IsNull(M_objrs!f_sp), "", M_objrs!f_sp)))
                StsRP = CStr(Trim(IIf(IsNull(M_objrs!f_rp), "", M_objrs!f_rp)))
                StsUC = CStr(Trim(IIf(IsNull(M_objrs!F_UC), "", M_objrs!F_UC)))
                StsSK = CStr(Trim(IIf(IsNull(M_objrs!f_SK), "", M_objrs!f_SK)))
                StsON = CStr(Trim(IIf(IsNull(M_objrs!f_ON), "", M_objrs!f_ON)))
                StsOS = CStr(Trim(IIf(IsNull(M_objrs!f_OS), "", M_objrs!f_OS)))
                LUserType = CStr(Trim(IIf(IsNull(M_objrs!usertype), "", M_objrs!usertype)))
                Bloked = Replace(IIf(IsNull(M_objrs!lockdarispv), "", M_objrs!lockdarispv), "@", "'")
                Stsblank = CStr(Trim(IIf(IsNull(M_objrs!F_blank), "", M_objrs!F_blank)))
                StsWO_Date = CStr(Trim(IIf(IsNull(M_objrs!f_WO_DATE), "", M_objrs!f_WO_DATE)))
                StsWO_2009 = CStr(Trim(IIf(IsNull(M_objrs!f_WO_2009), "", M_objrs!f_WO_2009)))
                StsWO_2010 = CStr(Trim(IIf(IsNull(M_objrs!f_WO_2010), "", M_objrs!f_WO_2010)))
                StsWO_2008 = CStr(Trim(IIf(IsNull(M_objrs!f_WO_2008), "", M_objrs!f_WO_2008)))
                StsWO_2007 = CStr(Trim(IIf(IsNull(M_objrs!f_WO_2007), "", M_objrs!f_WO_2007)))
                StsWO_2006 = CStr(Trim(IIf(IsNull(M_objrs!f_WO_2006), "", M_objrs!f_WO_2006)))
                StsWO_2005 = CStr(Trim(IIf(IsNull(M_objrs!f_WO_2005), "", M_objrs!f_WO_2005)))
                StsWO_2004 = CStr(Trim(IIf(IsNull(M_objrs!f_WO_2004), "", M_objrs!f_WO_2004)))
                StsWO_2003 = CStr(Trim(IIf(IsNull(M_objrs!f_WO_2003), "", M_objrs!f_WO_2003)))
                StsWO_2002 = CStr(Trim(IIf(IsNull(M_objrs!f_WO_2002), "", M_objrs!f_WO_2002)))
                StsWO_2001 = CStr(Trim(IIf(IsNull(M_objrs!f_WO_2001), "", M_objrs!f_WO_2001)))
                StsWO_2000 = CStr(Trim(IIf(IsNull(M_objrs!f_WO_2000), "", M_objrs!f_WO_2000)))
                StsWO_1999 = CStr(Trim(IIf(IsNull(M_objrs!F_WO_1999), "", M_objrs!F_WO_1999)))
                M_objrs.MoveNext
            Wend
            Set M_objrs = Nothing
             StsAll = StsVl + StsPR + StsPTP + StsBP + StsPOP + StsSP + StsRP + StsUC + StsON + StsSK + StsOS
            
            
        If StsAll <> "" Then
            If LUserType = "1" Then
                    If StsUC = "UC" Then
                       If Bloked <> "" Then
                       
                            F_CEK = " + Bloked + "
                       Else
                            F_CEK = "(substring(F_CEK_NEW,1,3)IN( '" + StsVl + "','" + StsPR + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "','" + StsRP + "','" + StsSK + "','" + StsON + "','" + StsOS + "') or F_CEK_NEW IS NULL) "
                       End If
                       
                    Else
                     If Bloked <> "" Then
                            F_CEK = "(" + Bloked + " )"
                    Else
                         F_CEK = "(substring(F_CEK_NEW,1,3)IN( '" + StsVl + "','" + StsPR + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "','" + StsRP + "','" + StsSK + "','" + StsON + "','" + StsOS + "') or F_CEK_NEW IS NULL) "
                    End If
                    
                    End If
                        
                End If
                Else
                
                 If Bloked <> "" Then
                            F_CEK = "(" + Bloked + " )"
                 End If
     End If

End Sub

Private Sub CmdSearchPTP_Click()
    '@@ 24 Januari 2012, Buat Search PTP
    Dim Cmdsql_user As String
    Dim M_objrs As ADODB.Recordset
    Dim M_WHERE As String
    Dim Status_PTP As String
    Dim ListItem As ListItem
    
    Dim totamount As Double
    Dim TOTCURBALANCE As Double
    Dim VOLUMEAMOUNT As Double
    
    M_WHERE = ""
    
    If UCase(Trim(MDIForm1.txtlevel.text)) = "AGENT" Then
        'Cek dulu Apakah agent tersebut Dapat Melihat Status All PTP atau Hanya Sebagian PTP
        Cmdsql_user = "select f_status_ptp from usertbl where userid='"
        Cmdsql_user = Cmdsql_user + Trim(MDIForm1.TxtUsername.text) + "'"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open Cmdsql_user, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        'Jika Data ditemukan
        If M_objrs.RecordCount > 0 Then
            Status_PTP = IIf(IsNull(M_objrs("f_status_ptp")), "", M_objrs("f_status_ptp"))
        End If
        Set M_objrs = Nothing
        
        'set kriteria SQL PTP
        M_WHERE = " where agent='" + Trim(MDIForm1.TxtUsername.text) + "'  "
        If Status_PTP = "" Then
            'M_WHERE = M_WHERE + " and custid in (select custid from reportptp where promisedate between "
            'M_WHERE = M_WHERE + "date(now()) and date(now())+3 ) "
            '@@ 03-04-2012, Diubah berdasarkan tanggal tagih
            M_WHERE = M_WHERE + " and date(tgl_tagih) between date(now()) and date(now())+3 "
        End If
        
        M_WHERE = M_WHERE + " and substring(f_cek_new,1,3)='PTP' "
    End If
    
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Then
        M_WHERE = " where agent in (select userid from usertbl where team='"
        M_WHERE = M_WHERE + MDIForm1.TxtUsername.text + "') and substring(f_cek_new,1,3)='PTP' "
        M_WHERE = M_WHERE + " and date(tgl_tagih) is not null "
    ElseIf UCase(Trim(MDIForm1.txtlevel.text)) = "ADMIN" Or UCase(Trim(MDIForm1.txtlevel.text)) = "ADMINISTRATOR" Or UCase(Trim(MDIForm1.txtlevel.text)) = "SUPERVISOR" Then
        M_WHERE = " where substring(f_cek_new,1,3)='PTP' "
        M_WHERE = M_WHERE + " and date(tgl_tagih) is not null "
    End If
    
    CMDSQL = " select * from mgm  " + M_WHERE
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
     MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
     Set M_objrs = Nothing
     Exit Sub
    End If
    
    LstVwSearchMgm.ListItems.clear
    
    While Not M_objrs.EOF
        Set ListItem = LstVwSearchMgm.ListItems.ADD(, , M_objrs.Bookmark)
        'statusprior = IIf(IsNull(M_Objrs("StatusPrior")), "", M_Objrs("StatusPrior"))
        ListItem.SubItems(1) = IIf(IsNull(M_objrs("CUSTID")), "", M_objrs("CUSTID"))
        ListItem.SubItems(2) = IIf(IsNull(M_objrs("PRIOR")), "", M_objrs("PRIOR"))
        ListItem.SubItems(3) = IIf(IsNull(M_objrs("NAME")), "", M_objrs("NAME"))
        ListItem.SubItems(4) = IIf(IsNull(M_objrs("RECSOURCE")), "", M_objrs("RECSOURCE"))
        ListItem.SubItems(5) = IIf(IsNull(M_objrs("NEXTACTDATE")), "", Format(M_objrs("NEXTACTDATE"), "dd/mm/yyyy hh:nn"))
        ListItem.SubItems(6) = IIf(IsNull(M_objrs("NEXTACT")), "", M_objrs("NEXTACT"))
        ListItem.SubItems(7) = IIf(IsNull(M_objrs("REMARKS")), "", M_objrs("REMARKS"))
        ListItem.SubItems(8) = CStr(IIf(IsNull(M_objrs("kethslkerja_new")), "", M_objrs("kethslkerja_new")) & " ")  'sPending)
        ListItem.SubItems(9) = IIf(IsNull(M_objrs("StatusCall")), "", M_objrs("StatusCall"))
        ListItem.SubItems(11) = IIf(IsNull(M_objrs("AGENT")), "", M_objrs("AGENT"))
        
        
        If UCase(MDIForm1.txtlevel.text) <> "SUPERVISOR" Then
            If Format(IIf(IsNull(M_objrs("flaglead")), 0, M_objrs("flaglead")), "##,###") = 1 Then
                ListItem.SubItems(12) = ""
            Else
                ListItem.SubItems(12) = Format(IIf(IsNull(M_objrs("Principal")), 0, M_objrs("Principal")), "##,###")
            End If
        Else
             ListItem.SubItems(12) = Format(IIf(IsNull(M_objrs("Principal")), 0, M_objrs("Principal")), "##,###")
        End If
        
        ListItem.SubItems(13) = Format(IIf(IsNull(M_objrs("AmountWo")), 0, M_objrs("AmountWo")), "##,###")
        totamount = totamount + IIf(IsNull(M_objrs("AmountWo")), 0, M_objrs("AmountWo"))
        
        
        ListItem.SubItems(14) = Format(IIf(IsNull(M_objrs("OpenDate")), "", M_objrs("OpenDate")), "yyyy/mm/dd")
        ListItem.SubItems(15) = Format(IIf(IsNull(M_objrs("B_D")), 0, M_objrs("B_D")))
        ListItem.SubItems(16) = Format(IIf(IsNull(M_objrs("Pay_Dt")), 0, M_objrs("Pay_Dt")), "yyyy/mm/dd")
        
        ListItem.SubItems(17) = Format(IIf(IsNull(M_objrs("lastpay")), 0, M_objrs("lastpay")), "##,###")
        
        ListItem.SubItems(18) = IIf(IsNull(M_objrs("TGLSTATUS")), "", Format(M_objrs("TGLSTATUS"), "YYYY/MM/DD"))
        ListItem.SubItems(19) = IIf(IsNull(M_objrs("TGLCALL")), "", Format(M_objrs("TGLCALL"), "YYYY/MM/DD"))
        ListItem.SubItems(20) = IIf(IsNull(M_objrs("Kethslkerja")), "", M_objrs("Kethslkerja"))
        ListItem.SubItems(21) = Format(IIf(IsNull(M_objrs("TGLINCOMING")), "", M_objrs("TGLINCOMING")), "YYYY/MM/DD")
        ListItem.SubItems(23) = IIf(IsNull(M_objrs("resultcpa")), "", M_objrs("resultcpa"))
        ListItem.SubItems(24) = IIf(IsNull(M_objrs("tglinsertfrmcpa")), "", M_objrs("tglinsertfrmcpa"))
        ListItem.SubItems(25) = Format(IIf(IsNull(M_objrs("curbal")), "", M_objrs("curbal")), "##,###")
        TOTCURBALANCE = TOTCURBALANCE + IIf(IsNull(M_objrs("curbal")), 0, M_objrs("curbal"))
       
        ListItem.SubItems(26) = IIf(IsNull(M_objrs("dob")), "", Format(M_objrs("dob"), "yyyy-mm-dd"))
        ListItem.SubItems(27) = IIf(IsNull(M_objrs("ktpno")), "", M_objrs("ktpno"))
        ListItem.SubItems(28) = IIf(IsNull(M_objrs("CUSTNO")), "", M_objrs("CUSTNO"))
            
SorryLompat:
        
        VOLUMEAMOUNT = VOLUMEAMOUNT + IIf(IsNull(M_objrs("AmountWo")), 0, M_objrs("AmountWo"))
        
        M_objrs.MoveNext
    Wend
    
    txtamount.text = Format(totamount, "##,###")
    txtcurbalance.text = Format(TOTCURBALANCE, "##,###")
    
    If LstVwSearchMgm.ListItems.Count = 0 Then
        TxtJmlDtMgm.text = "Tidak Ada Data"
        TxtJmlVolMgm.text = "0"
    Else
        TxtJmlDtMgm.text = "Total " + CStr(M_objrs.RecordCount) + " Records"
        TxtJmlVolMgm.text = Format(VOLUMEAMOUNT, "##,###")
    End If
    
    Set M_objrs = Nothing
End Sub

Private Sub Combo1_Click(Index As Integer)

Dim m_data As New CLS_FRMSEARCH
Dim M_objrs As ADODB.Recordset
Select Case Index
Case 0
        If Combo1(0).text = "LUNAS" Then
        Combo1(0).text = Empty
        Combo1(1).text = Empty
        Exit Sub
        End If
    Set M_objrs = m_data.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + Combo1(Index).text + "'")
    If M_objrs.RecordCount <> 0 Then
        Combo1(0).text = M_objrs("USERID")
        Combo1(1).text = M_objrs("AGENT")
    Else
        Combo1(0).text = Empty
        Combo1(1).text = Empty
    End If
    
   

Case 1
    Set M_objrs = m_data.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + Combo1(Index).text + "'")
    If M_objrs.RecordCount <> 0 Then
        Combo1(0).text = M_objrs("USERID")
        Combo1(1).text = M_objrs("AGENT")
    Else
        Combo1(0).text = Empty
        Combo1(1).text = Empty
    End If
Case 2
   
    Set M_objrs = m_data.QUERY_DATASOURCE(M_OBJCONN, "KODEDS = '" + Combo1(Index).text + "'")
    If M_objrs.RecordCount <> 0 Then
        Combo1(2).text = M_objrs("KODEDS")
        Combo1(3).text = cnull(M_objrs("KETERANGAN"))
    Else
        Combo1(2).text = Empty
        Combo1(3).text = Empty
    End If
Case 3
Set M_objrs = m_data.QUERY_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).text + "'")
    If M_objrs.RecordCount <> 0 Then
        Combo1(2).text = M_objrs("KODEDS")
        Combo1(3).text = M_objrs("KETERANGAN")
    Else
        Combo1(2).text = Empty
        Combo1(3).text = Empty
    End If
Case 4
    Set M_objrs = m_data.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + Combo1(Index).text + "'")
    If M_objrs.RecordCount <> 0 Then
        Combo1(4).text = M_objrs("USERID")
        Combo1(5).text = M_objrs("AGENT")
    Else
        Combo1(4).text = Empty
        Combo1(5).text = Empty
    End If
Case 5
    Set M_objrs = m_data.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + Combo1(Index).text + "'")
    If M_objrs.RecordCount <> 0 Then
        Combo1(4).text = M_objrs("USERID")
        Combo1(5).text = M_objrs("AGENT")
    Else
        Combo1(4).text = Empty
        Combo1(5).text = Empty
    End If
End Select
Set m_data = Nothing
Set M_objrs = Nothing
End Sub
Private Sub isi_combo_agent()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
    'Menegecek jenis User yang login, jika dia agent
    'maka combo nama agent di kunci
    
    If UCase(MDIForm1.txtlevel.text) = "AGENT" Then
        cmb_kdagent.text = MDIForm1.TxtUsername.text
        cmb_nmagent.text = MDIForm1.txtnama.text
        cmb_kdagent.Enabled = False
        cmb_nmagent.Enabled = False
    End If
    
    'Jika yang login=administrator
    If MDIForm1.txtlevel.text = "Administrator" Or MDIForm1.txtlevel.text = "Admin" Then
            CMDSQL = "SELECT * FROM USERTBL where   AKTIF='1'  order by  KDLEVEL='1' DESC,agent "
    End If
    'Jika yang login Supervisor
    If MDIForm1.txtlevel.text = "Supervisor" Then
        CMDSQL = "select  userid,agent from usertbl where   AKTIF='1' and ( spvcode='"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "' or userid ='" + MDIForm1.TxtUsername.text + "')"
        CMDSQL = CMDSQL + " order by  KDLEVEL='1' DESC,agent "
    ElseIf MDIForm1.txtlevel.text = "Agent" Then
        CMDSQL = "select userid,agent from usertbl where userid='" + MDIForm1.TxtUsername.text + "'"
    End If
    'Jika yang login TeamLeader
    
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    'Jika tidak ada data agent/TL maka tutup form viewmgmdata
  
    
    While Not M_objrs.EOF
        
        cmb_kdagent.AddItem IIf(IsNull(M_objrs("userid")), "", M_objrs("userid"))
        cmb_nmagent.AddItem IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
        M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
Dim sSearchText As String
Dim lReturn As Long
Select Case Index
Case 0, 1, 2, 3
KeyAscii = 0
If KeyAscii = 13 Then
   Combo1_Click (Index)
   KeyAscii = 0
Else
   sSearchText = Left$(Combo1(Index).text, Combo1(Index).SelStart) & Chr$(KeyAscii)
   lReturn = SendMessage(Combo1(Index).hwnd, CB_FINDSTRING, -1, ByVal sSearchText)
   If lReturn <> CB_ERR Then
      mbIgnoreListClick = True
      Combo1(Index).ListIndex = lReturn
      mbIgnoreListClick = False
      Combo1(Index).text = Combo1(Index).list(lReturn)
      Combo1(Index).SelStart = Len(sSearchText)
      Combo1(Index).SelLength = Len(Combo1(Index).text)
      KeyAscii = 0
   End If
End If
End Select
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    Dim m_data As New CLS_FRMSEARCH
    Dim M_objrs As ADODB.Recordset
    Select Case Index
    Case 0
        Set M_objrs = m_data.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + Combo1(Index).text + "'")
        If M_objrs.RecordCount <> 0 Then
            Combo1(0).text = M_objrs("USERID")
            Combo1(1).text = M_objrs("AGENT")
        Else
            Combo1(0).text = Empty
            Combo1(1).text = Empty
        End If
    Case 1
        Set M_objrs = m_data.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + Combo1(Index).text + "'")
        If M_objrs.RecordCount <> 0 Then
            Combo1(0).text = M_objrs("USERID")
            Combo1(1).text = M_objrs("AGENT")
        Else
            Combo1(0).text = Empty
            Combo1(1).text = Empty
        End If
    Case 2
    Set M_objrs = m_data.QUERY_DATASOURCE(M_OBJCONN, "KODEDS = '" + Combo1(Index).text + "'")
        If M_objrs.RecordCount <> 0 Then
            Combo1(2).text = M_objrs("KODEDS")
            Combo1(3).text = cnull(M_objrs("KETERANGAN"))
        Else
            Combo1(2).text = Empty
            Combo1(3).text = Empty
        End If
    Case 3
    Set M_objrs = m_data.QUERY_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).text + "'")
        If M_objrs.RecordCount <> 0 Then
            Combo1(2).text = M_objrs("KODEDS")
            Combo1(3).text = M_objrs("KETERANGAN")
        Else
            Combo1(2).text = Empty
            Combo1(3).text = Empty
        End If
    End Select
    Set m_data = Nothing
    Set M_objrs = Nothing
End Sub
Private Sub Combo2_DropDown()
    Call load_statuscall
End Sub
Public Sub load_statuscall()
    Dim sStrsql As String
    Dim M_objrs As ADODB.Recordset
    sStrsql = "  SELECT tblstatuscall_keterangan FROM tblstatuscall WHERE tblstatuscall_kdstatus='1' order by tblstatuscall_id,grp_call "
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        Combo2.clear
        While Not M_objrs.EOF
                Combo2.AddItem IIf(IsNull(M_objrs!tblstatuscall_keterangan), "", M_objrs!tblstatuscall_keterangan)
                M_objrs.MoveNext
        Wend
    Set M_objrs = Nothing
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Combo4_Click(Index As Integer)
  Combo4(1).clear
  Select Case Combo4(0).text
  
  Case "DPD"
        Combo4(1).AddItem "OUTSTANDING"
        Combo4(1).text = "OUTSTANDING"
  Case "OUTSTANDING"
        Combo4(1).AddItem "DPD"
        Combo4(1).text = "DPD"
  End Select
End Sub

Private Sub Command2_Click()
    Dim objExcel As New Excel.Application
    Dim objExcelSheet As Excel.Worksheet
    Dim col, row As Integer
    Dim a As String
    If LstVwSearchMgm.ListItems.Count > 0 Then
        objExcel.Workbooks.ADD
        Set objExcelSheet = objExcel.Worksheets.ADD
     
        For col = 1 To LstVwSearchMgm.ColumnHeaders.Count
            objExcelSheet.Cells(1, col).Value = LstVwSearchMgm.ColumnHeaders(col)
        Next
     
        For row = 2 To LstVwSearchMgm.ListItems.Count + 1
            For col = 1 To LstVwSearchMgm.ColumnHeaders.Count
            If col = 1 Then
                    objExcelSheet.Cells(row, col).Value = LstVwSearchMgm.ListItems(row - 1).text
            Else
                '" 'cararandy 29032016 "
                Dim hasil1 As String
                    If col <> 12 And col <> 14 Then
                        hasil1 = "'" + LstVwSearchMgm.ListItems(row - 1).SubItems(col - 1)
                        objExcelSheet.Cells(row, col).Value = hasil1
                    Else
                        hasil1 = LstVwSearchMgm.ListItems(row - 1).SubItems(col - 1)
                        objExcelSheet.Cells(row, col).Value = hasil1
                    End If
                End If
            Next
        Next
     
        objExcelSheet.Columns.AutoFit
        Cd_save.ShowOpen
        a = Cd_save.FileName
     
        objExcelSheet.SaveAs a & ".xls"
        MsgBox "Export Completed", vbInformation, Me.Caption
     
        objExcel.Workbooks.Open a & ".xls"
        objExcel.Visible = True
    Else
        MsgBox "No data to export", vbInformation, Me.Caption
    End If

End Sub

Private Sub Form_Activate()
    'Call tampil_waktu
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Exit Sub
End Sub

Private Sub Form_Load()
    Dim M_objrs As ADODB.Recordset
    Dim m_data  As New CLS_FRMSEARCH
    Dim xxx     As Integer
    
    Dim ss As Integer

    Call HEADER_VIEW_mgm
    Call HEADER_VIEW_Refferall
    Call isi_combo_agent
    
    Call list_client(Combo99)
    txtpage.Locked = True
    txtcountpage.Locked = True
    
    If MDIForm1.txtlevel.text = "Admin" Then
        Command2.Visible = True
'        label1(17).Visible = True
'        Combo99.Visible = True
        Label1(17).Visible = False
        Combo99.Visible = False

    End If
    
    jmlpage = GetSetting("anto", "textboxes", "text3", "")
    If Val(jmlpage) = 0 Then
        txtjmllimit.text = 10
    Else
        txtjmllimit.text = Val(jmlpage)
    End If
    txtjmllimit.text = "100"
    
    If UCase(MDIForm1.txtlevel.text) = "AGENT" Then
        CekDtDistribute.Visible = False
        Label1(13).Visible = False
        cmd_review.Visible = False
        cmd_claimback_acc.Visible = False
        Cmd_listrequestdecease.Visible = False
        cmb_kdagent.text = MDIForm1.TxtUsername.text
        cmb_nmagent.text = MDIForm1.txtnama.text
        'cmb_kdagent.Enabled = False
        'cmb_nmagent.Enabled = False
        Label1(18).Visible = False
        Label1(19).Visible = False
        Combo4(0).Visible = False
        Combo4(1).Visible = False
    Else
        CekDtDistribute.Visible = True
    End If

    If UCase(MDIForm1.txtlevel.text) = "TEAMLEADER" Then
        cmd_review.Visible = False
        cmd_claimback_acc.Visible = False
    End If

    SSTab1.Tab = 0
    SSTab2.Tab = 0
    
    Me.Width = 11880
    Me.Height = 9945
    'Me.Height = 6105
    SSTab1.TabVisible(1) = False
    Me.Top = 500
    Me.Left = 1000
    
    DTimeLastCall(0).Value = "00:00"
    DTimeLastCall(1).Value = "23:59"
    StsmgmSchedule = False

    Set M_objrs = m_data.QUERY_AGENT_JADWAL(M_OBJCONN, "")
    While Not M_objrs.EOF
        Combo1(0).AddItem cnull(M_objrs("USERID"))
        Combo1(1).AddItem cnull(M_objrs("AGENT"))
        Combo1(4).AddItem cnull(M_objrs("USERID"))
        Combo1(5).AddItem cnull(M_objrs("AGENT"))
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing


'=======================================
'pemisahaan
    If UCase(MDIForm1.txtlevel.text) = "ADMIN" Or UCase(MDIForm1.txtlevel.text) = "ADMINISTRATOR" Then
        Set M_objrs = m_data.QUERY_DATASOURCE(M_OBJCONN, "kodeds in (select distinct recsource from mgm)")
    ElseIf UCase(MDIForm1.txtlevel.text) = "AGENT" Then
        'Set M_objrs = m_data.QUERY_DATASOURCE(M_OBJCONN, "kodeds in (select distinct recsource from mgm where agent = '" & MDIForm1.TxtUsername.text & "' )")
        '=====asep21032020===='
        Set M_objrs = m_data.QUERY_DATASOURCE(M_OBJCONN, "kodeds in (select distinct recsource from mgm where agent = '' )")
        '====================='
    Else
        Set M_objrs = m_data.QUERY_DATASOURCE(M_OBJCONN, "kodeds in (select distinct recsource from mgm where agent in (select userid from usertbl where team = '" & MDIForm1.TxtUsername.text & "') or agent =  '" & MDIForm1.TxtUsername.text & "')")
    End If
'=======================================

    While Not M_objrs.EOF
        Combo1(2).AddItem M_objrs("KODEDS")
        Combo1(3).AddItem cnull(M_objrs("KETERANGAN"))
        M_objrs.MoveNext
    Wend
    
    If UCase(MDIForm1.Text3.text) = "ADMIN" Then
        Label1(5).Visible = True
        txtnocard.Visible = True
    End If
    
    Set M_objrs = Nothing
    Set m_data = Nothing
    
    'Frame2.Left = (Screen.Width - Frame2.Width) / 2
    Frame2.Width = Screen.Width
    'Frame1.Width = Screen.Width - Frame9.Width
    LstVwSearchMgm.Width = (Screen.Width - Frame9.Width) - 120
    Frame8.Left = ((LstVwSearchMgm.Width - Frame8.Width) / 2) + LstVwSearchMgm.Left
    
    '====asep21032020==='
'    Dim fopen As String
'    Dim m_cust1 As New ADODB.Recordset
'    Set m_cust1 = New ADODB.Recordset
'    m_cust1.CursorLocation = adUseClient
'    CMDSQL = "select agent from mgm "
'    m_cust1.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
'
'    fopen = cnull(m_cust1("agent"))
'    If fopen = "" Then
'        If FrmCC_Colection.CBOACCOUNT = "PTP" Then
'            M_OBJCONN.Execute "update mgm set agent = '" & MDIForm1.TxtUsername.text & "', nama_agent = '" & MDIForm1.txtnama.text & "' where custid = '" & LstVwSearchMgm.SelectedItem.SubItems(1) & "'"
'            M_OBJCONN.Execute "update mgm set agent = '" & MDIForm1.TxtUsername.text & "', nama_agent = '" & MDIForm1.txtnama.text & "' where custid = '" & LstVwSearchMgm.SelectedItem.SubItems(1) & "'"
'            FrmCC_Colection.Show 1
'        End If
'    ElseIf fopen = MDIForm1.TxtUsername.text Then
'        FrmCC_Colection.Show 1
'    ElseIf fopen <> MDIForm1.TxtUsername.text Then
'        MsgBox "Data milik agen " & fopen & ", mohon pilih data yang lain", vbCritical + vbOKOnly
'        Exit Sub
'    End If
    '==================='
    lstautodial.ColumnHeaders.ADD 1, , "No", 5 * TXT
    lstautodial.ColumnHeaders.ADD 2, , "ID", 5 * TXT
    lstautodial.ColumnHeaders.ADD 3, , "ID CUST", 15 * TXT
    lstautodial.ColumnHeaders.ADD 4, , "PHONE", 15 * TXT
End Sub

Private Sub show_Search_mgmData()
    Dim harga As Double
    Dim ListItem As ListItem
    Dim Lcustid1 As String
    Dim Lcustid2 As String
    Dim LCall As String
    Dim i As Integer
    Dim CMDSQL As String
    Dim sPending As String
    Dim M_objrs As ADODB.Recordset
    Dim VOLUMEAMOUNT As Double
    Dim statusprior As String
    Dim exp%
    Dim totamount As Double
    Dim TOTCURBALANCE As Double
    Dim tgl_server As String
    Dim expired_claim As Integer
    Dim sts_data_acc() As String
    Dim rs_cek_data As ADODB.Recordset
    
    Dim tgl_janji As Date
    Dim tgl_bayar As Date
    
    Dim tgl_expired As Date
    Dim tgl_app_claim As Date
    
    Dim number_count As Integer

    i = 1
'On Error GoTo HELL
        
    datajml = m_cari.RecordCount
    LstVwSearchMgm.ListItems.clear
    Me.MousePointer = vbHourglass
    ProgressBar2.Max = m_cari.RecordCount + 1
    
    '@@19-11-10 ///////// Ini tambahan buat mencatat custid per session yang dicolect agent pd blok data
        'cek status id lock session di tabel usertbl
        Dim CekSession As String
        Dim M_Objrs_Session As ADODB.Recordset
        Dim NilaiSession As String
        
        Dim CmdWaktuServer As String
        Dim m_ObjrsWktServer As ADODB.Recordset
        Dim WaktuServer As Date
        
        Dim SimpanDtSession As String
        
        'Ambil Id Session Start
        CekSession = "select f_idsessstart from usertbl where userid='"
        CekSession = CekSession + Trim(MDIForm1.TxtUsername.text) + "'"
        Set M_Objrs_Session = New ADODB.Recordset
        M_Objrs_Session.CursorLocation = adUseClient
        M_Objrs_Session.Open CekSession, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs_Session.RecordCount <> 0 Then
         NilaiSession = IIf(IsNull(M_Objrs_Session("f_idsessstart")), "", M_Objrs_Session("f_idsessstart"))
        Else
         NilaiSession = ""
        End If
        Set M_Objrs_Session = Nothing
        
        'Ambil Waktu Server Terkini
        CmdWaktuServer = "select now() "
        Set m_ObjrsWktServer = New ADODB.Recordset
        m_ObjrsWktServer.CursorLocation = adUseClient
        m_ObjrsWktServer.Open CmdWaktuServer, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        WaktuServer = Format(m_ObjrsWktServer(0), "yyyy-mm-dd hh:mm:ss")
        Set m_ObjrsWktServer = Nothing
    
    '@@20-11-10 ///////// Ini tambahan buat mencatat custid per session yang dicolect agent pd blok data
    number_count = 1
    If txtpage.text > 1 Then
        number_count = ((Val(txtpage.text) * txtjmllimit) - Val(txtjmllimit.text)) + 1
    End If
    
    tgl_server = waktu_server_sekarang
    
    While Not m_cari.EOF
      
        '@@20-11-10 ///Ini buat mencatat custid dan status awal yang dikerjain agent per session ////
            If NilaiSession <> Empty Then
                '@@ 04-10-2011, Log session log data dinonaktifkan terlebih dahulu, karena lama
'                SimpanDtSession = "insert into tblperformpersessionlock(idlock,"
'                SimpanDtSession = SimpanDtSession + "startlock,agent,custid,name,"
'                SimpanDtSession = SimpanDtSession + "f_ceklalu,tgl_f_ceklalu) values ('"
'                SimpanDtSession = SimpanDtSession + Trim(NilaiSession) + "','"
'                SimpanDtSession = SimpanDtSession + CStr(Format(WaktuServer, "yyyy-mm-dd hh:mm:ss")) + "','"
'                SimpanDtSession = SimpanDtSession + Trim(m_cari!agent) + "','"
'                SimpanDtSession = SimpanDtSession + Trim(m_cari!CustId) + "','"
'                SimpanDtSession = SimpanDtSession + Trim(m_cari!Name) + "','"
'                SimpanDtSession = SimpanDtSession + IIf(IsNull(m_cari!f_cek_new), "", Trim(m_cari!f_cek_new)) + "',"
'                SimpanDtSession = SimpanDtSession + IIf(IsNull(m_cari!tglcall), "null", "'" + CStr(Format(m_cari!tglcall, "yyyy-mm-dd hh:mm:ss")) + "'") + ")"
'                M_OBJCONN.Execute SimpanDtSession
            End If
        '@@20-11-10 ///Ini buat mencatat custid dan status awal yang dikerjain agent per session ////
      'Me.Refresh
        ProgressBar2.Value = m_cari.Bookmark
        Lcustid1 = CStr(IIf(IsNull(m_cari!CustId), "", m_cari!CustId))
        'sPending = CStr(Trim(IIf(IsNull(m_cari!f_Pending), "", m_cari!f_Pending)))
        
'        END CEK CLAIM ACC -------------------------------------------------------------------------------------
        
        'Set listItem = LstVwSearchMgm.ListItems.ADD(, , m_cari.Bookmark)
        Set ListItem = LstVwSearchMgm.ListItems.ADD(, , number_count)
        
        Dim interval As Integer
        Dim K As Integer
        'Dim tgl_server As String
        
        'Jika tidak dikerjakan selama 3 hari
        
        statusprior = IIf(IsNull(m_cari("StatusPrior")), "", m_cari("StatusPrior"))
        ListItem.SubItems(1) = IIf(IsNull(m_cari("CUSTID")), "", m_cari("CUSTID"))
        ListItem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
        ListItem.SubItems(3) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
        ListItem.SubItems(4) = IIf(IsNull(m_cari("RECSOURCE")), "", m_cari("RECSOURCE"))
        ListItem.SubItems(5) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "dd/mm/yyyy hh:nn"))
        ListItem.SubItems(6) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
        ListItem.SubItems(7) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS"))
        ListItem.SubItems(9) = CStr(IIf(IsNull(m_cari("kethslkerja_new")), "", m_cari("kethslkerja_new")) & " " & sPending)
        ListItem.SubItems(8) = IIf(IsNull(m_cari("StatusCall")), "", m_cari("StatusCall"))
        ListItem.SubItems(11) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
        '===remark asep19042020'
'        If UCase(MDIForm1.txtlevel.text) <> "SUPERVISOR" Then
'            If Format(IIf(IsNull(m_cari("flaglead")), 0, m_cari("flaglead")), "##,###") = 1 Then
'                   ListItem.SubItems(12) = ""
'            Else
'                ListItem.SubItems(12) = Format(IIf(IsNull(m_cari("Principal")), 0, m_cari("Principal")), "##,###")
'            End If
'        Else
'             ListItem.SubItems(12) = Format(IIf(IsNull(m_cari("Principal")), 0, m_cari("Principal")), "##,###")
'        End If
        '============================='
        ListItem.SubItems(13) = Format(IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo")), "##,###")
        totamount = totamount + IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo"))
        
        ListItem.SubItems(14) = Format(IIf(IsNull(m_cari("OpenDate")), "", m_cari("OpenDate")), "yyyy/mm/dd")
        ListItem.SubItems(15) = Format(IIf(IsNull(m_cari("B_D")), 0, m_cari("B_D")))
        ListItem.SubItems(16) = Format(IIf(IsNull(m_cari("Pay_Dt")), 0, m_cari("Pay_Dt")), "yyyy/mm/dd")
        
        ListItem.SubItems(17) = Format(IIf(IsNull(m_cari("lastpay")), 0, m_cari("lastpay")), "##,###")
        
        ListItem.SubItems(18) = IIf(IsNull(m_cari("TGLSTATUS")), "", Format(m_cari("TGLSTATUS"), "YYYY/MM/DD"))
        ListItem.SubItems(19) = IIf(IsNull(m_cari("TGLCALL")), "", Format(m_cari("TGLCALL"), "YYYY/MM/DD"))
        ListItem.SubItems(20) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
        ListItem.SubItems(21) = Format(IIf(IsNull(m_cari("TGLINCOMING")), "", m_cari("TGLINCOMING")), "YYYY/MM/DD")
        ListItem.SubItems(23) = IIf(IsNull(m_cari("resultcpa")), "", m_cari("resultcpa"))
        ListItem.SubItems(24) = IIf(IsNull(m_cari("tglinsertfrmcpa")), "", m_cari("tglinsertfrmcpa"))
        'remark asep19042020==='
        'ListItem.SubItems(25) = Format(IIf(IsNull(m_cari("curbal")), "", m_cari("curbal")), "##,###")
        'TOTCURBALANCE = TOTCURBALANCE + IIf(IsNull(m_cari("curbal")), 0, m_cari("curbal"))
       
        '@@16032011 Tambahan DOB dan No KTP
        'ListItem.SubItems(26) = IIf(IsNull(m_cari("dob")), "", Format(m_cari("dob"), "yyyy-mm-dd"))
        '==============================='
        ListItem.SubItems(27) = IIf(IsNull(m_cari("ktpno")), "", m_cari("ktpno"))
        ListItem.SubItems(28) = IIf(IsNull(m_cari("REGION")), "", m_cari("REGION"))
        'MERUBAH WARNA JIKA TIDAK DI CALL SELAMA 3HARI
        
        If m_cari("TGLCALL") <> "" Then
            interval = DateDiff("d", Format(m_cari("TGLCALL"), "yyyy-mm-dd"), Format(tgl_server, "yyyy-mm-dd"))
        Else
            interval = 0
        End If

          number_count = number_count + 1
SorryLompat:
        'listitem.SubItems(19) = IIf(IsNull(m_cari("F_CEK")), "", m_cari("F_CEK"))
        VOLUMEAMOUNT = VOLUMEAMOUNT + IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo"))
        m_cari.MoveNext
        
    Wend
    
    '@@ 20-11-10 update f_idsesstart ke null jika nilaisession ada (tidak null)
      If NilaiSession <> "" Or IsNull(NilaiSession) = False Or NilaiSession <> Empty Then
        Dim UpdateSess As String
        
        UpdateSess = "update usertbl set f_idsessstart=null,f_pesanlockauto=null where userid='"
        UpdateSess = UpdateSess + Trim(MDIForm1.TxtUsername.text) + "'"
        M_OBJCONN.Execute UpdateSess
        NilaiSession = ""
      End If
    '@@ 20-11-10 update f_idsesstart ke null jika nilaisession ada (tidak null)
  
    
    txtamount.text = Format(totamount, "##,###")
    txtcurbalance.text = Format(TOTCURBALANCE, "##,###")
        
    If LstVwSearchMgm.ListItems.Count = 0 Then
       TxtJmlDtMgm.text = "Tidak Ada Data"
       TxtJmlVolMgm.text = "0"
    Else
       TxtJmlDtMgm.text = "Total " + CStr(m_cari.RecordCount) + " Records"
       TxtJmlVolMgm.text = "Total " + CStr(Format(VOLUMEAMOUNT, "##,###"))
    End If
    LstVwSearchMgm.SortKey = 2
    LstVwSearchMgm.Sorted = True
    ProgressBar1.Value = 0
    ProgressBar1.Visible = False
    MousePointer = vbNormal
    Set m_cari = Nothing
    Set m_cari2 = Nothing
    Exit Sub
hell:
    Me.MousePointer = vbNormal
    MsgBox Err.Description
End Sub

Private Sub HEADER_VIEW_mgm()
    LstVwSearchMgm.ColumnHeaders.ADD 1, , "No", 5 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 2, , "Customer No", 15 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 3, , "Priority", 0
    LstVwSearchMgm.ColumnHeaders.ADD 4, , "Nama Customer", 15 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 5, , "Batch", 15 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 6, , "Tgl FollowUp", 13 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 7, , "Visit", 0
    LstVwSearchMgm.ColumnHeaders.ADD 8, , "History Call", 15 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 9, , "Statuscall", 13 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 10, , "Status Telp", 0
    LstVwSearchMgm.ColumnHeaders.ADD 11, , "Call Initial", 0
    LstVwSearchMgm.ColumnHeaders.ADD 12, , "Agent", 7 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 13, , "Principle", 0 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 14, , "WO Amount", 13 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 15, , "Open Date", 13 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 16, , "WO Date", 13 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 17, , "LPD", 13 * TXT
    
    'LstVwSearchMgm.ColumnHeaders.ADD 18, , "DataBase", 0
    '@@ 13-09-2011, Nomor 18 diganti dengan LPA
    LstVwSearchMgm.ColumnHeaders.ADD 18, , "LPA", 13 * TXT
    
    LstVwSearchMgm.ColumnHeaders.ADD 19, , "Tgl Status", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 20, , "Tgl Call", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 21, , "Sts Account", 0
    LstVwSearchMgm.ColumnHeaders.ADD 22, , "PTP Date", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 23, , "id", 0 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 24, , "STS", 0 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 25, , "Tanggal status CPA", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 26, , "Current Balance", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 27, , "DOB", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 28, , "KTP", 0
    LstVwSearchMgm.ColumnHeaders.ADD 29, , "Region", 10 * TXT
End Sub

Sub WaitSecs(Seconds As Single)
    Dim a As Long
    Seconds = Seconds + Timer
    While Seconds > Timer
        a = DoEvents
    Wend
End Sub

Private Sub cmd_Click(Index As Integer)
    Select Case Index
    Case 0
        'txtpage.Text = 1
        If txtpage.text = 1 Then cmd(3).Enabled = False
        
        If Val(txtcountpage) > Val(txtpage.text) Then
            txtpage.text = Val(txtpage.text) + 1
        End If
        If txtcountpage.text <> "" Then
            Command1(0).DoClick
        End If
   
    Case 1
        If Val(txtpage.text) > 1 Then
            txtpage.text = Val(txtpage.text) - 1
        End If
      
        If txtcountpage.text <> "" Then
            Command1(0).DoClick
        End If
        
    Case 2
        txtpage.text = txtcountpage.text
        If txtcountpage.text <> "" Then
            Command1(0).DoClick
        End If
        If txtcountpage.text = 0 Then txtpage.text = 1
    Case 3
        If txtcountpage.text <> "" Then
            txtpage.text = 1
            Command1(0).DoClick
        End If
    End Select
   LstVwSearchMgm.SortKey = IndexColumnHEader
   LstVwSearchMgm.Sorted = True
     
    If txtpage.text = 1 Then
        cmd(3).Enabled = False: cmd(1).Enabled = False
    Else
        cmd(3).Enabled = True: cmd(1).Enabled = True
    End If
    
    If txtpage.text = txtcountpage.text Then
        cmd(2).Enabled = False: cmd(0).Enabled = False
    Else
        cmd(2).Enabled = True: cmd(0).Enabled = True
    End If

End Sub

Private Sub Command1_Click(Index As Integer)
    Dim NAMACUST As String
    Dim statuscall As String
    Dim NamaAgent As String
    Dim DATASOURCE As String
    Dim TGLLAHIR As String
    Dim OFFPHONE As String
    Dim OFFPHONE2 As String
    Dim HOMEPHONE As String
    Dim HOMEPHONE2 As String
    Dim MOBILEPHONE As String
    Dim MOBILEPHONE2 As String
    Dim FAXPHONE As String
    Dim Lcustid As String
    Dim FAXPHONE2 As String
    Dim KETHSLKERJA As String
    Dim lLastCallDate As String
    Dim lStatusCek As String
    Dim sPending As String
    Dim FCEKSTATUS As String
    Dim strverify As String
    Dim strapprovel As String
    Dim m_data As New CLS_FRMSEARCH
    Dim M_objrs As New ADODB.Recordset
    Dim PANJANG As Integer
    Dim nmagentprev As String
    Dim strReject As String
    Dim strSukses As String
    Dim strapprovelyet As String
    Dim strinject As String
    Dim blokeddatamarkup As String
    Dim BlokedPTPNoPayment As String
    Dim STSLOCKTL As String
    Dim STSfromaccount As String
    '@@ 15 Agustus 2011, Bloked Payment request gaby
    Dim BlokedPayment As String
    Dim RSTEMP As New ADODB.Recordset
    Dim STRSQL As String
    Dim dblLimitpage As Double
    Dim i As Integer
    Dim xx As Integer
    'jejaktian(tambahantian)
    Dim AHOMENO As String
    Dim AHOMENO2 As String
    Dim AOFFICENO As String
    Dim AOFFICENO2 As String
    Dim extoffice As String
    Dim extoffice2 As String
    Dim homenoadd1 As String
    Dim ahomenoadd1 As String
    Dim homenoadd2 As String
    Dim ahomenoadd2 As String
    Dim officenoadd1 As String
    Dim aofficenoadd2 As String
    Dim officenoadd2 As String
    Dim mobilenoadd1 As String
    Dim mobilenoadd2 As String
    Dim ec_telp As String
    Dim alamatrumah As String
    Dim alamatkantor As String
    Dim alamatec As String
    '===============================
    Dim CmdSql_Info As String
    Dim M_ObjrsInfo As ADODB.Recordset
    Dim pesan As String
    Dim m_objrsCekSess As ADODB.Recordset
    Dim CmdCekSess As String
    Dim idsess As String
    Dim Lcustno As String
    Dim client As String
    Dim tglsource As String
    Me.Refresh
    Select Case Index
    Case 0
        Command1(0).Enabled = False
        F_CEK = Empty
        WO_DATE = Empty
        
'        M_objrs.CursorLocation = adUseClient
'        CMDSQL = "SELECT *  FROM usertbl WHERE USERID = '" + MDIForm1.TxtUsername.text + "'"
'        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
'
'        If Not M_objrs.EOF Then
'            strinject = IIf(IsNull(M_objrs!lockdarispv), "", M_objrs!lockdarispv)
'            If strinject = "" Then
'                ' Jika bukan akses all baru di enable
'                If cek_aksesall = "0" Then
'                    CmdSearchPTP.Enabled = True
'                End If
'                Bloked = ""
'            Else
'    '            CmdSearchPTP.Enabled = False
'                Bloked = IIf(IsNull(M_objrs!lockdarispv), "", Replace(M_objrs!lockdarispv, "@", "'"))
'            End If
'            '@@140710 Bloked Entry data
'            BlokedEntry = IIf(IsNull(M_objrs!lock_entry_lpd), "", M_objrs!lock_entry_lpd)
'            blokeddatamarkup = IIf(IsNull(M_objrs!lockmarkup), "", M_objrs!lockmarkup)
'
'            '@@15 Agustus 2011 Bloked Data Payment request gaby
'            BlokedPayment = IIf(IsNull(M_objrs!lockpayment), "", M_objrs!lockpayment)
'
'            '@@ 21 April 2014 Bloked Data PTP-NoPayment Request Joko
'            BlokedPTPNoPayment = IIf(IsNull(M_objrs!lock_ptp_payment), "", M_objrs!lock_ptp_payment)
'        End If
        
       
'        If STSLOCKTL <> Empty Then cmb_kdagent.text = "": cmb_kdagent.Enabled = False: cmb_nmagent.Enabled = False: GoTo CUY
'            Set M_objrs = Nothing
'            StsAll = StsVl + StsPR + StsBP + StsPOP + StsUC + StsON + StsSK + StsOS
'
'            If StsAll <> "" Then
'            If LUserType = "1" Then
'            If StsUC = "UC" Then
'                If Bloked <> "" Then
'                    F_CEK = "(" + Bloked + " )"
'                Else
'                    F_CEK = " substring(F_CEK_NEW,1,3)  IN('" + StsVl + "','" + StsPR + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "','" + StsSK + "', '" + StsON + "','" + StsOS + "','') "
'                End If
'                Else
'                    If Bloked <> "" Then
'                        F_CEK = "(" + Bloked + " )"
'                    Else
'                        F_CEK = " substring(F_CEK_NEW,1,3)  IN('" + StsVl + "','" + StsPR + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "','" + StsSK + "', '" + StsON + "','" + StsOS + "','') "
'                    End If
'                End If
'            End If
'        End If
      
      
        If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Or UCase(MDIForm1.txtlevel.text) = "ADMIN" Or UCase(MDIForm1.txtlevel.text) = "ADMINISTRATOR" Then
            GoTo CUY
        End If
      
        If Trim(Text1(0).text) = Empty And Trim(cmb_kdagent.text) = Empty And Combo1(2).text = Empty And Len(TDBMask2.Value) < 1 And Len(TDBMask1.text) < 1 And TdDob.ValueIsNull And Len(txtnocard.text) < 3 _
            And cmbStsLastCall(0).text = "" And CmbStatusCek.text = "" And DtLastCall(0).ValueIsNull And CekDtDistribute.Value = 0 And Combo3.text = "" And TxtAlamat.text = "" And Combo2.text = "" And Combo1(2).text = "" Then
            
            MsgBox "Masukan Kriteria Customer Yang Akan Dicari...!!!", vbCritical + vbOKOnly, "Peringatan"
            Command1(0).Enabled = True
            Text1(0).SetFocus
            Set m_data = Nothing
            Exit Sub
        Else
CUY:
            LstVwSearchMgm.ListItems.clear
            Frame3.Visible = True
    
            If CekDtDistribute.Value = 1 Then
               NamaAgent = "AGENT is null"
            Else
                If txtnocard.text <> Empty Then ' cek no custid
                    Lcustid = "CUSTID LIKE " + "'%" + UBAH_QUOTE(Trim(txtnocard.text)) + "%'"
                End If
                If txtregion.text <> Empty Then ' cek region
                    Lcustno = "region LIKE " + "'%" + UBAH_QUOTE(Trim(txtregion.text)) + "%'"
                End If
                If Text1(0).text <> Empty Then ' cek nama customer
                    NAMACUST = "name LIKE " + "'%" + UBAH_QUOTE(Trim(Text1(0).text)) + "%'"
                End If
                
                If cmb_kdagent.text <> Empty Then
                    NamaAgent = "AGENT = '" + Trim(cmb_kdagent.text) + "'"
                End If
                
                If Combo2.text <> Empty Then
                    If Combo2.text = "New Data" Then
                        statuscall = "coalesce(STATUSCALL,'')='' AND coalesce(F_CEK_NEW,'')='' "
                    Else
                        statuscall = "STATUSCALL = '" + Trim(Combo2.text) + "'"
                    End If
                End If
                
                If Combo1(2).text <> Empty Then
                    DATASOURCE = "RECSOURCE ilike '%" + Trim(Combo1(2).text) + "%'"
                End If
                
                If UCase(MDIForm1.txtlevel) = "AGENT" Then
                    tglsource = "date(tglsource) = date(now()) "
                End If
                
                Dim recx As String
                
                
                If Combo99.text <> Empty Then
                    recx = Left(Combo99.text, 1) & "X" & Right(Combo99.text, Len(Combo99.text) - 2)
                    client = "(left(RECSOURCE,3) <> 'EX_' and RECSOURCE ilike '%" + Trim(Combo99.text) + "%') or RECSOURCE ilike '%" & Trim(recx) & "%'"
                End If
                
                If TdDob.ValueIsNull = False Then
                    TGLLAHIR = "DOB = '" + Format(TdDob.text, "yyyy/mm/dd") + "'"
                End If
                
                '==========asep05052020=============='
                Dim qpub As String
                If Combo4(0).text <> Empty Then
                    If Combo4(0).text = "OUTSTANDING" Then
                        qpub = qpub + " order by oustanding desc "
                    End If

                    If Combo4(0).text = "DPD" Then
                        qpub = qpub + " order by delq_amt_by_x desc "
                    End If

                    If Combo4(1).text = "OUTSTANDING" Then
                        qpub = qpub + " , oustanding desc "
                    End If

                    If Combo4(1).text = "DPD" Then
                        qpub = qpub + " ,delq_amt_by_x desc "
                    End If
                End If
                '=================================='
            
                If Len(TDBMask1.text) > 1 Then
                    OFFPHONE = "OFFICENO Like '%" + TDBMask1.text + "%'"
                    OFFPHONE2 = "OFFICENO2 Like '%" + TDBMask1.text + "%'"
                    HOMEPHONE = "HOMENO Like '%" + TDBMask1.text + "%'"
                    HOMEPHONE2 = "HOMENO2 Like '%" + TDBMask1.text + "%'"
                    FAXPHONE = "FAXNO Like '%" + TDBMask1.text + "%'"
                    FAXPHONE2 = "FAXNO2 Like '%" + TDBMask1.text + "%'"
                    MOBILEPHONE = "MOBILENO like '%" + TDBMask1.text + "%'"
                    MOBILEPHONE2 = "MOBILENO2 like '%" + TDBMask1.text + "%'"
                    AHOMENO = "ahomeno Like '%" + TDBMask1.text + "%'"
                    AHOMENO2 = "ahomeno2 Like '%" + TDBMask1.text + "%'"
                    AOFFICENO = "aofficeno Like '%" + TDBMask1.text + "%'"
                    AOFFICENO2 = "aofficeno2 Like '%" + TDBMask1.text + "%'"
                    extoffice = "extoffice Like '%" + TDBMask1.text + "%'"
                    extoffice2 = "extoffice2 Like '%" + TDBMask1.text + "%'"
                    homenoadd1 = "homenoadd1 Like '%" + TDBMask1.text + "%'"
                    ahomenoadd1 = "ahomenoadd1 Like '%" + TDBMask1.text + "%'"
                    homenoadd2 = "homenoadd2 Like '%" + TDBMask1.text + "%'"
                    ahomenoadd2 = "ahomenoadd2 Like '%" + TDBMask1.text + "%'"
                    officenoadd1 = "officenoadd1 Like '%" + TDBMask1.text + "%'"
                    aofficenoadd2 = "aofficenoadd2 Like '%" + TDBMask1.text + "%'"
                    officenoadd2 = "officenoadd2 Like '%" + TDBMask1.text + "%'"
                    mobilenoadd1 = "mobilenoadd1 Like '%" + TDBMask1.text + "%'"
                    mobilenoadd2 = "mobilenoadd2 Like '%" + TDBMask1.text + "%'"
                    ec_telp = "ec_telp Like '%" + TDBMask1.text + "%'"
                End If
            
                If Len(TDBMask2.Value) > 1 Then
                    MOBILEPHONE = "MOBILENO like '%" + TDBMask2.Value + "%'"
                    MOBILEPHONE2 = "MOBILENO2 like '%" + TDBMask2.Value + "%'"
                End If
                
                If Len(TxtAlamat.text) > 1 Then
                    alamatrumah = "AddrNow like '%" + TxtAlamat.text + "%'"
                    alamatkantor = "addrpt like '%" + TxtAlamat.text + "%'"
                    alamatec = "ecaddr like '%" + TxtAlamat.text + "%'"
                End If
                
                If Left(Combo3.text, 3) = "ALL" Then
                    strverify = " stscpa=1"
                End If
            
                If DtLastCall(0).ValueIsNull = False Then
                    lLastCallDate = "TGLSTATUS BETWEEN '" + Format(DtLastCall(0).Value, "MM/DD/YYYY") & " " & CStr(DTimeLastCall(0).Value) + "' AND '" + Format(DtLastCall(1).Value, "MM/DD/YYYY") & " " & CStr(DTimeLastCall(1).Value) + "'"
                End If
                
            End If
        End If
          
        'Unload FRM_SEARCH
        If Check1.Value = 0 Then
        Else
            If blokeddatamarkup <> "" Then
               ' F_CEK = ""
               ' WO_DATE = ""
               ' BlokedEntry = ""
               ' Bloked = ""
            End If
            
            ' TAMBAHAN BlokedPTPNoPayment 21 APRIL 2014
             Set m_cari = m_data.QUERY_SEARCH_CONDITION_mgm(M_OBJCONN, NAMACUST, NamaAgent, DATASOURCE, TGLLAHIR, _
                                                          OFFPHONE, OFFPHONE2, HOMEPHONE, HOMEPHONE2, MOBILEPHONE, _
                                                         MOBILEPHONE2, FAXPHONE, FAXPHONE2, MDIForm1.Text3.text, _
                                                         AHOMENO, AHOMENO2, AOFFICENO, AOFFICENO2, extoffice, extoffice2, homenoadd1, ahomenoadd1, _
                                                         homenoadd2, ahomenoadd2, officenoadd1, aofficenoadd2, officenoadd2, mobilenoadd1, mobilenoadd2, _
                                                         ec_telp, alamatrumah, alamatkantor, alamatec, Lcustid, Bloked, lLastCallDate, lStatusCek, sPending, FCEKSTATUS, WO_DATE, strverify, strapprovel, strapprovelyet, strReject, strSukses, Bloked, BlokedEntry, blokeddatamarkup, nmagentprev, "", qpub, tglsource, BlokedPayment, BlokedPTPNoPayment, Val(txtpage.text), 10000, statuscall, Lcustno, client)
            
            
             SaveSetting "anto", "textboxes", "text3", txtjmllimit.text
             jmlpage = GetSetting("anto", "textboxes", "text3", "")
             If Val(jmlpage) = 0 Then
                 dblLimitpage = 10
             Else
                 dblLimitpage = Val(jmlpage)
             End If
            
             Set totalrows = m_data.QUERY_SEARCH_jmlrow(M_OBJCONN, NAMACUST, NamaAgent, DATASOURCE, TGLLAHIR, _
                                                          OFFPHONE, OFFPHONE2, HOMEPHONE, HOMEPHONE2, MOBILEPHONE, _
                                                         MOBILEPHONE2, FAXPHONE, FAXPHONE2, MDIForm1.Text3.text, Lcustid, Bloked, lLastCallDate, lStatusCek, sPending, FCEKSTATUS, WO_DATE, strverify, strapprovel, strapprovelyet, strReject, strSukses, Bloked, BlokedEntry, blokeddatamarkup, nmagentprev, "", BlokedPayment, BlokedPTPNoPayment, Val(txtpage.text), 0, statuscall)

        End If
        
        txtpage.Locked = True
        txtcountpage.Locked = True
        
    '   LBLCOUNT.Caption = "Jumlah Record :" + CStr(m_cari.RecordCount)
        txttotal.text = CStr(totalrows.RecordCount)
        'Text3.Text = CStr(m_cari.RecordCount)
        Text3.text = txttotal.text
        If txtpage.text <= txtcountpage.text Then
            cmd(2).Enabled = True
            cmd(0).Enabled = True
        End If
    
        If txtpage.text <> "" Then
            txtcountpage.text = Ceiling(txttotal.text / dblLimitpage)
            If txtpage.text > txtcountpage.text And txtcountpage.text <> "0" Then
                txtpage.text = 1
                txtcountpage.text = 0
                GoTo CUY
            End If
        End If
 
        If m_cari.RecordCount = 0 Then
            MsgBox "Data Anda Tidak Ditemukan!", vbInformation + vbOKOnly, "Aplikasi"
            Command1(0).Enabled = True
                
                
                'Jika data tidak ditemukan maka jika idsessstart maka di null-in lagi idessstarnya
            If txtnocard.text = Empty _
                  And TDBMask1.text = Empty _
                  And Text1(0).text = Empty _
                  And Combo2.text = Empty _
                  And IsNull(DtLastCall(0).Value) _
                  And IsNull(DtLastCall(1).Value) _
                  And Combo1(2).text = Empty _
                  And Combo1(3).text = Empty _
                  And Combo3.text = Empty _
                  And cmb_kdagent.text = Trim(MDIForm1.TxtUsername.text) Then
                    
                    
                CmdCekSess = "select f_idsessstart from usertbl where userid='"
                CmdCekSess = CmdCekSess + Trim(MDIForm1.TxtUsername.text) + "'"
                Set m_objrsCekSess = New ADODB.Recordset
                m_objrsCekSess.CursorLocation = adUseClient
                m_objrsCekSess.Open CmdCekSess, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                'idsess = IIf(IsNull(CStr(m_objrsCekSess("f_idsessstart"))), "", CStr(m_objrsCekSess("f_idsessstart")))
                'idsess = CStr(m_objrsCekSess("f_idsessstart"))
                If IsNull(m_objrsCekSess("f_idsessstart")) Then
                    idsess = ""
                Else
                    idsess = CStr(m_objrsCekSess("f_idsessstart"))
                End If
                Set m_objrsCekSess = Nothing
                    
                If idsess <> "" Or idsess <> Empty Then
                    'Kasih informasi ke agent kenapa datanya kosong
                    
                    CmdSql_Info = "select * from tbltemplockacc_current where id='"
                    CmdSql_Info = CmdSql_Info + Trim(idsess) + "'"
                    Set M_ObjrsInfo = New ADODB.Recordset
                    M_ObjrsInfo.CursorLocation = adUseClient
                    M_ObjrsInfo.Open CmdSql_Info, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                    
                    pesan = "Kondisi Lock data anda saat ini adalah:" + Chr(13)
                    pesan = pesan + "Start Lock: " + CStr(M_ObjrsInfo("start_lock")) + Chr(13)
                    pesan = pesan + "End Lock: " + CStr(M_ObjrsInfo("end_lock")) + Chr(13)
                    pesan = pesan + "Account yang di lock: " + M_ObjrsInfo("account_lock") + Chr(13)
                    pesan = pesan + "Di lock oleh: " + M_ObjrsInfo("lock_by") + Chr(13)
                    pesan = pesan + "Status yang di lock: " + M_ObjrsInfo("status_lock") + Chr(13)
                    pesan = pesan + Chr(13)
                    pesan = pesan + "System tidak dapat menemukan data sesuai lock anda, " + Chr(13)
                    pesan = pesan + "anda dapat menghubungi TL/SPV/Orang yang melock data anda(lihat lock by) " + Chr(13)
                    pesan = pesan + "untuk merelease data anda! Terima kasih."
                    
                    MsgBox pesan, vbOKOnly + vbInformation, "Informasi"
                    Set M_ObjrsInfo = Nothing
                    
                    CmdCekSess = "update usertbl set f_idsessstart=null,f_idsessend=null where userid='"
                    CmdCekSess = CmdCekSess + Trim(MDIForm1.TxtUsername.text) + "'"
                    M_OBJCONN.Execute CmdCekSess
                    
                    '@@20022013 tambahan jika data tidak di temukan
                End If
                    
            End If
                  
            Set m_data = Nothing
            'Call CariDataAll
            Exit Sub
        Else
            search_ok = True
            If Check1.Value = 1 Then
                'kalau found refferall data
                'Unload FRM_PRESCREEN
                'FRM_PRESCREEN.Caption = "Search Non mgm Data"
                'FRM_PRESCREEN.Show
                SSTab1.Tab = 0
'                    Call show_UCDATA
                ' Untuk ngecek apakah data sudah 5 hari belum dikerjakan juga // By Izuddin
                'Call cek_lama_account
                ' -------------------------------------------------------------------------
                Call show_Search_mgmData
            Else
                ' kalau found mgm data
                SSTab1.Tab = 1

            End If
                
            '@@12022013, buat akses akun yang bareng2 khusus TL dan agent
            If UCase(MDIForm1.txtlevel.text) = "AGENT" Or UCase(MDIForm1.txtlevel.text) = "TEAMLEADER" Then
                'Call CariDataAll
            End If
        End If

        'Call Warna_Row_Listview(VIEW_MGMDATA, LstVwSearchMgm, vbWhite, vbWhite)
        Set M_objrs = Nothing
        Command1(0).Enabled = True
    Case 1
        MDIForm1.LstGrade.ListItems.clear
        Unload Me
        
    Case 2
        txtnocard.text = Empty
        txtregion.text = Empty
        Text1(0).text = Empty
        TdDob.text = Empty
        If MDIForm1.txtlevel.text <> "Agent" Then
            cmb_kdagent.text = Empty
            cmb_nmagent.text = Empty
        End If
        Combo1(1).text = Empty
        Combo1(2).text = Empty
        Combo1(3).text = Empty
        TDBMask1.text = Empty
        TDBMask2.text = Empty
        cmbStsLastCall(0).text = Empty
        DtLastCall(0).Value = Empty
        DtLastCall(1).Value = Empty
        CmbStatusCek.text = Empty
        
End Select
Set m_data = Nothing


' Frame3.Visible = False
End Sub

Sub cek_lama_account()
    Dim rs_cek As ADODB.Recordset
    Dim tglserver As Date
    Dim interval As Integer
    Dim interval2 As Integer
    Dim cek_available As Integer
    Dim TL_Review As String
    
    If UCase(Left(MDIForm1.txtlevel.text, 5)) = "AGENT" Then
        cek_available = 0
        Set rs_cek = New ADODB.Recordset
        rs_cek.CursorLocation = adUseClient
        rs_cek.ActiveConnection = M_OBJCONN
        rs_cek.CursorType = adOpenDynamic
        rs_cek.LockType = adLockOptimistic
        
        If rs_cek.State = 1 Then rs_cek.Close
        rs_cek.Open "SELECT userid FROM usertbl WHERE lower(userid) LIKE 'review%' AND team IN (SELECT team FROM usertbl WHERE userid='" & Trim(MDIForm1.TxtUsername.text) & "')"
        'TL_Review = IIf(IsNull(rs_cek!USERID), "", rs_cek!USERID)
        
        If rs_cek.State = 1 Then rs_cek.Close
        rs_cek.Open "SELECT now() as tgl_server"
        tglserver = Format(rs_cek!tgl_server, "yyyy-mm-dd")
        
        If rs_cek.State = 1 Then rs_cek.Close
        rs_cek.Open "SELECT id,custid,tglsource,tglcall,spv_allow FROM mgm WHERE tglcall is null AND spv_allow is null AND agent='" & Trim(MDIForm1.TxtUsername.text) & "'"
        If rs_cek.RecordCount > 0 Then
            Do Until rs_cek.EOF
                Dim K As Integer
                Dim tgltelpon As String
                Dim arrayLV() As Integer
                
                interval = DateDiff("d", Format(rs_cek!tglsource, "yyyy-mm-dd"), Format(tglserver, "yyyy-mm-dd"))
                

                ' Jika kelewat 5 hari dari tgl upload
                If interval > 5 Then
                    cek_available = cek_available + 1
                    ' 04 Agustus 2014 - MASUKKIN KE LOG
                    M_OBJCONN.Execute "INSERT INTO tbl_log_acc_review(custid,agent,keterangan) values('" & rs_cek!CustId & "','" & Trim(MDIForm1.TxtUsername.text) & "','5HARI NOT FOLLOW')"
                    ' =================================
                    M_OBJCONN.Execute "UPDATE mgm SET agent_asli=agent WHERE id=" & rs_cek!ID & ""
                    M_OBJCONN.Execute "UPDATE mgm SET agent='" & TL_Review & "' WHERE id=" & rs_cek!ID & ""
                End If
                rs_cek.MoveNext
            Loop
            
            If cek_available > 0 Then
                'M_OBJCONN.Execute "UPDATE usertbl SET f_blok='1' WHERE userid='" & Trim(mdiform1.txtusername.text) & "'"
                MsgBox cek_available & " Data(s) Masuk ke coding TL REVIEW karena Data lebih dari 5 hari belum dikerjakan", vbCritical + vbOKOnly, "Akun data 5 hari"
                'End
            End If
        End If
        
        Set rs_cek = Nothing
        
    End If
End Sub

Sub UPDATE_BP()
'Dim TGLSYS As ADODB.Recordset
'Dim SPTP As ADODB.Recordset
'Dim CMDSQL As String
'Dim TGLKOMP As Date
'Set TGLSYS = New ADODB.Recordset
'TGLSYS.CursorLocation = adUseClient
'TGLSYS.Open "SELECT TGLSYSTEM FrOM VWCALLCFG1", M_OBJCONN, adOpenDynamic, adLockOptimistic
'TGLKOMP = TGLSYS!TGLSYSTEM
'Set SPTP = New ADODB.Recordset
'CMDSQL = "UPDATE mgm SET KETHSLKErJA='BP-BROKEN PROMISE',F_CEK='BP-',KETHSLKErJADESC='' "
'CMDSQL = CMDSQL + "WHERE CUSTID IN (SELECT CUSTID FrOM mgm WHErE DATEDIFF(day,TdbdatePTP,TGLKOMP)>5)"

End Sub

Private Sub show_UCDATA()
    Dim sdata As ADODB.Recordset
    Dim ListItem As ListItem
    Dim i%
    Set sdata = New ADODB.Recordset
    sdata.CursorLocation = adUseClient
    CMDSQL = "select CUSTID,KETHSLKERJA,AGENT FROM mgm WHERE left(KETHSLKERJA,2) in ('WN','NK','MV') AND AGENT='" & VIEW_MGMDATA.cmb_kdagent.text & "'"
    sdata.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    i = 1
    MDIForm1.LstGrade.ListItems.clear
    Do While Not sdata.EOF
        Set ListItem = MDIForm1.LstGrade.ListItems.ADD(, , i)
        ListItem.SubItems(1) = IIf(IsNull(sdata!CustId), "", sdata!CustId)
        ListItem.SubItems(2) = IIf(IsNull(sdata!CustId), "", sdata!KETHSLKERJA)
        ListItem.SubItems(3) = IIf(IsNull(sdata!AGENT), "", sdata!AGENT)
        sdata.MoveNext
        i = i + 1
    Loop
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MDIForm1.m_targetview = False
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
   ListView1.SortKey = ColumnHeader.Index - 1
   ListView1.Sorted = True
End Sub

Private Sub ListView1_DblClick()
Dim M_objrs As ADODB.Recordset

If ListView1.ListItems.Count = 0 Then
    Exit Sub
End If
Status_Form = 1
    If UCase(MDIForm1.txtlevel.text) = "TEAMLEADER" Then
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open "SELECT USERID FROM usertbl WHERE SPVCODE ='" + MDIForm1.TxtUsername.text + "' AND USERID = '" + ListView1.SelectedItem.SubItems(9) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_objrs.RecordCount <> 0 Then
        Else
            MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "Aplikasi"
            Set M_objrs = Nothing
            Exit Sub
        End If
        Set M_objrs = Nothing
    Else
        If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Or UCase(MDIForm1.txtlevel.text) = "ADMINISTRATOR" Then
        Else
            If Trim(UCase(MDIForm1.TxtUsername.text)) = Trim(UCase(ListView1.SelectedItem.SubItems(9))) Then
            Else
                MsgBox "Data Ini Milik Agent Yang Lain", vbCritical + vbOKOnly, "Aplikasi"
                Set M_objrs = Nothing
                Exit Sub
            End If
        End If
    End If
    FrmCC_Colection.Show vbModal
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
Dim M_objrs As ADODB.Recordset
Dim m_msgbox As Variant
If ListView1.ListItems.Count = 0 Then
    Exit Sub
End If
If KeyAscii = 13 Then
    Call ListView1_DblClick
End If
If UCase(MDIForm1.txtlevel.text) <> "AGENT" Then
    If KeyAscii = 100 Or KeyAscii = 68 Then
        If UCase(MDIForm1.txtlevel.text) = "TEAMLEADER" Then
            Set M_objrs = New ADODB.Recordset
            M_objrs.CursorLocation = adUseClient
            M_objrs.Open "SELECT USERID FROM usertbl WHERE SPVCODE ='" + MDIForm1.TxtUsername.text + "' AND USERID = '" + ListView1.SelectedItem.SubItems(9) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If M_objrs.RecordCount <> 0 Then
                m_msgbox = MsgBox("Yakin Akan Di Hapus??...", vbExclamation + vbYesNo, "Informasi")
                If m_msgbox = vbYes Then
                    M_OBJCONN.Execute "Delete From cc_custtbl where Nomor = " + ListView1.SelectedItem.SubItems(17) + " And RECSOURCE <> 'infomedia1' and left(recsource,3)<>'inf'"
                    ListView1.ListItems.Remove ListView1.SelectedItem.Index
                End If
            Else
                MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "Aplikasi"
                Set M_objrs = Nothing
                Exit Sub
            End If
            Set M_objrs = Nothing
        Else
            m_msgbox = MsgBox("Yakin Akan Di Hapus??...", vbExclamation + vbYesNo, "Informasi")
            If m_msgbox = vbYes Then
                'M_OBJCONN.Execute "Delete From cc_custtbl where custid ='" + ListView1.SelectedItem.SubItems(1) + "'"
                M_OBJCONN.Execute "Delete From cc_custtbl where Nomor = " + ListView1.SelectedItem.SubItems(17) + " And RECSOURCE <> 'infomedia1'"
                ListView1.ListItems.Remove ListView1.SelectedItem.Index
            End If
        End If
    End If
End If
End Sub

Private Sub LstVwSearchmgm_Click()
    If b_pindah = True Then
        FrmCustIdTransfer.List1.AddItem LstVwSearchMgm.SelectedItem.SubItems(1)
    End If
End Sub

Private Sub LstVwSearchmgm_ColumnClick(ByVal ColumnHeader As ColumnHeader)
   LstVwSearchMgm.SortKey = ColumnHeader.Index - 1
   IndexColumnHEader = ColumnHeader.Index - 1
   LstVwSearchMgm.Sorted = True
End Sub
Private Sub LstVwSearchMgm_DblClick()
Dim STRSQL  As String
Dim MOBJRSKISRUT As New ADODB.Recordset
Dim CmdsqlCek As String
Dim M_ObjrsCekAkses As ADODB.Recordset
Dim M_Objrs_Akses_Acc As ADODB.Recordset
'On Error GoTo ke

Dim M_objrs As ADODB.Recordset
If LstVwSearchMgm.ListItems.Count = 0 Then
    Exit Sub
End If

glexp = LstVwSearchMgm.SelectedItem.SubItems(4)
Status_Form = 2
If LstVwSearchMgm.ListItems.Count = 0 Then
    Exit Sub
End If
    
    If UCase(MDIForm1.txtlevel.text) = "AGENT" Then
        If UCase(MDIForm1.TxtUsername.text) <> Trim(UCase(LstVwSearchMgm.SelectedItem.SubItems(11))) Then
            'Dim Cmdsql As String
            Dim M_Objrs_Cek As ADODB.Recordset
            Dim Vcek As Boolean
    
            '@@16032011 Tambahan jika CH tersebut memiliki data Visa, tapi punya agent lain tetep bisa dibuka sama agent tsb
            'Cek dulu punya data no.ktp apa ngga
            If LstVwSearchMgm.SelectedItem.SubItems(27) <> "" Then
                CMDSQL = "select custid,agent from mgm where (name='"
                CMDSQL = CMDSQL + Trim(LstVwSearchMgm.SelectedItem.SubItems(3)) + "' and dob='"
                CMDSQL = CMDSQL + Format(LstVwSearchMgm.SelectedItem.SubItems(26), "yyyy-mm-dd") + "' or ktpno='"
                CMDSQL = CMDSQL + Trim(LstVwSearchMgm.SelectedItem.SubItems(27)) + "') and custid<>'"
                CMDSQL = CMDSQL + Trim(LstVwSearchMgm.SelectedItem.SubItems(1)) + "'"
            Else
                CMDSQL = "select custid,agent from mgm where name='"
                CMDSQL = CMDSQL + Trim(LstVwSearchMgm.SelectedItem.SubItems(3)) + "' and dob='"
                CMDSQL = CMDSQL + Format(LstVwSearchMgm.SelectedItem.SubItems(26), "yyyy-mm-dd") + "' and custid <>'"
                CMDSQL = CMDSQL + Trim(LstVwSearchMgm.SelectedItem.SubItems(1)) + "'"
            End If
            Set M_Objrs_Cek = New ADODB.Recordset
            M_Objrs_Cek.CursorLocation = adUseClient
            M_Objrs_Cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            Vcek = False
            While Not M_Objrs_Cek.EOF
                If UCase(MDIForm1.TxtUsername.text) = UCase(Trim(M_Objrs_Cek("agent"))) Then
                    Vcek = True
                End If
                M_Objrs_Cek.MoveNext
            Wend
            Set M_Objrs_Cek = Nothing
            
    
            '@@02082012 Cek Coding nih......
            CMDSQL = "select * from "
            CMDSQL = CMDSQL + "(select spvcode from usertbl where userid='"
            CMDSQL = CMDSQL + CStr(Trim(MDIForm1.TxtUsername.text))
            CMDSQL = CMDSQL + "') as a, "
            CMDSQL = CMDSQL + " (select spvcode as spvcode_new,sts_akses_agent as sts_akses_agent_new "
            CMDSQL = CMDSQL + " from usertbl where userid='"
            CMDSQL = CMDSQL + CStr(Trim(LstVwSearchMgm.SelectedItem.SubItems(11)))
            CMDSQL = CMDSQL + "') as b "
            CMDSQL = CMDSQL + " where a.SPVCODE = b.spvcode_new "
            Set M_Objrs_Cek = New ADODB.Recordset
                M_Objrs_Cek.CursorLocation = adUseClient
                M_Objrs_Cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If M_Objrs_Cek.RecordCount > 0 Then
                If IsNull(M_Objrs_Cek("sts_akses_agent_new")) = True Then
                    Vcek = False
                ElseIf CStr(Trim(M_Objrs_Cek("sts_akses_agent_new"))) = "1" Then
                    Vcek = True
                End If
            End If
            Set M_Objrs_Cek = Nothing
    
            If Vcek = False Then
                MsgBox "Anda Tidak Berhak Untuk Mengedit Data Ini", vbCritical + vbOKOnly, "Aplikasi"
                Exit Sub
            Else
                GoTo lanjut
            End If
        End If
    End If
lanjut:
    If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Then
    'If UCase(mdiform1.txtlevel.text) = "TEAMLEADER" Then
        Dim PO_AGENT As String
        If VIEW_MGMDATA.cmb_kdagent.text = "PULLOUT" Then
            Set M_objrs = New ADODB.Recordset
                M_objrs.CursorLocation = adUseClient
            CMDSQL = "SELECT PO_Agent FROM mgm where CUSTID='" & LstVwSearchMgm.SelectedItem.SubItems(11) & "'"
            M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If Not M_objrs.EOF Then
                PO_AGENT = M_objrs!PO_AGENT
            End If
            Set M_objrs = Nothing
        Else
            PO_AGENT = VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11)
        End If
    Else

        Set M_objrs = New ADODB.Recordset
            M_objrs.CursorLocation = adUseClient
            
        If UCase(Trim(MDIForm1.txtlevel.text)) = "SUPERVISOR" Then
            CMDSQL = "SELECT USERID FROM usertbl WHERE  USERID = '" + Trim(LstVwSearchMgm.SelectedItem.SubItems(11)) + "'"
        ElseIf UCase(MDIForm1.txtlevel.text) = "TEAMLEADER" Then
            CMDSQL = "SELECT USERID FROM usertbl WHERE TEAM ='" + MDIForm1.TxtUsername.text + "' AND USERID = '" + Trim(LstVwSearchMgm.SelectedItem.SubItems(11)) + "'"
        ElseIf UCase(Trim(MDIForm1.txtlevel.text)) = "ADMINISTRATOR" Or UCase(MDIForm1.txtlevel.text) = "ADMIN" Then
            CMDSQL = "SELECT USERID FROM usertbl "
        End If
        
        '@@ 19 Juli 2010 .. Ini pengalihan error buka data oleh agent
        On Error GoTo Salah
        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
selanjutnya:
        If M_objrs.RecordCount = 0 Then
        STRSQL = "SELECT * FROM USERTBL WHERE  USERID IN (SELECT  agentprev FROM MGM WHERE CUSTID ='" + Trim(LstVwSearchMgm.SelectedItem.SubItems(1)) + "') AND TEAM ='" + MDIForm1.TxtUsername.text + "'"
            Set MOBJRSKISRUT = New ADODB.Recordset
                MOBJRSKISRUT.CursorLocation = adUseClient
                MOBJRSKISRUT.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
            If MOBJRSKISRUT.RecordCount > 0 Then
                GoTo ke
            End If
    
            MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "Aplikasi"
            Set MOBJRSKISRUT = Nothing
            Set M_objrs = Nothing
            Exit Sub
        End If
        Set M_objrs = Nothing
    End If
    
    '=====asep20200502==='
    F_OPENCUST = True
    FrmCC_Colection.bBuka = False
    F_OPENCUST = False
    '======================='
ke:
    Me.MousePointer = vbHourglass
    Flag_mgm = False
    'Matikan main timer activity By Izuddin 16042013
    main_timer_activity = 0
    'MDIForm1.Timer7.Enabled = False
    '--
    'FrmCC_Colection.Show vbModal
    'SET WAKTU LOGOUT
    M_OBJCONN.Execute "UPDATE usertbl SET last_logout='now()' WHERE userid='" + Trim(MDIForm1.TxtUsername.text) + "'"
    '--
    If MDIForm1.txtlevel.text = "Agent" Then
        FrmCC_Colection.Show vbModal
    Else
        FrmCC_Colection.Show 'vbModal
    End If
    
    If LstVwSearchMgm.ListItems.Count <> 0 Then
        strStatusCpa = LstVwSearchMgm.SelectedItem.SubItems(23)
    End If
    Me.MousePointer = vbNormal
    Exit Sub

Salah:
    CMDSQL = "select * from mgm where "
    CMDSQL = CMDSQL + "agent = '" + MDIForm1.TxtUsername.text + "' and custid='"
    CMDSQL = CMDSQL + Trim(LstVwSearchMgm.SelectedItem.SubItems(1)) + "'"
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    GoTo selanjutnya
    Exit Sub
    
End Sub

Private Sub LstVwSearchmgm_KeyPress(KeyAscii As Integer)
Dim M_objrs As ADODB.Recordset
If KeyAscii = 13 Then
    Call LstVwSearchMgm_DblClick
    Exit Sub
End If
If UCase(MDIForm1.txtlevel.text) <> "AGENT1" Then
    If KeyAscii = 112 Or KeyAscii = 80 Then
        If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Then
        'If UCase(mdiform1.txtlevel.text) = "TEAMLEADER" Then
            Set M_objrs = New ADODB.Recordset
            M_objrs.CursorLocation = adUseClient
            M_objrs.Open "SELECT USERID FROM usertbl WHERE TEAM ='" + MDIForm1.TxtUsername.text + "' AND USERID = '" + LstVwSearchMgm.SelectedItem.SubItems(11) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If M_objrs.RecordCount <> 0 Then
            Else
                MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "Aplikasi"
                Set M_objrs = Nothing
                Exit Sub
            End If
            Set M_objrs = Nothing
        Else

        End If
    End If
    

    If KeyAscii = 73 Or KeyAscii = 105 Then
'        b_pindah = True
'        FrmCustIdTransfer.Show
        If UCase(Left(MDIForm1.txtlevel.text, 5)) = "ADMIN" Then
            ' ====== 12 Agustus 2014
            ' b_pindah = True
            ' FrmCustIdTransfer.Show
        ElseIf UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Or UCase(MDIForm1.txtlevel.text) = "MANAGER" Then
                    CMDSQL = "SELECT USERID FROM usertbl WHERE USERID = '" + LstVwSearchMgm.SelectedItem.SubItems(11) + "'"
                    'CMDSQL = "SELECT USERID FROM usertbl WHERE USERID = 'ADM01'"
                    
        '@@18092012 Teamleader Tidak diperbolehkan melakukan transfer
        ElseIf UCase(MDIForm1.txtlevel.text) = "TEAMLEADER" Or UCase(MDIForm1.txtlevel.text) = "MANAGER" Then
                     MsgBox "Mohon maaf, pemindahan account data saat ini tidak diperbolehkan!", vbOKOnly + vbExclamation, "Peringatan"
                     
                     Exit Sub
                     CMDSQL = "SELECT USERID FROM usertbl WHERE TEAM ='" + MDIForm1.TxtUsername.text + "' AND USERID = '" + LstVwSearchMgm.SelectedItem.SubItems(11) + "'"
        ElseIf UCase(MDIForm1.txtlevel.text) = "AGENT" Then
                CMDSQL = "SELECT USERID FROM usertbl where   userid='REVIEW'"
        End If
        
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_objrs.RecordCount <> 0 Then
            If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Or UCase(MDIForm1.txtlevel.text) = "MANAGER" Or UCase(MDIForm1.txtlevel.text) = "ADMINISTRATOR" Then
                b_pindah = True
                'FrmCustIdTransfer.Show
            Else
            MsgBox "Mohon maaf, Anda Tidak Berhak Melakukan Transfer Account!", vbOKOnly + vbExclamation, "Peringatan"
            'MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "Aplikasi"
            Set M_objrs = Nothing
            Exit Sub
            End If
        End If
        Set M_objrs = Nothing
    End If
    
    If KeyAscii = 88 Or KeyAscii = 120 Then
        Dim n%
        
        If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" And Combo1(2).text = "" Then
            Set M_objrs = New ADODB.Recordset
            M_objrs.CursorLocation = adUseClient
            M_objrs.Open "SELECT USERID FROM usertbl WHERE TEAM ='" + MDIForm1.TxtUsername.text + "' AND USERID = '" + LstVwSearchMgm.SelectedItem.SubItems(11) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If M_objrs.RecordCount <> 0 Then
            Else
                MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "Aplikasi"
                Set M_objrs = Nothing
                Exit Sub
            End If
            Set M_objrs = Nothing
        End If
    End If
End If
End Sub
Private Sub Timer1_Timer()
'strsql = "select * from usertbl where tgl update between "
End Sub

Public Sub renderdonk()
Dim NAMACUST As String
Dim NamaAgent As String
Dim DATASOURCE As String
Dim TGLLAHIR As String
Dim nmagentprev As String
Dim OFFPHONE As String
Dim OFFPHONE2 As String
Dim HOMEPHONE As String
Dim HOMEPHONE2 As String
Dim MOBILEPHONE As String
Dim MOBILEPHONE2 As String
Dim FAXPHONE As String
Dim Lcustid As String
Dim Lcustno As String
Dim FAXPHONE2 As String
Dim KETHSLKERJA As String
Dim lLastCallDate As String
Dim lStatusCek As String
Dim sPending As String
Dim FCEKSTATUS As String
Dim strverify As String
Dim strapprovel As String
Dim m_data As New CLS_FRMSEARCH
Dim M_objrs As New ADODB.Recordset
Dim PANJANG As Integer
Dim strReject As String
Dim strSukses As String
Dim strapprovelyet As String
Dim strinject As String
Dim strmarkup As String
Dim BlokedEntry As String
Dim STSLOCKTL As String
Dim STSfromaccount As String
'jejaktian(tambahantian)
    Dim AHOMENO As String
    Dim AHOMENO2 As String
    Dim AOFFICENO As String
    Dim AOFFICENO2 As String
    Dim extoffice As String
    Dim extoffice2 As String
    Dim homenoadd1 As String
    Dim ahomenoadd1 As String
    Dim homenoadd2 As String
    Dim ahomenoadd2 As String
    Dim officenoadd1 As String
    Dim aofficenoadd2 As String
    Dim officenoadd2 As String
    Dim mobilenoadd1 As String
    Dim mobilenoadd2 As String
    Dim ec_telp As String
    Dim alamatrumah As String
    Dim alamatkantor As String
    Dim alamatec As String
    Dim tglsource As String
    '===============================

    F_CEK = Empty
    WO_DATE = Empty
    If UCase(Left(MDIForm1.txtlevel.text, 5)) = "ADMIN" Or UCase(Left(MDIForm1.txtlevel.text, 5)) = "SUPER" Or UCase(Left(MDIForm1.txtlevel.text, 5)) = "TEAML" Then
    Else
    Call CEK_STATUS_F_CEK
    End If
'    Call UPDATE_BP
        M_objrs.CursorLocation = adUseClient
        CMDSQL = "SELECT * FROM usertbl WHERE USERID = '" + MDIForm1.TxtUsername.text + "'"
        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic

        If Not M_objrs.EOF Then
           strinject = IIf(IsNull(M_objrs!lockdarispv), "", M_objrs!lockdarispv)
           If strinject = "" Then
              Bloked = ""
           Else
            Bloked = IIf(IsNull(M_objrs!lockdarispv), "", Replace(M_objrs!lockdarispv, "@", "'"))
            BlokedEntry = IIf(IsNull(M_objrs!lock_entry_lpd), "", M_objrs!lock_entry_lpd)
           End If
         strmarkup = IIf(IsNull(M_objrs!lockmarkup), "", M_objrs!lockmarkup)
        End If
        
        While Not M_objrs.EOF
                 StsVl = CStr(Trim(IIf(IsNull(M_objrs!f_VL), "", M_objrs!f_VL)))
                StsON = CStr(Trim(IIf(IsNull(M_objrs!f_ON), "", M_objrs!f_ON)))
                StsOS = CStr(Trim(IIf(IsNull(M_objrs!f_OS), "", M_objrs!f_OS)))
                StsSK = CStr(Trim(IIf(IsNull(M_objrs!f_SK), "", M_objrs!f_SK)))
                StsPR = CStr(Trim(IIf(IsNull(M_objrs!f_PR), "", M_objrs!f_PR)))
                StsPTP = CStr(Trim(IIf(IsNull(M_objrs!f_ptp), "", M_objrs!f_ptp)))
                StsBP = CStr(Trim(IIf(IsNull(M_objrs!f_bp), "", M_objrs!f_bp)))
                StsPOP = CStr(Trim(IIf(IsNull(M_objrs!f_pop), "", M_objrs!f_pop)))
                StsSP = CStr(Trim(IIf(IsNull(M_objrs!f_sp), "", M_objrs!f_sp)))
                StsUC = CStr(Trim(IIf(IsNull(M_objrs!F_UC), "", M_objrs!F_UC)))
                StsWO_Date = CStr(Trim(IIf(IsNull(M_objrs!f_WO_DATE), "", M_objrs!f_WO_DATE)))
                StsWO_2009 = CStr(Trim(IIf(IsNull(M_objrs!f_WO_2009), "", M_objrs!f_WO_2009)))
                StsWO_2008 = CStr(Trim(IIf(IsNull(M_objrs!f_WO_2008), "", M_objrs!f_WO_2008)))
                StsWO_2007 = CStr(Trim(IIf(IsNull(M_objrs!f_WO_2007), "", M_objrs!f_WO_2007)))
                StsWO_2006 = CStr(Trim(IIf(IsNull(M_objrs!f_WO_2006), "", M_objrs!f_WO_2006)))
                StsWO_2005 = CStr(Trim(IIf(IsNull(M_objrs!f_WO_2005), "", M_objrs!f_WO_2005)))
                StsWO_2004 = CStr(Trim(IIf(IsNull(M_objrs!f_WO_2004), "", M_objrs!f_WO_2004)))
                StsWO_2003 = CStr(Trim(IIf(IsNull(M_objrs!f_WO_2003), "", M_objrs!f_WO_2003)))
                StsWO_2002 = CStr(Trim(IIf(IsNull(M_objrs!f_WO_2002), "", M_objrs!f_WO_2002)))
                StsWO_2001 = CStr(Trim(IIf(IsNull(M_objrs!f_WO_2001), "", M_objrs!f_WO_2001)))
                StsWO_2000 = CStr(Trim(IIf(IsNull(M_objrs!f_WO_2000), "", M_objrs!f_WO_2000)))
                StsWO_1999 = CStr(Trim(IIf(IsNull(M_objrs!F_WO_1999), "", M_objrs!F_WO_1999)))
                LUserType = CStr(Trim(IIf(IsNull(M_objrs!usertype), "", M_objrs!usertype)))
                STSLOCKTL = CStr(Trim(IIf(IsNull(M_objrs!lockdarispvbuattl), "", M_objrs!lockdarispvbuattl)))
                STSfromaccount = CStr(Trim(IIf(IsNull(M_objrs!fromaccount), "", M_objrs!fromaccount)))
                
                M_objrs.MoveNext
        Wend
        Set M_objrs = Nothing
        StsAll = StsVl + StsPR + StsBP + StsPOP + StsUC + StsSK + StsON + StsOS
       
       If STSLOCKTL <> Empty Then
        If Left(cmb_kdagent.text, 5) = "LUNAS" Then
                If STSfromaccount = "LUNAS PENDING" Then
                    STSLOCKTL = STSLOCKTL
                ElseIf STSfromaccount = "LUNAS COMPLETE" Then
                      STSLOCKTL = STSLOCKTL
                Else
                     STSLOCKTL = ""
                End If
                
        Else
                STSLOCKTL = ""
        End If
        End If
        
     If StsAll <> "" Then
            If LUserType = "1" Then
                    If StsUC = "UC" Then
                     '       F_CEK = "substring(F_CEK_NEW,1,3)  IN('NK-','MV-','WN-','" + StsNa + "','" + StsOP + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "','')"
                     '   Else
                            F_CEK = "substring(F_CEK_NEW,1,3)  IN( '" + StsVl + "','" + StsPR + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "','" + StsSK + "','" + StsON + "','" + StsOS + "','')"
                        End If
                        
                    End If
     End If
        
      If StsWO_Date = "1" Then
            If LUserType = "1" Then
                WO_DATE = "substring(B_D,1,4) in ('" + StsWO_2009 + "','" + StsWO_2008 + "','" + StsWO_2007 + "','" + StsWO_2006 + "','" + StsWO_2005 + "', "
                WO_DATE = WO_DATE + "'" + StsWO_2004 + "', '" + StsWO_2003 + "', '" + StsWO_2002 + "', '" + StsWO_2001 + "','" + StsWO_2000 + "','" + StsWO_1999 + "')"
            End If
      End If
        If Trim(Text1(0).text) = Empty And Trim(cmb_kdagent.text) = Empty And Combo1(2).text = Empty And Len(TDBMask2.Value) < 1 And Len(TDBMask1.text) < 1 And TdDob.ValueIsNull And Len(txtnocard.text) < 3 _
        And cmbStsLastCall(0).text = "" And CmbStatusCek.text = "" And DtLastCall(0).ValueIsNull And CekDtDistribute.Value = 0 And Combo3.text = "" Then
            MsgBox "Masukan Kriteria Customer Yang Akan Dicari...!!!", vbCritical + vbOKOnly, "Peringatan"
            Text1(0).SetFocus
            Set m_data = Nothing
            Exit Sub
        Else
        
         LstVwSearchMgm.ListItems.clear
         Frame3.Visible = True
         If CekDtDistribute.Value = 1 Then
            NamaAgent = "AGENT is null"
         Else
            If txtnocard.text <> Empty Then
                Lcustid = "CUSTID LIKE " + "'%" + UBAH_QUOTE(txtnocard.text) + "%'"
            ElseIf txtregion.text <> Empty Then
                    Lcustno = "region LIKE " + "'%" + UBAH_QUOTE(txtregion.text) + "%'"
            Else
                If Text1(0).text <> Empty Then
                    NAMACUST = "name LIKE " + "'%" + UBAH_QUOTE(Text1(0).text) + "%'"
                End If
                If cmb_kdagent.text <> Empty Then
                    If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Then
                        NamaAgent = "AGENT in(select userid from usertbl where team='" + Trim(cmb_kdagent.text) + "')"
                    ElseIf UCase(MDIForm1.txtlevel.text) = "AGENT" Or UCase(MDIForm1.txtlevel.text) = "ADMIN" Then
                        NamaAgent = "AGENT = '" + Trim(cmb_kdagent.text) + "'"
                    Else
                    
                    End If
                
                
                    If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Then
                        nmagentprev = "agentprev IN (SELECT USERID FROM USERTBL WHERE TEAM='" + MDIForm1.TxtUsername.text + "' )"
                ElseIf UCase(MDIForm1.txtlevel.text) = "AGENT" Then
                        nmagentprev = "agentprev ='" + MDIForm1.TxtUsername.text + "' "
                Else
                        nmagentprev = ""
                End If
              End If
                If Combo1(2).text <> Empty Then
                    DATASOURCE = "RECSOURCE = '" + Trim(Combo1(2).text) + "'"
                End If
                If TdDob.ValueIsNull Then
                Else
                    TGLLAHIR = "DOB = '" + Format(TdDob.text, "yyyy/mm/dd") + "'"
                End If
                If Len(TDBMask1.text) > 1 Then
                    OFFPHONE = "OFFICENO Like '%" + TDBMask1.text + "%'"
                    OFFPHONE2 = "OFFICENO2 Like '%" + TDBMask1.text + "%'"
                    HOMEPHONE = "HOMENO Like '%" + TDBMask1.text + "%'"
                    HOMEPHONE2 = "HOMENO2 Like '%" + TDBMask1.text + "%'"
                    FAXPHONE = "FAXNO Like '%" + TDBMask1.text + "%'"
                    FAXPHONE2 = "FAXNO2 Like '%" + TDBMask1.text + "%'"
                End If
                
                If Len(TDBMask2.Value) > 1 Then
                    MOBILEPHONE = "MOBILENO like '%" + TDBMask2.Value + "%'"
                    MOBILEPHONE2 = "MOBILENO2 like '%" + TDBMask2.Value + "%'"
                End If
                
                
                If Left(Combo3.text, 1) = 6 Then
                    strverify = "intverify=0 and  stscpa=1 and (resultcpa is null or resultcpa='')"
                End If
                
                If Left(Combo3.text, 1) = 1 Then
                  strapprovel = " (intapprovel=0 or intapprovel is null )  and intverify=1  and (resultcpa is null or resultcpa='')  "
                End If
                
                If Left(Combo3.text, 1) = 4 Then
                  strapprovelyet = " (intapprovel=0 or intapprovel is null )  and (intverify=0 or intverify isnull) and stscpa=1 and (resultcpa is null or resultcpa='')  "
                End If
                
                If Left(Combo3.text, 1) = 2 Then
                  strReject = "   resultcpa ='GAGAL'  "
                End If
                
                If Left(Combo3.text, 1) = 3 Then
                  strSukses = "   resultcpa ='SUKSES'  "
                End If
                
                If DtLastCall(0).ValueIsNull Then
                Else
                    lLastCallDate = "TGLSTATUS BETWEEN '" + Format(DtLastCall(0).Value, "MM/DD/YYYY") & " " & CStr(DTimeLastCall(0).Value) + "' AND '" + Format(DtLastCall(1).Value, "MM/DD/YYYY") & " " & CStr(DTimeLastCall(1).Value) + "'"
                End If
        End If
        End If
          
                'Unload FRM_SEARCH
                If Check1.Value = 0 Then
                    Set m_cari = m_data.QUERY_SEARCH_CONDITION(M_OBJCONN, NAMACUST, NamaAgent, DATASOURCE, TGLLAHIR, _
                                                            OFFPHONE, OFFPHONE2, HOMEPHONE, HOMEPHONE2, MOBILEPHONE, _
                                                            MOBILEPHONE2, FAXPHONE, FAXPHONE2, MDIForm1.Text3.text, Lcustid, F_CEK, lLastCallDate, lStatusCek)
                Else
                   If strmarkup <> "" Then
                    F_CEK = ""
                    WO_DATE = ""
                    BlokedEntry = ""
                    Bloked = ""
                End If
                    Set m_cari = m_data.QUERY_SEARCH_CONDITION_mgm(M_OBJCONN, NAMACUST, NamaAgent, DATASOURCE, TGLLAHIR, _
                                                             OFFPHONE, OFFPHONE2, HOMEPHONE, HOMEPHONE2, MOBILEPHONE, _
                                                            MOBILEPHONE2, FAXPHONE, FAXPHONE2, MDIForm1.Text3.text, _
                                                            AHOMENO, AHOMENO2, AOFFICENO, AOFFICENO2, extoffice, extoffice2, homenoadd1, ahomenoadd1, homenoadd2, ahomenoadd2, officenoadd1, aofficenoadd2, officenoadd2, mobilenoadd1, _
                                                            mobilenoadd2, ec_telp, alamatrumah, alamatkantor, alamatec, _
                                                            Lcustid, F_CEK, lLastCallDate, lStatusCek, sPending, FCEKSTATUS, WO_DATE, strverify, strapprovel, strapprovelyet, strReject, strSukses, Bloked, BlokedEntry, strmarkup, nmagentprev, STSLOCKTL, qpub, tglsource, "", "", , , , Lcustno, "")
                End If
        
            If m_cari.RecordCount = 0 Then
                MsgBox "Data Tidak Ditemukan", vbInformation + vbOKOnly, "Aplikasi"
                Set m_data = Nothing
                Exit Sub
            Else
               
                search_ok = True
                If Check1.Value = 1 Then
                    'kalau found refferall data
                    SSTab1.Tab = 0
'                    Call show_UCDATA
                    Call show_Search_mgmData
                    
                Else
                    ' kalau found mgm data
                    SSTab1.Tab = 1

                End If
            End If
        End If

End Sub

'=================================@@11022013 Buat Nambahin script Cari Data ALL ================
Private Sub CariDataAll()
    Dim harga As Double
    Dim ListItem As ListItem
    Dim Lcustid1 As String
    Dim Lcustid2 As String
    Dim LCall As String
    Dim i, K As Integer
    Dim CMDSQL As String
    Dim sPending As String
    Dim M_objrs As ADODB.Recordset
    Dim VOLUMEAMOUNT As Double
    Dim statusprior As String
    Dim exp%
    Dim totamount As Double
    Dim TOTCURBALANCE As Double
    Dim kdprofile_aksesall As String
    
    Dim M_Objrs_Cek As ADODB.Recordset
    'Dim CMDSQL As String
    
    '@@19022013
    Dim M_ObjrsCekStatus As ADODB.Recordset
    
    'Cek statusnya langsung dari usertbl aja
    ' UPDATE 21 MEI 2013 IZUDDIN
    If UCase(MDIForm1.txtlevel.text) = "AGENT" Or _
        UCase(MDIForm1.txtlevel.text) = "TEAMLEADER" Then
        CMDSQL = "select f_akses_all_acc,profile_akses_all from usertbl where userid='"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "'"
        Set M_ObjrsCekStatus = New ADODB.Recordset
        M_ObjrsCekStatus.CursorLocation = adUseClient
        M_ObjrsCekStatus.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_ObjrsCekStatus.RecordCount > 0 Then
            AksesAllAcc = IIf(IsNull(M_ObjrsCekStatus("f_akses_all_acc")), "", M_ObjrsCekStatus("f_akses_all_acc"))
            kdprofile_aksesall = IIf(IsNull(M_ObjrsCekStatus("profile_akses_all")), "", M_ObjrsCekStatus("profile_akses_all"))
        End If
        Set M_ObjrsCekStatus = Nothing
    End If
    
    If AksesAllAcc = "1" Then
        CMDSQL = "SELECT * FROM mgm WHERE custid in "
        CMDSQL = CMDSQL + " (select b.custid from tbl_profile_aksesall a, tbl_cust_aksesall b  "
        CMDSQL = CMDSQL + " where a.kd_profile=b.kd_profile AND a.kd_profile='"
        CMDSQL = CMDSQL + kdprofile_aksesall + "' "
        CMDSQL = CMDSQL + " AND a.waktu_akhir > now() AND  "
        CMDSQL = CMDSQL + " a.waktu_awal <= now()) AND agent='AKSESALL'"
    Else
        ' Balikkin ke agent sebelumnya 03 Juni 2014
        M_OBJCONN.Execute "UPDATE mgm SET agent=agent_asli WHERE agent='AKSESALL' AND agent_asli IS NOT NULL AND custid not in (SELECT custid FROM tbl_cust_aksesall);"
        Exit Sub
    End If
    
    i = 1
    
    Set M_Objrs_Cek = New ADODB.Recordset
    M_Objrs_Cek.CursorLocation = adUseClient
    M_Objrs_Cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
    datajml = M_Objrs_Cek.RecordCount
    
    Me.MousePointer = vbHourglass
    
    ProgressBar1.Max = M_Objrs_Cek.RecordCount + 1
    
    If M_Objrs_Cek.RecordCount > 0 Then
        ' SET FOCUS ACCOUNT 13 MEI 2014 --------------
        If UCase(MDIForm1.txtlevel.text) = "AGENT" Then
            LstVwSearchMgm.ListItems.clear
        End If
        ' --------------------------------------
        MsgBox "INFO: Anda memiliki tambahan account sebanyak: " & M_Objrs_Cek.RecordCount & "  yang dapat di collect bersama. Lihat baris paling bawah dan bertanda merah!", vbOKOnly + vbInformation, "Informasi"
        While Not M_Objrs_Cek.EOF
            ProgressBar1.Value = M_Objrs_Cek.Bookmark
            Lcustid1 = CStr(IIf(IsNull(M_Objrs_Cek!CustId), "", M_Objrs_Cek!CustId))
            sPending = CStr(Trim(IIf(IsNull(M_Objrs_Cek!f_Pending), "", M_Objrs_Cek!f_Pending)))
            
            Set ListItem = LstVwSearchMgm.ListItems.ADD(, , M_Objrs_Cek.Bookmark)
            
            If MDIForm1.txtlevel.text = "TeamLeader" Then
                If IIf(IsNull(M_Objrs_Cek("stscpa")), "0", M_Objrs_Cek("stscpa")) = 1 Then
                    ListItem.ForeColor = vbRed
                End If
                
                If IIf(IsNull(M_Objrs_Cek("intapprovel")), "0", M_Objrs_Cek("intapprovel")) = 1 Then
                  ListItem.ForeColor = vbBlue
                End If
            End If
                    
            statusprior = IIf(IsNull(M_Objrs_Cek("StatusPrior")), "", M_Objrs_Cek("StatusPrior"))
            ListItem.SubItems(1) = IIf(IsNull(M_Objrs_Cek("CUSTID")), "", M_Objrs_Cek("CUSTID"))
            ListItem.SubItems(2) = IIf(IsNull(M_Objrs_Cek("PRIOR")), "", M_Objrs_Cek("PRIOR"))
            ListItem.SubItems(3) = IIf(IsNull(M_Objrs_Cek("NAME")), "", M_Objrs_Cek("NAME"))
            ListItem.SubItems(4) = IIf(IsNull(M_Objrs_Cek("RECSOURCE")), "", M_Objrs_Cek("RECSOURCE"))
            ListItem.SubItems(5) = IIf(IsNull(M_Objrs_Cek("NEXTACTDATE")), "", Format(M_Objrs_Cek("NEXTACTDATE"), "dd/mm/yyyy hh:nn"))
            ListItem.SubItems(6) = IIf(IsNull(M_Objrs_Cek("NEXTACT")), "", M_Objrs_Cek("NEXTACT"))
            ListItem.SubItems(7) = IIf(IsNull(M_Objrs_Cek("REMARKS")), "", M_Objrs_Cek("REMARKS"))
            ListItem.SubItems(8) = CStr(IIf(IsNull(M_Objrs_Cek("kethslkerja_new")), "", M_Objrs_Cek("kethslkerja_new")) & " " & sPending)
            ListItem.SubItems(9) = IIf(IsNull(M_Objrs_Cek("StatusCall")), "", M_Objrs_Cek("StatusCall"))
            ListItem.SubItems(11) = IIf(IsNull(M_Objrs_Cek("AGENT")), "", M_Objrs_Cek("AGENT"))
            
            If UCase(MDIForm1.txtlevel.text) <> "SUPERVISOR" Then
                If Format(IIf(IsNull(M_Objrs_Cek("flaglead")), 0, M_Objrs_Cek("flaglead")), "##,###") = 1 Then
                    ListItem.SubItems(12) = ""
                Else
                    ListItem.SubItems(12) = Format(IIf(IsNull(M_Objrs_Cek("Principal")), 0, M_Objrs_Cek("Principal")), "##,###")
                End If
            Else
                 ListItem.SubItems(12) = Format(IIf(IsNull(M_Objrs_Cek("Principal")), 0, M_Objrs_Cek("Principal")), "##,###")
            End If

            ListItem.SubItems(13) = Format(IIf(IsNull(M_Objrs_Cek("AmountWo")), 0, M_Objrs_Cek("AmountWo")), "##,###")
            totamount = totamount + IIf(IsNull(M_Objrs_Cek("AmountWo")), 0, M_Objrs_Cek("AmountWo"))
            
            ListItem.SubItems(14) = Format(IIf(IsNull(M_Objrs_Cek("OpenDate")), "", M_Objrs_Cek("OpenDate")), "yyyy/mm/dd")
            ListItem.SubItems(15) = Format(IIf(IsNull(M_Objrs_Cek("B_D")), 0, M_Objrs_Cek("B_D")))
            ListItem.SubItems(16) = Format(IIf(IsNull(M_Objrs_Cek("Pay_Dt")), 0, M_Objrs_Cek("Pay_Dt")), "yyyy/mm/dd")
            
            ListItem.SubItems(17) = Format(IIf(IsNull(M_Objrs_Cek("lastpay")), 0, M_Objrs_Cek("lastpay")), "##,###")
            
            ListItem.SubItems(18) = IIf(IsNull(M_Objrs_Cek("TGLSTATUS")), "", Format(M_Objrs_Cek("TGLSTATUS"), "YYYY/MM/DD"))
            ListItem.SubItems(19) = IIf(IsNull(M_Objrs_Cek("TGLCALL")), "", Format(M_Objrs_Cek("TGLCALL"), "YYYY/MM/DD"))
            ListItem.SubItems(20) = IIf(IsNull(M_Objrs_Cek("Kethslkerja")), "", M_Objrs_Cek("Kethslkerja"))
            ListItem.SubItems(21) = Format(IIf(IsNull(M_Objrs_Cek("TGLINCOMING")), "", M_Objrs_Cek("TGLINCOMING")), "YYYY/MM/DD")
            ListItem.SubItems(23) = IIf(IsNull(M_Objrs_Cek("resultcpa")), "", M_Objrs_Cek("resultcpa"))
            ListItem.SubItems(24) = IIf(IsNull(M_Objrs_Cek("tglinsertfrmcpa")), "", M_Objrs_Cek("tglinsertfrmcpa"))
            ListItem.SubItems(25) = Format(IIf(IsNull(M_Objrs_Cek("curbal")), "", M_Objrs_Cek("curbal")), "##,###")
            TOTCURBALANCE = TOTCURBALANCE + IIf(IsNull(M_Objrs_Cek("curbal")), 0, M_Objrs_Cek("curbal"))
           
            '@@16032011 Tambahan DOB dan No KTP
            ListItem.SubItems(26) = IIf(IsNull(M_Objrs_Cek("dob")), "", Format(M_Objrs_Cek("dob"), "yyyy-mm-dd"))
            ListItem.SubItems(27) = IIf(IsNull(M_Objrs_Cek("ktpno")), "", M_Objrs_Cek("ktpno"))
                 
SorryLompat:
            
            VOLUMEAMOUNT = VOLUMEAMOUNT + IIf(IsNull(M_Objrs_Cek("AmountWo")), 0, M_Objrs_Cek("AmountWo"))
            
            ListItem.ForeColor = vbRed
            For K = 1 To 26
                ListItem.ListSubItems(K).ForeColor = vbRed
            Next K
            M_Objrs_Cek.MoveNext
        Wend
        
    Else

        ' UPDATE 19 AGUSTUS 2014 agent_asli dihilangkan
        CMDSQL = "UPDATE mgm SET agent=agent_asli WHERE " & _
                " agent='AKSESALL' AND custid in(SELECT custid FROM tbl_cust_aksesall a,tbl_profile_aksesall b WHERE " & _
                " a.kd_profile=b.kd_profile AND b.waktu_akhir < now()) AND agent_asli is not null"

        M_OBJCONN.Execute CMDSQL
        CMDSQL = "DELETE FROM tbl_cust_aksesall "
        CMDSQL = CMDSQL & " WHERE kd_profile in (SELECT kd_profile FROM tbl_profile_aksesall WHERE waktu_akhir < now()) "
        M_OBJCONN.Execute CMDSQL
        
                CMDSQL = "UPDATE usertbl SET profile_akses_all=null,f_akses_all_acc=null,f_pesanresetauto=null WHERE profile_akses_all in (SELECT kd_profile FROM tbl_profile_aksesall WHERE waktu_akhir < now());"
                M_OBJCONN.Execute CMDSQL
'            End If
            AksesAllAcc = ""
'        End If
'        Set M_ObjrsCekStatus = Nothing
    End If
        
    MousePointer = vbNormal
    Set M_Objrs_Cek = Nothing
End Sub


