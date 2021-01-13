VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form form_payment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment"
   ClientHeight    =   3405
   ClientLeft      =   8025
   ClientTop       =   3375
   ClientWidth     =   5385
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5385
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Caption         =   "Custid"
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "form_payment.frx":0000
         Left            =   1320
         List            =   "form_payment.frx":000A
         TabIndex        =   13
         Top             =   2250
         Width           =   3495
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   1560
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   450
         Calculator      =   "form_payment.frx":001E
         Caption         =   "form_payment.frx":003E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_payment.frx":00AA
         Keys            =   "form_payment.frx":00C8
         Spin            =   "form_payment.frx":0112
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999999999999
         MinValue        =   -99999999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.CommandButton Command1 
         Caption         =   "PAY"
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   480
         Width           =   3495
      End
      Begin TDBDate6Ctl.TDBDate cmbDateSch 
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Top             =   1175
         Width           =   3540
         _Version        =   65536
         _ExtentX        =   6244
         _ExtentY        =   556
         Calendar        =   "form_payment.frx":013A
         Caption         =   "form_payment.frx":0252
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_payment.frx":02BE
         Keys            =   "form_payment.frx":02DC
         Spin            =   "form_payment.frx":033A
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
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
      Begin VB.Label Label1 
         Caption         =   "Status Paymen"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   12
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Agent"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Payment"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Paydate"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nama Ch"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Custid"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "form_payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub sets()
    Text1.text = FrmCC_Colection.lblCustId.text
    Text2.text = FrmCC_Colection.lblNama.text
    Text3.text = FrmCC_Colection.lblaoc.Caption
End Sub

Private Sub Command1_Click()
    Call inserttbllunas
    Call FrmCC_Colection.isi_datapayment
End Sub

Private Sub Form_Load()
    Call sets
End Sub

Private Sub inserttbllunas()
    strins = "insert into tbllunas (custid,paydate,payment,agent,tglinsert,name_ch,datafrom) values " & vbCrLf
    strins = strins & "('" & Text1.text & "', '" & Format(cmbDateSch.Value, "yyyy-mm-dd") & "', '" & TDBNumber1.Value & "','" & Text3.text & "', now(), '" & Text2.text & "', '" & Combo1.text & "');"
    M_OBJCONN.Execute strins
    MsgBox "Inserted"
    Unload Me
End Sub
