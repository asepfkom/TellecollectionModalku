VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form frmextensioncc 
   BorderStyle     =   0  'None
   Caption         =   "&H00ABE18E&"
   ClientHeight    =   4425
   ClientLeft      =   14070
   ClientTop       =   2370
   ClientWidth     =   4320
   LinkTopic       =   "Form2"
   ScaleHeight     =   4425
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FCFCFC&
      Caption         =   "OTHER INFO"
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   4275
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   7
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   3090
         Width           =   1875
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   6
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   2790
         Width           =   1875
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   0
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1245
         Width           =   1875
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   5
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2190
         Width           =   1875
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   4
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1875
         Width           =   1875
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   3
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3975
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   2
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1560
         Width           =   1875
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   1
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   3690
         Visible         =   0   'False
         Width           =   1875
      End
      Begin TDBDate6Ctl.TDBDate instalment_complite 
         Height          =   255
         Left            =   1815
         TabIndex        =   15
         Top             =   660
         Width           =   1860
         _Version        =   65536
         _ExtentX        =   3281
         _ExtentY        =   450
         Calendar        =   "frmextensioncc.frx":0000
         Caption         =   "frmextensioncc.frx":0118
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmextensioncc.frx":0184
         Keys            =   "frmextensioncc.frx":01A2
         Spin            =   "frmextensioncc.frx":0200
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483645
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
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   3.54027066542603E-316
         CenturyMode     =   0
      End
      Begin TDBNumber6Ctl.TDBNumber lblLastPay 
         Height          =   255
         Left            =   1815
         TabIndex        =   18
         Top             =   2505
         Width           =   1875
         _Version        =   65536
         _ExtentX        =   3307
         _ExtentY        =   450
         Calculator      =   "frmextensioncc.frx":0228
         Caption         =   "frmextensioncc.frx":0248
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmextensioncc.frx":02B4
         Keys            =   "frmextensioncc.frx":02D2
         Spin            =   "frmextensioncc.frx":031C
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483645
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
         ReadOnly        =   -1
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1638405
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBDate6Ctl.TDBDate dpdate 
         Height          =   255
         Left            =   1815
         TabIndex        =   22
         Top             =   945
         Width           =   1875
         _Version        =   65536
         _ExtentX        =   3307
         _ExtentY        =   450
         Calendar        =   "frmextensioncc.frx":0344
         Caption         =   "frmextensioncc.frx":045C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmextensioncc.frx":04C8
         Keys            =   "frmextensioncc.frx":04E6
         Spin            =   "frmextensioncc.frx":0544
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483645
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
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   3.54027066542603E-316
         CenturyMode     =   0
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Status Loan"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   165
         TabIndex        =   24
         Top             =   3105
         Width           =   1515
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   150
         TabIndex        =   23
         Top             =   900
         Width           =   1560
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Admin Fee"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   165
         TabIndex        =   20
         Top             =   2805
         Width           =   1515
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Outstand Loan"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   135
         TabIndex        =   19
         Top             =   2505
         Width           =   1515
      End
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Monthly Insterest L"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   120
         TabIndex        =   17
         Top             =   375
         Width           =   1590
      End
      Begin VB.Label Label34 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1815
         TabIndex        =   16
         Top             =   375
         Width           =   1860
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000007&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3825
         TabIndex        =   14
         Top             =   90
         Width           =   375
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Outstand Late Fee"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   2235
         Width           =   1680
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Paid Late Fee"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   1905
         Width           =   1560
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "X Loan Code"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   210
         TabIndex        =   5
         Top             =   3960
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Paid Interest"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1575
         Width           =   1560
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Add Late Fee"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   3690
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Outstand Principle"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   1230
         Width           =   1560
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Inst. Complt. Date"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   615
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmextensioncc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_Load()
    Call isi
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Label2_Click()
    Unload Me
End Sub

Private Sub isi()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
    CMDSQL = "SELECT * FROM mgm where custid = '" & FrmCC_Colection.lblCustId.text & "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    txta(0).text = cnull(M_objrs!out_principle)
    txta(1).text = cnull(M_objrs!add_latefee)
    txta(2).text = cnull(M_objrs!paid_interest)
    'txta(3).text = cnull(M_objrs!x_loan_code)
    txta(4).text = cnull(M_objrs!paid_latefee)
    txta(5).text = cnull(M_objrs!late_fee)
    txta(6).text = cnull(M_objrs!admin_fee)
    txta(7).text = cnull(M_objrs!payment_status)
    instalment_complite.Value = cnull(M_objrs!tgllunas)
    'instalment_complite.Value = IIf(IsNull(m_cust!tgllunas), "", Format(m_cust!tgllunas, "yyyy-mm-dd"))
    Label34.Caption = cnull(M_objrs("discpersen"))
    lblLastPay.Value = cnull(M_objrs("oustanding"))
    'dpdate.Value = cnull(M_objrs!delq_amt_by_x)

    Set M_objrs = Nothing
End Sub

Private Sub txta_Click(Index As Integer)
    
    
    If Index = 6 Then
        FrmCC_Colection.txtgetnomor.text = txta(6).text
        FrmCC_Colection.Txtperiod.Caption = txta(6).text
        FrmCC_Colection.CmbPhone.text = "CON MOBILE 2"
        Unload Me
    End If
    
End Sub
