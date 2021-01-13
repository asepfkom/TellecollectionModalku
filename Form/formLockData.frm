VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form formLockData 
   BackColor       =   &H80000015&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lock Data System"
   ClientHeight    =   8940
   ClientLeft      =   3060
   ClientTop       =   810
   ClientWidth     =   14085
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   14085
   Begin VB.Frame Frame8 
      BackColor       =   &H80000015&
      Height          =   3615
      Left            =   3405
      TabIndex        =   31
      Top             =   2160
      Visible         =   0   'False
      Width           =   9495
      Begin VB.CommandButton Command8 
         Caption         =   "Release"
         Height          =   495
         Left            =   6960
         TabIndex        =   34
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Exit"
         Height          =   495
         Left            =   8160
         TabIndex        =   33
         Top             =   2880
         Width           =   1095
      End
      Begin MSComctlLib.ListView lv4 
         Height          =   2460
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   9060
         _ExtentX        =   15981
         _ExtentY        =   4339
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
         Appearance      =   1
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
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000A&
      Height          =   8175
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   7815
      Begin VB.Frame Frame7 
         Height          =   5055
         Left            =   0
         TabIndex        =   26
         Top             =   3120
         Width           =   7695
         Begin VB.CommandButton Command6 
            Caption         =   "MnRelease"
            Height          =   735
            Left            =   120
            TabIndex        =   30
            Top             =   4200
            Width           =   975
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Run"
            Enabled         =   0   'False
            Height          =   735
            Left            =   6600
            TabIndex        =   29
            Top             =   4200
            Width           =   975
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Check sebelum Run"
            Height          =   735
            Left            =   5400
            TabIndex        =   28
            Top             =   4200
            Width           =   975
         End
         Begin MSComctlLib.ListView lv3 
            Height          =   3780
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   7380
            _ExtentX        =   13018
            _ExtentY        =   6668
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
            Appearance      =   1
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
         Height          =   735
         Left            =   0
         TabIndex        =   19
         Top             =   2160
         Width           =   7815
         Begin TDBDate6Ctl.TDBDate tgl1 
            Height          =   315
            Left            =   1680
            TabIndex        =   21
            Top             =   240
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   556
            Calendar        =   "formLockData.frx":0000
            Caption         =   "formLockData.frx":0118
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "formLockData.frx":0184
            Keys            =   "formLockData.frx":01A2
            Spin            =   "formLockData.frx":0200
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
            Left            =   4605
            TabIndex        =   22
            Top             =   240
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   556
            Calendar        =   "formLockData.frx":0228
            Caption         =   "formLockData.frx":0340
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "formLockData.frx":03AC
            Keys            =   "formLockData.frx":03CA
            Spin            =   "formLockData.frx":0428
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
            Height          =   375
            Left            =   3120
            TabIndex        =   24
            Top             =   240
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   661
            Caption         =   "formLockData.frx":0450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "formLockData.frx":04BC
            Spin            =   "formLockData.frx":050C
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
            Height          =   375
            Left            =   6000
            TabIndex        =   25
            Top             =   240
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   661
            Caption         =   "formLockData.frx":0534
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "formLockData.frx":05A0
            Spin            =   "formLockData.frx":05F0
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
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4095
            TabIndex        =   23
            Top             =   360
            Width           =   225
         End
         Begin VB.Label Label5 
            Caption         =   "Running Date"
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
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Enabled         =   0   'False
         Height          =   855
         Left            =   0
         TabIndex        =   15
         Top             =   1080
         Width           =   7815
         Begin VB.CheckBox Check1 
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   600
            TabIndex        =   18
            Top             =   240
            Width           =   1455
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Range Outstanding"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2400
            TabIndex        =   17
            Top             =   240
            Width           =   2775
         End
         Begin VB.CheckBox Check3 
            Caption         =   "CustID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6000
            TabIndex        =   16
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "formLockData.frx":0618
         Left            =   2160
         List            =   "formLockData.frx":061A
         TabIndex        =   14
         Top             =   300
         Width           =   4455
      End
      Begin VB.Line Line4 
         BorderWidth     =   3
         X1              =   0
         X2              =   7800
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   0
         X2              =   7800
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   0
         X2              =   7800
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label3 
         Caption         =   "Bank/Fintech"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "By Custid"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   8040
      TabIndex        =   6
      Top             =   6720
      Width           =   5775
      Begin VB.ComboBox cbosheet 
         Height          =   315
         Left            =   1590
         TabIndex        =   9
         Top             =   870
         Width           =   3165
      End
      Begin VB.CommandButton cmdbrowse 
         BackColor       =   &H00C0FFC0&
         Caption         =   "...."
         Height          =   315
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   465
         Width           =   555
      End
      Begin VB.TextBox txtlocation 
         Height          =   315
         Left            =   1590
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   3165
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label4 
         Caption         =   "Sheet"
         Height          =   255
         Left            =   390
         TabIndex        =   11
         Top             =   900
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "Location"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   795
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Range Outstanding"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8040
      TabIndex        =   2
      Top             =   5160
      Width           =   5775
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   5
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "To"
         Height          =   255
         Left            =   2400
         TabIndex        =   4
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Status"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   8040
      TabIndex        =   0
      Top             =   480
      Width           =   5775
      Begin MSComctlLib.ListView lv1 
         Height          =   4020
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   7091
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
         Appearance      =   1
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
      Begin VB.Line Line1 
         X1              =   0
         X2              =   0
         Y1              =   4920
         Y2              =   4935
      End
   End
End
Attribute VB_Name = "formLockData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public M_XLSCONN As New ADODB.Connection
Public qpub As String

Private Sub cbosheet_Change()
    If txtlocation.text <> "" Then
        If M_XLSCONN.State = adStateOpen Then M_XLSCONN.Close
        M_XLSCONN.Open "Provider = Microsoft.Jet.OleDb.4.0;data source = " & CommonDialog1.FileName & ";Extended Properties=Excel 8.0;"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        ssql = "SELECT * FROM [" & cbosheet.text & "] "
        M_objrs.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
        Set M_objrs = Nothing
    End If
End Sub

Private Sub Check1_Click()
    If Check1.Value = vbChecked Then
        Frame1.Enabled = True
    Else
        Frame1.Enabled = False
    End If
End Sub

Private Sub Check2_Click()
    If Check2.Value = vbChecked Then
        Frame2.Enabled = True
    Else
        Frame2.Enabled = False
    End If
End Sub

Private Sub Check3_Click()
    If Check3.Value = vbChecked Then
        Frame3.Enabled = True
    Else
        Frame3.Enabled = False
    End If
End Sub

Private Sub enabledbwah()
    If Combo1.text = "" Then
        Frame5.Enabled = False
        Frame6.Enabled = False
        Frame7.Enabled = False
    Else
        Frame5.Enabled = True
        Frame6.Enabled = True
        Frame7.Enabled = True
    End If
End Sub

Private Sub cmdbrowse_Click()
    With CommonDialog1
        .DialogTitle = "Import From File"
        .Filter = "Excel Files|*.xls"
        .ShowOpen
    End With
    txtlocation.text = ""
    If CommonDialog1.FileName = "" Then Exit Sub
    txtlocation.text = CommonDialog1.FileName
    If M_XLSCONN.State = adStateOpen Then M_XLSCONN.Close
    M_XLSCONN.Open "Provider = Microsoft.Jet.OleDb.4.0;data source = " & CommonDialog1.FileName & ";Extended Properties=Excel 8.0;"
    Set M_objrs = M_XLSCONN.OpenSchema(adSchemaTables)
    cbosheet.clear
    If M_objrs.EOF And M_objrs.BOF Then Exit Sub
    While Not M_objrs.EOF
        cbosheet.AddItem IIf(IsNull(M_objrs!table_name), "", M_objrs!table_name)
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    'Set M_XLSCONN = Nothing

End Sub

Private Sub Combo1_Change()
    Call enabledbwah
End Sub

Private Sub Combo1_Click()
    Call enabledbwah
End Sub

Private Sub Command1_Click()
    a = 1
    For i = 1 To lv1.ListItems.Count
        If lv1.ListItems(i).Checked = True Then
            lv2.ListItems(a).text = lv1.ListItems(i).text
            lv2.ListItems(a).SubItems(1) = lv1.ListItems(i).SubItems(1)
            lv2.ListItems(a).SubItems(2) = lv1.ListItems(i).SubItems(2)
            a = a + 1
        End If
    Next i
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Command4_Click()
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open "select * from tbl_lock_log_" & MDIForm1.TxtUsername.text & " where bank = '" & Combo1.text & "' and jam_awal < now() and jam_akhir > now();", M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If M_objrs.RecordCount > 0 Then
        MsgBox "Ada Lock yang sedang berlangsung, Process tidak bisa dilanjutkan."
        Exit Sub
    End If
    
    Call check
    Command5.Enabled = True
End Sub

Private Sub Command5_Click()
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open "select tanda hitung from tbl_lock_log_" & MDIForm1.TxtUsername.text & " order by id desc ", M_OBJCONN, adOpenDynamic, adLockOptimistic

    If M_objrs.RecordCount > 50 Then
        qbd = "insert into tbl_lock_log_" & MDIForm1.TxtUsername.text & "_log select * from tbl_lock_log_" & MDIForm1.TxtUsername.text & " where id not in (select id from tbl_lock_log_" & MDIForm1.TxtUsername.text & " order by 1 desc limit 50);" & vbCrLf
        qbd = qbd & "delete from tbl_lock_log_" & MDIForm1.TxtUsername.text & " where id not in (select id from tbl_lock_log_" & MDIForm1.TxtUsername.text & " order by 1 desc limit 50);"
        M_OBJCONN.Execute qbd
    End If

    If M_objrs.RecordCount = 0 Then
        tandax = 1
    Else
        tandax = M_objrs!Hitung + 1
    End If
    
    filterx = ""
    If Check1.Value = 1 Then
        filterx = filterx & "STATUS,"
    End If
    If Check2.Value = 1 Then
        filterx = filterx & "OUTSTANDING,"
    End If
    If Check3.Value = 1 Then
        filterx = filterx & "CUSTID"
    End If
    
    tgl1x = Format(tgl1.Value, "yyyy-mm-dd") & " " & jam1.Value
    tgl2x = Format(tgl2.Value, "yyyy-mm-dd") & " " & jam2.Value
    
    For i = 1 To lv3.ListItems.Count
        If lv3.ListItems(i).Checked = True Then
            qins = "insert into tbl_lock_log_" & MDIForm1.TxtUsername.text & " (tanda, agent,filter,jam_awal, jam_akhir,execute_by,bank) values "
            qins = qins & "(" & tandax & ", '" & lv3.ListItems(i).SubItems(1) & "' , '" & filterx & "', '" & tgl1x & "', '" & tgl2x & "', '" & MDIForm1.TxtUsername.text & "', '" & Combo1.text & "'); "
            M_OBJCONN.Execute qins
        End If
    Next i
    
    qins = "insert into tbl_lock_custid_" & MDIForm1.TxtUsername.text & " (custid)  select custid from ( " & qpub & " ) abc;"
    M_OBJCONN.Execute qins
    
    qupd = "update tbl_lock_custid_" & MDIForm1.TxtUsername.text & " set tanda = " & tandax & " where coalesce(tanda,0) = 0;"
    M_OBJCONN.Execute qupd
    
    MsgBox "Lock Sudah Di Process"

End Sub

Private Sub Command6_Click()
    Frame1.Enabled = False
    Frame2.Enabled = False
    Frame3.Enabled = False
    Frame4.Enabled = False
    Frame5.Enabled = False
    Frame6.Enabled = False
    Frame7.Enabled = False
    
    Frame8.Visible = True
    Call callrelease
End Sub

Private Sub callrelease()
    Set M_objrs = Nothing
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open "select * from tbl_lock_log_" & MDIForm1.TxtUsername.text & " where bank = '" & Combo1.text & "' and jam_awal < now() and jam_akhir > now();", M_OBJCONN, adOpenDynamic, adLockOptimistic
    'M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic

    lv4.ListItems.clear
    
    While Not M_objrs.EOF
        Set ListItem = lv4.ListItems.ADD(, , M_objrs("id"))
        ListItem.SubItems(1) = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
        ListItem.SubItems(2) = IIf(IsNull(M_objrs("filter")), "", M_objrs("filter"))
        ListItem.SubItems(3) = Format(IIf(IsNull(M_objrs("jam_awal")), "", M_objrs("jam_awal")), "dd-mm-yyyy hh:nn:ss")
        ListItem.SubItems(4) = Format(IIf(IsNull(M_objrs("jam_akhir")), "", M_objrs("jam_akhir")), "dd-mm-yyyy hh:nn:ss")
        ListItem.SubItems(5) = IIf(IsNull(M_objrs("bank")), "", M_objrs("bank"))
        M_objrs.MoveNext
    Wend
    
    If lv4.ListItems.Count = 0 Then
        MsgBox "Tidak ada Lock"
    End If
    
    Set M_objrs = Nothing
End Sub

Private Sub Command7_Click()
    If Combo1.text = "" Then
        Combo1.Enabled = True
    End If
    Frame4.Enabled = True
    Frame5.Enabled = True
    Frame6.Enabled = True
    Frame7.Enabled = True
    
    Frame8.Visible = False
End Sub

Private Sub Command8_Click()
    For i = 1 To lv4.ListItems.Count
        If lv4.ListItems(i).Checked = True Then
            qup = "update tbl_lock_log_" & MDIForm1.TxtUsername.text & " set jam_akhir = now() where id = " & lv4.ListItems(i).text & ""
            M_OBJCONN.Execute qup
        End If
    Next i
    
    MsgBox "Sudah Direlease"
    Command7_Click
End Sub

Private Sub Form_Activate()
    Call lv1st
    MsgBox "Pilih Bank/Fintech untuk membuka akses Status/Range Outstanding/Custid"
End Sub

Private Sub lv1st()
    Dim M_objrs As New ADODB.Recordset
    Dim CMDSQL As String
    Dim ListItem As ListItem
    
    lv1.ColumnHeaders.clear
    lv1.ColumnHeaders.ADD 1, , "No", 7 * 120
    lv1.ColumnHeaders.ADD 2, , "Cust ID", 10 * 120
    lv1.ColumnHeaders.ADD 3, , "Status", 20 * 120
    
    lv3.ColumnHeaders.clear
    lv3.ColumnHeaders.ADD 1, , "No", 10 * 120
    lv3.ColumnHeaders.ADD 2, , "AGENT", 10 * 120
    lv3.ColumnHeaders.ADD 3, , "DATA", 10 * 120
    
    lv4.ColumnHeaders.clear
    lv4.ColumnHeaders.ADD 1, , "ID", 2 * 140
    lv4.ColumnHeaders.ADD 2, , "AGENT", 10 * 120
    lv4.ColumnHeaders.ADD 3, , "FILTER", 20 * 120
    lv4.ColumnHeaders.ADD 4, , "JAM AWAL", 20 * 120
    lv4.ColumnHeaders.ADD 5, , "JAM AKHIR", 20 * 120
    lv4.ColumnHeaders.ADD 6, , "BANK", 20 * 120
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    CMDSQL = "select tblstatuscall_kdstscall,tblstatuscall_keterangan from tblstatuscall  where tblstatuscall_kdstatus = '1' order by 1"
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    lv1.ListItems.clear
    
    While Not M_objrs.EOF
        Set ListItem = lv1.ListItems.ADD(, , M_objrs.Bookmark)
        ListItem.SubItems(1) = IIf(IsNull(M_objrs("tblstatuscall_kdstscall")), "", M_objrs("tblstatuscall_kdstscall"))
        ListItem.SubItems(2) = IIf(IsNull(M_objrs("tblstatuscall_keterangan")), "", M_objrs("tblstatuscall_keterangan"))
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
End Sub


Private Sub getbankspv()
    If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Then
        q = "select client from tbl_list_client_indium  order by 1"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        Combo1.clear
        For i = 1 To M_objrs.RecordCount
            q1 = "select distinct recsource from mgm where agent in (select userid from usertbl where team = '" & MDIForm1.TxtUsername.text & "' or userid = '" & MDIForm1.TxtUsername.text & "') and recsource ilike '%" & cnull(M_objrs!client) & "%'   "
            Set M_OBJRS1 = New ADODB.Recordset
            M_OBJRS1.CursorLocation = adUseClient
            M_OBJRS1.Open q1, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            If M_OBJRS1.RecordCount > 0 Then
                Combo1.AddItem cnull(M_objrs!client)
            End If
            M_objrs.MoveNext
        Next i
    Else
        Call list_client(Combo1)
    End If
End Sub

Private Sub Form_Load()
    Call getbankspv
    Call createtbl
End Sub

Private Sub createtbl()
    qs = "select * from information_schema.columns where table_name = 'tbl_lock_log_" & LCase(MDIForm1.TxtUsername.text) & "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

    If M_objrs.RecordCount = 0 Then
        qc = "create table tbl_lock_log_" & MDIForm1.TxtUsername.text & " (id serial, tanda integer, agent varchar, filter varchar, jam_awal timestamp without time zone, jam_akhir timestamp without time zone, execute_by varchar, tanggal timestamp without time zone default now(), bank varchar, stsopenagent smallint);"
        M_OBJCONN.Execute qc
        
        qc = "create table tbl_lock_log_" & MDIForm1.TxtUsername.text & "_log (id serial, tanda integer, agent varchar, filter varchar, jam_awal timestamp without time zone, jam_akhir timestamp without time zone, execute_by varchar, tanggal timestamp without time zone default now(), bank varchar, stsopenagent smallint);"
        M_OBJCONN.Execute qc
        
        qc = "create table tbl_lock_custid_" & MDIForm1.TxtUsername.text & " (tanda integer, custid varchar, tanggal timestamp without time zone default now());"
        M_OBJCONN.Execute qc
        
        qc = "create table tbl_lock_custid_" & MDIForm1.TxtUsername.text & "_log (tanda integer, custid varchar, tanggal timestamp without time zone default now());"
        M_OBJCONN.Execute qc
    End If
End Sub


Private Sub check()
    'status
    Dim stscall, campaign, custidx, hit1, hit2 As String
    
    stscall = ""
    campaign = ""
    custidx = ""
    
    If Check1.Value = 1 Then
        For i = 1 To lv1.ListItems.Count
            If lv1.ListItems(i).Checked = True Then
                stscall = stscall & "'" & lv1.ListItems(i).SubItems(2) & "',"
            End If
        Next i
        stscall = Left(stscall, Len(stscall) - 1)
    End If
    
    If Check2.Value = 1 Then
        hit1 = Replace(Text1.text, ",", "")
        hit2 = Replace(Text2.text, ",", "")
    End If
    
    If Check3.Value = 1 Then
        ssql = "SELECT * FROM [" & cbosheet.text & "]   "
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
            
        Set rsTemporary = New ADODB.Recordset
        rsTemporary.CursorLocation = adUseClient
        rsTemporary.CursorType = adOpenDynamic
        rsTemporary.ActiveConnection = M_OBJCONN
        rsTemporary.LockType = adLockOptimistic
            
        rs.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
    
        While Not rs.EOF
            custidx = custidx & "'" & rs(0) & "',"
            rs.MoveNext
        Wend
            custidx = Left(custidx, Len(custidx) - 1)
    End If
    
    If Combo1.text Like "*UANG*" Then
        campaign = "UANG"
    ElseIf Combo1.text Like "*PLUS*" Then
        campaign = "PLUS"
    Else
        campaign = Combo1.text
    End If
    
    qpub = "select custid from mgm where recsource ilike '%" & campaign & "%' and coalesce(agent,'') <> '' "
    
    q = "select agent, count(custid) from mgm where recsource ilike '%" & campaign & "%' and coalesce(agent,'') <> '' "
    
    If stscall <> "" Then
        q = q + " and statuscall in (" & stscall & ")"
        qpub = qpub + " and statuscall in (" & stscall & ")"
    End If
    
    If hit1 <> "" And hit2 <> "" Then
        q = q + " and (curbal >= " & hit1 & " and curbal <= " & hit2 & ")"
        qpub = qpub + " and (curbal >= " & hit1 & " and curbal <= " & hit2 & ")"
    End If
    
    If custidx <> "" Then
        q = q + " and custid in (" & custidx & ")"
        qpub = qpub + " and custid in (" & custidx & ")"
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

Private Sub Text1_LostFocus()
    Text1.text = Format(Text1.text, "##,###")
End Sub

Private Sub Text2_LostFocus()
     Text2.text = Format(Text2.text, "##,###")
End Sub
