VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frm_list_autodialer_break 
   BorderStyle     =   0  'None
   Caption         =   "Reason"
   ClientHeight    =   1515
   ClientLeft      =   8070
   ClientTop       =   5505
   ClientWidth     =   4035
   LinkTopic       =   "Form5"
   ScaleHeight     =   1515
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Height          =   1500
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4005
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   165
         TabIndex        =   5
         Top             =   360
         Width           =   3630
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Height          =   420
         ItemData        =   "form_autodial_off.frx":0000
         Left            =   1950
         List            =   "form_autodial_off.frx":0013
         TabIndex        =   4
         Top             =   735
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Timer Timer_autdialer_break 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1470
         Top             =   735
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Left            =   2445
         TabIndex        =   2
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
         _Version        =   196610
         ForeColor       =   16777215
         BackColor       =   32768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Resume Calling"
         ButtonStyle     =   4
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "Break Reason"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   165
         TabIndex        =   6
         Top             =   105
         Width           =   1605
      End
      Begin VB.Label LblCount_durasi 
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1035
         TabIndex        =   3
         Top             =   750
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "Timer : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   1
         Top             =   750
         Width           =   735
      End
   End
End
Attribute VB_Name = "frm_list_autodialer_break"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StatusAwal As String

Private Sub Combo1_Click()
    
    'On Error GoTo cek_break

    'If AutoDialerON = False Then
    If F_AutoDial = False Then
        Timer_autdialer_break.Enabled = True
        LblCount_durasi.Caption = "0"
        cmdsqlphone = " insert into tbl_autodialer_agent_break(agent,status_break,waktu_start,sessionid,ip_login) values"
        cmdsqlphone = cmdsqlphone + "('" + MDIForm1.txtusername.text + "','" + Combo1.text + "',now(), '" & MDIForm1.txtusername.text & "_break_" & Format(FungsiWaktuServer, "YYYY-MM-DD_HH") & "','" & MDIForm1.Winsock1.LocalIP & "')"
        M_OBJCONN.Execute cmdsqlphone
        Combo1.Enabled = False
    Else
        Timer_autdialer_break.Enabled = True
        MDIForm1.TimerAutoDial.Enabled = False
        Combo1.Enabled = False
    End If
    LblCount_durasi.Caption = "0"
'cek_break:
'    MsgBox Err.Description
End Sub

Private Sub Combo1_DropDown()
    Combo1.clear
    Combo1.AddItem "Lunch"
    Combo1.AddItem "Meeting"
    Combo1.AddItem "Pray"
    Combo1.AddItem "Toilet"
    Combo1.AddItem "Coaching"
End Sub

Private Sub Form_Load()
    Timer_autdialer_break.Enabled = True
    MDIForm1.TimerAutoDial = False
    StatusAwal = ""
End Sub

Private Sub List1_Click()
Dim cmdsqlBreak As String

Select Case List1.text
    Case "Meeting", "Pray", "Lunch", "Toilet", "Coaching"
        If StatusAwal = "" Then
        Else
            StatusAwal = List1.text
        End If
        Timer_autdialer_break.Enabled = True
    End Select
   LblCount_durasi.Caption = "0"
End Sub

Private Sub SSCommand1_Click()
    'If AutoDialerON = True Then
    If F_AutoDial = True Then
        MDIForm1.TimerAutoDial.Enabled = True
        'AutoDialerBreak = False
        Timer_autdialer_break.Enabled = False
        MDIForm1.TimerAutoDial.Enabled = False
        Combo1.Enabled = True
        Unload Me
    Else
'        cmdsqlphone = " insert into tbl_autodialer_agent_break(agent,status_break,durasi) values"
'        cmdsqlphone = cmdsqlphone + "('" + agent + "','" + reason_stop + "','" + CStr(durasi) + "')"
        cmdsqlphone = "update tbl_autodialer_agent_break set durasi = '" & LblCount_durasi.Caption & "', waktu_end = now() where id in "
        cmdsqlphone = cmdsqlphone + "(select max(id) from tbl_autodialer_agent_break where agent = '" & MDIForm1.txtusername.text & "' and status_break not in ('ManualDial','start_autodialer','AutoDial','form break show'))"
        M_OBJCONN.Execute cmdsqlphone
        Timer_autdialer_break.Enabled = False
        Combo1.Enabled = True
        MDIForm1.TimerAutoDial = True
        Unload Me
    End If
    break_time = False
End Sub

Private Sub Timer_autdialer_break_Timer()
    If LblCount_durasi.Caption <> "" Then
        'LblCount_durasi.Caption = ""
        LblCount_durasi.Caption = Val(LblCount_durasi.Caption) + 1
        If LblCount_durasi.Caption = "5" And Combo1.text = "" Then
            'Timer_autdialer_break.Enabled = False
            MsgBox "Harap Pilih Alasan Break", vbCritical + vbOKOnly
            Exit Sub
        Else
            'Timer_autdialer_break.Enabled = True
            LblCount_durasi.Caption = Val(LblCount_durasi.Caption) + 1
        End If
    End If

End Sub
Public Function FungsiWaktuServer()
 'Fungsi Untuk mengambil waktu dan tanggal di server database
 Dim CMDSQL As String
 Dim M_objrs As ADODB.Recordset
 
 CMDSQL = "select now() as waktu"
 
 Set M_objrs = New ADODB.Recordset
 M_objrs.CursorLocation = adUseClient
 
 M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
 
 WaktuServer = Format(M_objrs(0), "yyyy-mm-dd hh:mm:00")
 FungsiWaktuServer = WaktuServer
 Set M_objrs = Nothing
End Function




