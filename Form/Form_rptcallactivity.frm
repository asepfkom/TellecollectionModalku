VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form_rptcallactivity 
   Caption         =   "Report Call Activity"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15330
   LinkTopic       =   "Form2"
   ScaleHeight     =   8550
   ScaleWidth      =   15330
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Criteria Report"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8460
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      Begin VB.TextBox txtlead 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   14175
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   3645
         Width           =   915
      End
      Begin VB.CommandButton SSCommand2 
         BackColor       =   &H00F1E5DB&
         Cancel          =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6525
         Picture         =   "Form_rptcallactivity.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2340
         Width           =   1620
      End
      Begin VB.CommandButton SSCommand1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Export to Excel"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   6525
         Picture         =   "Form_rptcallactivity.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1230
         Width           =   1620
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
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
         Height          =   3255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   6330
         Begin VB.CheckBox CheckAll_Agent 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check All"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   4440
            Width           =   1455
         End
         Begin VB.TextBox txtJmlAgent 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7680
            Locked          =   -1  'True
            TabIndex        =   7
            Text            =   "0"
            Top             =   4440
            Width           =   975
         End
         Begin VB.CommandButton Cmd_Refersh3 
            BackColor       =   &H00F1E5DB&
            Caption         =   "REFRESH"
            Height          =   255
            Left            =   3960
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   4440
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox CheckAll_MGR 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check All"
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
            Left            =   120
            TabIndex        =   5
            Top             =   2910
            Width           =   1455
         End
         Begin VB.TextBox txt_jumlah_acc 
            Enabled         =   0   'False
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
            Left            =   5280
            TabIndex        =   4
            Top             =   2910
            Width           =   975
         End
         Begin MSComctlLib.ListView LVAgent 
            Height          =   2580
            Left            =   90
            TabIndex        =   9
            Top             =   270
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   4551
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
         Begin MSComctlLib.ProgressBar ProgressBar3 
            Height          =   255
            Left            =   2040
            TabIndex        =   10
            Top             =   4440
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
         End
         Begin VB.Label Label30 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Jumlah User :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   5640
            TabIndex        =   11
            Top             =   4440
            Width           =   2055
         End
      End
      Begin VB.CommandButton cmdCari 
         Height          =   375
         Left            =   6525
         Picture         =   "Form_rptcallactivity.frx":0DAC
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   795
         Width           =   1605
      End
      Begin VB.TextBox TxtPath 
         Enabled         =   0   'False
         Height          =   285
         Left            =   0
         TabIndex        =   1
         Top             =   -255
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4305
         Left            =   120
         TabIndex        =   14
         Top             =   3990
         Width           =   14970
         _ExtentX        =   26405
         _ExtentY        =   7594
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSComDlg.CommonDialog Cd_save 
         Left            =   6495
         Top             =   165
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "*.xlsx"
      End
      Begin TDBDate6Ctl.TDBDate tgl_call 
         Height          =   315
         Index           =   0
         Left            =   1290
         TabIndex        =   15
         Top             =   255
         Width           =   1365
         _Version        =   65536
         _ExtentX        =   2408
         _ExtentY        =   556
         Calendar        =   "Form_rptcallactivity.frx":139A
         Caption         =   "Form_rptcallactivity.frx":14B2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Form_rptcallactivity.frx":151E
         Keys            =   "Form_rptcallactivity.frx":153C
         Spin            =   "Form_rptcallactivity.frx":159A
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
      Begin TDBDate6Ctl.TDBDate tgl_call 
         Height          =   315
         Index           =   1
         Left            =   3255
         TabIndex        =   16
         Top             =   255
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   556
         Calendar        =   "Form_rptcallactivity.frx":15C2
         Caption         =   "Form_rptcallactivity.frx":16DA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Form_rptcallactivity.frx":1746
         Keys            =   "Form_rptcallactivity.frx":1764
         Spin            =   "Form_rptcallactivity.frx":17C2
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
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Jml Lead"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   13245
         TabIndex        =   20
         Top             =   3690
         Width           =   915
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
         Height          =   315
         Left            =   2745
         TabIndex        =   18
         Top             =   285
         Width           =   825
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tanggal"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   165
         TabIndex        =   17
         Top             =   285
         Width           =   1245
      End
   End
End
Attribute VB_Name = "Form_rptcallactivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckAll_MGR_Click()
    If CheckAll_MGR.Value = 1 Then
        If LVAgent.ListItems.Count <> 0 Then
            For i = 1 To LVAgent.ListItems.Count
                LVAgent.ListItems(i).Checked = True
            Next i
        End If
    ElseIf CheckAll_MGR.Value = 0 Then
        If LVAgent.ListItems.Count <> 0 Then
            For i = 1 To LVAgent.ListItems.Count
                LVAgent.ListItems(i).Checked = False
            Next i
        End If
    End If
End Sub

Private Sub CmdCari_Click()
Dim STRSQL As String
Dim MOBJ As ADODB.Recordset

    If (tgl_call(0).ValueIsNull) Or (tgl_call(1).ValueIsNull) Then
        MsgBox "Tanggal Call Harus Diisi", vbInformation, "Informasi"
        Exit Sub
    End If
    
    If LVAgent.ListItems.Count = 0 Then
        MsgBox "Agent Tidak Tersedia", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For K = 1 To LVAgent.ListItems.Count
        If LVAgent.ListItems(K).Checked = True Then
            cek = cek + 1
        End If
    Next K
    
    If cek = 0 Then
        MsgBox "Pilih Agent Terlebih Dahulu!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    For i = 1 To txt_jumlah_acc.text
        If LVAgent.ListItems(i).Checked = True Then
            ListCustId = ListCustId & "'" & LVAgent.ListItems(i).SubItems(1) & "',"
        End If
    Next i
    
    ListCustId = Mid(ListCustId, 1, Len(ListCustId) - 1)
    STRSQL = "select x.*, round(x.""PTP""::numeric / x.""Call Initiated""::numeric,2) as ""Persentage PTP"" from("
    STRSQL = STRSQL + " select a.agent ,d.jumlah_data as ""Jumlah Data"",coalesce(b.jml,0) as ""New Data"", c.callattempt as ""Call Atempt"" ,a.total as ""Call Initiated"""
    STRSQL = STRSQL + " ,""BPH"",""CALL BACK"",""PAID"",""MSG"",""NO PTP"",""PAID"",""PTD"",""PTP"",""RTP"",""UTC1"",""UTC2"",c.durasi as ""Talktime"" from("
    STRSQL = STRSQL + "select agent ,sum(""BPH"") as ""BPH"" ,sum(""CALL BACK"") as ""CALL BACK"" ,sum(""MSG"") as ""MSG"" ,sum(""NO PTP"") as ""NO PTP"" ,sum(""PAID"") as ""PAID"""
    STRSQL = STRSQL + " ,sum(""PTD"") as ""PTD"" ,sum(""PTP"") as ""PTP"" ,sum(""RTP"") as ""RTP"" ,sum(""UTC1"") as ""UTC1"" ,sum(""UTC2"") as ""UTC2"", sum(total) as total from ("
    STRSQL = STRSQL + " select *,""BPH""+""CALL BACK""+""MSG""+""NO PTP""+""PAID""+""PTD""+""PTP""+""RTP""+""UTC1""+""UTC2"" as Total from ("
    STRSQL = STRSQL + " select agent ,"
    STRSQL = STRSQL + " case when kodeds = 'BPH' then 1 else 0 end as ""BPH"" ,"
    STRSQL = STRSQL + " case when kodeds = 'CALL BACK' then 1 else 0 end as ""CALL BACK"" ,"
    STRSQL = STRSQL + " case when kodeds = 'MSG' then 1 else 0 end as ""MSG"" ,"
    STRSQL = STRSQL + " case when kodeds = 'NO PTP' then 1 else 0 end as ""NO PTP"" ,"
    STRSQL = STRSQL + " case when kodeds = 'PAID' then 1 else 0 end as ""PAID"" ,"
    STRSQL = STRSQL + " case when kodeds = 'PTD' then 1 else 0 end as ""PTD"" ,"
    STRSQL = STRSQL + " case when kodeds = 'PTP' then 1 else 0 end as ""PTP"" ,"
    STRSQL = STRSQL + " case when kodeds = 'RTP' then 1 else 0 end as ""RTP"" ,"
    STRSQL = STRSQL + " case when kodeds = 'UTC1' then 1 else 0 end as ""UTC1"" ,"
    STRSQL = STRSQL + " case when kodeds = 'UTC2' then 1 else 0 end as ""UTC2"""
    STRSQL = STRSQL + " from ("
    STRSQL = STRSQL + " select agent, custid, kodeds from mgm_hst where date(tgl) between '" + Format(tgl_call(0).Value, "yyyy-mm-dd") + "' and '" + Format(tgl_call(1).Value, "yyyy-mm-dd") + "' and agent in (" & ListCustId & ") ) hst  ) abc  ) a group by agent  )a"
    STRSQL = STRSQL + " left join ("
    STRSQL = STRSQL + " select agent, count(f_cek_new) as jml from mgm where coalesce(statuscall,'')= '' group by agent) b on a.agent=b.agent"
    STRSQL = STRSQL + " Left Join ("
    STRSQL = STRSQL + " select agent, count(agent)as callattempt, sum(durasi_billsec) as durasi  from mgm_hst where date(tgl) between '" + Format(tgl_call(0).Value, "yyyy-mm-dd") + "' and '" + Format(tgl_call(1).Value, "yyyy-mm-dd") + "' and agent in (" & ListCustId & ") group by agent)c on a.agent=c.agent"
    STRSQL = STRSQL + " Left Join ("
    STRSQL = STRSQL + " select agent,count(id) as jumlah_data from mgm group by agent)d on a.agent =d.agent)x"
    
    Set MOBJ = New ADODB.Recordset
    MOBJ.CursorLocation = adUseClient
    MOBJ.Open STRSQL, M_OBJCONN, adOpenKeyset, adLockOptimistic
    txtlead.text = MOBJ.RecordCount
    Set DataGrid1.DATASOURCE = MOBJ
   
End Sub

Private Sub Form_Load()
    
    CheckAll_MGR.Value = 1
        
    Call HeaderLvAgent
    Call ISIAGENT
    Call CheckAll_MGR_Click

End Sub

Private Sub ISIAGENT()
    Dim sQuery As String
    Dim Rs_Agent As ADODB.Recordset
    Dim Nomor As Double
    Dim list As ListItem
    
     If MDIForm1.txtlevel.text = "Supervisor" Or MDIForm1.txtlevel.text = "TeamLeader" Then
        sQuery = "SELECT * FROM usertbl WHERE aktif = '1' AND kdlevel='1' and spvcode = '" + MDIForm1.TxtUsername + "' order by agent "
    Else
        sQuery = "SELECT * FROM usertbl WHERE aktif = '1' AND kdlevel='1' order by agent "
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
            Rs_Agent.MoveNext
        Wend
    End If
    
    txt_jumlah_acc = Rs_Agent.RecordCount
End Sub

Private Sub HeaderLvAgent()
    LVAgent.ColumnHeaders.ADD 1, , "No", 600
    LVAgent.ColumnHeaders.ADD 2, , "AGENT", 5000
End Sub

Private Sub SSCommand1_Click()
    export_data
End Sub
Private Sub export_data()
    If (tgl_call(0).ValueIsNull) And (tgl_call(1).ValueIsNull) Then
        MsgBox "Tanggal Call Harus Diisi", vbInformation, "Informasi"
        Exit Sub
    End If
    
    If LVAgent.ListItems.Count = 0 Then
        MsgBox "Agent Tidak Tersedia", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For K = 1 To LVAgent.ListItems.Count
        If LVAgent.ListItems(K).Checked = True Then
            cek = cek + 1
        End If
    Next K
    
    If cek = 0 Then
        MsgBox "Pilih Agent Terlebih Dahulu!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    For i = 1 To txt_jumlah_acc.text
        If LVAgent.ListItems(i).Checked = True Then
            ListCustId = ListCustId & "'" & LVAgent.ListItems(i).SubItems(1) & "',"
        End If
    Next i
    
    ListCustId = Mid(ListCustId, 1, Len(ListCustId) - 1)
    STRSQL = "select x.*, round(x.""PTP""::numeric / x.""Call Initiated""::numeric,2) as ""Persentage PTP"" from("
    STRSQL = STRSQL + " select a.agent ,d.jumlah_data as ""Jumlah Data"",coalesce(b.jml,0) as ""New Data"", c.callattempt as ""Call Atempt"" ,a.total as ""Call Initiated"""
    STRSQL = STRSQL + " ,""BPH"",""CALL BACK"",""PAID"",""MSG"",""NO PTP"",""PAID"",""PTD"",""PTP"",""RTP"",""UTC1"",""UTC2"",c.durasi as ""Talktime"" from("
    STRSQL = STRSQL + "select agent ,sum(""BPH"") as ""BPH"" ,sum(""CALL BACK"") as ""CALL BACK"" ,sum(""MSG"") as ""MSG"" ,sum(""NO PTP"") as ""NO PTP"" ,sum(""PAID"") as ""PAID"""
    STRSQL = STRSQL + " ,sum(""PTD"") as ""PTD"" ,sum(""PTP"") as ""PTP"" ,sum(""RTP"") as ""RTP"" ,sum(""UTC1"") as ""UTC1"" ,sum(""UTC2"") as ""UTC2"", sum(total) as total from ("
    STRSQL = STRSQL + " select *,""BPH""+""CALL BACK""+""MSG""+""NO PTP""+""PAID""+""PTD""+""PTP""+""RTP""+""UTC1""+""UTC2"" as Total from ("
    STRSQL = STRSQL + " select agent ,"
    STRSQL = STRSQL + " case when kodeds = 'BPH' then 1 else 0 end as ""BPH"" ,"
    STRSQL = STRSQL + " case when kodeds = 'CALL BACK' then 1 else 0 end as ""CALL BACK"" ,"
    STRSQL = STRSQL + " case when kodeds = 'MSG' then 1 else 0 end as ""MSG"" ,"
    STRSQL = STRSQL + " case when kodeds = 'NO PTP' then 1 else 0 end as ""NO PTP"" ,"
    STRSQL = STRSQL + " case when kodeds = 'PAID' then 1 else 0 end as ""PAID"" ,"
    STRSQL = STRSQL + " case when kodeds = 'PTD' then 1 else 0 end as ""PTD"" ,"
    STRSQL = STRSQL + " case when kodeds = 'PTP' then 1 else 0 end as ""PTP"" ,"
    STRSQL = STRSQL + " case when kodeds = 'RTP' then 1 else 0 end as ""RTP"" ,"
    STRSQL = STRSQL + " case when kodeds = 'UTC1' then 1 else 0 end as ""UTC1"" ,"
    STRSQL = STRSQL + " case when kodeds = 'UTC2' then 1 else 0 end as ""UTC2"""
    STRSQL = STRSQL + " from ("
    STRSQL = STRSQL + " select agent, custid, kodeds from mgm_hst where date(tgl) between '" + Format(tgl_call(0).Value, "yyyy-mm-dd") + "' and '" + Format(tgl_call(1).Value, "yyyy-mm-dd") + "' and agent in (" & ListCustId & ") ) hst  ) abc  ) a group by agent  )a"
    STRSQL = STRSQL + " left join ("
    STRSQL = STRSQL + " select agent, count(f_cek_new) as jml from mgm where coalesce(statuscall,'')= '' group by agent) b on a.agent=b.agent"
    STRSQL = STRSQL + " Left Join ("
    STRSQL = STRSQL + " select agent, count(agent)as callattempt, sum(durasi_billsec) as durasi  from mgm_hst where date(tgl) between '" + Format(tgl_call(0).Value, "yyyy-mm-dd") + "' and '" + Format(tgl_call(1).Value, "yyyy-mm-dd") + "' and agent in (" & ListCustId & ") group by agent)c on a.agent=c.agent"
    STRSQL = STRSQL + " Left Join ("
    STRSQL = STRSQL + " select agent,count(id) as jumlah_data from mgm group by agent)d on a.agent =d.agent)x"
    isi_data (STRSQL)
    
End Sub
Private Sub SSCommand2_Click()
    Unload Me
End Sub

Private Sub isi_data(STRSQL As String)
On Error GoTo Salah
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim ListItem As ListItem
    Dim cmdsql_update As String
    Dim objExcel        As Excel.Application
    Dim objBook         As Excel.Workbook
    Dim objSheet        As Excel.Worksheet
    Dim i As Integer
    Dim m_msgbox As String
    
    i = 1
    

    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
  
   
Form_Save:
    Cd_save.ShowSave
    TxtPath.text = Cd_save.FileName
    
    'Cek apakah user menekan tombol cancel pada dialog save
    If TxtPath.text = Empty Then
        'Tanyakan ke user.. apakah benar2 akan membatalkan proses download???
        m_msgbox = MsgBox("Anda ingin Download dibatalkan?", vbYesNo + vbQuestion, "Konfirmasi")
        'Jika user benar-benar akan membatalkan proses download, keluar dari fungsi ini!
        If m_msgbox = vbYes Then
              MsgBox "Download dibatalkan!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If
        If m_msgbox = vbNo Then '-> jika user tidak membatalkan proses download
          GoTo Form_Save        '-> maka goto form_save
        End If
    End If

 Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.ADD
    Set objSheet = objBook.ActiveSheet
        
    
'    lblstatus.Caption = "Status download: Mengisi field... silahkan tunggu!"
    
    
    On Error GoTo Salah
    'Proses pengsisian nama field ke excel
    Dim x, Y    As Integer
        If M_objrs.State = 1 Then
            x = 0
            Y = M_objrs.fields().Count - 1
            Do Until x > Y
                DoEvents
                objSheet.Cells(1, i).Value = CStr(M_objrs.fields(x).Name)
                i = i + 1
                x = x + 1
            Loop
        End If
    
   ' lblstatus.Caption = "Status download: Membuat file excel... silahkan tunggu!"
    objSheet.Range("A2").CopyFromRecordset M_objrs '-> Proses pengisian data dimulai dari Cell A2
    objBook.SaveAs TxtPath.text, xlWorkbookNormal
    objExcel.Quit
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    Set M_objrs = Nothing
 
Salah:
    Exit Sub
End Sub





