VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_pindah_agent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SUBTITUTION AGENT"
   ClientHeight    =   7020
   ClientLeft      =   5220
   ClientTop       =   1890
   ClientWidth     =   10020
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   10020
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5415
      Left            =   0
      TabIndex        =   7
      Top             =   2040
      Width           =   10095
      Begin VB.CommandButton Command3 
         BackColor       =   &H000000FF&
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton cmdclear 
         BackColor       =   &H00FF8080&
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "EXTRACT CAMPAIGN BARU"
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
         Left            =   2880
         TabIndex        =   21
         Top             =   1680
         Width           =   3855
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "-------"
         Height          =   375
         Left            =   7080
         TabIndex        =   19
         Top             =   1200
         Width           =   2655
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "-------"
         Height          =   375
         Left            =   4440
         TabIndex        =   18
         Top             =   1200
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "-------"
         Height          =   375
         Left            =   1800
         TabIndex        =   17
         Top             =   1200
         Width           =   2655
      End
      Begin VB.ComboBox Combo3 
         Appearance      =   0  'Flat
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
         Left            =   1800
         TabIndex        =   10
         Top             =   600
         Width           =   5415
      End
      Begin MSComctlLib.ListView LvPTP 
         Height          =   1860
         Left            =   480
         TabIndex        =   20
         Top             =   2280
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   3281
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
      Begin VB.Label Label9 
         Caption         =   "PEGANGAN"
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
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label7 
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
         Left            =   7200
         TabIndex        =   11
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "Team Leader"
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
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "DIPINDAHKAN KE :"
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
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "-------"
         Height          =   375
         Left            =   7080
         TabIndex        =   26
         Top             =   1320
         Width           =   2655
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "-------"
         Height          =   375
         Left            =   4440
         TabIndex        =   25
         Top             =   1320
         Width           =   2655
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "-------"
         Height          =   375
         Left            =   1800
         TabIndex        =   24
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1320
         Width           =   2655
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
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
         Left            =   1800
         TabIndex        =   4
         Top             =   720
         Width           =   5415
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
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
         Left            =   1800
         TabIndex        =   3
         Top             =   120
         Width           =   5415
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000A&
         Caption         =   "PEGANGAN"
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
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000A&
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
         Left            =   7200
         TabIndex        =   6
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000A&
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
         Left            =   7200
         TabIndex        =   5
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000A&
         Caption         =   "Team Leader"
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
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000A&
         Caption         =   "Agent"
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
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form_pindah_agent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub leader(zzz As ComboBox)
    zzz.clear
    sStrsql = "select agent, userid from usertbl where aktif  = 1 and usertype = 2"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If M_objrs.RecordCount > 0 Then
        For i = 1 To M_objrs.RecordCount
            zzz.AddItem cnull(M_objrs!AGENT)
            M_objrs.MoveNext
        Next i
    End If
End Sub

Private Sub create_table_log()
    querysel = "select * from information_schema.columns  where table_name = 'tbl_changed_team_log'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open querysel, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If M_objrs.RecordCount = 0 Then
        qcreate = "create table tbl_changed_team_log (userid varchar,leader_ex varchar,leader_now varchar,campaign_ex varchar,campaign_now varchar,user_update varchar,tgl_update timestamp default now(),id serial not null);"
        M_OBJCONN.Execute qcreate
    End If
End Sub

Private Sub CmdClear_Click()
    Combo1.text = ""
    Combo2.text = ""
    Label4.Caption = ""
    Label3.Caption = ""
    Text1.text = ""
    Text2.text = ""
    Text3.text = ""
    Combo3.text = ""
    Label7.Caption = ""
    Option1.Caption = "-------"
    Option2.Caption = "-------"
    Option3.Caption = "-------"
    Option4.Caption = "-------"
    Option5.Caption = "-------"
    Option6.Caption = "-------"
    Option1.Value = False
    Option2.Value = False
    Option3.Value = False
    Option4.Value = False
    Option5.Value = False
    Option6.Value = False
    LvPTP.ListItems.clear
End Sub

Private Sub Combo1_Click()
    'call_leader
    sStrsql = "select agent, userid from usertbl where aktif  = 1 and usertype = 2 and agent = '" & Combo1.text & "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If M_objrs.RecordCount > 0 Then
        Label4.Caption = cnull(M_objrs!USERID)
    End If

    'call_agent
    Combo2.clear
    sStrsql = "select agent, userid from usertbl where aktif  = 1 and usertype = 1 and team = '" & Label4.Caption & "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If M_objrs.RecordCount > 0 Then
        For i = 1 To M_objrs.RecordCount
            Combo2.AddItem cnull(M_objrs!AGENT)
            M_objrs.MoveNext
        Next i
    End If
End Sub

Private Sub Combo1_DropDown()
    Call leader(Combo1)
End Sub

Private Sub Combo2_Click()
    'CmdClear_Click
    Text1.text = ""
    Text2.text = ""
    Text3.text = ""
    Option4.Caption = "-------"
    Option5.Caption = "-------"
    Option6.Caption = "-------"
    
    sStrsql = "select agent, userid from usertbl where aktif  = 1 and usertype = 1 and agent = '" & Combo2.text & "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If M_objrs.RecordCount > 0 Then
        Label3.Caption = cnull(M_objrs!USERID)
        get_rekan (1)
    End If
End Sub

Private Sub Combo3_Click()
    sStrsql = "select agent, userid from usertbl where aktif  = 1 and usertype = 2 and agent = '" & Combo3.text & "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If M_objrs.RecordCount > 0 Then
        Label7.Caption = cnull(M_objrs!USERID)
        Call get_rekan(2)
    End If
End Sub

Private Sub Combo3_DropDown()
    Call leader(Combo3)
End Sub

Private Sub get_rekan(l As Integer)
    If l = 1 Then
        q = "select distinct recsource from mgm where agent = '" & Label3.Caption & "'"
    ElseIf l = 2 Then
        q = "select distinct recsource from mgm where agent in (select userid from usertbl where team = '" & Label7.Caption & "')"
    End If
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    Dim zz As String
    Dim sss As String
    
    qs = "select * from tbl_list_client_indium order by 1"
    Set M_objrsc = New ADODB.Recordset
    M_objrsc.CursorLocation = adUseClient
    M_objrsc.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    sss = ""
    zz = ""
    While Not M_objrs.EOF
            For i = 1 To M_objrsc.RecordCount
                a = M_objrsc!client
                If M_objrs!recsource Like "*" & a & "*" Then
                    If M_objrsc!client = "PLUS" Then
                        If sss Like "*PLUS*" Then
                        Else
                            If Option1.Caption = "-------" Then
                                If l = 1 Then
                                    Text1.text = "RUPIAH PLUS"
                                    Option4.Caption = "RUPIAH PLUS"
                                Else
                                    Option1.Caption = "RUPIAH PLUS"
                                End If
                                sss = sss & " PLUS "
                            ElseIf Option2.Caption = "-------" Then
                                If l = 1 Then
                                    Text2.text = "RUPIAH PLUS"
                                    Option5.Caption = "RUPIAH PLUS"
                                Else
                                    Option2.Caption = "RUPIAH PLUS"
                                End If
                                sss = sss & " PLUS "
                            ElseIf Option3.Caption = "-------" Then
                                If l = 1 Then
                                    Text3.text = "RUPIAH PLUS"
                                    Option6.Caption = "RUPIAH PLUS"
                                Else
                                    Option3.Caption = "RUPIAH PLUS"
                                End If
                                sss = sss & " PLUS "
                            End If
                        End If
                    ElseIf M_objrsc!client = "EXPRES" Then
                        If sss Like "*EXPRES*" Then
                        Else
                            If Option1.Caption = "-------" Then
                                If l = 1 Then
                                    Text1.text = "UANGEXPRESS"
                                Else
                                    Option1.Caption = "UANGEXPRESS"
                                End If
                                sss = sss & " EXPRES "
                            ElseIf Option2.Caption = "-------" Then
                                If l = 1 Then
                                    Text2.text = "UANGEXPRESS"
                                Else
                                    Option2.Caption = "UANGEXPRESS"
                                End If
                                sss = sss & " EXPRES "
                            ElseIf Option3.Caption = "-------" Then
                                If l = 1 Then
                                    Text3.text = "UANGEXPRESS"
                                Else
                                    Option3.Caption = "UANGEXPRESS"
                                End If
                                sss = sss & " EXPRES "
                            End If
                        End If
                    ElseIf M_objrsc!client = "GLOBAL" Then
                        If sss Like "*GLOBAL*" Then
                        Else
                            If Option1.Caption = "-------" Then
                                If l = 1 Then
                                    Text1.text = "GLOBALINDO"
                                Else
                                    Option1.Caption = "GLOBALINDO"
                                End If
                                sss = sss & " GLOBAL "
                            ElseIf Option2.Caption = "-------" Then
                                If l = 1 Then
                                    Text2.text = "GLOBALINDO"
                                Else
                                    Option2.Caption = "GLOBALINDO"
                                End If
                                sss = sss & " GLOBAL "
                            ElseIf Option3.Caption = "-------" Then
                                If l = 1 Then
                                    Text3.text = "GLOBALINDO"
                                Else
                                    Option3.Caption = "GLOBALINDO"
                                End If
                                sss = sss & " GLOBAL "
                            End If
                        End If
                    Else
                        If sss Like "*" & M_objrsc!client & "*" Then
                        Else
                            If Option1.Caption = "-------" Then
                                If Text1.text = "" Or l <> 1 Then
                                    If l = 1 Then
                                        Text1.text = M_objrsc!client
                                        Option4.Caption = M_objrsc!client
                                    Else
                                        Option1.Caption = M_objrsc!client
                                    End If
                                    zz = zz & " " & M_objrsc!client
                                    sss = sss & " " & M_objrsc!client & " "
                                Else
                                    GoTo turun2
                                End If
                            ElseIf Option2.Caption = "-------" Then
turun2:
                                If Text2.text = "" Or l <> 1 Then
                                    If l = 1 Then
                                        Text2.text = M_objrsc!client
                                        Option5.Caption = M_objrsc!client
                                    Else
                                        Option2.Caption = M_objrsc!client
                                    End If
                                    zz = zz & " " & M_objrsc!client
                                    sss = sss & " " & M_objrsc!client & " "
                                Else
                                    GoTo turun3
                                End If
                            ElseIf Option3.Caption = "-------" Then
turun3:
                                If Text3.text = "" Or l <> 1 Then
                                    If l = 1 Then
                                        Text3.text = M_objrsc!client
                                        Option6.Caption = M_objrsc!client
                                    Else
                                        Option3.Caption = M_objrsc!client
                                    End If
                                    zz = zz & " " & M_objrsc!client
                                    sss = sss & " " & M_objrsc!client & " "
                                End If
                            End If
                        End If
                    End If
                End If
                M_objrsc.MoveNext
            Next i
        M_objrsc.MoveFirst
        M_objrs.MoveNext
    Wend
End Sub

Private Sub header()
    LvPTP.ColumnHeaders.clear
    With LvPTP.ColumnHeaders
        .ADD 1, , "CAMPAIGN CODE", 10 * 420
        .ADD 2, , "EXTRACTED CAMPAIGN CODE", 10 * 440
    End With
End Sub

Private Sub Command1_Click()
    LvPTP.ListItems.clear
    
    If Option1.Value = False And Option2.Value = False And Option3.Value = False Then
        MsgBox "Harap pilih Pegangan yang akan dipindahkan"
        Exit Sub
    End If

    If Option1.Value = True Then
        zzz = Option1.Caption
    ElseIf Option2.Value = True Then
        zzz = Option2.Caption
    ElseIf Option3.Value = True Then
        zzz = Option3.Caption
    End If
    
    If Option4.Value = True Then
        aaa = Option4.Caption
    ElseIf Option5.Value = True Then
        aaa = Option5.Caption
    ElseIf Option6.Value = True Then
        aaa = Option6.Caption
    End If
    
    If aaa <> "" Then
        ccc = " left('" & aaa & "',1)||'X'||substring('" & aaa & "','3',length('" & aaa & "')) "
    
        quext = " select ex, ('EX_'||split_part(ex,'" & aaa & "',1)||" & ccc & "||'_'||'" & zzz & "'||split_part(ex,'" & aaa & "',2))::varchar new from ( " & vbCrLf
        quext = quext & " select distinct (recsource)::varchar ex from mgm where agent = '" & Label3.Caption & "' and recsource ilike '%" & aaa & "%' " & vbCrLf
        quext = quext & " ) a " & vbCrLf
        'quext = " select distinct ('EX_'||recsource)::varchar from mgm where agent = '" & Label3.Caption & "' and recsource ilike '%" & Text1.text & "%' "
    End If
'    If Text2.text <> "" Then
'        ccc = " left('" & Text2.text & "',1)||'X'||substring('" & Text2.text & "','3',length('" & Text2.text & "')) "
'
'        quext = " select ex, ('EX_'||split_part(ex,'" & Text2.text & "',1)||" & ccc & "||'_'||'" & zzz & "'||split_part(ex,'" & Text2.text & "',2))::varchar new from ( " & vbCrLf
'        quext = quext & " select distinct (recsource)::varchar ex from mgm where agent = '" & Label3.Caption & "' and recsource ilike '%" & Text2.text & "%'" & vbCrLf
'        quext = quext & " ) a " & vbCrLf
'    End If
'    If Text3.text <> "" Then
'        ccc = " left('" & Text3.text & "',1)||'X'||substring('" & Text3.text & "','3',length('" & Text3.text & "')) "
'
'        quext = " select ex, ('EX_'||split_part(ex,'" & Text3.text & "',1)||" & ccc & "||'_'||'" & zzz & "'||split_part(ex,'" & Text3.text & "',2))::varchar new from ( " & vbCrLf
'        quext = quext & " select distinct (recsource)::varchar ex from mgm where agent = '" & Label3.Caption & "' and recsource ilike '%" & Text3.text & "%' " & vbCrLf
'        quext = quext & " ) a " & vbCrLf
'    End If
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open quext, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    While Not M_objrs.EOF
        Set listv = LvPTP.ListItems.ADD(, , IIf(IsNull(M_objrs!ex), "", M_objrs!ex))
        listv.SubItems(1) = IIf(IsNull(M_objrs!New), "", M_objrs!New)
        M_objrs.MoveNext
    Wend
End Sub

Private Sub Command3_Click()
    Call create_table_log
    
    If LvPTP.ListItems.Count = 0 Then
        MsgBox "Harap extract campaign terlebih dahulu"
        Exit Sub
    End If
    
    If Option4.Value = True Then
        aaa = Option4.Caption
    ElseIf Option5.Value = True Then
        aaa = Option5.Caption
    ElseIf Option6.Value = True Then
        aaa = Option6.Caption
    End If
    
    qexecute = " update usertbl set team = '" & Label7.Caption & "', spvcode = '" & Label7.Caption & "' where userid = '" & Label3.Caption & "';" & vbCrLf ' and recsource ilike '%" & Text1.text & "%'; "
    If Text1.text <> "" Then
        For i = 1 To LvPTP.ListItems.Count
             q = "select * from datasourcetbl where kodeds = '" & LvPTP.ListItems(1).SubItems(1) & "'"
             Set M_objrs = New ADODB.Recordset
             M_objrs.CursorLocation = adUseClient
             M_objrs.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic
             
             d = 0
             If M_objrs.RecordCount = 0 Then
                 d = 1
             End If
            
             qexecute = qexecute & " update mgm set recsource = '" & LvPTP.ListItems(i).SubItems(1) & "' where agent = '" & Label3.Caption & "' and recsource ilike '%" & aaa & "%' and recsource = '" & LvPTP.ListItems(i).text & "'; " & vbCrLf
             qexecute = qexecute & " insert into tbl_changed_team_log values ('" & Label3.Caption & "','" & Label4.Caption & "','" & Label7.Caption & "', '" & LvPTP.ListItems(i).text & "', '" & LvPTP.ListItems(i).SubItems(1) & "','" & MDIForm1.TxtUsername.text & "');" & vbCrLf
             
             If d = 1 Then
                 qexecute = qexecute & " insert into datasourcetbl (kodeds,status,campaign_ket,tglentry) values ('" & LvPTP.ListItems(i).SubItems(1) & "','1',to_char(now(),'yyyymmdd'), now()); " & vbCrLf
             End If
        Next i
    End If
'    If Text2.text <> "" Then
'        q = "select * from datasourcetbl where kodeds = '" & LvPTP.ListItems(2).SubItems(1) & "'"
'        Set M_objrs = New ADODB.Recordset
'        M_objrs.CursorLocation = adUseClient
'        M_objrs.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic
'
'        d = 0
'        If M_objrs.RecordCount = 0 Then
'            d = 1
'        End If
'
'        qexecute = qexecute & " update mgm set recsource = '" & LvPTP.ListItems(2).SubItems(1) & "' where agent = '" & Label3.Caption & "' and recsource ilike '%" & Text2.text & "%'; " & vbCrLf
'        qexecute = qexecute & " insert into tbl_changed_team_log values ('" & Label3.Caption & "','" & Label4.Caption & "','" & Label7.Caption & "', '" & LvPTP.ListItems(2).text & "', '" & LvPTP.ListItems(2).SubItems(1) & "','" & MDIForm1.TxtUsername.text & "');" & vbCrLf
'
'        If d = 1 Then
'            qexecute = qexecute & " insert into datasourcetbl (kodeds,status,campaign_ket,tglentry) values ('" & LvPTP.ListItems(2).SubItems(1) & "','1',to_char(now(),'yyyymmdd'), now()); " & vbCrLf
'        End If
'    End If
'    If Text3.text <> "" Then
'        q = "select * from datasourcetbl where kodeds = '" & LvPTP.ListItems(3).SubItems(1) & "'"
'        Set M_objrs = New ADODB.Recordset
'        M_objrs.CursorLocation = adUseClient
'        M_objrs.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic
'
'        d = 0
'        If M_objrs.RecordCount = 0 Then
'            d = 1
'        End If
'
'        qexecute = qexecute & " update mgm set recsource = '" & LvPTP.ListItems(3).SubItems(1) & "' where agent = '" & Label3.Caption & "' and recsource ilike '%" & Text3.text & "%'; " & vbCrLf
'        qexecute = qexecute & " insert into tbl_changed_team_log values ('" & Label3.Caption & "','" & Label4.Caption & "','" & Label7.Caption & "', '" & LvPTP.ListItems(3).text & "', '" & LvPTP.ListItems(3).SubItems(1) & "','" & MDIForm1.TxtUsername.text & "');" & vbCrLf
'
'        If d = 1 Then
'            qexecute = qexecute & " insert into datasourcetbl (kodeds,status,campaign_ket,tglentry) values ('" & LvPTP.ListItems(3).SubItems(1) & "','1',to_char(now(),'yyyymmdd'), now()); " & vbCrLf
'        End If
'    End If
    M_OBJCONN.Execute qexecute
    
    MsgBox "Done", vbOKOnly, "Pesan"
    Call CmdClear_Click
End Sub

Private Sub Form_Load()
    Call header
End Sub

Private Sub extract_campaign()
    q = "select distinct recsource from mgm where agent = '" & Label3.Caption & "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If M_objrs.RecordCount > 0 Then
        For i = 1 To M_objrs.RecordCount
        ''BELOMAN''
            'sdfdsf
            M_objrs.MoveNext
        Next i
    End If
End Sub

Private Sub Option1_Click()
    If Option1.Caption = "-------" Then
        MsgBox "Tidak bisa memilih yang kosong"
        Option1.Value = False
        Exit Sub
    End If
End Sub

Private Sub Option2_Click()
    If Option2.Caption = "-------" Then
        MsgBox "Tidak bisa memilih yang kosong"
        Option2.Value = False
        Exit Sub
    End If
End Sub

Private Sub Option3_Click()
    If Option3.Caption = "-------" Then
        MsgBox "Tidak bisa memilih yang kosong"
        Option3.Value = False
        Exit Sub
    End If
End Sub

Private Sub Option4_Click()
    If Option4.Caption = "-------" Then
        MsgBox "Tidak bisa memilih yang kosong"
        Option4.Value = False
        Exit Sub
    End If
End Sub

Private Sub Option5_Click()
    If Option5.Caption = "-------" Then
        MsgBox "Tidak bisa memilih yang kosong"
        Option5.Value = False
        Exit Sub
    End If
End Sub

Private Sub Option6_Click()
    If Option6.Caption = "-------" Then
        MsgBox "Tidak bisa memilih yang kosong"
        Option6.Value = False
        Exit Sub
    End If
End Sub
