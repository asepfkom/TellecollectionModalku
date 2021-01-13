VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form_rptcall_activity 
   Caption         =   "Report Call Activity "
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   19140
   LinkTopic       =   "Form2"
   ScaleHeight     =   5130
   ScaleWidth      =   19140
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000B&
      Caption         =   "Dashboard"
      Height          =   5070
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   19035
      Begin VB.Frame Frame4 
         Caption         =   "Search"
         Height          =   4455
         Left            =   15690
         TabIndex        =   1
         Top             =   210
         Width           =   3255
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "Form_rptcall_activity.frx":0000
            Left            =   1695
            List            =   "Form_rptcall_activity.frx":0002
            TabIndex        =   9
            Top             =   1020
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Search"
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   1065
            Width           =   1455
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Export"
            Height          =   375
            Left            =   105
            TabIndex        =   3
            Top             =   1470
            Width           =   1455
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Touch per Custid per Agent"
            Height          =   435
            Left            =   1575
            TabIndex        =   2
            Top             =   1455
            Visible         =   0   'False
            Width           =   1455
         End
         Begin TDBDate6Ctl.TDBDate TDBDate3 
            Height          =   285
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   1470
            _Version        =   65536
            _ExtentX        =   2593
            _ExtentY        =   503
            Calendar        =   "Form_rptcall_activity.frx":0004
            Caption         =   "Form_rptcall_activity.frx":011C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Form_rptcall_activity.frx":0188
            Keys            =   "Form_rptcall_activity.frx":01A6
            Spin            =   "Form_rptcall_activity.frx":0204
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   12648447
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
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   6815745
            Value           =   39876
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate TDBDate4 
            Height          =   285
            Left            =   1680
            TabIndex        =   6
            Top             =   720
            Width           =   1470
            _Version        =   65536
            _ExtentX        =   2593
            _ExtentY        =   503
            Calendar        =   "Form_rptcall_activity.frx":022C
            Caption         =   "Form_rptcall_activity.frx":0344
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Form_rptcall_activity.frx":03B0
            Keys            =   "Form_rptcall_activity.frx":03CE
            Spin            =   "Form_rptcall_activity.frx":042C
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   12648447
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
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   6815745
            Value           =   39876
            CenturyMode     =   0
         End
         Begin MSComDlg.CommonDialog CD_save 
            Left            =   2760
            Top             =   90
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label7 
            Caption         =   "Date"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   615
         End
      End
      Begin MSComctlLib.ListView LvAgent 
         Height          =   4440
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   15480
         _ExtentX        =   27305
         _ExtentY        =   7832
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "Form_rptcall_activity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command6_Click()

    If (TDBDate3.ValueIsNull) And (TDBDate3.ValueIsNull) Then
        MsgBox "Tanggal Call Harus Diisi", vbInformation, "Informasi"
        Exit Sub
    End If
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    CMDSQL = " select tblstatuscall_kdstscall as stts from tblstatuscall where tblstatuscall_kdstatus = '1' order by tblstatuscall_keterangan"
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvAgent.ColumnHeaders.clear
    LvAgent.ColumnHeaders.ADD 1, , "No", 10 * 120
    '===asep==='
    LvAgent.ColumnHeaders.ADD 2, , "New Data", 10 * 120
    LvAgent.ColumnHeaders.ADD 3, , "Jumlah Data", 10 * 120
    LvAgent.ColumnHeaders.ADD 4, , "Call", 10 * 120
    LvAgent.ColumnHeaders.ADD 5, , "Durasi", 10 * 120
    'LvAgent.ColumnHeaders.ADD 5, , "AgentStatus", 10 * 120
    LvAgent.ColumnHeaders.ADD 6, , "Agent", 10 * 120
    '==========='
    z = 7
    While Not M_objrs.EOF
        LvAgent.ColumnHeaders.ADD z, , "" & M_objrs!stts & "", 7 * 120
        M_objrs.MoveNext
        z = z + 1
    Wend
    
    
    LvAgent.ColumnHeaders.ADD z, , "TOTAL", 10 * 120

    
    'isi
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    CMDSQL = " select tblstatuscall_keterangan as stts from tblstatuscall where tblstatuscall_kdstatus = '1' order by 1"
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    a = ""
    B = ""
    c = ""
    
    While Not M_objrs.EOF
        a = a + " ,case when kodeds = '" & "" & M_objrs!stts & "" & "' then 1 else 0 end as """ & "" & M_objrs!stts & """"
        B = B + """" & "" & M_objrs!stts & """+"
        c = c + " ,sum( """ & "" & M_objrs!stts & """" & " ) as """ & "" & M_objrs!stts & """"
        M_objrs.MoveNext
    Wend
        B = Left(B, Len(B) - 1)
        c = c
    
'    q = " select agent" & "" & c & ", sum(total) as total from ("
'    q = q + "select *," & "" & B & " as Total from ("
'    q = q & "select agent " & "" & a & ""
'    q = q & "from (select a.agent, a.custid, a.kodeds, b.recsource from mgm_hst a inner join mgm b on a.custid = b.custid "
'    q = q & " where tgl between '" & Format(TDBDate3.Value, "yyyy-mm-dd") & " 00:00:00' and '" & Format(TDBDate4.Value, "yyyy-mm-dd") & " 23:59:59' "

    '=========asep19/01/2020======'
    q = " select b.jml as ""New Data"" ,d.jumlah_data as ""Jumlah Data"", c.callattempt as ""Call"",c.durasi as ""Durasi"""
    q = q + " ,a.* from("
    q = q + "select agent" & "" & c & ", sum(total) as total from ("
    q = q + "select *," & "" & B & " as Total from ("
    q = q & "select agent " & "" & a & ""
    q = q & "from (select agent, custid, kodeds from mgm_hst"
    q = q & " where tgl between '" & Format(TDBDate3.Value, "yyyy-mm-dd") & " 00:00:00' and '" & Format(TDBDate4.Value, "yyyy-mm-dd") & " 23:59:59' "

'    If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Then
'        q = q & " and in (select distinct recsource from mgm where agent in (select userid from usertbl where team = '" & MDIForm1.txtusername.text & "' or userid = '" & MDIForm1.txtusername.text & "' )) "
'    End If

    If Combo1.text = "RUPIAH PLUS" Then
        q = q & " ) hst "
    ElseIf Combo1.text = "UANGEXPRESS" Then
        q = q & " ) hst "
    ElseIf Combo1.text = "GLOBALINDO" Then
        q = q & " ) hst "
    Else
        q = q & " ) hst "
    End If

    '=========================='
'    If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Then
'        q = q & " and b.recsource in (select distinct recsource from mgm where agent in (select userid from usertbl where team = '" & MDIForm1.TxtUsername.text & "' or userid = '" & MDIForm1.TxtUsername.text & "' )) "
'    End If
'
'    If Combo1.text = "RUPIAH PLUS" Then
'        q = q & " and recsource ilike '%PLUS%') hst "
'    ElseIf Combo1.text = "UANGEXPRESS" Then
'        q = q & " and recsource ilike '%EXPRESS%') hst "
'    ElseIf Combo1.text = "GLOBALINDO" Then
'        q = q & " and recsource ilike '%GLOBAL%') hst "
'    Else
'        q = q & " and recsource ilike '%" & Combo1.text & "%') hst "
'    End If

    q = q & " ) abc "
    q = q & " ) a group by agent "
    q = q + " )a left join"
    '=======tambahan asep===='
    q = q + " (select agent, count(statuscall) as jml from mgm where coalesce(statuscall,'')= 'New Data' group by agent) b"
    q = q + " on a.agent=b.agent Left Join"
    q = q + " (select agent, count(agent)as callattempt, sum(durasi_billsec) as durasi  from mgm_hst where tgl between '" & Format(TDBDate3.Value, "yyyy-mm-dd") & " 00:00:00' and '" & Format(TDBDate4.Value, "yyyy-mm-dd") & " 23:59:59'"
    q = q + " group by agent)c on a.agent=c.agent Left Join"
    q = q + " (select agent,count(id) as jumlah_data from mgm group by agent)d on a.agent =d.agent "
    '=======end======'
    Set M_objrs = New ADODB.Recordset
    
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    LvAgent.ListItems.clear
    While Not M_objrs.EOF
        Set ListItem = LvAgent.ListItems.ADD(, , M_objrs.Bookmark)
        For i = 1 To z - 1
            ListItem.SubItems(i) = IIf(IsNull(M_objrs(i - 1)), "", M_objrs(i - 1))
        Next i
        M_objrs.MoveNext
    Wend
End Sub

Private Sub Command7_Click()
    Dim objExcel As New Excel.Application
Dim objExcelSheet As Excel.Worksheet
Dim col, row As Integer
Dim a As String
If LvAgent.ListItems.Count > 0 Then
    objExcel.Workbooks.ADD
    Set objExcelSheet = objExcel.Worksheets.ADD
 

    For col = 1 To LvAgent.ColumnHeaders.Count
        objExcelSheet.Cells(1, col).Value = LvAgent.ColumnHeaders(col)
    Next
 
    For row = 2 To LvAgent.ListItems.Count + 1
        For col = 1 To LvAgent.ColumnHeaders.Count
        If col = 1 Then
                objExcelSheet.Cells(row, col).Value = LvAgent.ListItems(row - 1).text
        Else
            '" 'cararandy 29032016 "
            Dim hasil1 As String
                hasil1 = "'" + LvAgent.ListItems(row - 1).SubItems(col - 1)
                objExcelSheet.Cells(row, col).Value = hasil1
            End If
        Next
    Next
 
    objExcelSheet.Columns.AutoFit
    CD_save.ShowOpen
    a = CD_save.FileName
 
    objExcelSheet.SaveAs a & ".xls"
    MsgBox "Export Completed", vbInformation, Me.Caption
 
    objExcel.Workbooks.Open a & ".xls"
    objExcel.Visible = True
Else
    MsgBox "No data to export", vbInformation, Me.Caption
End If
End Sub
