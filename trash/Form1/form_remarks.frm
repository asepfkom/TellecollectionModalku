VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form form_remarks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remarks"
   ClientHeight    =   6945
   ClientLeft      =   1770
   ClientTop       =   2310
   ClientWidth     =   10455
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   10455
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "form_remarks.frx":0000
      Left            =   1680
      List            =   "form_remarks.frx":001F
      TabIndex        =   1
      Top             =   285
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      Begin VB.CommandButton Command2 
         Caption         =   "Export"
         Height          =   375
         Left            =   7680
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin MSComctlLib.ListView LvPTP 
         Height          =   5340
         Left            =   0
         TabIndex        =   5
         Top             =   1440
         Width           =   10380
         _ExtentX        =   18309
         _ExtentY        =   9419
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
      Begin TDBDate6Ctl.TDBDate tgl1 
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Top             =   720
         Width           =   1485
         _Version        =   65536
         _ExtentX        =   2619
         _ExtentY        =   556
         Calendar        =   "form_remarks.frx":0067
         Caption         =   "form_remarks.frx":017F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_remarks.frx":01EB
         Keys            =   "form_remarks.frx":0209
         Spin            =   "form_remarks.frx":0267
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
         Left            =   3765
         TabIndex        =   8
         Top             =   720
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   556
         Calendar        =   "form_remarks.frx":028F
         Caption         =   "form_remarks.frx":03A7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_remarks.frx":0413
         Keys            =   "form_remarks.frx":0431
         Spin            =   "form_remarks.frx":048F
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
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Tanggal"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   10440
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "BANK/FinTech"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog CD_save 
      Left            =   0
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "form_remarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub header()
    LvPTP.ColumnHeaders.clear
    With LvPTP.ColumnHeaders
        .ADD 1, , "CUSTID"
        .ADD 2, , "AGENT"
        .ADD 3, , "KODEDS"
        .ADD 4, , "STATUSCALL"
        .ADD 5, , "PHONE"
        .ADD 6, , "HST"
        .ADD 7, , "UNIQUE"
        .ADD 8, , "TANGGAL"
    End With
End Sub

Private Sub search()
    If Combo1.text = "" Then
        MsgBox "Harap Pilih Rekan"
        Combo1.SetFocus
        Exit Sub
    End If
    
    If tgl1.ValueIsNull = True Or tgl2.ValueIsNull = True Then
        MsgBox "Harap Pilih Tanggal"
        Exit Sub
    End If
    
    c_tgl1 = Format(tgl1.Value, "yyyy-mm-dd")
    c_tgl2 = Format(tgl2.Value, "yyyy-mm-dd")
    
    sStrsql = "select custid, agent, hst, kodeds, tgl, phoneno, statuscall, unique_id from mgm_hst "
    sStrsql = sStrsql & " where custid in (select custid from mgm where recsource ilike '%" & Combo1.text & "%') "
    sStrsql = sStrsql & " and tgl between '" & c_tgl1 & " 00:00:00' and '" & c_tgl2 & " 23:59:59'"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    LvPTP.ListItems.clear
    While Not M_objrs.EOF
        Set ListItem = LvPTP.ListItems.ADD(, , cnull(M_objrs("custid")))
            ListItem.SubItems(1) = cnull(M_objrs("agent"))
            ListItem.SubItems(2) = cnull(M_objrs("kodeds"))
            ListItem.SubItems(3) = cnull(M_objrs("statuscall"))
            ListItem.SubItems(4) = cnull(M_objrs("phoneno"))
            ListItem.SubItems(5) = cnull(M_objrs("hst"))
            ListItem.SubItems(6) = cnull(M_objrs("unique_id"))
            ListItem.SubItems(7) = Format(M_objrs("tgl"), "yyyy-mm-dd hh:nn:ss")
        M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing
End Sub

Private Sub Command1_Click()
    Call search
End Sub

Private Sub Command2_Click()
    Dim objExcel As New Excel.Application
    Dim objExcelSheet As Excel.Worksheet
    Dim col, row As Integer
    Dim a As String
    If LvPTP.ListItems.Count > 0 Then
        objExcel.Workbooks.ADD
        Set objExcelSheet = objExcel.Worksheets.ADD
     
    
        For col = 1 To LvPTP.ColumnHeaders.Count
            objExcelSheet.Cells(1, col).Value = LvPTP.ColumnHeaders(col)
        Next
     
        For row = 2 To LvPTP.ListItems.Count + 1
            For col = 1 To LvPTP.ColumnHeaders.Count
            If col = 1 Then
                    objExcelSheet.Cells(row, col).Value = "'" + LvPTP.ListItems(row - 1).text
            Else
                '" 'cararandy 29032016 "
                Dim hasil1 As String
                    If col <> 12 And col <> 14 Then
                        hasil1 = "'" + LvPTP.ListItems(row - 1).SubItems(col - 1)
                        objExcelSheet.Cells(row, col).Value = hasil1
                    Else
                        hasil1 = "'" + LvPTP.ListItems(row - 1).SubItems(col - 1)
                        objExcelSheet.Cells(row, col).Value = hasil1
                    End If
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

Private Sub Form_Load()
    Call header
    Call list_client(Combo1)
End Sub
