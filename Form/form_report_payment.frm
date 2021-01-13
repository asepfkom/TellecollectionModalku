VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form form_report_payment 
   Caption         =   "Report Payment"
   ClientHeight    =   7380
   ClientLeft      =   4380
   ClientTop       =   1260
   ClientWidth     =   11865
   LinkTopic       =   "Form2"
   ScaleHeight     =   7380
   ScaleWidth      =   11865
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.CommandButton Command2 
         Caption         =   "Export"
         Height          =   375
         Left            =   4995
         TabIndex        =   6
         Top             =   630
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         Height          =   375
         Left            =   5010
         TabIndex        =   5
         Top             =   165
         Width           =   1575
      End
      Begin TDBDate6Ctl.TDBDate tgl1 
         Height          =   315
         Left            =   1365
         TabIndex        =   1
         Top             =   165
         Width           =   1485
         _Version        =   65536
         _ExtentX        =   2619
         _ExtentY        =   556
         Calendar        =   "form_report_payment.frx":0000
         Caption         =   "form_report_payment.frx":0118
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_report_payment.frx":0184
         Keys            =   "form_report_payment.frx":01A2
         Spin            =   "form_report_payment.frx":0200
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
         Left            =   3510
         TabIndex        =   2
         Top             =   165
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   556
         Calendar        =   "form_report_payment.frx":0228
         Caption         =   "form_report_payment.frx":0340
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_report_payment.frx":03AC
         Keys            =   "form_report_payment.frx":03CA
         Spin            =   "form_report_payment.frx":0428
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
      Begin MSComctlLib.ListView LvPTP 
         Height          =   5580
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   11580
         _ExtentX        =   20426
         _ExtentY        =   9843
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
      Begin MSComDlg.CommonDialog CD_save 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   0
         X2              =   11880
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label2 
         Caption         =   "Tanggal"
         Height          =   255
         Left            =   570
         TabIndex        =   4
         Top             =   165
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   255
         Left            =   3060
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "form_report_payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call search
End Sub

Private Sub Form_Load()
    Call header
    'Call list_client(Combo1)
End Sub

Private Sub header()
    LvPTP.ColumnHeaders.clear
    With LvPTP.ColumnHeaders
        .ADD 1, , "No"
        .ADD 2, , "Nama Customer"
        .ADD 3, , "CN"
        .ADD 4, , "Tanggal Pembayaran"
        .ADD 5, , "Jumlah Pembayaran"
        .ADD 6, , "Status Pembayaran"
        .ADD 7, , "Agent"
    End With
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

Private Sub search()
'    If Combo1.text = "" Then
'        MsgBox "Harap Pilih Rekan"
'        Combo1.SetFocus
'        Exit Sub
'    End If
    
'    If tgl1.ValueIsNull = True Or tgl2.ValueIsNull = True Then
'        MsgBox "Harap Pilih Tanggal"
'        Exit Sub
'    End If
    
    c_tgl1 = Format(tgl1.Value, "yyyy-mm-dd")
    c_tgl2 = Format(tgl2.Value, "yyyy-mm-dd")
    
    getmgm = " (select custid, name from mgm where custid in (select custid from tbllunas where paydate between " & vbCrLf
    getmgm = getmgm & "'" & c_tgl1 & " 00:00:00' and '" & c_tgl2 & " 23:59:59') ) a" & vbCrLf
    
    gettbllunas = "(select custid,paydate,payment,datafrom,agent from tbllunas where paydate between " & vbCrLf
    gettbllunas = gettbllunas & " '" & c_tgl1 & " 00:00:00' and '" & c_tgl2 & " 23:59:59') b" & vbCrLf
    
    fusion = "select name,a.custid,paydate,payment,datafrom,agent from " & getmgm & ", " & gettbllunas & " where a.custid = b.custid"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open fusion, M_OBJCONN, adOpenDynamic, adLockOptimistic
    LvPTP.ListItems.clear
    a = 1
    While Not M_objrs.EOF
        Set ListItem = LvPTP.ListItems.ADD(, , a)
            ListItem.SubItems(1) = cnull(M_objrs("name"))
            ListItem.SubItems(2) = cnull(M_objrs("custid"))
            ListItem.SubItems(3) = Format(M_objrs("paydate"), "yyyy-mm-dd hh:nn:ss")
            ListItem.SubItems(4) = cnull(M_objrs("payment"))
            ListItem.SubItems(5) = cnull(M_objrs("datafrom"))
            ListItem.SubItems(6) = cnull(M_objrs!AGENT)
            a = a + 1
        M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing
End Sub

