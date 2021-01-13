VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form form_maintenance_db 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maintenance DB"
   ClientHeight    =   6795
   ClientLeft      =   855
   ClientTop       =   1785
   ClientWidth     =   14010
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   14010
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13935
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   375
         Left            =   3960
         TabIndex        =   3
         Top             =   6240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Top             =   6240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Export"
         Height          =   375
         Left            =   12360
         TabIndex        =   1
         Top             =   6240
         Width           =   1575
      End
      Begin MSComctlLib.ListView LvPTP 
         Height          =   6060
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   4860
         _ExtentX        =   8573
         _ExtentY        =   10689
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   6060
         Left            =   5640
         TabIndex        =   5
         Top             =   120
         Width           =   8340
         _ExtentX        =   14711
         _ExtentY        =   10689
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
         X1              =   5160
         X2              =   5160
         Y1              =   0
         Y2              =   6840
      End
   End
End
Attribute VB_Name = "form_maintenance_db"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub search()
    LvPTP.ColumnHeaders.clear
    With LvPTP.ColumnHeaders
        .ADD 1, , "TABLE", 5000
    End With
    
    sStrsql = "select distinct table_name from information_schema.columns  where table_name ilike '%backupdistribute__%' or table_name ilike '%backuppullout_%' order by 1"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    LvPTP.ListItems.clear
    While Not M_objrs.EOF
        Set ListItem = LvPTP.ListItems.ADD(, , cnull(M_objrs("table_name")))
        M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing
End Sub

Private Sub view()
    Dim c1 As Integer

    sStrsql = "select column_name from information_schema.columns  where table_name = '" & LvPTP.SelectedItem.text & "' order by ordinal_position"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    ListView1.ListItems.clear
    ListView1.ColumnHeaders.clear
    If M_objrs.RecordCount > 0 Then
        'ListView1.ColumnHeaders.clear
        c1 = M_objrs.RecordCount
        With ListView1.ColumnHeaders
            For i = 1 To c1
                .ADD i, , cnull(M_objrs!column_name)
                M_objrs.MoveNext
            Next i
        End With
    End If
    Set M_objrs = Nothing
    
    sStrsql = "select * from " & LvPTP.SelectedItem.text
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    While Not M_objrs.EOF
        
        Set ListItem = ListView1.ListItems.ADD(, , cnull(M_objrs("custid")))
            For i = 1 To c1 - 1
                ListItem.SubItems(i) = cnull(M_objrs(i))
            Next i
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
    If ListView1.ListItems.Count > 0 Then
        objExcel.Workbooks.ADD
        Set objExcelSheet = objExcel.Worksheets.ADD
     
    
        For col = 1 To ListView1.ColumnHeaders.Count
            objExcelSheet.Cells(1, col).Value = ListView1.ColumnHeaders(col)
        Next
     
        For row = 2 To ListView1.ListItems.Count + 1
            For col = 1 To ListView1.ColumnHeaders.Count
            If col = 1 Then
                    objExcelSheet.Cells(row, col).Value = ListView1.ListItems(row - 1).text
            Else
                '" 'cararandy 29032016 "
                Dim hasil1 As String
                    If col <> 12 And col <> 14 Then
                        hasil1 = "'" + ListView1.ListItems(row - 1).SubItems(col - 1)
                        objExcelSheet.Cells(row, col).Value = hasil1
                    Else
                        hasil1 = ListView1.ListItems(row - 1).SubItems(col - 1)
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

Private Sub Command3_Click()
    Dim a As Integer
    a = 0
    For i = 1 To LvPTP.ListItems.Count
        If LvPTP.ListItems(i).Checked = True Then
            query = "Drop table " & LvPTP.ListItems(i).text
            M_OBJCONN.Execute query
            a = a + 1
        End If
    Next i
    MsgBox "Table terhapus sebanyak " & a
    Call search
End Sub

Private Sub Form_Load()
    Call search
End Sub

Private Sub LvPTP_DblClick()
    Call view
End Sub
