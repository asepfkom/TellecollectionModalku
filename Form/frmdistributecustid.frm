VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmdistributecustid 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Distribute per CustID"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5340
   LinkTopic       =   "Form2"
   ScaleHeight     =   6720
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Information"
      Height          =   4455
      Left            =   0
      TabIndex        =   13
      Top             =   2280
      Width           =   9255
      Begin MSComctlLib.ListView LvPTP 
         Height          =   4020
         Left            =   60
         TabIndex        =   14
         Top             =   240
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   7091
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "600"
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000E&
         Caption         =   "Cara Pakai"
         Height          =   2175
         Left            =   1560
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   3735
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmdistributecustid.frx":0000
            Height          =   2055
            Left            =   120
            TabIndex        =   3
            Top             =   180
            Width           =   2895
         End
         Begin VB.Label Label7 
            BackColor       =   &H8000000E&
            Caption         =   "_"
            BeginProperty Font 
               Name            =   "Showcard Gothic"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3120
            TabIndex        =   2
            Top             =   120
            Width           =   495
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Distribute"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3960
         TabIndex        =   8
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtlocation 
         Height          =   315
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   3165
      End
      Begin VB.CommandButton cmdbrowse 
         BackColor       =   &H00C0FFC0&
         Caption         =   "...."
         Height          =   315
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   585
         Width           =   555
      End
      Begin VB.ComboBox cbosheet 
         Height          =   315
         Left            =   1350
         TabIndex        =   5
         Top             =   990
         Width           =   3165
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Check"
         Height          =   495
         Left            =   2520
         TabIndex        =   4
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Timer Timer1 
         Interval        =   600
         Left            =   11160
         Top             =   360
      End
      Begin VB.Label Label3 
         Caption         =   "History"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         Caption         =   "Cara Pakai"
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
         Left            =   4080
         TabIndex        =   9
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Choose File"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Location"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "Sheet"
         Height          =   255
         Left            =   150
         TabIndex        =   10
         Top             =   1020
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmdistributecustid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public M_XLSCONN As New ADODB.Connection
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
    cbosheet.CLEAR
    If M_objrs.EOF And M_objrs.BOF Then Exit Sub
    While Not M_objrs.EOF
        cbosheet.AddItem IIf(IsNull(M_objrs!table_name), "", M_objrs!table_name)
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    'Set M_XLSCONN = Nothing

End Sub

Private Sub Command1_Click()
    qs = "select now() as tanggal"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    abc = Format(M_objrs!Tanggal, "yyyy_mm_ddi__hhi_mi_ss")
    
    qc = "create table backupdistribute___" & abc & " as (select mgm.custid,mgm.agent as agent_lama, a.agent as agent_baru from mgm, tbltemp_distributepercust a where mgm.custid = a.custid);" & vbCrLf
    qc = qc + "alter table backupdistribute___" & abc & " add column distributeby varchar ;" & vbCrLf
    qc = qc + "alter table backupdistribute___" & abc & " add column tanggal timestamp without time zone default now();" & vbCrLf
    qc = qc + "update backupdistribute___" & abc & " set distributeby = '" & MDIForm1.TxtUsername.text & "';" & vbCrLf & vbCrLf
    
    qc = qc + "update mgm set agent = a.agent from tbltemp_distributepercust a where mgm.custid = a.custid;"
    M_OBJCONN.Execute qc
    
    MsgBox "Distribute Success"
    
End Sub

Private Sub Command2_Click()
    Dim str_sql As String
    
    qs = "select * from information_schema.columns where table_name = 'tbltemp_distributepercust'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
    If M_objrs.RecordCount = 0 Then
        qc = "create table tbltemp_distributepercust ( id serial, custid varchar, agent varchar );"
        M_OBJCONN.Execute qc
    End If
        qd = "delete from tbltemp_distributepercust;"
        M_OBJCONN.Execute qd
        
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
        
            str_sql = "INSERT INTO tbltemp_distributepercust (custid,agent) Values ( '" + rs(0) + "', '" + rs(1) + "' );"
            M_OBJCONN.Execute str_sql
            
            rs.MoveNext
        Wend
                    
                
        a = rs.RecordCount
                
        qs = "select mgm.custid,mgm.name,mgm.agent as agentlama, tbltemp_distributepercust.agent as agentbaru from mgm,tbltemp_distributepercust  where mgm.custid = tbltemp_distributepercust.custid"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        LvPTP.ListItems.CLEAR
        
        While Not M_objrs.EOF
            Set ListItem = LvPTP.ListItems.ADD(, , cnull(M_objrs("custid")))
                ListItem.SubItems(1) = cnull(M_objrs("name"))
                ListItem.SubItems(2) = cnull(M_objrs("agentlama"))
                ListItem.SubItems(3) = cnull(M_objrs("agentbaru"))
            M_objrs.MoveNext
        Wend
        
        B = M_objrs.RecordCount
        
        
        Dim teks As String
        teks = "Jumlah Data Excel : " & a & vbCrLf & "Jumlah Data Didatabase setelah dicheck : " & B & vbCrLf
        
        If a <> B Then
            teks = teks & "Status : Tidak Sesuai"
        Else
            teks = teks & "Status : Sesuai"
        End If
        MsgBox teks
        Command1.Enabled = True
End Sub

Private Sub Form_Load()
    Call header
End Sub

Private Sub Label3_Click()
    formlogdistribute.Show
End Sub

Private Sub Label5_Click()
    Frame2.Visible = True
End Sub

Private Sub Label7_Click()
    Frame2.Visible = False
End Sub

Private Sub Timer1_Timer()
    If Label5.BackColor = &H8000000D Then
        Label5.BackColor = &H8000000F
    Else
        Label5.BackColor = &H8000000D
    End If
End Sub

Private Sub header()
    LvPTP.ColumnHeaders.CLEAR
    With LvPTP.ColumnHeaders
        .ADD 1, , "CUSTID"
        .ADD 2, , "NAMA NASABAH"
        .ADD 3, , "AGENT LAMA"
        .ADD 4, , "AGENT DISTRIBUSI"
    End With
End Sub
