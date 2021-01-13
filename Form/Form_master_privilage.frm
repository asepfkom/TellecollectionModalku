VERSION 5.00
Begin VB.Form Form_master_privilage 
   BackColor       =   &H80000011&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Privilage"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11850
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   11850
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Setup Auto Dialer"
      Height          =   255
      Index           =   16
      Left            =   6720
      TabIndex        =   30
      Top             =   2415
      Width           =   2430
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Caption         =   "Master Privilage"
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      Begin VB.CheckBox Check1 
         Caption         =   "Report History Upload"
         Height          =   255
         Index           =   5
         Left            =   4095
         TabIndex        =   29
         Top             =   960
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Report Activity Call"
         Height          =   255
         Index           =   9
         Left            =   4095
         TabIndex        =   28
         Top             =   2310
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Report Break Time"
         Height          =   255
         Index           =   10
         Left            =   4095
         TabIndex        =   27
         Top             =   2640
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   375
         Left            =   9600
         TabIndex        =   26
         Top             =   7080
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Follow Up"
         Height          =   255
         Index           =   18
         Left            =   9720
         TabIndex        =   19
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Maintenance DB"
         Height          =   255
         Index           =   17
         Left            =   6600
         TabIndex        =   18
         Top             =   2625
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Reminder"
         Height          =   255
         Index           =   20
         Left            =   9720
         TabIndex        =   17
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Broadcast"
         Height          =   255
         Index           =   19
         Left            =   9720
         TabIndex        =   16
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Upload Data Customer"
         Height          =   255
         Index           =   15
         Left            =   6585
         TabIndex        =   15
         Top             =   1965
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Transfer Data"
         Height          =   255
         Index           =   14
         Left            =   6600
         TabIndex        =   14
         Top             =   1620
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Distribusi Data"
         Height          =   255
         Index           =   13
         Left            =   6600
         TabIndex        =   13
         Top             =   1275
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Approval Additional Number"
         Height          =   255
         Index           =   12
         Left            =   6600
         TabIndex        =   12
         Top             =   945
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Report Payment"
         Height          =   255
         Index           =   11
         Left            =   4095
         TabIndex        =   11
         Top             =   2970
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Report Remarks Last Day"
         Height          =   255
         Index           =   8
         Left            =   4095
         TabIndex        =   10
         Top             =   1980
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Result Report"
         Height          =   255
         Index           =   7
         Left            =   4095
         TabIndex        =   9
         Top             =   1650
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Distribute Report"
         Height          =   255
         Index           =   6
         Left            =   4095
         TabIndex        =   8
         Top             =   1305
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Status Call"
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   7
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Admin"
         Height          =   255
         Index           =   3
         Left            =   2400
         TabIndex        =   6
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Supervisor"
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   5
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Agent"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   4
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Set Password"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form_master_privilage.frx":0000
         Left            =   720
         List            =   "Form_master_privilage.frx":000D
         TabIndex        =   2
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Menu Utama"
         Height          =   255
         Left            =   9720
         TabIndex        =   25
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Tools"
         Height          =   255
         Left            =   6600
         TabIndex        =   24
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Report"
         Height          =   255
         Left            =   4080
         TabIndex        =   23
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Officer"
         Height          =   255
         Left            =   2400
         TabIndex        =   22
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Master"
         Height          =   255
         Left            =   2040
         TabIndex        =   21
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "File"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form_master_privilage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim menus As String
Dim menu1 As String

Private Sub lists(c As Integer)
    Dim s
    's = Array("MnFile(3)", "mnagent", "mntl", "mnadmin", "mnNact", "mnrdistribut", "mnrresult", "rrld", "mnrpayment", "tigaA", "mndistribut", "mndpc", "mnrecycle", "nmuploadcustomer", "nmuploadpayment", "mnPO", "mntd", "mnaddclient", "mnLDS", "mnMPD", "mnais", "mntarikremarks", "mnmaintenancedb", "SSCommand1(0)", "SSCommand1(10)", "SSCommand1(8)")
    s = Array("MnFile(3)", "mnagent", "mntl", "mnadmin", "mnNact", "mnhu", "mnrdistribut", "mnrresult", "mnra", "mnrhd", "mnrh", "mnrpayment", "tigaA", "mndistribut", "mnrecycle", "nmuploadcustomer", "mnADS", "mnmaintenancedb", "SSCommand1(0)", "SSCommand1(10)", "SSCommand1(8)")
    menus = s(c)
End Sub

Private Sub createtbl()
    CMDSQL = "select * from information_schema.columns  where table_name = 'tbl_privilage'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If M_objrs.RecordCount = 0 Then
        M_OBJCONN.Execute "create table tbl_privilage (id serial not null, menu varchar, status smallint, users varchar);"
    End If
End Sub

Private Sub inserttbl()
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    
    M_OBJCONN.Execute "delete from tbl_privilage where users = '" & Combo1.text & "'"
    
    STRSQL = "select * from tbl_privilage where users = '" & Combo1.text & "' order by id"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If rs.RecordCount = 0 Then
        'For i = 0 To 25
        For i = 0 To 20
            rs.AddNew
            Call lists(i)
            rs!Menu = menus
            sets = Check1(i).Value
            rs!STATUS = cnull(sets)
            rs!Users = Combo1.text
            rs.update
        Next i
    Else
        'For i = 0 To 25
        For i = 0 To 20
            Call lists(i)
            rs!Menu = menus
            sets = Check1(i).Value
            rs!STATUS = cnull(sets)
            rs!Users = Combo1.text
            rs.update
        Next i
    End If
    MsgBox "Done"
    
End Sub

Private Sub Combo1_Click()
    Call shows
End Sub

Private Sub Command1_Click()
    Call inserttbl
End Sub

Private Sub shows()
    If Combo1.text <> "" Then
        'For i = 0 To 25
        For i = 0 To 20
            Check1(i).Value = 0
        Next i
        
        STRSQL = "select * from tbl_privilage where users = '" & Combo1.text & "' order by id"
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        If rs.RecordCount > 0 Then
            'For i = 0 To 25
            For i = 0 To 20
                Check1(i).Value = rs!STATUS
                rs.MoveNext
            Next i
        End If
    End If
End Sub

Private Sub Form_Load()
    Call createtbl
End Sub
