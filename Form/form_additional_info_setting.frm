VERSION 5.00
Begin VB.Form form_additional_info_setting 
   Caption         =   "Form Additional Info Setting"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5655
   LinkTopic       =   "Form2"
   ScaleHeight     =   3975
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Caption         =   "Additional Info"
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.CommandButton Command1 
         Caption         =   "Save Settting"
         Height          =   375
         Left            =   3840
         TabIndex        =   19
         Top             =   3480
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3720
         TabIndex        =   18
         Top             =   120
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         ItemData        =   "form_additional_info_setting.frx":0000
         Left            =   1800
         List            =   "form_additional_info_setting.frx":0002
         TabIndex        =   16
         Top             =   3000
         Width           =   3495
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         ItemData        =   "form_additional_info_setting.frx":0004
         Left            =   1800
         List            =   "form_additional_info_setting.frx":0006
         TabIndex        =   15
         Top             =   2640
         Width           =   3495
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         ItemData        =   "form_additional_info_setting.frx":0008
         Left            =   1800
         List            =   "form_additional_info_setting.frx":000A
         TabIndex        =   14
         Top             =   2280
         Width           =   3495
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         ItemData        =   "form_additional_info_setting.frx":000C
         Left            =   1800
         List            =   "form_additional_info_setting.frx":000E
         TabIndex        =   13
         Top             =   1920
         Width           =   3495
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         ItemData        =   "form_additional_info_setting.frx":0010
         Left            =   1800
         List            =   "form_additional_info_setting.frx":0012
         TabIndex        =   12
         Top             =   1560
         Width           =   3495
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         ItemData        =   "form_additional_info_setting.frx":0014
         Left            =   1800
         List            =   "form_additional_info_setting.frx":0016
         TabIndex        =   11
         Top             =   1200
         Width           =   3495
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         ItemData        =   "form_additional_info_setting.frx":0018
         Left            =   1800
         List            =   "form_additional_info_setting.frx":001A
         TabIndex        =   10
         Top             =   840
         Width           =   3495
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         ItemData        =   "form_additional_info_setting.frx":001C
         Left            =   1800
         List            =   "form_additional_info_setting.frx":0038
         TabIndex        =   9
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Height          =   375
         Index           =   7
         Left            =   360
         TabIndex        =   8
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Height          =   375
         Index           =   6
         Left            =   360
         TabIndex        =   7
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Height          =   375
         Index           =   5
         Left            =   360
         TabIndex        =   6
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   5
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   4
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rekan"
         Height          =   255
         Left            =   3120
         TabIndex        =   17
         Top             =   120
         Width           =   615
      End
   End
End
Attribute VB_Name = "form_additional_info_setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_DropDown(Index As Integer)
    Call combo_next(Index)
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Combo2_Click()
    Call clear
    Call checkrekansetting
End Sub

Private Sub Command1_Click()
    Call insert
End Sub

Private Sub Form_Load()
    Call createtable
    Call list_client(Combo2)
End Sub

Private Sub combo_next(min As Integer)
    For i = 0 To 6
        If Combo1(i).text <> "" And Combo1(i + 1).text = "" Then
            Combo1(i + 1).clear
            For a = 0 To 7
                c1 = Combo1(i).list(a)
                If c1 <> Combo1(i).text Then
                    Combo1(i + 1).AddItem Combo1(i).list(a)
                End If
            Next a
        End If
    Next i
End Sub

Private Sub createtable()
Dim crs As New ADODB.Recordset
Dim c_str As String
    c_str = " SELECT * From information_schema.Columns WHERE table_name='tbl_additional_info'"
    Set crs = New ADODB.Recordset
    crs.CursorLocation = adUseClient
    crs.Open c_str, M_OBJCONN, adOpenKeyset, adLockReadOnly
    
    If crs.RecordCount = 0 Then
        c_str = "create table tbl_additional_info ( id serial, users varchar, rekan varchar, caption varchar, dbname varchar, tgl date default now());"
        M_OBJCONN.Execute c_str
    End If
Set crs = Nothing
End Sub

Private Sub insert()
    If Combo2.text = "" Then
        MsgBox "Rekan harus dipilih"
        Exit Sub
    End If
    
    For i = 0 To 7
        If Text2(i).text <> "" Then
            j = i
        End If
    Next i
    
    c_str = " SELECT * From tbl_additional_info where rekan = '" & Combo2.text & "' "
    Set crs = New ADODB.Recordset
    crs.CursorLocation = adUseClient
    crs.Open c_str, M_OBJCONN, adOpenKeyset, adLockReadOnly
    
    If crs.RecordCount > 0 Then
        If MsgBox("Additional Info sebelumnya sudah ada, Ingin set ulang? ", vbYesNo + vbQuestion, "Konfirmasi") = vbYes Then
            M_OBJCONN.Execute "Delete from tbl_additional_info where rekan = '" & Combo2.text & "'"
            GoTo bawah
        End If
    Else
bawah:
        For K = 0 To j
            c_str = "insert into tbl_additional_info (users, rekan, caption, dbname) values "
            c_str = c_str & "('" & MDIForm1.TxtUsername.text & "','" & Combo2.text & "','" & Text2(K).text & "','" & Combo1(K).text & "')"
            M_OBJCONN.Execute c_str
        Next K
        MsgBox "Additional Info Rekan " & Combo2.text & " Berhasil Di-Set"
    End If
End Sub

Private Sub checkrekansetting()
    c_str = " SELECT * From tbl_additional_info where rekan = '" & Combo2.text & "' order by 1"
    Set crs = New ADODB.Recordset
    crs.CursorLocation = adUseClient
    crs.Open c_str, M_OBJCONN, adOpenKeyset, adLockReadOnly
    
    If crs.RecordCount > 0 Then
        For i = 1 To crs.RecordCount
            Text2(i - 1).text = crs!Caption
            Combo1(i - 1).text = cnull(crs!dbname)
        Next i
    End If
End Sub

Private Sub clear()
    For i = 0 To 7
        Text2(i).text = ""
        Combo1(i).text = ""
    Next i
End Sub
