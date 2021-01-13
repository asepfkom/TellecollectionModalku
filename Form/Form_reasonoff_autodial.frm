VERSION 5.00
Begin VB.Form Form_off_dial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reason Off Auto Dial"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Height          =   375
      Left            =   3960
      Picture         =   "Form_reasonoff_autodial.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1350
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   2
      Left            =   1110
      TabIndex        =   2
      Top             =   960
      Width           =   4395
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Reason"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   2
      Left            =   180
      TabIndex        =   3
      Top             =   990
      Width           =   1725
   End
   Begin VB.Image Image3 
      Height          =   18630
      Index           =   2
      Left            =   0
      Picture         =   "Form_reasonoff_autodial.frx":0626
      Top             =   780
      Width           =   26295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Reason Off Auto Dial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   300
      Width           =   3585
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   1
      Left            =   105
      Picture         =   "Form_reasonoff_autodial.frx":7C30
      Stretch         =   -1  'True
      Top             =   240
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   825
      Left            =   -1890
      Picture         =   "Form_reasonoff_autodial.frx":873A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19980
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Setup Admin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   0
      Left            =   570
      TabIndex        =   0
      Top             =   300
      Width           =   3585
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   0
      Left            =   120
      Picture         =   "Form_reasonoff_autodial.frx":DBA5
      Stretch         =   -1  'True
      Top             =   240
      Width           =   420
   End
End
Attribute VB_Name = "Form_off_dial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
If Combo1(2).Text = "" Then
MsgBox "Anda Harus Set Reason calll off on dial", vbInformation + vbOKOnly, App.Title
Exit Sub
End If



STRSQL = " update tbllog_autodial_activity set stopcall =now(), keterangan='" + Combo1(2).Text + "'"
STRSQL = STRSQL + " where startcall is not null and   agent ='" + MDIForm1.txtUserName.Text + "' and   stopcall is null"
M_OBJCONN.Execute (STRSQL)


MsgBox "Data telah disimpan", vbInformation + vbOKOnly
Form_searching.stop_call.Enabled = True
Form_searching.SSCommand2.Enabled = False

F_AutoDial = False
Unload Me
End Sub
Private Sub Combo1_DropDown(Index As Integer)
Select Case Index
Case 2
sstrsql = "select * from tblreason_autodial "
Set M_OBJRS = New ADODB.Recordset
    M_OBJRS.CursorLocation = adUseClient
    M_OBJRS.Open sstrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    Combo1(2).Clear
    While Not M_OBJRS.EOF
        'cmbcampaigncode.AddItem IIf(IsNull(M_OBJRS!tbldatasource_campaign_code), "", M_OBJRS!tbldatasource_campaign_code) & "!" & IIf(IsNull(M_OBJRS!tbldatasource_keterangan), "", M_OBJRS!tbldatasource_keterangan)
        Combo1(2).AddItem IIf(IsNull(M_OBJRS!keterangan), "", M_OBJRS!keterangan)
        M_OBJRS.MoveNext
    Wend
Set M_OBJRS = Nothing
End Select

End Sub
