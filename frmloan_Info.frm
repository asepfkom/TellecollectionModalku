VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmloan_Info 
   Caption         =   "Loan Info"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9075
   LinkTopic       =   "Form2"
   ScaleHeight     =   6870
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "INSTALMENT"
      Height          =   4485
      Left            =   135
      TabIndex        =   2
      Top             =   2280
      Width           =   8745
      Begin MSComctlLib.ListView listview_instalment 
         Height          =   4095
         Index           =   0
         Left            =   90
         TabIndex        =   3
         Top             =   225
         Width           =   8520
         _ExtentX        =   15028
         _ExtentY        =   7223
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   12632256
         BorderStyle     =   1
         Appearance      =   0
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
      Caption         =   "LOAN"
      Height          =   2190
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   8745
      Begin MSComctlLib.ListView listview_loan 
         Height          =   1845
         Index           =   2
         Left            =   90
         TabIndex        =   1
         Top             =   225
         Width           =   8520
         _ExtentX        =   15028
         _ExtentY        =   3254
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   12632256
         BorderStyle     =   1
         Appearance      =   0
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
End
Attribute VB_Name = "frmloan_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function Show_loan_info()
    Dim rs As New ADODB.Recordset
    Dim ListItem As ListItem
    rs.CursorLocation = adUseClient
    
    rs.Open "select x_loan_code,product_desc, status_loan,tenor,amount_disbursed,curbal from mgm ", M_OBJCONN, adOpenDynamic, adLockOptimistic
 
        listview_loan(2).ListItems.clear
    While Not rs.EOF
 
         Set ListItem = listview_loan(2).ListItems.ADD(, , rs.Bookmark)
            ListItem.SubItems(1) = IIf(IsNull(rs("x_loan_code")), "", rs("x_loan_code"))
            ListItem.SubItems(2) = IIf(IsNull(rs("product_desc")), "", rs("product_desc"))
            ListItem.SubItems(3) = IIf(IsNull(rs("status_loan")), "", rs("status_loan"))
            ListItem.SubItems(4) = IIf(IsNull(rs("tenor")), "", rs("tenor"))
            ListItem.SubItems(5) = IIf(IsNull(rs("amount_disbursed")), "", rs("amount_disbursed"))
            ListItem.SubItems(6) = IIf(IsNull(rs("curbal")), "", rs("curbal"))
            rs.MoveNext
         Wend
 

End Function
Private Sub HEADER_MAPPING_LOAN()
listview_loan(2).ColumnHeaders.ADD 1, , "No", 5 * TXT
    listview_loan(2).ColumnHeaders.ADD 2, , "Loan Code", 10 * TXT
    listview_loan(2).ColumnHeaders.ADD 3, , "Product", 10 * TXT
    listview_loan(2).ColumnHeaders.ADD 4, , "Status", 10 * TXT
    listview_loan(2).ColumnHeaders.ADD 5, , "Tenor", 10 * TXT
    listview_loan(2).ColumnHeaders.ADD 6, , "DIsbursement Date", 10 * TXT
    listview_loan(2).ColumnHeaders.ADD 7, , "Amount", 10 * TXT
    
End Sub


Private Sub Form_Load()
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
    
    Call HEADER_MAPPING_LOAN
End Sub
