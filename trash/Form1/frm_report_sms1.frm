VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frm_report_sms1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report SMS"
   ClientHeight    =   3165
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   11400
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Kriteria Penarikan"
      Height          =   885
      Left            =   5610
      TabIndex        =   12
      Top             =   600
      Width           =   5655
      Begin TDBDate6Ctl.TDBDate TDBDate1 
         Height          =   315
         Index           =   1
         Left            =   2550
         TabIndex        =   13
         Top             =   360
         Width           =   1440
         _Version        =   65536
         _ExtentX        =   2540
         _ExtentY        =   556
         Calendar        =   "frm_report_sms1.frx":0000
         Caption         =   "frm_report_sms1.frx":0118
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm_report_sms1.frx":0184
         Keys            =   "frm_report_sms1.frx":01A2
         Spin            =   "frm_report_sms1.frx":0200
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
         Format          =   "dd-mmm-yyyy"
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
         Text            =   "__-___-____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   37468
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate TDBDate1 
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   14
         Top             =   360
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   556
         Calendar        =   "frm_report_sms1.frx":0228
         Caption         =   "frm_report_sms1.frx":0340
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm_report_sms1.frx":03AC
         Keys            =   "frm_report_sms1.frx":03CA
         Spin            =   "frm_report_sms1.frx":0428
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
         Format          =   "dd-mmm-yyyy"
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
         Text            =   "__-___-____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   37468
         CenturyMode     =   0
      End
      Begin TDBTime6Ctl.TDBTime DTimeLastCall 
         Height          =   300
         Index           =   0
         Left            =   1440
         TabIndex        =   15
         Top             =   375
         Width           =   885
         _Version        =   65536
         _ExtentX        =   1561
         _ExtentY        =   529
         Caption         =   "frm_report_sms1.frx":0450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frm_report_sms1.frx":04BC
         Spin            =   "frm_report_sms1.frx":050C
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "hh:nn:ss"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "hh:nn:ss"
         HighlightText   =   0
         Hour12Mode      =   1
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxTime         =   0.999988425925926
         MidnightMode    =   0
         MinTime         =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__:__:__"
         ValidateMode    =   0
         ValueVT         =   2010382337
         Value           =   2.12482692446619E-314
      End
      Begin TDBTime6Ctl.TDBTime DTimeLastCall 
         Height          =   300
         Index           =   1
         Left            =   3975
         TabIndex        =   16
         Top             =   360
         Width           =   885
         _Version        =   65536
         _ExtentX        =   1561
         _ExtentY        =   529
         Caption         =   "frm_report_sms1.frx":0534
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frm_report_sms1.frx":05A0
         Spin            =   "frm_report_sms1.frx":05F0
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "hh:nn"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "hh:nn"
         HighlightText   =   0
         Hour12Mode      =   1
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxTime         =   0.99999
         MidnightMode    =   0
         MinTime         =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__:__"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   0.870289351851852
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   360
      Index           =   1
      Left            =   10140
      TabIndex        =   1
      Top             =   1575
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Preview"
      Height          =   360
      Index           =   0
      Left            =   8925
      TabIndex        =   0
      Top             =   1575
      Width           =   1125
   End
   Begin Crystal.CrystalReport RPT 
      Left            =   6135
      Top             =   3975
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2475
      Left            =   60
      TabIndex        =   2
      Top             =   495
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   4366
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   5460
      TabIndex        =   3
      Top             =   2730
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   3735
      Left            =   0
      Top             =   450
      Width           =   11385
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   5520
      TabIndex        =   11
      Top             =   0
      Width           =   5745
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      Height          =   300
      Index           =   4
      Left            =   8850
      TabIndex        =   10
      Top             =   1965
      Width           =   270
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date :"
      Height          =   300
      Index           =   5
      Left            =   5595
      TabIndex        =   9
      Top             =   1980
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From Batch :"
      Height          =   300
      Index           =   0
      Left            =   5580
      TabIndex        =   8
      Top             =   1650
      Width           =   930
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      Height          =   255
      Index           =   1
      Left            =   8850
      TabIndex        =   7
      Top             =   2355
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comparator Date :"
      Height          =   420
      Index           =   3
      Left            =   5640
      TabIndex        =   6
      Top             =   2445
      Width           =   930
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      Height          =   300
      Index           =   7
      Left            =   8880
      TabIndex        =   5
      Top             =   2565
      Width           =   270
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   5
      Left            =   0
      Picture         =   "frm_report_sms1.frx":0618
      Stretch         =   -1  'True
      Top             =   60
      Width           =   420
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "Report SMS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   540
      TabIndex        =   4
      Top             =   30
      Width           =   3915
   End
   Begin VB.Image Image2 
      Height          =   435
      Index           =   8
      Left            =   0
      Picture         =   "frm_report_sms1.frx":1122
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20700
   End
End
Attribute VB_Name = "frm_report_sms1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Select Case Index
    Case 0
        Select Case ListView1.SelectedItem.Text
      Case 1
        Dim OBJRS As New ADODB.Recordset
        Dim objracess As New ADODB.Recordset
        Set OBJRS = New ADODB.Recordset
        OBJRS.CursorLocation = adUseClient
        OBJRS.Open " select DISTINCT(OPERATOR),BIAYA from tbl_tarif ", M_OBJCONN1, adOpenDynamic, adLockOptimistic
        While Not OBJRS.EOF
         TRF = IIf(IsNull(OBJRS!BIAYA), "", OBJRS!BIAYA)
         JNSOPERATOR = Trim(IIf(IsNull(OBJRS!operator), "", OBJRS!operator))
         If JNSOPERATOR = "LAINNYA" Then
            M_OBJCONN1.Execute " update sentitems set tarif =" + CStr(TRF) + " where TARIF IS NULL AND  substring(destinationnumber,1,4) not  in (SELECT NO FROM TBL_TARIF WHERE OPERATOR NOT IN (SELECT OPERATOR FROM TBL_TARIF WHERE OPERATOR='" + JNSOPERATOR + "')) and substring(creatorid,1,2)<>'AW'  and substring(creatorid,1,3)<>'PIL'"
         Else
             M_OBJCONN1.Execute " update sentitems set tarif =" + CStr(TRF) + " where TARIF IS NULL AND  substring(destinationnumber,1,4)  in (select NO from tbl_tarif where operator='" + JNSOPERATOR + "') and substring(creatorid,1,2)<>'AW'  and substring(creatorid,1,3)<>'PIL'"
         End If
           OBJRS.MoveNext
       Wend
       
       Set OBJRS = Nothing
       
       
       M_RPTCONN.Execute "delete from tbl_gammu "
       Set objracess = New ADODB.Recordset
       objracess.CursorLocation = adUseClient
       objracess.Open "select * from tbl_gammu", M_RPTCONN, adOpenDynamic, adLockOptimistic
       If TDBDate1(0).ValueIsNull And TDBDate1(1).ValueIsNull Then
            STRSQL = "select sendingdatetime as tglkirim, destinationnumber as notujuan,textdecoded as isisms,status, split_part(creatorid,'-',1) as custid,split_part(creatorid,'-',2) as agent,split_part(creatorid,'-',3) as nmcust,tarif from sentitems where substring(creatorid,1,2)<>'AW'  and substring(creatorid,1,3)<>'PIL' "
       Else
            STRSQL = "select sendingdatetime as tglkirim, destinationnumber as notujuan,textdecoded as isisms,status, split_part(creatorid,'-',1) as custid,split_part(creatorid,'-',2) as agent,split_part(creatorid,'-',3) as nmcust,tarif from sentitems where substring(creatorid,1,2)<>'AW'  and substring(creatorid,1,3)<>'PIL' and sendingdatetime between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  "
       End If
       
        Set OBJRS = New ADODB.Recordset
        OBJRS.CursorLocation = adUseClient
        OBJRS.Open STRSQL, M_OBJCONN1, adOpenDynamic, adLockOptimistic
        
        ProgressBar1.Max = OBJRS.RecordCount + 1



        While Not OBJRS.EOF
            DoEvents
            ProgressBar1.Value = OBJRS.Bookmark
            objracess.AddNew
            objracess!tglkirim = CStr(IIf(IsNull(OBJRS!tglkirim), Null, Format(OBJRS!tglkirim, "yyyy-mm-dd hh:mm:ss")))
            objracess!notujuan = IIf(IsNull(OBJRS!notujuan), "", OBJRS!notujuan)
            objracess!isipesan = IIf(IsNull(OBJRS!isisms), "", Replace(OBJRS!isisms, "'", ""))
            objracess!STATUS = IIf(IsNull(OBJRS!STATUS), "", OBJRS!STATUS)
            objracess!CustId = Trim(Replace(IIf(IsNull(OBJRS!CustId), "", OBJRS!CustId), "AW", ""))
            objracess!agent = Trim(IIf(IsNull(OBJRS!agent), "", OBJRS!agent))
            objracess!nmcust = Trim(IIf(IsNull(OBJRS!nmcust), "", OBJRS!nmcust))
            objracess!tarif = Trim(IIf(IsNull(OBJRS!tarif), 0, OBJRS!tarif))
            objracess.update
            OBJRS.MoveNext
        Wend
            WaitSecs (2)
            RPT.Reset
            'RPT.Formulas(1) = "@User = totext('" + CStr(mdiform1.txtusername.text) + "')"
            'RPT.Formulas(1) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            'RPT.Formulas(2) = "@TglAkhirShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\Rptsentitems.rpt"
            Call SHOW_PRN
            
    Case 2
       M_RPTCONN.Execute "delete from tbl_gammu"
       Set objracess = New ADODB.Recordset
       objracess.CursorLocation = adUseClient
       objracess.Open "select * from tbl_gammu", M_RPTCONN, adOpenDynamic, adLockOptimistic
       If TDBDate1(0).ValueIsNull And TDBDate1(1).ValueIsNull Then
            STRSQL = STRSQL + " SELECT * FROM ("
            STRSQL = STRSQL + " select date(receivingdatetime) as tglmasuk ,tblbaru.textdecoded as isisms,notlp as nomasuk,split_part(creatorid,'-',1) as custid,split_part(creatorid,'-',2) as agent,split_part(creatorid,'-',3) as nmcust  from ( "
            STRSQL = STRSQL + " select * from ( "
            STRSQL = STRSQL + " select replace(sendernumber,'+62','0') as notlp,textdecoded,receivingdatetime ,custid from inbox) as a ) tblbaru right join "
            STRSQL = STRSQL + " (select * from sentitems where substring(creatorid,1,2)<>'AW'  and substring(creatorid,1,3)<>'PIL') c on  c.destinationnumber=tblbaru.notlp where notlp  is not null "
            STRSQL = STRSQL + "   ) AS TBLNEW GROUP BY TGLMASUK,ISISMS,NOMASUK,CUSTID,AGENT"
       Else
            STRSQL = STRSQL + "  SELECT TGLMASUK,ISISMS,NOMASUK,CUSTID,AGENT FROM ("
            STRSQL = STRSQL + " select date(receivingdatetime) as tglmasuk ,tblbaru.textdecoded as isisms,notlp as nomasuk,split_part(creatorid,'-',1) "
            STRSQL = STRSQL + " as custid,split_part(creatorid,'-',2) as agent  from (  select * from "
            STRSQL = STRSQL + " (  select replace(sendernumber,'+62','0') as notlp,textdecoded,receivingdatetime ,custid from inbox) as a ) tblbaru right join "
            STRSQL = STRSQL + " (select * from sentitems where substring(creatorid,1,2)<>'AW'  and substring(creatorid,1,3)<>'PIL') c on "
            STRSQL = STRSQL + " c.destinationnumber=tblbaru.notlp where notlp  is not null and receivingdateTIME    between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "') AS TBLNEW"
            STRSQL = STRSQL + " GROUP BY NOMASUK,TGLMASUK,ISISMS,NOMASUK,CUSTID,AGENT"
  
'            STRSQL = STRSQL + " select date(receivingdatetime) as tglmasuk ,tblbaru.textdecoded as isisms,notlp as nomasuk,split_part(creatorid,'-',1) as custid,split_part(creatorid,'-',2) as agent,split_part(creatorid,'-',3) as nmcust  from ( "
'            STRSQL = STRSQL + " select * from ( "
'            STRSQL = STRSQL + " select replace(sendernumber,'+62','0') as notlp,textdecoded,receivingdatetime ,custid from inbox) as a ) tblbaru right join"
'            STRSQL = STRSQL + " (select * from sentitems where substring(creatorid,1,2)<>'AW') c on  c.destinationnumber=tblbaru.notlp where notlp  is not null and receivingdateTIME   between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "' "
            
       End If
       
       
        Set OBJRS = New ADODB.Recordset
        OBJRS.CursorLocation = adUseClient
        OBJRS.Open STRSQL, M_OBJCONN1, adOpenDynamic, adLockOptimistic
        ProgressBar1.Max = OBJRS.RecordCount + 1
        While Not OBJRS.EOF
            objracess.AddNew
            DoEvents
            ProgressBar1.Value = OBJRS.Bookmark
            objracess!TGLMASUK = CStr(IIf(IsNull(OBJRS!TGLMASUK), Null, Format(OBJRS!TGLMASUK, "yyyy-mm-dd hh:mm:ss")))
            objracess!notujuan = IIf(IsNull(OBJRS!NOMASUK), "", OBJRS!NOMASUK)
            objracess!isipesan = IIf(IsNull(OBJRS!isisms), "", Replace(OBJRS!isisms, "'", ""))
            objracess!CustId = Trim(Replace(IIf(IsNull(OBJRS!CustId), "", OBJRS!CustId), "AW", ""))
            objracess!agent = Trim(IIf(IsNull(OBJRS!agent), "", OBJRS!agent))
           
            objracess.update
            OBJRS.MoveNext
        Wend
        Set OBJRS = Nothing
            WaitSecs (2)
            RPT.Reset
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptINBOX.rpt"
            Call SHOW_PRN
       Case 3
        RptSentitemsUploadExcel
          WaitSecs (2)
          RPT.Reset
          'RPT.Formulas(1) = "@User = totext('" + CStr(mdiform1.txtusername.text) + "')"
            'RPT.Formulas(1) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            'RPT.Formulas(2) = "@TglAkhirShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\Rptsentitems.rpt"
            Call SHOW_PRN
       Case 4
        InboxSmsBlastExcel
        WaitSecs (2)
        RPT.Reset
        RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptINBOX.rpt"
        Call SHOW_PRN
  End Select
   
            
Case 1
        Unload Me
End Select


End Sub
Private Sub SHOW_PRN()
    RPT.RetrieveDataFiles
    RPT.WindowLeft = 0
    RPT.WindowTop = 0
    RPT.WindowState = crptMaximized
    RPT.WindowShowPrintBtn = True
    RPT.WindowShowRefreshBtn = True
    RPT.WindowShowSearchBtn = True
    RPT.WindowShowPrintSetupBtn = True
    RPT.WindowControls = True
    RPT.PrintReport
    RPT.Reset
 'RPT.Action = 1
 
     
End Sub

Private Sub Form_Load()
Call header
Set listitem = ListView1.ListItems.ADD(, , "1")
    listitem.SubItems(1) = "Report Sentitems"
'Set listitem = ListView1.ListItems.ADD(, , "")
'    listitem.SubItems(1) = "------------------------------------------"
Set listitem = ListView1.ListItems.ADD(, , "2")
        listitem.SubItems(1) = "Report Inbox"
        
Set listitem = ListView1.ListItems.ADD(, , "3")
        listitem.SubItems(1) = "Report Sentitems SMS Upload"
        
Set listitem = ListView1.ListItems.ADD(, , "4")
        listitem.SubItems(1) = "Report Inbox SMS Upload"

End Sub
Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "No", 4 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Report", 50 * TXT
End Sub

Private Sub RptSentitemsUploadExcel()
    Dim OBJRS As New ADODB.Recordset
    Dim objracess As New ADODB.Recordset
        
        Set OBJRS = New ADODB.Recordset
        OBJRS.CursorLocation = adUseClient
        OBJRS.Open " select DISTINCT(OPERATOR),BIAYA from tbl_tarif ", M_OBJCONN1, adOpenDynamic, adLockOptimistic
        
        While Not OBJRS.EOF
         TRF = IIf(IsNull(OBJRS!BIAYA), "", OBJRS!BIAYA)
         JNSOPERATOR = Trim(IIf(IsNull(OBJRS!operator), "", OBJRS!operator))
         If JNSOPERATOR = "LAINNYA" Then
            M_OBJCONN1.Execute " update sentitems set tarif =" + CStr(TRF) + " where TARIF IS NULL AND  substring(destinationnumber,1,4) not  in (SELECT NO FROM TBL_TARIF WHERE OPERATOR NOT IN (SELECT OPERATOR FROM TBL_TARIF WHERE OPERATOR='" + JNSOPERATOR + "')) and substring(creatorid,1,2)<>'AW'"
         Else
             M_OBJCONN1.Execute " update sentitems set tarif =" + CStr(TRF) + " where TARIF IS NULL AND  substring(destinationnumber,1,4)  in (select NO from tbl_tarif where operator='" + JNSOPERATOR + "') and substring(creatorid,1,2)<>'AW'"
         End If
           OBJRS.MoveNext
       Wend
       
       Set OBJRS = Nothing
       
       
       M_RPTCONN.Execute "delete from tbl_gammu "
       Set objracess = New ADODB.Recordset
       objracess.CursorLocation = adUseClient
       objracess.Open "select * from tbl_gammu", M_RPTCONN, adOpenDynamic, adLockOptimistic
       If TDBDate1(0).ValueIsNull And TDBDate1(1).ValueIsNull Then
            STRSQL = "select sendingdatetime as tglkirim, destinationnumber as notujuan,textdecoded as isisms,status, split_part(creatorid,'-',1) as custid,split_part(creatorid,'-',2) as agent,split_part(creatorid,'-',4) as nmcust,tarif from sentitems where substring(creatorid,1,2)<>'AW' and split_part(creatorid,'-',2)='BlastExcelCard'  "
       Else
            STRSQL = "select sendingdatetime as tglkirim, destinationnumber as notujuan,textdecoded as isisms,status, split_part(creatorid,'-',1) as custid,split_part(creatorid,'-',2) as agent,split_part(creatorid,'-',4) as nmcust,tarif from sentitems where substring(creatorid,1,2)<>'AW' and sendingdatetime between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  and split_part(creatorid,'-',2)='BlastExcelCard' "
       End If
       
        Set OBJRS = New ADODB.Recordset
        OBJRS.CursorLocation = adUseClient
        OBJRS.Open STRSQL, M_OBJCONN1, adOpenDynamic, adLockOptimistic
        
        ProgressBar1.Max = OBJRS.RecordCount + 1



        While Not OBJRS.EOF
            DoEvents
            ProgressBar1.Value = OBJRS.Bookmark
            objracess.AddNew
            objracess!tglkirim = CStr(IIf(IsNull(OBJRS!tglkirim), Null, Format(OBJRS!tglkirim, "yyyy-mm-dd hh:mm:ss")))
            objracess!notujuan = IIf(IsNull(OBJRS!notujuan), "", OBJRS!notujuan)
            objracess!isipesan = IIf(IsNull(OBJRS!isisms), "", Replace(OBJRS!isisms, "'", ""))
            objracess!STATUS = IIf(IsNull(OBJRS!STATUS), "", OBJRS!STATUS)
            objracess!CustId = Trim(Replace(IIf(IsNull(OBJRS!CustId), "", OBJRS!CustId), "AW", ""))
            objracess!agent = Trim(IIf(IsNull(OBJRS!agent), "No Agent", OBJRS!agent))
            objracess!nmcust = Trim(IIf(IsNull(OBJRS!nmcust), "", OBJRS!nmcust))
            objracess!tarif = Trim(IIf(IsNull(OBJRS!tarif), 0, OBJRS!tarif))
            objracess.update
            OBJRS.MoveNext
        Wend
           
End Sub

Private Sub InboxSmsBlastExcel()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
    
     If TDBDate1(0).ValueIsNull And TDBDate1(1).ValueIsNull Then
         CMDSQL = "select distinct tglmasuk,isisms,pengirim,custid from ("
         CMDSQL = CMDSQL + "select i.receivingdatetime as tglmasuk,i.textdecoded as isisms, i.sendernumber as pengirim, "
         CMDSQL = CMDSQL + " split_part(s.creatorid,'-',1) as custid "
         CMDSQL = CMDSQL + "from inbox as i,sentitems as s where "
         CMDSQL = CMDSQL + " split_part(s.creatorid,'-',2)='BlastExcelCard' and "
         CMDSQL = CMDSQL + " replace(i.sendernumber,'+62','0')=s.destinationnumber) as a"
     Else
         CMDSQL = "select distinct tglmasuk,isisms,pengirim,custid from ("
         CMDSQL = CMDSQL + "select i.receivingdatetime as tglmasuk,i.textdecoded as isisms, i.sendernumber as pengirim, "
         CMDSQL = CMDSQL + " split_part(s.creatorid,'-',1) as custid "
         CMDSQL = CMDSQL + "from inbox as i,sentitems as s where "
         CMDSQL = CMDSQL + " split_part(s.creatorid,'-',2)='BlastExcelCard' and "
         CMDSQL = CMDSQL + " replace(i.sendernumber,'+62','0')=s.destinationnumber and "
         CMDSQL = CMDSQL + " i.receivingdatetime between '"
         CMDSQL = CMDSQL + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "') as a"
     End If
     
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
    If M_objrs.RecordCount <> 0 Then
        M_RPTCONN.Execute "delete from tbl_gammu "
        ProgressBar1.Max = M_objrs.RecordCount
        While Not M_objrs.EOF
            ProgressBar1.Value = M_objrs.Bookmark
            CMDSQL = "insert into tbl_gammu (tglmasuk,isipesan,custid,notujuan) values ('"
            CMDSQL = CMDSQL + CStr(Format(M_objrs("tglmasuk"), "yyyy-mm-dd hh:mm:ss")) + "','"
            CMDSQL = CMDSQL + IIf(IsNull(M_objrs("isisms")), "", Replace(M_objrs("isisms"), "'", "")) + "','"
            CMDSQL = CMDSQL + M_objrs("custid") + "','"
            CMDSQL = CMDSQL + M_objrs("pengirim") + "')"
            M_RPTCONN.Execute CMDSQL
            M_objrs.MoveNext
        Wend
    End If
    Set M_objrs = Nothing
     
End Sub
