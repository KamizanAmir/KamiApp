Private Sub MDIForm_Load()
Dim strDSN, strUID, strPWD As String
strDSN = "AICSSQLDB"
strUID = "ITAPP"
strPWD = "APP"
Set wsDB = New ADODB.Connection
wsDB.ConnectionString = "PROVIDER=MSDASQL;dsn=" & strDSN & ";uid=" & strUID & ";pwd=" & strPWD & ";"
wsDB.CursorLocation = adUseServer
wsDB.Open
Set wsDB_CLS = New ADODB.Connection
wsDB_CLS.ConnectionString = "PROVIDER=MSDASQL;dsn=" & strDSN & ";uid=" & strUID & ";pwd=" & strPWD & ";"
wsDB_CLS.CursorLocation = adUseServer
wsDB_CLS.Open
login_id = ""
MesSysName = "AIMS (SQL)"
MesSysVer = "(v4.9.98)"
CoreVersion = 4000998
If Left(Trim(App.Path), 20) = "\\aicwksvr2016\deploy" Then
    MsgBox "Error: System running from Deploy folder.", vbExclamation, "Incorrect Path Setup"
    Unload Me
    End
End If
Dim m_hWnd As Long
m_hWnd = FindWindow(MesSysName & " (*")
If m_hWnd > 0 Then
    MsgBox "Please close all AIMS program before starting a new one.", vbExclamation, "AIMS already started"
    Unload Me
    End
    Exit Sub
End If
Me.Caption = MesSysName & " " & MesSysVer
   User_Login.Show
   If Me.StatusBar1.Panels(1).Text = "" Then Call Timer1_disp_Timer
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
Call exitlogic(1)
End Sub
Private Sub Timer1_disp_Timer()
Call svrdatetime(dsp_serverdate, dsp_servertime, dsp_shifttype, dsp_proddate)
Me.StatusBar1.Panels(1).Text = "Server DateTime : " & dsp_serverdate & " (1 min. refresh)"
If Trim(dsp_shifttype) = "1" Then
    Me.StatusBar1.Panels(2).Text = "Production Date : " & dsp_proddate & " Shift : " & dsp_shifttype & "/Night"
Else
    Me.StatusBar1.Panels(2).Text = "Production Date : " & dsp_proddate & " Shift : " & dsp_shifttype & "/Morning"
End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "EXIT":
            Call exitlogic(2)
        Case "LOGIN":
            Call exitlogic(3)
    End Select
End Sub
Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
 chkupdate
    Select Case ButtonMenu.Key
        Case "SCOM":
            Call DisableAllMenu
            Setup_Master.Show
        Case "SPRD":
            Call DisableAllMenu
            Setup_ProdMaster.Show
        Case "SRTE":
            Call DisableAllMenu
            Setup_Route.Show
        Case "STST":
            Call DisableAllMenu
            Setup_TestMaster.Show
        Case "DBLK":
            Call DisableAllMenu
            Setup_BlockDevice.Show
        Case "NBPS":
            Call DisableAllMenu
            Amphenol_Port.Show
        Case "QAIN":
            Call DisableAllMenu
            Prd_QA_Screen.Show
        Case "BGRD":
            Call DisableAllMenu
            Prd_QA_BackGrind.Show
        Case "ARMS":
            Call DisableAllMenu
            Prd_Rms_Assy.Show
        Case "CRLT":
            Call DisableAllMenu
            Db_CreateLot.Show
        Case "DSPL":
            Call DisableAllMenu
            Db_SplitLot.Show
        Case "PAPT":
            Call DisableAllMenu
            Db_PrintApt.Show
        Case "RAPT":
            Call DisableAllMenu
            Db_ReprintApt.Show
        Case "DMVO":
            Call DisableAllMenu
            Db_Mvou.Show
        Case "DLST":
            Call DisableAllMenu
            Db_MVList.Show
        Case "TERM":
            Call DisableAllMenu
            Prd_Terminate.lbl_title.Caption = "TERMINATE LOT"
            Prd_Terminate.Show
        Case "RTER":
            Call DisableAllMenu
            Rvse_Trn.lbl_title.Caption = "REVERSE TERMINATE"
            Rvse_Trn.Show
        Case "OSTR":
            Call DisableAllMenu
            Prd_Terminate.lbl_title.Caption = "TERMINATE OSRAM LOT"
            Prd_Terminate.Show
        Case "ZRLT"
            Call DisableAllMenu
            Prd_Zerorize.Show
        Case "PREL"
            Call DisableAllMenu
            Db_PreLoad.Show
        Case "CRST":
            Call DisableAllMenu
            ST_Db_CreateLot.Show
        Case "SPST":
            Call DisableAllMenu
            ST_Db_SplitLot.Show
        Case "PTST":
            Call DisableAllMenu
            ST_Db_PrintApt.Show
        Case "MVST":
            Call DisableAllMenu
            ST_Db_Mvou.Show
        Case "REST":
            Call DisableAllMenu
            ST_Db_ReprintApt.Show
        Case "LPST":
            Call DisableAllMenu
            ST_DB_List_Preload.Show
        Case "MVOU":
            Call DisableAllMenu
            Prd_Mvou.Show
        Case "BYLT":
            Call DisableAllMenu
            Prd_CV.Show
        Case "SPLT":
            Call DisableAllMenu
            Prd_Split.Show
        Case "MRLT":
            Call DisableAllMenu
            Prd_MultiMerge.Show
        Case "HOLD":
            Call DisableAllMenu
            Prd_Hold.Show
        Case "PROC":
            Call DisableAllMenu
            Prd_TestRoute.Show
        Case "TTRV":
            Call DisableAllMenu
            Prd_TestTraveller.Show
        Case "RMST":
            Call DisableAllMenu
            Prd_RMSTT.Show
        Case "BSPL":
            Call DisableAllMenu
            Prd_BinSplit.Show
        Case "SSPL":
            Call DisableAllMenu
        Prd_FT_SAMPLE_SPLIT.Show
        Case "TRNS":
            Call DisableAllMenu
            Prd_Transfer.Show
        Case "KTIN"
        Case "KTOU"
            Call DisableAllMenu
            Kitting_Wdraw.Show
        Case "MATCNT":
            Call DisableAllMenu
            Prd_StockCount.Show
        Case "MCPC"
            Call DisableAllMenu
            Microchip_FOL_Combine.Show
        Case "SRCV":
            Call DisableAllMenu
            Sub_Receive.Show
        Case "DRAW":
            Call DisableAllMenu
            Sub_Withdraw.Show
        Case "SHLT":
            Call DisableAllMenu
            Prd_Shlt.Show
        Case "RSHT":
            Call DisableAllMenu
            Rvse_Trn.lbl_title.Caption = "REVERSE SHLT"
            Rvse_Trn.Show
        Case "FHLD":
            Call DisableAllMenu
            Prd_Future_Hold.Show
        Case "RLSE":
            Call DisableAllMenu
            Prd_Release.Show
        Case "XDOC":
            Call DisableAllMenu
            Prd_Attach_Doc.Show
        Case "CGTP":                                                
            Call DisableAllMenu
            Prd_Chg_LotInfo.lbl_title.Caption = "CHANGE TEST PROGRAM"
            Prd_Chg_LotInfo.txt_transtype.Text = "CGTP"
            Prd_Chg_LotInfo.new_testprog.Enabled = True
            Prd_Chg_LotInfo.new_testprog.BackColor = &HC0FFFF
            Prd_Chg_LotInfo.Show
        Case "CGRT":                                                
            Call DisableAllMenu
            Prd_Chg_LotInfo.lbl_title.Caption = "CHANGE PROCESS ROUTE"
            Prd_Chg_LotInfo.txt_transtype.Text = "CGRT"
            Prd_Chg_LotInfo.new_processroute.Enabled = True
            Prd_Chg_LotInfo.new_processroute.BackColor = &HC0FFFF
            Prd_Chg_LotInfo.new_routetype.Enabled = True
            Prd_Chg_LotInfo.new_routetype.BackColor = &HC0FFFF
            Prd_Chg_LotInfo.Show
        Case "CGTD":                                                
            Call DisableAllMenu
            Prd_Chg_LotInfo.lbl_title.Caption = "CHANGE TARGET DEVICE"
            Prd_Chg_LotInfo.txt_transtype.Text = "CGTD"
            Prd_Chg_LotInfo.new_targetdevice.Enabled = True
            Prd_Chg_LotInfo.new_targetdevice.BackColor = &HC0FFFF
            Prd_Chg_LotInfo.Show
        Case "CGSF": 
            Call DisableAllMenu
            Prd_Chg_LotInfo.lbl_title.Caption = "CHANGE SHIP-FORM"
            Prd_Chg_LotInfo.txt_transtype.Text = "CGSF"
            Prd_Chg_LotInfo.new_shipform.Enabled = True
            Prd_Chg_LotInfo.new_shipform.BackColor = &HC0FFFF
            Prd_Chg_LotInfo.Show
        Case "CGSD":
            Call DisableAllMenu
            Prd_Chg_ShipDate.Show
        Case "ENGH":
            Call DisableAllMenu
            Prd_Hold_Status.Show
        Case "RMVO":
            Call DisableAllMenu
            Rvse_Trn.lbl_title.Caption = "REVERSE MVOU"
            Rvse_Trn.Show
        Case "RPRO":
            Call DisableAllMenu
            Rvse_Trn.lbl_title.Caption = "REVERSE TEST PROC"
            Rvse_Trn.Show
        Case "RSCR":
            Call DisableAllMenu
            Rvse_Trn.lbl_title.Caption = "RESCREEN"
            Rvse_Trn.Show
        Case "RVDP":
            Call DisableAllMenu
            Rvse_Trn.lbl_title.Caption = "REVERSE DEPT PROCESS"
            Rvse_Trn.Show
        Case "RMRG":
            MsgBox "Function Not Ready", vbCritical, "Message"
        Case "ASMG":
            Call DisableAllMenu
            Prd_Atmel_Merge.Show
        Case "STBI":
            Call DisableAllMenu
            STMICRO_BINCONVERT.Show
        Case "SCAC":
            Call DisableAllMenu
            Prd_Terminate.lbl_title = "TERMINATE FAIRCHILD RESIDUE - FOM LOTS"
            Prd_Terminate.Show
        Case "FSMR":
            Call DisableAllMenu
            Prd_Fairchild_Mrlt.Show
        Case "PODB":
            Call UpdateInvoiceSummary
            MsgBox "Database Updated", vbInformation, "Invoice Summary"
        Case "POCL":
            Call DisableAllMenu
            PO_Closure.Show
        Case "RMSR":
            Call DisableAllMenu
            Rpt_RMS.Show
        Case "RMSR":
            Call DisableAllMenu
            Rpt_RMS.Show
        Case "B2B$":
            Call DisableAllMenu
           B2B_Update.Show
        Case "PMSS":
            Call DisableAllMenu
            Prd_Mvou_STSCRAP.Show
        Case "STBIS":
            Call DisableAllMenu
            STMICRO_BINCONVERT_SCRAP.Show
        Case "VTRN":
            Call DisableAllMenu
            Rpt_ViewTrans.Show
        Case "LSRH":
            Call DisableAllMenu
            Rpt_LotSearch.Show
        Case "HLOT":
            Call DisableAllMenu
            Rpt_HoldWip.Show
        Case "WIPR":
            Call DisableAllMenu
            Rpt_Wip.Show
        Case "OUTR":
            Call DisableAllMenu
            Rpt_Output.Show
        Case "SUBW":
            Call DisableAllMenu
            Rpt_GESub.Show
        Case "ASYR":
            Call DisableAllMenu
            Rpt_Assy.Show
        Case "TSTR":
            Call DisableAllMenu
            Rpt_Test.Show
        Case "VMIR":
            Call DisableAllMenu
            Rpt_Vision.Show
        Case "QRPT":
            Call DisableAllMenu
            Rpt_QA.Show
        Case "CSRP":
            Call DisableAllMenu
            Rpt_CS.Show
        Case "POTR"
            Call DisableAllMenu
            Rpt_PoTracking.Show
        Case "RBLK":
            Call DisableAllMenu
            Rpt_DBlocked.Show
        Case "MCOU"
            Call DisableAllMenu
            Rpt_MCbyEmp.Show
        Case "MACL"
            Call DisableAllMenu
            Rpt_MachineByLot.Show
        Case "LDEF"
            Call DisableAllMenu
            Rpt_DefectModeByLot.Show
        Case "MAPF"
            Call DisableAllMenu
            Prd_MapFile_Find.Show
        Case "ALPH"
            Call DisableAllMenu
            Rpt_Alps.Show
        Case "SQLR"
            Call DisableAllMenu
            SQL_Report.Show
        Case "RMSU"
            Call DisableAllMenu
            Prd_Rms_Remark.Show
        End Select
End Sub