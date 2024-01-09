Dim RSQ As ADODB.Recordset
Dim WS_MDB As ADODB.Connection
Dim Refno As String, WO As String, PO As String, Order As Date, IPN As String
Dim CPN As String, EPN As String, DATECODE As String
Dim waferFlag As Boolean
Dim prefix As String
Dim oricuslotno
Public WAFERALLOW
Public waferExpBypass
Public liExpBypass
'Dim UttestStat As String
Private Function GetForm(ByRef sName As String) As Form
  For Each GetForm In Forms
    If GetForm.Name = sName Then Exit Function
  Next GetForm
    Set GetForm = Forms.Add(sName)
  End Function

Private Sub Addinfo_Click()

'    If Left(custnameselect.TEXT, 4) = "RACE" Then
'        MsgBox "RACE loading... Please inform IT"
'        'Exit Sub
'    End If
'


'public variable
'2010-10-25
namedash = InStr(Trim(cboCust.TEXT), " - (")
gvCustName = Left(Trim(cboCust.TEXT), namedash)
gvCustpo = lblPONo
gvCustTgtDev = TARGET_DEVICE_TXT
    
    
    Dim ANRS As ADODB.Recordset
    Dim TWNRS As Recordset
    Dim sqlstr As String
    Dim progModule As Form
    Dim formStr As String
        
    XX01 = dd1
    XX02 = mm1
    XX03 = yy1
        
        
    txtSLUT = FoundSl
    aic_refno = Trim(refno_TXT)
         
    Set TWNRS = New ADODB.Recordset
    Set ANRS = New ADODB.Recordset
    sqltext = "SELECT * FROM AIC_LOADING_INSTRUCTION WHERE REFNO='" & aic_refno & "'"
    ANRS.Open sqltext, wsDB
    
'    sqlstr = "SELECT PROG_MODULE FROM AIC_LI_CUST_MASTER WHERE CUSTNO = '" & Left(Right(Trim(cboCust.Text), 4), 3) & "'"
'    TWNRS.Open sqlstr, wsDB
'TNRBI_CHILD
'    If Not TWNRS.EOF Then
'      formStr = Trim(TWNRS("prog_module"))
'    Else
'      MsgBox "Program Module did not specified !"
'    End If
    
    formStr = "LI_Addinfo"
    
    If ANRS.EOF = False Then
    If Not IsNull(ANRS!PACKAGE__LEAD) Then
        If Not IsNull(ANRS!SLNO) Then SLNOx = Trim(ANRS!SLNO)
        
        If Trim(ANRS!Status) = "R" Then
            MsgBox "L.I. already released! Cannot update data!", vbCritical, "ERROR"
            AIC_STAT = "R"
            LI_General.Visible = False
            GetForm(formStr).Show
        ElseIf Trim(ANRS!Status) = "C" Then
            MsgBox "L.I. already Cancelled! Cannot update data!", vbCritical, "ERROR"
            AIC_STAT = "C"
            LI_General.Visible = False
            GetForm(formStr).Show
        ElseIf Trim(ANRS!Status) = "N" Then
            AIC_STAT = "N"
            PRODFG = Trim(bdcombo)
            If PRODFG = "" Then PRODFG = Trim(lbldeviceno)
            '------------------------------------TESTED/UNTESTED------------------>
            Dim rsbd As ADODB.Recordset
            Set rsbd = New ADODB.Recordset
            If PRODFG <> "" Then
                internal_device_no_txt = PRODFG
            Else
                internal_device_no_txt = bdcombo
            End If
      '      CSQLSTRING = "SELECT * FROM WIPPRD WHERE  WPRD_PROD = '" & Trim(PRODFG) & "' "
'AIMS
            CSQLSTRING = "SELECT * FROM BAIC_PRODMAST WHERE  PDM_DEVICENO = '" & Trim(PRODFG) & "' "
            
            rsbd.Open CSQLSTRING, wsDB
            If Not rsbd.EOF Then
                package_lead_txt = rsbd("PDM_PACKAGELEAD")
                TARGET_DEVICE_TXT = Trim(rsbd("PDM_TARGETDEVICE"))
                
                lblTUT = Trim(rsbd("PDM_PROCESS_REF"))

'                lblFroute = Trim(rsbd!WPRD_FRST_RTE)
'                lblLroute = Trim(rsbd!WPRD_LST_RTE)
                
                lblFroute = Trim(rsbd!PDM_ASSY_ROUTE)
                lblLroute = Trim(rsbd!PDM_ASSY_ROUTE)

                Dim xroute As ADODB.Recordset
                Set xroute = New ADODB.Recordset
                sqltext = "select min(rtg_oper) firstoper, max(rtg_oper) lastoper from baic_routing where rtg_route='" & Trim(rsbd!PDM_ASSY_ROUTE) & "'"
                
                xroute.Open sqltext, wsDB, adOpenDynamic, adLockOptimistic
                If xroute.EOF = False Then
                    lblOpr1 = xroute!firstoper
                    lblOpr2 = xroute!lastoper
'                    lblTUT = xroute!WRTE_RT_GRP_1
                End If
                xroute.Close
                Set xroute = Nothing
            End If
            rsbd.Close
            '-----------------------------------END------------------------------->
            UttestStat = Trim(lblTUT)
            'CHK TEST PROCESS
            markingJ1 = Trim(topx(2))
            markingJ2 = Trim(topx(3))
            markingJ = markingJ1 & "-" & markingJ2
            PLDF = Trim(package_txt) & Trim(ld_txt)
            If Trim(markingJ) = "" Then
                markingJ = "NA"
            End If
            If Trim(markingJ) = "-" Then
                markingJ = "NA"
            End If
            Dim ProdRS As ADODB.Recordset
            Set ProdRS = New ADODB.Recordset
            sqltext = "SELECT PDM_TARGETDEVICE FROM BAIC_PRODMAST WHERE PDM_DEVICENO='" & PRODFG & "'"
            
            ProdRS.Open sqltext, wsDB
            If ProdRS.EOF = False Then
                TDev = Trim(ProdRS!pdm_targetdevice)
            End If
            ProdRS.Close
            Set ProdRS = Nothing
            'END CHK
            
            'PASS PARAMETER
            Call PassParameter
            LI_General.Visible = False
            GetForm(formStr).Show
       End If
    Else
        MsgBox "LI Not Completed!!! Please Completely Save All Data First.", vbOKOnly, "Loading Instruction"
    End If
    Else
        MsgBox "LI No Does Not Exist!!! Please Save Data First.", vbOKOnly, "Loading Instruction"
    End If
End Sub

Private Sub PassParameter()
    'Refresh parameter value
    custx = ""
    CUSLOTx = ""
    CODEORAx = ""
    TARGET_DEVICEX = ""
    PACKAGEx = ""
    PONOx = ""
    LDx = ""
    BDx = ""
    DEVICEx = ""
    NOWAFERLOTx = 0
    SLNOx = ""
    CATALOGNOx = ""
    optBDx = False
    optTDx = False
    
    IBIDEVICEx = ""
    IBIPACKAGELDx = ""
    IBITARGET_DEVICEx = ""
    IBIBDx = ""
    IBIBD1x = ""
    
    CUSLOTx = Trim(txtCusLot)
    custx = cboCust.List(cboCust.ListIndex)
    CODEORAx = lblCustomerCodeOra.Caption
    TARGET_DEVICEX = Trim(TARGET_DEVICE_TXT)
    PACKAGEx = Trim(package_txt)
    LDx = Trim(ld_txt)
    BDx = Trim(bonding_diagram_txt)
    PONOx = Trim(lblPONo)
    CATALOGNOx = Trim(txtCatalogNo)
    
    optBDx = optBD.Value
    optTDx = optTargetDevice.Value
    DEVICEx = IIf(IsNull(Trim(lbldeviceno.Caption)), bdcombo.List(bdcombo.ListIndex), Trim(lbldeviceno.Caption))

    IBIDEVICEx = Trim(internal_device_no_txt)
    IBIPACKAGELDx = Trim(package_lead_txt)
    IBITARGET_DEVICEx = Trim(TARGET_DEVICE_TXT)
    IBIBDx = Trim(bd_no_txt)
    IBIBD1x = Trim(bd_no_txt1)
    If LenB(txtNoWaferLot.TEXT) <> 0 Then
        NOWAFERLOTx = txtNoWaferLot
    End If
    WAFERX = WAFER.TEXT
    QTYx = Val(qty.TEXT)
End Sub


Private Sub bdcombo_CLICK()
'bom_verify = ""
    
'Quah add 20211013, Device re-selected, LI_BOM deleted, need to verify again.
ssql = " delete from BAIC_LI_BOM where BOM_LIREF = '" & Trim(lbl_refno) & "'"
wsDB.Execute ssql
    
    
'20201026 Quah/Ko compare BD vs PRODMAST vs BOM must have same BD revision.  Not checking part content. Request by ChoyYing.
If (Right(Trim(bdcombo.TEXT), 6) <> "TTRPBF" And Right(Trim(bdcombo.TEXT), 3) <> "TTR") And (Left(Trim(bdcombo.TEXT), 4) <> "TROJ" _
   And Right(Trim(bdcombo.TEXT), 3) <> "RWK") And (Left(Trim(bdcombo.TEXT), 4) <> "SODR") And Left(Trim(bdcombo.TEXT), 5) <> "TSODR" And Right(Trim(bdcombo.TEXT), 6) <> "RWKPBF" And _
   Right(Trim(bdcombo.TEXT), 6) <> "PBFRWK" And Right(Trim(bdcombo.TEXT), 6) <> "RMAPBF" And Right(Trim(bdcombo.TEXT), 6) <> "RESPBF" And Right(Trim(bdcombo.TEXT), 3) <> "TST" And _
   Right(Trim(bdcombo.TEXT), 6) <> "TSTPBF" And Left(Trim(bdcombo.TEXT), 2) <> "OS" Then
'ASYRAF 20231017 ADD CHECKING DISABLE DATE
'        ssql = " select distinct Prd.PDM_INTERNAL_BD 'LIBD', Bom.BD_NUMBER 'BOMBD', Prd.PDM_DEVICENO from BAIC_PRODMAST Prd " & _
'               " left outer join AIC_BOM_COMPONENT BOM on Bom.DEVICE_NO=Prd.pdm_deviceno and Bom.REMARK='A' and Bom.DISABLE_DATE is null " & _
'               " where Prd.PDM_DEVICENO='" & Trim(bdcombo.TEXT) & "' "
        ssql = " select distinct Prd.PDM_INTERNAL_BD 'LIBD', Bom.BD_NUMBER 'BOMBD', Prd.PDM_DEVICENO from BAIC_PRODMAST Prd " & _
               " left outer join AIC_BOM_COMPONENT BOM on Bom.DEVICE_NO=Prd.pdm_deviceno and Bom.REMARK='A' AND GETDATE() < BOM.DISABLE_DATE  " & _
               " where Prd.PDM_DEVICENO='" & Trim(bdcombo.TEXT) & "' "
        Debug.Print ssql
        Set rs = New ADODB.Recordset
        rs.Open ssql, wsDB
        If rs.EOF = False Then
           If Trim(rs!LIBD) <> Trim(rs!BOMBD) Then
               MsgBox "Unmatch BD-Revision. Please check....." & vbCrLf & "Product Master: " & Trim(rs!LIBD) & vbCrLf & "Bom List: " & Trim(rs!BOMBD), vbCritical, "Message"
               VALID_FLAGx = False
               Exit Sub
           End If
        End If
        rs.Close
  End If
  'end... compare BD revision.
  
  'ASYRAF 230214
                        If Left(Trim(refno_TXT), 2) = "CH" Then
                            '20110714 AMICCOM
                            xyymmdd = Right(yy1, 2) & mm1 & dd1
                            topx(2).TEXT = Trim(topInfo(2).TEXT)
                            topx(3).TEXT = Trim(topInfo(3).TEXT)
                            topx(2).TEXT = Replace(topx(2).TEXT, "YYWW", ww_txt)
                            topx(2).TEXT = Replace(topx(2).TEXT, "YYMMDD", xyymmdd)
                            topx(2).TEXT = Replace(topx(2).TEXT, "YYMMDD", xyymmdd)
                            topx(2).TEXT = Replace(topx(2).TEXT, "(CustLot#)", "")  '2011-08-11 change to (CustLot#)  :refer Beh.
                            topx(3).TEXT = Replace(topx(3).TEXT, "YYWW", ww_txt)
                            topx(3).TEXT = Replace(topx(3).TEXT, "YYMMDD", xyymmdd)
                            topx(3).TEXT = Replace(topx(3).TEXT, "(CustLot#)", "") '2011-08-11 change to (CustLot#)  :refer Beh.
                        End If
    
    
mark_spec_txt = ""  '2014-11-18
'DIANAX
'Quah re-open 2014-12-23

If Right(Trim(bdcombo.TEXT), 3) = "TST" And custnameselect.TEXT = "STMICRO" Then 'koko add 20220809
ElseIf (Right(Trim(bdcombo.TEXT), 3) = "TST" Or Right(Trim(bdcombo.TEXT), 3) = "RWK") Then
    mark_spec_txt = "NA"
    For Xr = 0 To 5
        topx(Xr).Locked = False
        bottom(Xr).Locked = False
        topx(Xr).Enabled = True
        bottom(Xr).Enabled = True
    Next Xr
End If

'yana2022
'20230912 AIN ADD FOR KAK ROSE TO ENABLE MARKING EDIT FOR SOT23.AQ24COM.TGRPBF DEVICE
If Trim(bdcombo.TEXT) = "SOIR14.NPA1C02142TNIP.RMAPBF" Or Trim(bdcombo.TEXT) = "SOIR14.NPA1C02142CNIP.RMAPBF" Or Trim(bdcombo.TEXT) = "SOT23.AQ24COM.TGRPBF" Then
    For Xr = 0 To 5
        topx(Xr).Locked = False
        bottom(Xr).Locked = False
        topx(Xr).Enabled = True
        bottom(Xr).Enabled = True
    Next Xr
End If

'YANA ADD FOR MAHADHIR TO ENABLE MARKING EDIT FOR ENGINERING DEVICE
If InStr(Trim(bdcombo.TEXT), "ENPA") Then
    For Xr = 0 To 5
        topx(Xr).Locked = False
        bottom(Xr).Locked = False
        topx(Xr).Enabled = True
        bottom(Xr).Enabled = True
    Next Xr
End If

'YANA ADD FOR KAK ROSE TO ENABLE MARKING EDIT FOR REWORK DEVICE
If InStr(Trim(bdcombo.TEXT), "RMA") Then
    For Xr = 0 To 5
        topx(Xr).Locked = False
        bottom(Xr).Locked = False
        topx(Xr).Enabled = True
        bottom(Xr).Enabled = True
    Next Xr
End If

'2011-09-29 check for Device Block
Dim devblock As ADODB.Recordset
Set devblock = New ADODB.Recordset
ssql = "select PDM_BLOCK, PDM_CATEGORY from BAIC_PRODMAST where pdm_deviceno='" & Trim(bdcombo.TEXT) & "'"
devblock.Open ssql, wsDB
If Not devblock.EOF Then
    xxcat = Trim(devblock!pdm_category)
    If Trim(devblock!pdm_block) = "BLOCK" Then
        MsgBox "Cannot proceed. Please refer to AIMS Blocked-Device.", vbCritical, "Message"
        bdcombo.TEXT = ""
        Exit Sub
    End If
End If
devblock.Close
Set devblock = Nothing


'2021-12-09 temporary disabled.
'2021-11-15 Ricky request block if BOM was created/maintain over 1 year ago.
'Set devblock = New ADODB.Recordset
'ssql = " select DEVICE_NO, max(CREATION_DATE) CREATION_DATE, MAX(maint_date) MAINT_DATE from AIC_BOM_COMPONENT " & _
'       " where DATEDIFF(day,CREATION_DATE,GETDATE()) >= 365 and DEVICE_NO='" & Trim(bdcombo.TEXT) & "' " & _
'       " group by DEVICE_NO "
'Debug.Print ssql
'devblock.Open ssql, wsDB
'If Not devblock.EOF Then
'    If IsNull(devblock!MAINT_DATE) = True Or xserverdate - devblock!MAINT_DATE >= 365 Then
'        MsgBox "BOM last update was over 1 year old." & vbCrLf & "Refer Material Planner for confirmation.", vbCritical, "Message"
'        bdcombo.TEXT = ""
'        Exit Sub
'    End If
'End If
'devblock.Close
'Set devblock = Nothing



'Quah 20160829
'Quah 20160915 separate NPA,NPX
If Left(refno_TXT, 2) = "GE" Or Left(refno_TXT, 2) = "TA" Or Left(refno_TXT, 2) = "NV" Then  '
    If xxcat = "NPA" Then
        txtCusLot = "#TMARK3#"
    Else
        txtCusLot = "#TMARK4#"
    End If
End If



'2015-10-15 check for NS Attribute, block if not yet setup, Planner to inform Finance.
'-------------------------------------------------------------------------------------
If Left(refno_TXT, 2) = "NN" Then
    Set devblock = New ADODB.Recordset
    ssql = "select COUNT(*) CNT, MAX(isnull(IPH_PRODUCT_DESC,'')) ATTR from baic_invoice_price_header where IPH_DEVICENO='" & Trim(bdcombo.TEXT) & "' group by IPH_DEVICENO"
    Debug.Print ssql
    devblock.Open ssql, wsDB
    If Not devblock.EOF Then
        If devblock!cnt = 1 Then
            If Trim(devblock!ATTR) = "" Then
                MsgBox "Error with Price Attribute. Inform Finance.", vbCritical, "Message"
                bdcombo.TEXT = ""
                Exit Sub
            Else
                'have something, ok.
            End If
        Else
            MsgBox "Error with Price Attribute. Inform Finance.", vbCritical, "Message"
            bdcombo.TEXT = ""
            Exit Sub
        End If
    Else
        'check with mary.
        MsgBox "Error with Price Attribute. Inform Finance.", vbCritical, "Message"
        bdcombo.TEXT = ""
        Exit Sub
    End If
    devblock.Close
    Set devblock = Nothing
End If



'2012-02-23 check if wrong customer selected.
'----------------------------------------------
Set devblock = New ADODB.Recordset
ssql = "select PDM_CUSTOMER from BAIC_PRODMAST where pdm_deviceno='" & Trim(bdcombo.TEXT) & "'"
devblock.Open ssql, wsDB
If Not devblock.EOF Then
    If Trim(devblock!PDM_CUSTOMER) <> custnameselect.TEXT Then
        MsgBox "Error. Deviceno not belong to this customer.", vbCritical, "Message"
        bdcombo.TEXT = ""
        Exit Sub
    End If
End If
devblock.Close
Set devblock = Nothing

'STDevicenoTST = Trim(bdcombo.TEXT)

'STDevicenoTST = Right(Trim(bdcombo.TEXT), 3) = "TST"
If custnameselect.TEXT = "TACTILIS" And InStr(bdcombo.TEXT, "DIE") > 0 Then
    checkingbom = "NO"
ElseIf custnameselect.TEXT = "STMICRO" And Right(Trim(bdcombo.TEXT), 3) = "TST" Then 'KO ADD 20220809
    checkingbom = "NO"
Else
 checkingbom = "YES"
End If

If checkingbom = "YES" Then
    '2011-10-11 check for blank short desc
    bompartcount = 0
    Set devblock = New ADODB.Recordset
    'SSQL = "select count(*) kount " & _
    '                         " from aic_bom_header A, aic_bom_component B  " & _
    '                         " Where a.bill_sequence_id = b.bill_sequence_id " & _
    '                         " and A.active_flag = 'Y' " & _
    '                         " and b.apt_print = 'Y' " & _
    '                         " and b.part_no not like 'DIE%' " & _
    '                         " and a.DEVICE_NO ='" & Trim(bdcombo.TEXT) & "'"
    ssql = "select count(*) kount " & _
                             " from aic_bom_header A, aic_bom_component B  " & _
                             " Where a.bill_sequence_id = b.bill_sequence_id " & _
                             " and A.active_flag = 'Y' " & _
                             " and a.DEVICE_NO ='" & Trim(bdcombo.TEXT) & "'"
                            Debug.Print ssql
    devblock.Open ssql, wsDB
    If Not devblock.EOF Then
        bompartcount = devblock!kount
    End If
    devblock.Close
    Set devblock = Nothing
    Debug.Print ssql
    If bompartcount > 0 Then
        Set devblock = New ADODB.Recordset
        'Quah 2013-08-23 exclude Capillary 052, no need to appear in APT, but APT_Print=Y
        ssql = "select b.PART_NO, b.PART_SHORT_DESC " & _
                                 " from aic_bom_header A, aic_bom_component B  " & _
                                 " Where a.bill_sequence_id = b.bill_sequence_id " & _
                                 " and A.active_flag = 'Y' " & _
                                 " and b.apt_print = 'Y' and b.remark='A' " & _
                                 " and b.PART_SHORT_DESC = '' " & _
                                 " and b.part_no not like 'DIE%' " & _
                                 " and b.part_no not like '052%' " & _
                                 " and a.DEVICE_NO ='" & Trim(bdcombo.TEXT) & "'"
        Debug.Print ssql
        devblock.Open ssql, wsDB
        If Not devblock.EOF Then
            MsgBox "Bom Part Desc is blank. Refer to Material Planner. ", vbCritical, "Message"
            bdcombo.TEXT = ""
            Exit Sub
        End If
        devblock.Close
        Set devblock = Nothing
    
    ElseIf bompartcount = 0 Then
            MsgBox "Bom Part not defined. Refer to Material Planner. ", vbCritical, "Message"
            bdcombo.TEXT = ""
            Exit Sub
    
    End If
End If      'checkingbom
    
    
mark_spec_txt.Locked = True
    Dim rsbd As ADODB.Recordset
    
    'Quah 20090626 - check for valid packagelead in WIPPRD, AIC_PACKAGELEAD_MASTER
'AIMS
'    Set rsbd = New ADODB.Recordset
'    CSQLSTRING = " SELECT WPRD_PRD_GRP_4 shortpackage FROM WIPPRD, AIC_PACKAGELEAD_MASTER WHERE  WPRD_PROD ='" & Trim(bdcombo.TEXT) & "' and WPRD_PRD_GRP_4=PACKAGLEAD_SHORT and WPRD_PRD_GRP_4 <> 'NA'"
'    rsbd.Open CSQLSTRING, wsDB
'    If rsbd.EOF Then
'            MsgBox "Product package in Workstream (VPRD) not match with AIC_PACKAGELEAD_MASTER." & Chr(13) & _
'            "Please check/register.", vbCritical, "Error"
'            bdcombo.TEXT = ""
'            Exit Sub
'    End If
'    rsbd.Close
    'Quah 20090626
    
    lblFroute = "": lblLroute = "": lblOpr1 = "": lblOpr2 = "": lblTUT = ""
    Set rsbd = New ADODB.Recordset
    internal_device_no_txt = bdcombo
  '  CSQLSTRING = "SELECT * FROM WIPPRD WHERE  WPRD_PROD = '" & Trim(bdcombo.TEXT) & "' "
'AIMS
    CSQLSTRING = "SELECT * FROM BAIC_PRODMAST WHERE  PDM_DEVICENO = '" & Trim(bdcombo.TEXT) & "' "
    
    Debug.Print CSQLSTRING
    rsbd.Open CSQLSTRING, wsDB
    If Not rsbd.EOF Then
    
        'Quah 20081028 check for Route attached.
'        If Trim(rsbd!WPRD_FRST_RTE) = "" Then
        If Trim(rsbd!PDM_ASSY_ROUTE) = "" Then
            MsgBox "Planner, please check." & Chr(13) & "No ROUTE is attached to this Device."
            bdcombo.TEXT = ""
            Exit Sub
        End If
        
        'Quah 20090317 check for Route Verification.
'aims...
'        If Trim(rsbd!wprd_unit_3) = "Y" Then
'            'okay
'        Else
'            MsgBox ("Planner, please check. ROUTE Not Verified Yet.")
'            bdcombo.TEXT = ""
'            Exit Sub
'        End If
        
        package_lead_txt = rsbd("pdm_packagelead")
        TARGET_DEVICE_TXT = Trim(rsbd("pdm_targetdevice"))
            
            lblTUT = Trim(rsbd("pdm_process_ref"))
            
        Dim rsDevice As ADODB.Recordset
        lblFroute = Trim(rsbd!PDM_ASSY_ROUTE)
        lblLroute = Trim(rsbd!PDM_ASSY_ROUTE)
        Dim xroute As ADODB.Recordset
        Set xroute = New ADODB.Recordset
'        sqltext = "SELECT * FROM WIPRTE WHERE WRTE_ROUTE='" & Trim(rsbd!WPRD_FRST_RTE) & "'"
'aims
        sqltext = "select min(rtg_oper) firstoper, max(rtg_oper) lastoper from baic_routing where rtg_route='" & Trim(rsbd!PDM_ASSY_ROUTE) & "'"
        Debug.Print sqltext
        xroute.Open sqltext, wsDB, adOpenDynamic, adLockOptimistic
        If xroute.EOF = False Then
'            lblOpr1 = xroute!WRTE_FRST_RT_OPR
'            lblOpr2 = xroute!WRTE_LST_RT_OPR
'            lblTUT = xroute!WRTE_RT_GRP_1
            lblOpr1 = xroute!firstoper
            lblOpr2 = xroute!lastoper
'            lblTUT = xroute!WRTE_RT_GRP_1
        
        End If
        xroute.Close
        Set xroute = Nothing
    End If
    rsbd.Close
    
    Set rsbd = New ADODB.Recordset
  '  CSQLSTRING = "SELECT * FROM WIPPRD WHERE WPRD_PROD = '" & Trim(bdcombo.TEXT) & "' "
'AIMS
    CSQLSTRING = "SELECT * FROM BAIC_PRODMAST WHERE PDM_DEVICENO = '" & Trim(bdcombo.TEXT) & "' "

    rsbd.Open CSQLSTRING, wsDB
    If rsbd.EOF = False Then
    'aims
        If IsNull(rsbd!PDM_INTERNAL_BD) Then
            bd_no_txt.TEXT = ""
        Else
            bd_no_txt.TEXT = Trim(rsbd!PDM_INTERNAL_BD)
        End If
        If Not IsNull(Trim(rsbd!pdm_catalog_no)) Then txtCatalogNo = Trim(rsbd!pdm_catalog_no)
    Else
        MsgBox "ERROR - Internal BD Not Found!", vbInformation
    End If
    rsbd.Close
    Set rsbd = Nothing
    
    dash1 = 0: DASH2 = 0: dash3 = 0
    dash1 = InStr(1, bd_no_txt, "-")
    DASH2 = InStr(dash1 + 1, bd_no_txt, "-")
    dash3 = InStr(DASH2 + 1, bd_no_txt, "-")
    
    
     'Quah 2010-01-28 for NN Retaping (TST), skip checking BD.
     '2010-05-14 INCL RWK
'     If Left(Trim(refno_TXT), 2) = "NN" And (Right(Trim(bdcombo.TEXT), 3) = "TST" Or Right(Trim(bdcombo.TEXT), 3) = "RWK") Then
     If (Right(Trim(bdcombo.TEXT), 3) = "TST" Or Right(Trim(bdcombo.TEXT), 3) = "RWK" Or Right(Trim(bdcombo.TEXT), 3) = "DIE") Then        'Add DIE for TACTILIS
        'skip bd
        sqltext = " select Remark2 markspec FROM AIC_DEVICE_CONTROL WHERE CUSTOMER='CUSBD' AND TARGETDEVICE='MARKSPEC' and remark7='" & Trim(TARGET_DEVICE_TXT) & "'" & _
                  " and REMARK3='OPEN' AND remark4='" & Trim(bdcombo) & "'  ORDER BY REMARK5 DESC"
                
        Debug.Print sqltext
        Dim rsmarking As ADODB.Recordset
        Set rsmarking = New ADODB.Recordset
        rsmarking.Open sqltext, wsDB
        If Not rsmarking.EOF Then
            mark_spec_txt = Trim(rsmarking!MarkSpec)
        End If
        rsmarking.Close
        
        Call MarkSpec
     Else
     
    Set rsbd = New ADODB.Recordset
    If dash3 <= 0 Then  'QUAH 20091111 for Osram, BD number only 2 segment due to no package code as prefix.
        CSQLSTRING = "SELECT * FROM AIC_BD_NO WHERE AICBD_CUSBD_NUMBER = '" & Trim(bonding_diagram_txt) & "' AND AICBD_BD_NUMBER LIKE '" & Left(Trim(bd_no_txt), DASH2 - 1) & "%'  AND AICBD_REVISION = '" & Right(Trim(bd_no_txt), 1) & "' AND AICBD_STATUS='INUSE'"
        Debug.Print CSQLSTRING
    Else
        CSQLSTRING = "SELECT * FROM AIC_BD_NO WHERE AICBD_CUSBD_NUMBER = '" & Trim(bonding_diagram_txt) & "' AND AICBD_BD_NUMBER LIKE '" & Left(Trim(bd_no_txt), dash3 - 1) & "%'  AND AICBD_REVISION = '" & Right(Trim(bd_no_txt), 1) & "' AND AICBD_STATUS='INUSE'"
    End If
    Debug.Print CSQLSTRING
    rsbd.Open CSQLSTRING, wsDB
    If rsbd.EOF = False Then
        If Not IsNull(Trim(rsbd!AICBD_DEVICE)) Then
            For iCnt = 1 To lvwDie.ListItems.Count
                Set itmx = lvwDie.ListItems(iCnt)
                    If Trim(itmx.TEXT) = "1" Then
                        Text1 = "1"
                        Text2 = Trim(rsbd!AICBD_DEVICE)
                        Text3 = itmx.SubItems(2)
                        Text4 = itmx.SubItems(3)
                        Text5 = itmx.SubItems(4)
                        Text6 = itmx.SubItems(5)
                    End If
            Next iCnt
            Text1 = "1"
            Text2 = Trim(rsbd!AICBD_DEVICE)
            chkShow.Value = 1
        End If
    
        'Quah 2012-03-30
        If Left(Trim(refno_TXT), 2) = "HT" Then
                Text3 = Me.txtCusLot
                Text4 = Me.qty
                Text5 = Me.qty
        End If
        'Quah 2012-04-05
        If Left(Trim(refno_TXT), 2) = "AA" Or Left(Trim(refno_TXT), 2) = "AA" Then
                Text3 = Me.WAFER
                Text4 = Me.qty
                Text5 = Me.qty
        End If
        
     
        
        If Left(Trim(refno_TXT), 2) <> "NN" And Left(Trim(refno_TXT), 2) <> "NZ" Then
        
        'autopull marking & markspec 20091222
        '===============================================================
        mark_spec_txt = ""
        INIT_MARK
        
'        sqltext = " select Remark2 markspec FROM AIC_DEVICE_CONTROL WHERE CUSTOMER='CUSBD' AND TARGETDEVICE='MARKSPEC' and remark7='" & Trim(TARGET_DEVICE_TXT) & "'" & _
'                  " and remark4='" & Trim(bdcombo) & "' and remark1='" & Trim(bonding_diagram_txt.TEXT) & "'"
        
        sqltext = " select Remark2 markspec FROM AIC_DEVICE_CONTROL WHERE CUSTOMER='CUSBD' AND TARGETDEVICE='MARKSPEC' and remark7='" & Trim(TARGET_DEVICE_TXT) & "'" & _
                  " and REMARK3='OPEN' AND remark4='" & Trim(bdcombo) & "'  ORDER BY REMARK5 DESC"
                
        Debug.Print sqltext
        Dim rsmark As ADODB.Recordset
        Set rsmark = New ADODB.Recordset
        rsmark.Open sqltext, wsDB
        If Not rsmark.EOF Then
            mark_spec_txt = Trim(rsmark!MarkSpec)
            'DIANAX
'''            For Xr = 0 To 5
'''                topx(Xr).Locked = True  'locked
'''                bottom(Xr).Locked = True 'locked
'''            Next Xr
        End If
        rsmark.Close
        '===============================================================
        'copy code from MarkSpec procedure
                If Left(Trim(refno_TXT), 2) = "AS" Then
                    Call MarkSpec_Sanjose
                Else
                    If Left(Trim(refno_TXT), 2) = "GM" Then '2010-05-10
                        mark_spec_txt.Locked = False
                    Else
                        Call MarkSpec
                        If Left(Trim(refno_TXT), 2) = "AD" Then
                            Call MarkSpec_AD
                        End If
                        'Quah 2013-01-09 AVT=NIKO
                        If Left(Trim(refno_TXT), 2) = "NK" Or Left(Trim(refno_TXT), 2) = "AO" Then
                            Call MarkSpec_NK
                        End If
                        If Left(Trim(refno_TXT), 2) = "II" Then
                            Call MarkSpec_II
                        End If
                        If Left(Trim(refno_TXT), 2) = "CT" Then
                            Call MarkSpec_CT
                        End If
                        If Left(Trim(refno_TXT), 2) = "GT" Then
                            Call MarkSpec_GT
                        End If
                        If Left(Trim(refno_TXT), 2) = "BE" And LenB(txtCusLot) = 0 Then  'Quah 20181204
'                            txtCusLot = "XXXXXXXYMMDD"   '20181304
                            txtCusLot = "XXXXXXXYMD"   '20181219
'''''''''                            bepoint = 0     'Quah 20091201 skip .1,.2 inserted by Planner to differentiate loading batch.
'''''''''                            bepoint = InStr(txtCusLot, ".")
'''''''''                            If bepoint > 0 Then
'''''''''                                topx(2).TEXT = Right(Mid(txtCusLot, 1, bepoint - 1), 2)
'''''''''                            Else
'''''''''                                topx(2).TEXT = Right(txtCusLot, 2)
'''''''''                            End If
'''''''''                            If InStr(topx(3).TEXT, "YWW") > 0 Then
'''''''''                                y_ww = Right(ww_txt, 3)
'''''''''                                topx(3).TEXT = Replace(topx(3).TEXT, "YWW", y_ww)
'''''''''                            End If
                        End If
                        If Left(Trim(refno_TXT), 2) = "MN" Then 'Quah 20090629 Magnachip auto marking
    '                        If Trim(topx(2).TEXT) = "YWLLLLGC" Then
                            If Left(Trim(topx(2).TEXT), 6) = "YWLLLL" Then 'Quah 2010-04-12 LL request to check only first 6 chars, last 2 is not fixed.
                                Set Rs03 = New ADODB.Recordset
                                CSQLSTRING = "select * from wwcal where ww_date='" & Format(Date, "DD-MMM-YYYY") & "'"
                                Debug.Print CSQLSTRING
                                Rs03.Open CSQLSTRING, wsDB
                                If Not Rs03.EOF Then
    '                                topx(2).TEXT = Trim(Rs03!ww_magnachip) & "(ASSY#L4#)GC"
                                    topx(2).TEXT = Trim(Rs03!ww_magnachip) & "(ASSY#L4#)" & Right(Trim(topx(2).TEXT), 2)
                                Else
                                    MsgBox "ERROR IN MAGNACHIP MARKING-DATECODE. PLEASE CHECK !!!!", vbCritical
                                    topx(2).TEXT = ""
                                    mark_spec_txt.TEXT = ""
                                    Exit Sub
                                End If
                                Rs03.Close
                            End If
                        End If
                        If Left(Trim(refno_TXT), 2) = "UL" Then 'Quah 20091008 Ultrachip auto marking
                            If Trim(topx(2).TEXT) = "YYMMWW" Then
                                Set Rs03 = New ADODB.Recordset
                                CSQLSTRING = "select * from wwcal where ww_date='" & Format(Date, "DD-MMM-YYYY") & "'"
                                Debug.Print CSQLSTRING
                                Rs03.Open CSQLSTRING, wsDB
                                If Not Rs03.EOF Then
                                    topx(2).TEXT = Trim(Rs03!ww_ultrachip)
                                Else
                                    MsgBox "ERROR IN ULTRACHIP MARKING-DATECODE. PLEASE CHECK !!!!", vbCritical
                                    topx(2).TEXT = ""
                                    mark_spec_txt.TEXT = ""
                                    Exit Sub
                                End If
                                Rs03.Close
                            End If
                        End If
                        '20110708
                        If Left(Trim(refno_TXT), 2) = "CM" Then
                            '20110714 AMICCOM
                            xyymmdd = Right(yy1, 2) & mm1 & dd1
                            topx(2).TEXT = Trim(topInfo(2).TEXT)
                            topx(3).TEXT = Trim(topInfo(3).TEXT)
                            topx(2).TEXT = Replace(topx(2).TEXT, "YYWW", ww_txt)
                            topx(2).TEXT = Replace(topx(2).TEXT, "YYMMDD", xyymmdd)
                            topx(2).TEXT = Replace(topx(2).TEXT, "(CustLot#)", "")  '2011-08-11 change to (CustLot#)  :refer Beh.
                            topx(3).TEXT = Replace(topx(3).TEXT, "YYWW", ww_txt)
                            topx(3).TEXT = Replace(topx(3).TEXT, "YYMMDD", xyymmdd)
                            topx(3).TEXT = Replace(topx(3).TEXT, "(CustLot#)", "") '2011-08-11 change to (CustLot#)  :refer Beh.
                        End If
                        
                        'ASYRAF 230214
                        If Left(Trim(refno_TXT), 2) = "CH" Then
                            '20110714 AMICCOM
                            xyymmdd = Right(yy1, 2) & mm1 & dd1
                            topx(2).TEXT = Trim(topInfo(2).TEXT)
                            topx(3).TEXT = Trim(topInfo(3).TEXT)
                            topx(2).TEXT = Replace(topx(2).TEXT, "YYWW", ww_txt)
                            topx(2).TEXT = Replace(topInfo(2).TEXT, "YYMMDD", xyymmdd)
                            topx(2).TEXT = Replace(topx(2).TEXT, "(CustLot#)", "")  '2011-08-11 change to (CustLot#)  :refer Beh.
                            topx(3).TEXT = Replace(topx(3).TEXT, "YYWW", ww_txt)
                            topx(3).TEXT = Replace(topx(3).TEXT, "YYMMDD", xyymmdd)
                            topx(3).TEXT = Replace(topx(3).TEXT, "(CustLot#)", "") '2011-08-11 change to (CustLot#)  :refer Beh.
                        End If
                        
                        'Quah 2013-09-18
                        'Quah 2014-04-02 add for TopSystem
                        'Quah 2014-05-09 add for Amphenol
                        If Left(Trim(refno_TXT), 2) = "GE" Or Left(Trim(refno_TXT), 2) = "TA" Or Left(Trim(refno_TXT), 2) = "NV" Then
                            topx(2).TEXT = Replace(topx(2).TEXT, "YYWW", ww_txt)
                            topx(3).TEXT = Replace(topx(3).TEXT, "YYWW", ww_txt)
                        End If
                        
                        
                        '20170404 Add for ASTI
                        If Left(Trim(refno_TXT), 2) = "DP" Then
                            topx(3).TEXT = Replace(topx(3).TEXT, "YYYY", yy1)
                            topx(3).TEXT = Replace(topx(3).TEXT, "WW", Right(ww_txt, 2))
                        End If
                        
                        
                        '20110913
                        If Left(Trim(refno_TXT), 2) = "TF" Then
                            dash = InStr(txtCusLot, "-")
                            If dash = 0 Then
                                tfmark = Right(Trim(txtCusLot), 4)
                            Else
                                tfmark = Right(Left(Trim(txtCusLot), dash - 1), 4)
                            End If
                            
                            topx(3).TEXT = Replace(topx(3).TEXT, "(last4waferlot#)", "")
                            topx(3).TEXT = Replace(topx(3).TEXT, "XXXX", tfmark)
                        End If
                        
                        
                        '20220512
                        If InStr(cbofullpackage, "PLCC") > 0 And InStr(cboCust, "MICROCHIP") Then
                           For Xr = 0 To 5
                                'boleh edit
                                topx(Xr).Locked = False
                                bottom(Xr).Locked = False
                                topx(Xr).Enabled = True
                                bottom(Xr).Enabled = True
                            Next Xr
                        End If
                        
                        
                        '
                    End If
                    
                    'Quah 20090211 add Fairchild-Pg (FP)
                    If Left(Trim(refno_TXT), 2) = "FS" Or Left(Trim(refno_TXT), 2) = "FP" Then     'Different prefix setting.
                        
                        txtCusLot.TEXT = "" 'added 2012-01-17 to initlize data click by FOM buttom.
                        
                        Set Rs03 = New ADODB.Recordset
                        CSQLSTRING = "SELECT * FROM AIC_SETUP_DATA WHERE TABLE_NAME='FSMARKING' and PROG_ITEM1='" & Trim(Me.TARGET_DEVICE_TXT) & "'"
                        Rs03.Open CSQLSTRING, wsDB
                        If Not Rs03.EOF Then
                            X1 = Trim(Rs03!prog_item2)
                            X2 = Trim(Rs03!prog_item3)
                            
                            'Added 2012-01-17
                            Me.topx(2) = X1 & Left(Trim(WAFER.TEXT), 6) & X2
                        Else
                            X1 = ""
                            X2 = ""
                        End If
                        Rs03.Close
                        'disabled 2012-01-17
'                        Me.topx(2) = X1 & Left(Trim(WAFER.TEXT), 6) & X2


'
'                        If Trim(Me.TARGET_DEVICE_TXT) = "LTA504SGZ" Then
'                            Me.topx(2) = "H1" & Left(Trim(WAFER.Text), 6) & "YYWWQ"
'                        ElseIf Trim(Me.TARGET_DEVICE_TXT) = "SG6961SZ" Then
'                            Me.topx(2) = "B1" & Left(Trim(WAFER.Text), 6) & "YYWWQ"
'                        ElseIf Trim(Me.TARGET_DEVICE_TXT) = "LTA504SGZF" Then
'                            Me.topx(2) = "F1" & Left(Trim(WAFER.Text), 6) & "YYWWQ"
'                        ElseIf Trim(Me.TARGET_DEVICE_TXT) = "SG5841DZ" Then
'                            Me.topx(2) = "D3" & Left(Trim(WAFER.Text), 6) & "YYWWQ"
'                        ElseIf Trim(Me.TARGET_DEVICE_TXT) = "SG5841JDZ" Then
'                            Me.topx(2) = "D3" & Left(Trim(WAFER.Text), 6) & "YYWWQ"
'                        ElseIf Trim(Me.TARGET_DEVICE_TXT) = "SG6848DZ1" Then
'                            Me.topx(2) = "B" & Left(Trim(WAFER.Text), 6) & "YYWWQ"
'                        ElseIf Trim(Me.TARGET_DEVICE_TXT) = "SG6846LSZ" Then
'                            Me.topx(2) = "F" & Left(Trim(WAFER.Text), 6) & "YYWWQ"
'                        ElseIf Trim(Me.TARGET_DEVICE_TXT) = "SG6846LDZ" Then
'                            Me.topx(2) = "F" & Left(Trim(WAFER.Text), 6) & "YYWWQ"
'                        ElseIf Trim(Me.TARGET_DEVICE_TXT) = "SG684965DZ" Then
'                            Me.topx(2) = "B" & Left(Trim(WAFER.Text), 6) & "YYWWQ"
'                        ElseIf Trim(Me.TARGET_DEVICE_TXT) = "SG6841DZ" Then
'                            Me.topx(2) = "K" & Left(Trim(WAFER.Text), 6) & "YYWWQ"
'                        ElseIf Trim(Me.TARGET_DEVICE_TXT) = "SG6841SZ" Then
'                            Me.topx(2) = "K" & Left(Trim(WAFER.Text), 6) & "YYWWQ"
'                        Else
'                            Me.topx(2) = "F" & Left(Trim(WAFER.Text), 6) & "YYWWQ"
'                        End If
                    End If
                
                End If
        
        End If
        
        '===============================================================
    Else
        MsgBox "ERROR - DIE PART NOT FOUND, Please verify with .com Customer BD and Internal BD already insert into Database!", vbInformation
    End If
    rsbd.Close
    Set rsbd = Nothing
   
    End If  'skip bd for NN TST (retaping)
   
   
    txtCusLot.SetFocus
End Sub

Private Sub be_confirm_Click()
    totalDat = be_listscribe.ListItems.Count
    Dim rnum As Integer
    be_total = ""
    
If totalDat = 0 Then
    MsgBox "No wafer# selected!"
Else
    For rnum = 1 To totalDat
        If be_listscribe.ListItems(rnum).Checked = True Then
            be_total = Val(be_total) + Val(be_listscribe.ListItems(rnum).SubItems(1))
            scribe = scribe & Trim(be_listscribe.ListItems(rnum)) & vbCrLf
            
            If be_listscribe.ListItems(rnum).SubItems(2) <> "" Then
                MsgBox "Scribe# " & Trim(be_listscribe.ListItems(rnum)) & " used in LI " & be_listscribe.ListItems(rnum).SubItems(2) & ""
                be_total = ""
                Exit Sub
            End If
        End If
    Next rnum
End If
End Sub

Private Sub be_waferlotno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    be_listscribe.ColumnHeaders.Clear
    be_listscribe.ListItems.Clear
    
    Dim sscribe As String
    sscribe = "select start_qty 'SCRIBE#',GOOD_QTY,remarks1 'LI_REF' from AIC_LABEL_REFERENCE where ASSYLOTNO = 'BOURNS IQA' AND CUSLOTNO ='" & Trim(be_waferlotno.TEXT) & "' ORDER BY START_QTY ASC" '& _                " and remarks1 = '' "
    
    Dim rsData As New ADODB.Recordset
    Dim lngIndex As Long
    Dim lngFieldCount As Long
    Dim lngWidth As Long
    Dim lstitmTemp As ListItem

    Set rsData = wsDB.Execute(sscribe)
    
    lngFieldCount = rsData.Fields.Count
    lngWidth = (be_listscribe.Width - 320) / lngFieldCount
    For lngIndex = 0 To lngFieldCount - 1
        be_listscribe.ColumnHeaders.Add , , rsData.Fields(lngIndex).Name, lngWidth
    Next
    
    Do While Not rsData.EOF
        Set lstitmTemp = be_listscribe.ListItems.Add(, , rsData.Fields(0).Value)
        For lngIndex = 1 To lngFieldCount - 1
            If IsNull(rsData.Fields(lngIndex).Value) Then
                lstitmTemp.SubItems(lngIndex) = ""
            Else
                lstitmTemp.SubItems(lngIndex) = rsData.Fields(lngIndex).Value
            End If
        Next
    
        rsData.MoveNext
    Loop
    
'    be_listscribe.ColumnHeaders(3).Alignment = lvwColumnRight
End If
End Sub

Private Sub beconf_Click()
If be_total = "" Then
    MsgBox "Please check total first!!!"
    Exit Sub
End If

'CHECK NEW LI
If stat = "C" Or stat = "R" Then
    MsgBox "Not allowed! LI released or cancelled!!!"
    be_waferlotno.Clear
    be_waferlotno.Locked = False
    be_listscribe.ListItems.Clear
    be_total = ""
    scribe = ""
    Frame_BE.Visible = False
    Check1.Value = Unchecked
    Exit Sub
End If

Call be_confirm_Click

ANS = 0
ANS = MsgBox("CONFIRM SELECT WAFER SCRIBE?", vbYesNo, "Message")
If ANS = vbYes Then

    funcDataList = ""
    Dim strOrig, strTemp As String
    strOrig = scribe
    strTemp = ""
    
    strTemp = Replace(strOrig, Chr(10), "")
    strTemp = Replace(strTemp, Chr(13), "','")
    strTemp = "'" & strTemp & "'"
    funcDataList = strTemp

        Set rs = New ADODB.Recordset
        SQL = "update AIC_LABEL_REFERENCE set remarks1='" & Trim(refno_TXT) & "' WHERE ASSYLOTNO = 'BOURNS IQA' AND CUSLOTNO ='" & Trim(be_waferlotno.TEXT) & "' " & _
                " AND start_qty IN (" & funcDataList & ")"
        Debug.Print SQL
        rs.Open SQL, wsDB
            MsgBox "Updated to LI#: " & Trim(refno_TXT) & ""

            be_waferlotno.Clear
            be_waferlotno.Locked = False
            be_listscribe.ListItems.Clear
            be_total = ""
            scribe = ""
            Check1.Value = Unchecked
            Frame_BE.Visible = False
Else
    be_waferlotno.Clear
    be_waferlotno.Locked = False
    be_listscribe.ListItems.Clear
    be_total = ""
    scribe = ""
    Frame_BE.Visible = False
    Check1.Value = Unchecked
End If
    
End Sub

Private Sub BOMLI_VIEW_Click()

  'Quah add 2021-Sept, for BOM screen.
  Dim chkbom As ADODB.Recordset
  Set chkbom = New ADODB.Recordset
  ssql = "select * from baic_li_bom where bom_liref='" & Trim(refno_TXT) & "'"
  Debug.Print ssql
  chkbom.Open ssql, wsDB
  If Not chkbom.EOF Then
    libomfound = True
  Else
    libomfound = False
  End If
  chkbom.Close
  
'  If libomfound = False Then
'  If login_id = "ITADM" Then
      frm_MfgInventory.Show vbModal
'  End If
  
  'KOKO ADD 20200627 ' FOR BOM LI
'  CrystalReportB.Connect = "PROVIDER=MSDASQL;dsn=AICSSQLDB;uid=ITAPP;pwd=APP;"
'  CrystalReportB.ReportFileName = App.Path & "\report\BOMLI-A.RPT"
'  CrystalReportB.SelectionFormula = "{BAIC_LI_BOM.BOM_LIREF}='" & Trim(refno_TXT) & "'"
 ' CrystalReportB.Destination = crptToWindow
 ' CrystalReportB.WindowState = crptMaximized
 ' CrystalReportB.Action = 1
            
End Sub

Private Sub bonding_diagram_txt_GotFocus()
'    Dim chkAllpkg As ADODB.Recordset
'    Set chkAllpkg = New ADODB.Recordset
'  '  sqltext = "SELECT * FROM WIPPRD WHERE  WPRD_PRD_GRP_4 = '" & Trim(package_txt) & Trim(ld_txt) & "'"
''AIMS
'    SQLTEXT = "SELECT * FROM BAIC_PRODMAST WHERE  PDM_PACKAGELEAD = '" & Trim(package_txt) & Trim(ld_txt) & "'"
'    chkAllpkg.Open SQLTEXT, wsDB
'    If chkAllpkg.EOF = True Or package_txt = "NA" Then
'        package_txt = vbNullString
'        ld_txt = vbNullString
'        MsgBox "PackageLead is not setup in Database. Check with Planner!!!"
'    End If
'    chkAllpkg.Close
'    Set chkAllpkg = Nothing
End Sub

Private Sub bonding_diagram_txt_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        
'      Dim chkAllpkg As ADODB.Recordset
'      Set chkAllpkg = New ADODB.Recordset
'    '  sqltext = "SELECT * FROM WIPPRD WHERE  WPRD_PRD_GRP_4 = '" & Trim(package_txt) & Trim(ld_txt) & "'"
'      'AIMS
'      SQLTEXT = "SELECT * FROM BAIC_PRODMAST WHERE  PDM_PACKAGELEAD = '" & Trim(package_txt) & Trim(ld_txt) & "'"
'      chkAllpkg.Open SQLTEXT, wsDB
'      If chkAllpkg.EOF = True Or package_txt = "NA" Then
'          package_txt = vbNullString
'          ld_txt = vbNullString
'          MsgBox "PackageLead is not setup in Database. Check with Planner!!!"
'          Exit Sub
'      End If
'      chkAllpkg.Close
'      Set chkAllpkg = Nothing
        
        
        TARGET_DEVICE_TXT = vbNullString
        target_device_txt1 = vbNullString
        cbobonding_diagram.Clear
        bd_no_txt.TEXT = vbNullString
        bd_no_txt1.TEXT = vbNullString
        If bonding_diagram_txt = "N/A" Or bonding_diagram_txt = "NA" Then
            lbldeviceno.Visible = False
            bdcombo.Visible = True
            bdcombo.Clear
            lbltestonly.Visible = True
            lbltestonly = "TEST ONLY"
            Dim DevRSZ As ADODB.Recordset
            Set DevRSZ = New ADODB.Recordset
            
            'Quah 2010-05-27 remove dot in % search.
'            CSQLSTRING = "SELECT A.* ,B.WRTE_ROUTE FROM WIPPRD A,WIPRTE B " & _
'                     "WHERE WPRD_PRD_GRP_2='" & Trim(package_txt) & "' AND WPRD_PRD_GRP_3='" & Trim(ld_txt) & "' " & _
'                     "AND B.WRTE_ROUTE=A.WPRD_FRST_RTE AND B.WRTE_RT_GRP_1='TEST' AND (A.WPRD_PROD LIKE '%TST%' OR A.WPRD_PROD LIKE '%DTP%' OR A.WPRD_PROD LIKE '%RTP%' OR A.WPRD_PROD LIKE '%RETEST%' OR A.WPRD_PROD LIKE '%UTTMPBF%' OR A.WPRD_PROD LIKE '%RWK%')"
'AIMS
'quah 2012-12-12 exclude inactive device
            CSQLSTRING = "SELECT PDM_DEVICENO FROM BAIC_PRODMAST WHERE  (pdm_inactive_date ='' or pdm_inactive_date is null) and PDM_CUSTOMER='" & Trim(custnameselect.TEXT) & "' AND PDM_PACKAGE='" & Trim(package_txt) & "' AND PDM_LEAD='" & Trim(ld_txt) & "' AND " & _
            "(PDM_DEVICENO LIKE '%TST%' or PDM_DEVICENO LIKE '%DIE%' OR PDM_DEVICENO LIKE '%RTP%' OR PDM_DEVICENO LIKE '%RETEST%' OR PDM_DEVICENO LIKE '%RWK%') " & _
            "ORDER BY PDM_DEVICENO"

            Debug.Print CSQLSTRING
            DevRSZ.Open CSQLSTRING, wsDB
            If Not DevRSZ.EOF Then
                Do While Not DevRSZ.EOF
                    V = DevRSZ("pdm_deviceno")
                    bdcombo.AddItem V
                    DevRSZ.MoveNext
                Loop
            End If
            DevRSZ.Close
            Set DevRSZ = Nothing
            bdcombo.SetFocus
        Else
            Dim rsbd As ADODB.Recordset
            Set rsbd = New ADODB.Recordset
            Dim rsDEV As ADODB.Recordset
            Set rsDEV = New ADODB.Recordset
            lbldeviceno.Visible = False
            bdcombo.Visible = True
            bdcombo.Clear
            
            CSQLSTRING = "SELECT * FROM aic_bd_no WHERE AICBD_CUSBD_NUMBER = '" & Trim(bonding_diagram_txt) & "' AND AICBD_STATUS='INUSE'"
            rsbd.Open CSQLSTRING, wsDB
            Debug.Print CSQLSTRING
              
            Do While Not rsbd.EOF
                '******** FOR REVISION WITH / **************
                BDREV = Trim(rsbd("AICBD_BD_NUMBER")) + "/" + Trim(rsbd("AICBD_REVISION"))
                Debug.Print BDREV
                
      '          CSQLSTRING = "SELECT * FROM WIPPRD WHERE  WPRD_PRD_GRP_2 = '" & Trim(package_txt) & "' AND  WPRD_PRD_GRP_3 = '" & Trim(ld_txt) & "' AND  WPRD_USRDF_SMDAT_2 = '" & BDREV & "'"
'AIMS
'Quah 2012-12-12 exclude inactive device
                CSQLSTRING = "SELECT PDM_DEVICENO WPRD_PROD FROM BAIC_PRODMAST WHERE  (pdm_inactive_date ='' or pdm_inactive_date is null) and PDM_PACKAGE = '" & Trim(package_txt) & "' AND  PDM_LEAD = '" & Trim(ld_txt) & "' AND  PDM_INTERNAL_BD = '" & BDREV & "' order by PDM_DEVICENO"
                rsDEV.Open CSQLSTRING, wsDB
                Debug.Print CSQLSTRING
                If Not rsDEV.EOF Then
                    Do While Not rsDEV.EOF
                        V = rsDEV("WPRD_PROD")
                        bdcombo.AddItem V
                        rsDEV.MoveNext
                    Loop
                End If
                rsDEV.Close
                '******** FOR REVISION WITH - **************
                BDREV = Trim(rsbd("AICBD_BD_NUMBER")) + "-" + Trim(rsbd("AICBD_REVISION"))
                Debug.Print BDREV
   '             CSQLSTRING = "SELECT * FROM WIPPRD WHERE  WPRD_PRD_GRP_2 = '" & Trim(package_txt) & "' AND  WPRD_PRD_GRP_3 = '" & Trim(ld_txt) & "' AND  WPRD_USRDF_SMDAT_2 = '" & BDREV & "'"
'aims
'Quah 2012-12-12 exclude inactive device
                CSQLSTRING = "SELECT pdm_deviceno wprd_prod FROM baic_prodmast WHERE  (pdm_inactive_date ='' or pdm_inactive_date is null) and pdm_package = '" & Trim(package_txt) & "' AND pdm_lead = '" & Trim(ld_txt) & "' AND  pdm_internal_bd = '" & BDREV & "'"
                
                rsDEV.Open CSQLSTRING, wsDB
                Debug.Print CSQLSTRING
                If Not rsDEV.EOF Then
                    Do While Not rsDEV.EOF
                        V = rsDEV("WPRD_PROD")
                        bdcombo.AddItem V
                        rsDEV.MoveNext
                    Loop
                End If

                rsDEV.Close
                bdcombo.SetFocus
                rsbd.MoveNext
            Loop
            rsbd.Close
            Set rsbd = Nothing
        End If
    End If
End Sub

Private Sub bottom_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        d = Index + 1
        If d < 6 Then
            If bottom(d).Enabled = True Then
                bottom(d).SetFocus
            End If
        End If
    End If
End Sub

Private Sub CANCEL_OPT_Click()
    If CANCEL_OPT.Value = True Then
        SAVEREC.Caption = "CANCEL LI"
        cmdUpdate.Enabled = False
        SAVEREC.Enabled = True
    End If
End Sub

Private Sub cbobonding_diagram_Click()
bdcombo.Clear
bonding_diagram_txt.TEXT = Trim(cbobonding_diagram.List(cbobonding_diagram.ListIndex))
bonding_diagram_txt.Visible = True
optBD.Value = True
optTargetDevice.Value = False
cbobonding_diagram.Visible = False
DEVICECOUNT = 0
        If bonding_diagram_txt = "N/A" Then
            lbldeviceno.Visible = False
            bdcombo.Visible = True
            bdcombo.Clear
            lbltestonly.Visible = True
            lbltestonly = "TEST ONLY"
            Dim DevRSZ As ADODB.Recordset
            Set DevRSZ = New ADODB.Recordset
'            CSQLSTRING = "SELECT A.* ,B.WRTE_ROUTE FROM WIPPRD A,WIPRTE B " & _
'                     "WHERE WPRD_PRD_GRP_2='" & Trim(package_txt) & "' AND WPRD_PRD_GRP_3='" & Trim(ld_txt) & "' " & _
'                     "AND B.WRTE_ROUTE=A.WPRD_FRST_RTE AND B.WRTE_RT_GRP_1='TEST' AND (A.WPRD_PROD LIKE '%TST%' OR A.WPRD_PROD LIKE '%.DTP%' OR A.WPRD_PROD LIKE '%.RTP%' OR A.WPRD_PROD LIKE '%.RETEST%' OR A.WPRD_PROD LIKE '%.UTTMPBF%' OR A.WPRD_PROD LIKE '%.RWK%')"
            
'aims
'Quah 2012-12-12 exclude inactive device
            CSQLSTRING = "SELECT * FROM baic_prodmast " & _
                     " WHERE pdm_package='" & Trim(package_txt) & "' AND pdm_lead='" & Trim(ld_txt) & "' " & _
                     " AND PDM_CUSTOMER='" & Trim(custnameselect.TEXT) & "'  and (pdm_inactive_date ='' or pdm_inactive_date is null)"

            DevRSZ.Open CSQLSTRING, wsDB
            If Not DevRSZ.EOF Then
                Do While Not DevRSZ.EOF
                    V = DevRSZ("PDM_DEVICENO")
                    bdcombo.AddItem V
                    DEVICECOUNT = DEVICECOUNT + 1
                    DevRSZ.MoveNext
                Loop
            End If
            DevRSZ.Close
            Set DevRSZ = Nothing
            bdcombo.SetFocus
        Else
            Dim rsbd As ADODB.Recordset
            Set rsbd = New ADODB.Recordset
            Dim rsDEV As ADODB.Recordset
            Set rsDEV = New ADODB.Recordset
            lbldeviceno.Visible = False
            bdcombo.Visible = True
            bdcombo.Clear
            
'            CSQLSTRING = "SELECT * FROM aic_bd_no WHERE AICBD_CUSBD_NUMBER = '" & Trim(bonding_diagram_txt) & "' AND AICBD_STATUS='INUSE' AND AICBD_TARGET_DEVICE='" & Trim(TARGET_DEVICE_TXT) & "'"
            '2011-01-24 add customer name to match
            CSQLSTRING = "SELECT * FROM aic_bd_no WHERE AICBD_CUSTOMER_NAME = '" & Trim(custnameselect.TEXT) & "' AND AICBD_CUSBD_NUMBER = '" & Trim(bonding_diagram_txt) & "' AND AICBD_STATUS='INUSE' AND AICBD_TARGET_DEVICE='" & Trim(TARGET_DEVICE_TXT) & "'"
            
                'Quah 20170801 testrun for MICROCHIOP
                 'CSQLSTRING = " SELECT * FROM aic_bd_no WHERE AICBD_CUSTOMER_NAME = 'MICROCHIP' and AICBD_CUSBD_NUMBER = 'W296579XUP REV.A' AND AICBD_STATUS='INUSE' AND AICBD_TARGET_DEVICE='AT29657-098T-10P'"
                 'CSQLSTRING = "  SELECT * FROM aic_bd_no WHERE AICBD_CUSTOMER_NAME = 'ATMEL COL DC' and AICBD_CUSBD_NUMBER = 'W296579XUP REV.A' AND AICBD_STATUS='INUSE' AND AICBD_TARGET_DEVICE='AT29657-_9_T-_ _P'"
            
            rsbd.Open CSQLSTRING, wsDB
            Debug.Print CSQLSTRING
            
            'Quah 2011-01-24 if cannot find, dont match by custname
            If rsbd.EOF Then
                rsbd.Close
                CSQLSTRING = "SELECT * FROM aic_bd_no WHERE AICBD_CUSBD_NUMBER = '" & Trim(bonding_diagram_txt) & "' AND AICBD_STATUS='INUSE' AND AICBD_TARGET_DEVICE='" & Trim(TARGET_DEVICE_TXT) & "'"
                rsbd.Open CSQLSTRING, wsDB
                Debug.Print CSQLSTRING
                
                'Quah 20090618 if cannot match by targetdevice, then match by bd only.
                If rsbd.EOF Then
                    rsbd.Close
                    CSQLSTRING = "SELECT * FROM aic_bd_no WHERE AICBD_CUSTOMER_NAME = '" & Trim(custnameselect.TEXT) & "' AND AICBD_CUSBD_NUMBER = '" & Trim(bonding_diagram_txt) & "' AND AICBD_STATUS='INUSE'"
                    Debug.Print CSQLSTRING
                    rsbd.Open CSQLSTRING, wsDB
                End If
            End If
              
              
            XAICBD_DEVICE = ""
            Do While Not rsbd.EOF
                '******** FOR REVISION WITH / **************
                BDREV = Trim(rsbd("AICBD_BD_NUMBER")) + "/" + Trim(rsbd("AICBD_REVISION"))
                XAICBD_DEVICE = Trim(rsbd("AICBD_DEVICE"))
    '            CSQLSTRING = "SELECT * FROM WIPPRD WHERE WPRD_DESC = '" & Trim(TARGET_DEVICE_TXT) & "'  AND WPRD_PRD_GRP_2 = '" & Trim(package_txt) & "' AND  WPRD_PRD_GRP_3 = '" & Trim(ld_txt) & "' AND  WPRD_USRDF_SMDAT_2 = '" & BDREV & "'"
'aims
'Quah 2012-12-12 exclude inactive device
                CSQLSTRING = "SELECT * FROM baic_prodmast WHERE pdm_targetdevice = '" & Trim(TARGET_DEVICE_TXT) & "'  AND pdm_package = '" & Trim(package_txt) & "' AND  pdm_lead = '" & Trim(ld_txt) & "' AND  pdm_internal_bd = '" & BDREV & "' and (pdm_inactive_date ='' or pdm_inactive_date is null)"

                Debug.Print CSQLSTRING
                rsDEV.Open CSQLSTRING, wsDB
                If Not rsDEV.EOF Then
                    Do While Not rsDEV.EOF
                        V = rsDEV("WPRD_PROD")
                        bdcombo.AddItem V
                        DEVICECOUNT = DEVICECOUNT + 1
                        rsDEV.MoveNext
                    Loop
                End If
                rsDEV.Close
                '******** FOR REVISION WITH - **************
                BDREV = Trim(rsbd("AICBD_BD_NUMBER")) + "-" + Trim(rsbd("AICBD_REVISION"))
   '             CSQLSTRING = "SELECT * FROM WIPPRD WHERE  WPRD_DESC = '" & Trim(TARGET_DEVICE_TXT) & "' AND WPRD_PRD_GRP_2 = '" & Trim(package_txt) & "' AND  WPRD_PRD_GRP_3 = '" & Trim(ld_txt) & "' AND  WPRD_USRDF_SMDAT_2 = '" & BDREV & "'"
'aims
'Quah 2012-12-12 exclude inactive device
                CSQLSTRING = "SELECT * FROM baic_prodmast WHERE  pdm_targetdevice = '" & Trim(TARGET_DEVICE_TXT) & "' AND pdm_package = '" & Trim(package_txt) & "' AND  pdm_lead = '" & Trim(ld_txt) & "' AND  pdm_internal_bd = '" & BDREV & "' and (pdm_inactive_date ='' or pdm_inactive_date is null)"
                
                
                Debug.Print CSQLSTRING
                rsDEV.Open CSQLSTRING, wsDB
                If Not rsDEV.EOF Then
                    Do While Not rsDEV.EOF
                        V = rsDEV("pdm_deviceno")
                        bdcombo.AddItem V
                        DEVICECOUNT = DEVICECOUNT + 1
                        rsDEV.MoveNext
                    Loop
                End If

                rsDEV.Close
                bdcombo.SetFocus
                rsbd.MoveNext
            Loop
            rsbd.Close
            Set rsbd = Nothing
            'End If
        End If
            
            If DEVICECOUNT = 0 Then
'                 MsgBox "No Product Found. Please check VPRD in Workstream.", vbCritical, "Error"
                 MsgBox "No Product Found. Please check Product Master Setup.", vbCritical, "Error"
            End If
            
              
End Sub

Private Sub cboCust_Click()
Dim rsCnt As ADODB.Recordset
Dim sqlstr As String
Dim adocounter As ADODB.Recordset
Dim sqlstr2 As String

Call RESET


If Trim(cboCust) <> "" Then
      Set rsCnt = New ADODB.Recordset
      sqlstr = "SELECT * FROM AIC_LI_CUST_MASTER " _
             & " WHERE CUSTNO = " & Left(Right(Trim(cboCust.TEXT), 4), 3)
      rsCnt.Open sqlstr, wsDB
    
        namedash = InStr(Trim(cboCust.TEXT), " - (")
        custnameselect.TEXT = Trim(Left(Trim(cboCust.TEXT), namedash))
        
      If Not rsCnt.EOF Then
        'prefix = Trim(rsCnt("prefix"))
        prefix = Trim(rsCnt!prefix)
        Me.custprefix = prefix          '2011-10-07
        lblCustomerCodeOra.Caption = Trim(rsCnt!ORA_CUSTCODE)
      Else
        MsgBox "Prefix not found in customer master!", vbCritical, "Message"
        Exit Sub
      End If
      rsCnt.Close
  
  
'2010-12-02 hide ww for TMTECH, to avoid Planner copying for marking.
If prefix = "TH" Then   'TMTECH
    Label27.Visible = False
    ww_txt.Visible = False
Else
    Label27.Visible = True
    ww_txt.Visible = True
End If
  
'20160811 Block for FITIPOWER-BJ, req  by LY, due to no forecast and label spec W1-C025 not ready.
If prefix = "FJ" Then
      MsgBox "Customer Spec W1-C025 not ready. Refer Test Engr. ", vbCritical, "Message"
      Call RESET
      refno_TXT.TEXT = vbNullString
      Exit Sub
End If
  
'20160817 Quah temp block for UBIQ, due to not yet implement for TT to show lots < 84 days for combine.
'20161220 Quah unblock for loading.
If prefix = "UB" Then
'      MsgBox "Pls refer IT, for 84 days combine lot control.", vbCritical, "Message"
'      Call RESET
'      refno_TXT.TEXT = vbNullString
'      Exit Sub
End If
  
  
'  If prefix = "GM" Or prefix = "MT" Or prefix = "SG" Or prefix = "SL" Then
  If prefix = "MT" Or prefix = "SG" Or prefix = "SL" Then
      MsgBox "This screen is not for this customer. Please select the correct screen from the top menu.", vbCritical, "Message"
      Call RESET
      refno_TXT.TEXT = vbNullString
      Exit Sub
  End If
  If prefix = "NN" Then
      MsgBox "For NS-Melaka, please key-in or scan-in the Customer Lot No. now.", vbInformation, "Message"
  End If
  
  Set adocounter = New ADODB.Recordset
  
  YDCT = Format(Date, "YY")
  MDCT = Format(Date, "MM")
  
  DDCT = Format(Date, "DD")
  sqlstr2 = "select * from AIC_LI_COUNTER" _
          & " where customer_no ='" & Left(Right(Trim(cboCust.TEXT), 4), 3) & "'" _
          & " AND M_COUNTER='" & MDCT & "' AND Y_COUNTER='" & YDCT & "'" _
          & " AND PREFIX = '" & prefix & "' "
  adocounter.Open sqlstr2, wsDB
  
  If Not adocounter.EOF Then
    refno_TXT = Trim(adocounter!prefix) & YDCT & MDCT & Format((adocounter!Counter + 1), "000")
  Else
    refno_TXT = prefix & YDCT & MDCT & "001"
  End If
  adocounter.Close
  Set adocounter = Nothing
    NEWFLAG = True
'Else
' MsgBox " Please enter customer no !", vbInformation
' Exit Sub

    '2011-06-08
    If prefix = "FS" Or prefix = "FP" Then
        fsfom.Visible = True
        fsfom.Enabled = True
    Else
        fsfom.Visible = False
        fsfom.Enabled = False
    End If
        
'Quah 2013-01-09 AVT=NIKO
If custnameselect.TEXT = "NIKO" Or custnameselect.TEXT = "AVT" Then
    txtCusLot = "(ASSY#L2#)"
End If

'Quah 2013-11-29 DIODES
'Quah 2014-01-16 removed
'If custnameselect.TEXT = "DIODES" Then
'    txtCusLot = "#AICLOT#"
'End If
    
End If

End Sub
Private Sub cboCust_KeyPress(KeyAscii As Integer)
Dim rsCnt As ADODB.Recordset
Dim sqlstr As String
Dim adocounter As ADODB.Recordset
Dim sqlstr2 As String

If Trim(cboCust) <> "" Then
  Set rsCnt = New ADODB.Recordset
  sqlstr = "SELECT PREFIX FROM AIC_LI_CUST_MASTER " _
         & " WHERE CUSTNO = " & Left(Right(Trim(cboCust.TEXT), 4), 3)
  rsCnt.Open sqlstr, wsDB

  If Not rsCnt.EOF Then
    prefix = Trim(rsCnt("prefix"))
  Else
    MsgBox "Prefix not found in customer master!"
  End If
  rsCnt.Close
  
  Set adocounter = New ADODB.Recordset
  
  YDCT = Format(Date, "YY")
  MDCT = Format(Date, "MM")
  DDCT = Format(Date, "DD")
  sqlstr2 = "select * from AIC_LI_COUNTER" _
          & " where customer_no ='" & Left(Right(Trim(cboCust.TEXT), 4), 3) & "'" _
          & " AND M_COUNTER='" & MDCT & "' AND Y_COUNTER='" & YDCT & "'" _
          & " AND PREFIX = '" & prefix & "' "
  adocounter.Open sqlstr2, wsDB
  
  If Not adocounter.EOF Then
    refno_TXT = Trim(adocounter!prefix) & YDCT & MDCT & Format((adocounter!Counter + 1), "000")
  Else
    refno_TXT = prefix & YDCT & MDCT & "001"
  End If
  adocounter.Close
  Set adocounter = Nothing
'Else
' MsgBox " Please enter customer no !", vbInformation
' Exit Sub
End If

End Sub

Private Sub cboCust_LostFocus()
    'Quah 2014-11-24 indicate for TWN/NON-TWN
    Set rsCnt = New ADODB.Recordset
    sqlstr = "select CUS_DESTINATION from BAIC_CUSTOMER where cus_shortname='" & custnameselect.TEXT & "'"
    rsCnt.Open sqlstr, wsDB
    If rsCnt.EOF = False Then
        txtCustomer.TEXT = Trim(rsCnt!CUS_DESTINATION)
    End If
    rsCnt.Close
    For Xr = 0 To 5
        If txtCustomer.TEXT = "TAIWAN" Or txtCustomer.TEXT = "CHINA" Or InStr(cboruntype, "Engineering") > 0 Then   '2014-11-18
            topx(Xr).Enabled = True
            bottom(Xr).Enabled = True
        Else
            topx(Xr).Enabled = False
            bottom(Xr).Enabled = False
        End If
    Next Xr
    
    'Quah 2014-11-24
    txtMarkingType.TEXT = ""
    If prefix = "BE" Then
        txtMarkingType.TEXT = "INPUT"
    End If

    If Left(refno_TXT, 2) = "MC" Then
        If target_device_txt1 = "" Then
            optTargetDevice = True
            target_device_txt1.TEXT = Trim(TARGET_DEVICE_TXT)
        End If
    End If


End Sub

Private Sub cbofullpackage_Click()
'Quah 20090223 auto initialze short package name
    Set rs = New ADODB.Recordset
    
'    sqltext = " select distinct WPRD_PRD_GRP_4, WPRD_PRD_GRP_2, WPRD_PRD_GRP_3 from wipprd where wprd_prd_grp_4 in " & _
'            " (select PACKAGLEAD_SHORT from aic_packagelead_master where packagelead_full='" & Trim(cbofullpackage) & "') AND " & _
'            " WPRD_PRD_GRP_2 <> 'NA'"
'aims
    sqltext = " select distinct pdm_packagelead WPRD_PRD_GRP_4, pdm_package WPRD_PRD_GRP_2, pdm_lead WPRD_PRD_GRP_3 from baic_prodmast where pdm_packagelead = " & _
            " '" & Trim(cbofullpackage) & "' AND " & _
            " pdm_package <> 'NA'"

    Debug.Print sqltext
    rs.Open sqltext, wsDB
    Debug.Print sqltext
    If Not rs.EOF Then
        package_txt.TEXT = Trim(rs!WPRD_PRD_GRP_2)
        ld_txt.TEXT = Trim(rs!WPRD_PRD_GRP_3)
    Else
        MsgBox "Product Package not registered in AIMS. Please check VPRD for this Product.", vbCritical, "Error"
        Exit Sub
        'Quah
    End If
    rs.Close
    Set rs = Nothing
    FG = Trim(package_txt) & Trim(ld_txt)
    If FG = "PDIP32" Or FG = "PDIP40" Then
        bottom(3).Enabled = False
        bottom(4).Enabled = False
        bottom(5).Enabled = False
     Else
        bottom(3).Enabled = True
        bottom(4).Enabled = True
        bottom(5).Enabled = True
     End If
'     bonding_diagram_txt.Visible = True
'     bonding_diagram_txt.SetFocus
'     target_device_txt1.Visible = False
'end. Quah 20090223 auto initialize short package name

End Sub

Private Sub cboruntype_Click()
'If Left(Trim(refno_TXT), 2) = "AN" Or Left(Trim(refno_TXT), 2) = "SN" Then
'If cboruntype.TEXT = "Mass Production" Then
'    If yy1 = "2009" Then
'        If Left(Trim(txtCusLot), 2) = "XL" Then
'            ' ok
'        Else
'            MsgBox ("ANPEC/SINOPOWER Customer Lot No. for year 2009 must start with XL !" & Chr(13) & "X=ANPEC, L=2009")
'            txtCusLot.TEXT = ""
'            Exit Sub
'        End If
'    Else
'            MsgBox ("ANPEC/SINOPWOER Customer Lot No. prefix not yet defined for this year  !" & Chr(13) & "X=ANPEC, L=2009")
'            txtCusLot.TEXT = ""
'            Exit Sub
'    End If
'End If

'If Trim(cboruntype) = "Engineering E1" Then
'    MsgBox "Please inform IT for E1 loading.", vbInformation, "Message"
'End If

'2012-03-13
If Left(Trim(refno_TXT), 2) = "FS" Or Left(Trim(refno_TXT), 2) = "FP" Then
    If cboruntype.TEXT = "Mass Production" Then
        fsfom.Enabled = True
    Else
        txtCusLot.TEXT = ""
        fsfom.Enabled = False
        MsgBox "For Fairchild engr lot (EON), use FAC Pid# from cust.", vbInformation, "Message"
    End If
End If

End Sub

Private Sub Check1_Click()
If Check1.Value = Checked Then
    If Left(Trim(refno_TXT), 2) <> "BE" Then
        MsgBox "OPTION FOR BOURNS ONLY!", vbCritical
    Else
        If Text3 = "" Then
            Dim bescc As ADODB.Recordset
            Set bescc = New ADODB.Recordset
            ssql = " SELECT distinct WAFERLOTNO FROM AIC_INVENTORY_MASTER " & _
                   " WHERE customername='BOURNS' and (STATUS like '%RECEIVED' or STATUS like '%BUYOFF%') " & _
                   " order by WAFERLOTNO ASC"
            bescc.Open ssql, wsDB
            Do While Not bescc.EOF
                be_waferlotno.AddItem Trim(bescc!WAFERLOTNO)
                bescc.MoveNext
            Loop
            bescc.Close
            Set bescc = Nothing
            
            be_listscribe.ListItems.Clear
            be_total = ""
            scribe = ""
            Frame_BE.Visible = True
        Else
            'autofill
            be_waferlotno = Text3
            be_waferlotno.Locked = True
            
            be_listscribe.ListItems.Clear
            be_total = ""
            scribe = ""
            Call be_waferlotno_KeyPress(13)
            Frame_BE.Visible = True
        End If
    End If
Else
    be_waferlotno.Locked = False
    be_waferlotno.Clear
    be_listscribe.ListItems.Clear
    be_total = ""
    scribe = ""
    Frame_BE.Visible = False
End If
End Sub

Private Sub chkShow_Click()
If chkShow.Value = Checked Then
    
    If RELEASE_OPT = False And CANCEL_OPT = False Then
        
        '2012-02-20 GET DBANK QTY FOR DISPLAY/VERIFICATIONS.
        listdbinv.ListItems.Clear
        Dim getdbrs As ADODB.Recordset
        Set getdbrs = New ADODB.Recordset
        ssql = " SELECT WAFERLOTNO, STATUS, DIEBANKQTY, IQA_REC_DATE FROM AIC_INVENTORY_MASTER " & _
               " WHERE customername='" & custnameselect.TEXT & "' and (STATUS like '%RECEIVED' or STATUS like '%BUYOFF%') and DIEBANKQTY IS NOT NULL " & _
               " order by WAFERLOTNO"
        Debug.Print ssql
        getdbrs.Open ssql, wsDB
        Do While Not getdbrs.EOF
            Set itmx = listdbinv.ListItems.Add(1, , lvwDieCnt)
            itmx.SubItems(1) = Trim(getdbrs!WAFERLOTNO)
            itmx.SubItems(2) = Trim(getdbrs!Status)
            itmx.SubItems(3) = Trim(getdbrs!diebankqty)
            itmx.SubItems(4) = Trim(getdbrs!IQA_REC_DATE)
            getdbrs.MoveNext
        Loop
    End If
    '---------------------------------------------------------------------
    
    Frame1.Visible = True
    Text1.SetFocus
Else
    Frame1.Visible = False
End If

                    'DIANA 2014-10-03 TME, topx2 get from workweek last 3 digit (for 2 t.device)
                    'DIANA 2014-10-10 TME, FIX BUG ok
                    If Left(Trim(refno_TXT), 2) = "TC" Then
                        If TARGET_DEVICE_TXT = "EPA2018A" Or TARGET_DEVICE_TXT = "EPA2010A" Then
                            'topx(3).TEXT = Right(Trim(ww_txt), 3) & "A"
                            WX = Right(Trim(ww_txt), 3)
                            topx(3).TEXT = Replace(topx(3).TEXT, "YWW", WX)
                        End If
                    End If
End Sub

Private Sub cmd_MFG_Click()
'  frm_MfgInventory.Show
End Sub

Private Sub cmdFullReport_Click()
CheckMaterialBalance "- ALL -"
End Sub

Private Sub cmdPrint_Click()
chkUpdate

Dim sql2 As String
Dim rsM As ADODB.Recordset
Dim jobno As String
Dim ORACODE As String

    If Trim(cboCust.TEXT) = "" Then
        MsgBox " PLEASE INSERT customer code!!!", vbCritical
        cboCust.SetFocus
        Exit Sub
    End If
    If Trim(refno_TXT) = "" Then
        MsgBox " PLEASE INSERT REFERENCE NO!!!", vbCritical
        refno_TXT.SetFocus
    
    Else
        'check li
        SQL$ = "select count(*) as ab from AIC_LOADING_INSTRUCTION where REFNO  = '" & Trim(refno_TXT) & "'"
        Set RSQ = New ADODB.Recordset
        RSQ.Open SQL$, wsDB
        If Val(RSQ!ab) < 1 Then
            MsgBox "LI Not Generated Yet!!", vbInformation, "Error"
            Exit Sub
        End If
        SQL$ = "select count(*) as ac from AIC_LOADING_INSTRUCTION_REMARK where REFNO  = '" & Trim(refno_TXT) & "'"
        Set RSQ = New ADODB.Recordset
        RSQ.Open SQL$, wsDB
        If Val(RSQ!ac) < 1 Then
            MsgBox "Please Key-In Additional Information Before Print!!", vbInformation, "Error"
            Exit Sub
        End If
        RSQ.Close
        
        SQL$ = " select count(*) as ad from aic_li_dual_die where refno = '" & Trim(refno_TXT) & "'"
        RSQ.Open SQL$, wsDB
        If Val(RSQ!ad) < 1 Then
          MsgBox " Please key in die information before print!!", vbInformation, "ERROR"
          Exit Sub
        End If
        RSQ.Close
        
        Dim liFile As ADODB.Recordset
        Dim AICX As ADODB.Recordset
        Set AICX = New ADODB.Recordset
        Set liFile = New ADODB.Recordset
        Dim aic As TableDef
        
        Set rsM = New ADODB.Recordset
     
        dtlx = App.Path & "\report\liDB.MDB"
        Set DTL = OpenDatabase(dtlx, False, False)
            
        DTL.Execute "DELETE FROM liTEMPX"
        DTL.Execute "DELETE FROM liADD"
        DTL.Execute "DELETE FROM MARK_SPEC"
        DTL.Execute "DELETE FROM TB_LABEL"
        DTL.Execute "DELETE FROM LABEL_INFO"
        DTL.Execute "DELETE FROM DUAL_DIE"
        DTL.Execute "DELETE FROM CUST_PO"
        
        sql2 = " SELECT ORA_CUSTCODE FROM AIC_LI_CUST_MASTER " & _
               " WHERE CUSTNO = '" & Left(Right(Trim(cboCust.TEXT), 4), 3) & "'"
        rsM.Open sql2, wsDB
        If Not rsM.EOF Then
          ORACODE = Trim(rsM("ora_custcode"))
        Else
          MsgBox " No oracle code found!"
        End If
        
        wsSqlString = "select * from AIC_LOADING_INSTRUCTION" _
          & " where REFNO  = '" & refno_TXT & "'"
        liFile.Open wsSqlString, wsDB
        If liFile.EOF = False Then
            If LenB(Trim(liFile!WAFER)) = 0 Then
                MsgBox "DIE INFO IS NOT COMPLETE.", vbInformation
                liFile.Close
                Exit Sub
            End If
            
            'SPECIAL CHARACTER CONTROL
            If LenB(Trim(liFile!SpecChar)) <> 0 Then
                STRSTRING = "INSERT INTO SPEC_CHAR(FIELD1, FIELD2, FIELD3, FIELD4, FIELD5) " & _
                            "values('" & Left(liFile!SpecChar, 1) & "', '" & Mid(liFile!SpecChar, 2, 1) & "', " & _
                            " '" & Mid(liFile!SpecChar, 3, 1) & "', '" & Mid(liFile!SpecChar, 4, 1) & "', " & _
                            "'" & Mid(liFile!SpecChar, 5, 1) & "')"
            DTL.Execute STRSTRING
            End If
            
dbpackage = liFile("PACKAGE__LEAD")
dbfullpackage = dbpackage

'Quah 20090223 auto initialze long package name
'disable 20100923
'    Set RS = New ADODB.Recordset
'    sqltext = " select packagelead_full from aic_packagelead_master where PACKAGLEAD_SHORT='" & Trim(dbpackage) & "'"
'    Debug.Print sqltext
'    RS.Open sqltext, wsDB
'    If Not RS.EOF Then
'        dbfullpackage = Trim(RS!PACKAGELEAD_FULL)
'    End If
'    RS.Close
'    Set RS = Nothing
'end. Quah 20090223 auto initialize long package name
'Kamizan check here for label crystal
            
            
            With aic

                STRSTRING = "INSERT INTO liTEMPX VALUES ('" & liFile("DEVICE_NO") & "','" & _
                                          dbfullpackage & "','" & _
                                          (liFile("TARGET_DEVICE")) & "','" & _
                                          (liFile("WAFER")) & "'," & _
                                          (liFile("QTY")) & ",'" & _
                                          (liFile("BD_NO")) & "','" & _
                                          (liFile("INTERNAL_BD")) & "','" & _
                                          (liFile("CUSTOMER_NO")) & "','" & _
                                          (liFile("CUSTOMER_NAME")) & "','" & _
                                          (liFile("MARKING_SPEC")) & "','" & _
                                          (liFile("TOP1")) & "','" & _
                                          (liFile("TOP2")) & "','" & _
                                          (liFile("TOP3")) & "','" & _
                                          (liFile("TOP4")) & "','" & _
                                          (liFile("TOP5")) & "','" & _
                                          (liFile("TOP6")) & "','" & _
                                          (liFile("BOTTOM1")) & "','" & _
                                          (liFile("BOTTOM2")) & "','" & _
                                          (liFile("BOTTOM3")) & "','" & _
                                          (liFile("BOTTOM4")) & "','" & _
                                          (liFile("BOTTOM5")) & "','" & liFile("BOTTOM6") & "','" & _
                                          (liFile("DATE_TRANX")) & "'," & liFile("TIME_TRANX") & ",'" & _
                                          (liFile("EMP_ID")) & "','" & liFile("STATUS") & "','" & liFile("REFNO") & "','" & _
                                          (liFile("WORK_WEEK")) & "','" & liFile("CUSLOTNO") & "','" & liFile("PO_NO") & "','" & _
                                           ORACODE & "','')"


                DTL.Execute STRSTRING
            End With
            
            jobno = liFile("cuslotno")
            'Printing Daily Transfer Logsheet
        End If
        liFile.Close
        
        'Add dual die information
        
        wsSqlString = "select * from AIC_LI_DUAL_DIE " _
                    & " where REFNO  = '" & refno_TXT & "'" _
                    & " order by seqno "
        liFile.Open wsSqlString, wsDB
        Dim LIDEX As String
        While Not liFile.EOF
          If IsNull(liFile!LIDIE_QTY) = True Then
            LIDIEX = "-"
          Else
            LIDIEX = Trim(liFile!LIDIE_QTY)
          End If
                  
        STRSTRING = " INSERT INTO DUAL_DIE " & _
                    " VALUES('" & liFile("refno") & "', " & _
                    " '" & liFile("partno") & "','" & liFile("waferno") & "', " & _
                    " " & liFile("wafer_qty") & "," & liFile("die_qty") & "," & liFile("seqno") & ", " & _
                    " '" & Trim(LIDIEX) & "')"
                                                                  
          DTL.Execute STRSTRING
          liFile.MoveNext
        Wend
        liFile.Close
        
        'add additional instruction
        wsSqlString = "select * from AIC_LOADING_INSTRUCTION_REMARK" _
          & " where REFNO  = '" & refno_TXT & "'"
        liFile.Open wsSqlString, wsDB
        If liFile.EOF = False Then
            If Trim(liFile!URGENTLOT_FLAG) = "Y" Then
                urgflag = "URGENT LOT"
            Else
                urgflag = " "
            End If
            If Trim(liFile!DS_FLAG) = "Y" Then
                dsxflag = "DS"
            Else
                dsxflag = " "
            End If
            If Trim(liFile!ENGLOT_FLAG) = "E" Then
                englotflag = "E"
            Else
                englotflag = "N/A"
            End If
            If Trim(liFile!QTALOT_FLAG) = "Q" Then
                qtalotflag = "Q"
            Else
                qtalotflag = "N/A"
            End If
            If Trim(liFile!TEST_FLAG) = "Y" Then
                testxflag = "TESTED"
            ElseIf Trim(liFile!TEST_FLAG) = "N" Then
                testxflag = "UNTESTED"
            End If
            With aic
                'INSERT INTO ACCESS TABLE
                STRSTRING = "INSERT INTO liADD VALUES ('" & liFile("refno") & "','" & _
                (liFile("REM1")) & "','" & (liFile("REM2")) & "','" & (liFile("REM3")) & "','" & (liFile("REM4")) & "','" & (liFile("REM5")) & "','" & _
                (liFile("REM6")) & "','" & (liFile("REM7")) & "','" & (liFile("REM8")) & "','" & (liFile("REM9")) & "','" & (liFile("REM10")) & "','" & _
                (urgflag) & "','" & (englotflag) & "','" & (qtalotflag) & "','" & (testxflag) & "','" & (dsxflag) & "','" & _
                (liFile("SINGMARK_FLAG")) & "','" & (liFile("Form")) & "','" & (liFile("SL_NO")) & "','" & (liFile("TEST_PROCESS")) & "','" & (liFile("BIN1_TOP1")) & "','" & _
                (liFile("BIN1_TOP2")) & "','" & (liFile("BIN1_TOP3")) & "','" & (liFile("BIN1_TOP4")) & "','" & (liFile("BIN1_TOP5")) & "','" & (liFile("BIN1_TOP6")) & "','" & _
                (liFile("BIN3_TOP1")) & "','" & (liFile("BIN3_TOP2")) & "','" & (liFile("BIN3_TOP3")) & "','" & (liFile("BIN3_TOP4")) & "','" & _
                (liFile("BIN3_TOP5")) & "','" & (liFile("BIN3_TOP6")) & "','" & (liFile("BIN5_TOP1")) & "','" & (liFile("BIN5_TOP2")) & "','" & (liFile("BIN5_TOP3")) & "','" & _
                (liFile("BIN5_TOP4")) & "','" & (liFile("BIN5_TOP5")) & "','" & (liFile("BIN5_TOP6")) & "','" & _
                (liFile("loadingdate")) & "','" & (liFile("wsstartdate")) & "','" & Trim(ROUTEx) & "')"
                DTL.Execute STRSTRING
            End With

            'Printing Daily Transfer Logsheet
        End If
        liFile.Close
        'Chk Mark Spec
        Dim MSpcRS As ADODB.Recordset
        Set MSpcRS = New ADODB.Recordset
        sqltext = "SELECT * FROM AIC_MARKING_SPEC_X WHERE SPECNO='" & Trim(mark_spec_txt) & "' AND APPROVAL = 'APPROVED'"
        MSpcRS.Open sqltext, wsDB
        If MSpcRS.EOF = False Then
            SPECNO = Trim(MSpcRS!SPECNO)
            TOP1 = Trim(MSpcRS!TOP1)
            TOP2 = Trim(MSpcRS!TOP2)
            TOP3 = Trim(MSpcRS!TOP3)
            TOP4 = Trim(MSpcRS!TOP4)
            TOP5 = Trim(MSpcRS!TOP5)
            TOP6 = Trim(MSpcRS!TOP6)
            BOTTOM1 = Trim(MSpcRS!BOTTOM1)
            BOTTOM2 = Trim(MSpcRS!BOTTOM2)
            BOTTOM3 = Trim(MSpcRS!BOTTOM3)
            BOTTOM4 = Trim(MSpcRS!BOTTOM4)
            BOTTOM5 = Trim(MSpcRS!BOTTOM5)
            BOTTOM6 = Trim(MSpcRS!BOTTOM6)
            sqltext = "INSERT INTO MARK_SPEC VALUES('" & Trim(refno_TXT) & "','" & SPECNO & "','" & TOP1 & "','" & TOP2 & "','" & TOP3 & "','" & TOP4 & "','" & TOP5 & "','" & TOP6 & "','" & BOTTOM1 & "','" & BOTTOM2 & "','" & BOTTOM3 & "','" & BOTTOM4 & "','" & BOTTOM5 & "','" & BOTTOM6 & "')"
            DTL.Execute sqltext
        End If
        MSpcRS.Close
        Set MSpcRS = Nothing
     
     
'insert customer po info  20101025
podatestr = Format(Me.txt_podate, "DD/MM/YYYY")
If total_poqty = "" Or total_poqty = "N/A" Then total_poqty = 0
ssql = "INSERT INTO CUST_PO (pomode,PONO,potgtdev,poqty,podate) " & _
        " VALUES('" & Trim(Me.txt_pomode) & "','" & Trim(Me.lblPONo) & "','" & Trim(Me.TARGET_DEVICE_TXT) & "'," & total_poqty & ", '" & podatestr & "')"
        Debug.Print ssql
DTL.Execute ssql
     
     
'CATHERINE 070606 - LABEL INFO CHANGE TABLE
'----------NEW INSERT LABEL INFO UPDATE 20070606-----------------------------------------------------------------
        SQL$ = "select * from AIC_LI_LABELINFO WHERE REFNO='" & Trim$(refno_TXT) & "' ORDER BY SEQNO"
        Set rsLABEL = New ADODB.Recordset
        rsLABEL.Open SQL$, wsDB, adOpenDynamic
        Do While rsLABEL.EOF = False
             ssql = "INSERT INTO LABEL_INFO(FIELD1, FIELD2, FIELD3) " & _
                    "VALUES( '" & Trim(rsLABEL!Label) & "', '" & Trim(rsLABEL!TEXT) & "',  " & _
                    " " & Trim(rsLABEL!seqno) & ")"
            DTL.Execute ssql
            rsLABEL.MoveNext
        Loop
        rsLABEL.Close
        
        Me.CrystalReport2.WindowTitle = "LOADING INSTRUCTION"
        CrystalReport2.ReportFileName = App.Path & "\Report\LoadingInstructionRpt_General.RPT"
        CrystalReport2.Destination = 1
        CrystalReport2.Action = 1
        Me.CrystalReport2.PrintReport
 
    End If
End Sub

Private Sub cmdRefresh_Click()
    Unload Me
    LI_General.Show
End Sub

Private Sub cmdUpdate_Click()

chkUpdate

'20211008
'If Trim(bom_verify) = "" Then
'    MsgBox "Please check and confirm BOM parts.", vbCritical, "Message"
'    Exit Sub
'End If


If Trim(lblPONo) = "QUAH" Then
    MsgBox "Invalid PO No.", vbCritical, "Message"
    Exit Sub
End If

'20201228 EK request, if device not running > 3 years, get ENGR confirmation first, for correct TEST PROGRAM.
Dim chklastload As ADODB.Recordset
Set chklastload = New ADODB.Recordset
ssql = "select startdate, DateDiff(Day, startdate, GETDATE()) diff from AIC_WIP_HEADER, BAIC_LOTMAST WHERE ASSYLOTNO=LTM_LOTNO AND LTM_STATUS <> 'TRLT' AND LTM_TARGETDEVICE ='" & Trim(TARGET_DEVICE_TXT.TEXT) & "' ORDER BY STARTDATE DESC"
Debug.Print ssql
chklastload.Open ssql, wsDB
If Not chklastload.EOF Then
    If chklastload!diff > 1095 Then
       MsgBox "Targetdevice last load > 3 yrs ago." & vbCrLf & "Pls confirm any changes with Engrs before loading.", vbCritical, "Message"
       If MsgBox("Targetdevice > 3 years." & vbCrLf & "Confirm to Proceed???", vbYesNo, "Message") = vbNo Then
           Exit Sub
       End If
    End If
End If
chklastload.Close

'20230523 AIN FORTUNE DATECODE
If (Left(refno_TXT, 2) = "FV" And InStr(package_lead_txt, "2X2") > 0) Then
    Dim year
    dc = Left(topx(2), 1) 'datecode 2023 must be J
    
    Dim chkdcyear As ADODB.Recordset
    Set chkdcyear = New ADODB.Recordset
    dcsql = "SELECT WW_YEAR, WW_DATECODE FROM WWCAL_FORTUNE WHERE WW_DATECODE = '" & dc & "' "
    Debug.Print dcsql
    chkdcyear.Open dcsql, wsDB
    If Not Right(chkdcyear!ww_year, 2) = Left(ww_txt, 2) Then
'        MsgBox "The datecode for " & yy1 & " is not " & dc, vbCritical, "Message"
        MsgBox "The datecode not tally with customer spec", vbCritical, "Message"
        Exit Sub
    End If

End If



'Quah 20171127, for NIKO, check the length of Cuslotno & Marking.
If (Trim(custnameselect.TEXT) = "NIKO") Then
    'Quah 2018081 skip RWK, simplify logic.
    If InStr(package_lead_txt, "2X2") > 0 Or InStr(internal_device_no_txt, "RMA") > 0 Or InStr(internal_device_no_txt, "RWK") > 0 Then
        'OK SKIP CHECKING.
        '2X2 use YWW conversion formula.
    Else
        If Len(txtCusLot) <> 17 And Len(txtCusLot) <> 9 Then
            MsgBox "Invalid length. Please check Custlotno !!", vbCritical, "Message"
            Exit Sub
        End If
    End If
End If
''''''    If InStr(package_lead_txt, "2X2") > 0 Then
''''''        'skip marking check.
''''''    ElseIf InStr(package_lead_txt, "SOIC") > 0 Then
''''''        If Len(topx(3)) <> 17 Then
''''''            If InStr(internal_device_no_txt, "RWK") > 0 Then
''''''                'OK. Skip check length
''''''            Else
''''''                MsgBox "Invalid length. Please check Marking !!", vbCritical, "Message"
''''''                Exit Sub
''''''            End If
''''''        End If
''''''    Else
''''''        If Len(topx(2)) <> 14 Then
''''''            If InStr(internal_device_no_txt, "RWK") > 0 Then
''''''                'OK. Skip check length
''''''            Else
''''''                MsgBox "Invalid length. Please check Marking !!", vbCritical, "Message"
''''''                Exit Sub
''''''            End If
''''''        End If
''''''    End If


'Quah 20170620 check for AMPHENOL unregistered APT for Port Type and Dispensing Type (for MASS PROD lots).
If Trim(custnameselect.TEXT) = "AMPHENOL" And InStr(cboruntype, "Mass Production") > 0 Then
    'Quah 20170718 skip checking for NPX & HSE, requested by Anita.
    'Quah 20170731 skip for CAP, req by Anita.
    If InStr(internal_device_no_txt, "HSE") > 0 Or InStr(internal_device_no_txt, "NPX") > 0 Or InStr(internal_device_no_txt, "CAP") > 0 Then
        'skip checking.
    Else
        Dim RS2 As ADODB.Recordset
        
        '20181101 add checking in PDM_CATEGORY
        Set RS2 = New ADODB.Recordset
        ssql = "SELECT * FROM BAIC_PRODMAST WHERE pdm_targetdevice='" & internal_device_no_txt & "' and PDM_CATEGORY='NPA'"
        Debug.Print ssql
        RS2.Open ssql, wsDB
        If Not RS2.EOF Then
            NPALOT = True
        Else
            NPALOT = False
        End If
        RS2.Close
        
        If NPALOT = True Then
            Set RS2 = New ADODB.Recordset
            ssql = "SELECT * FROM BAIC_COMTBL WHERE TBL_REC_TYPE='GEPT'" & _
                   " AND TBL_KEY_A20= '" & Trim(TARGET_DEVICE_TXT) & "'"
            Debug.Print ssql
            RS2.Open ssql, wsDB
            If RS2.EOF = True Then
                MsgBox "AMPHENOL Device Not Registered For APT." & vbCrLf & "Pls inform CUSTOMER SERVICE.", vbCritical, "Message"
                Exit Sub
            End If
            RS2.Close
        End If
    End If
End If


Dim SQL As String

''ZUL 2012-03-05
'Dim pro As ADODB.Recordset
'Set pro = New ADODB.Recordset
'SQL = "SELECT * FROM AIC_LI_LABELINFO WHERE REFNO = '" & Trim(refno_TXT) & "'"
'pro.Open SQL, wsDB
'If Not pro.EOF Then
'    MsgBox "Please verify Label Info at second page~!!", vbCritical, "Message"
'End If
'pro.Close
'Set pro = Nothing

'Quah 2012-07-24
Dim LIBlock As ADODB.Recordset
Set LIBlock = New ADODB.Recordset
ssql = "select * from BAIC_CUSTOMER_ADDSETUP where CUS_RECTYPE='BLOCK_LI' and CUS_SHORTNAME='" & Trim(custnameselect.TEXT) & "'"
Debug.Print ssql
LIBlock.Open ssql, wsDB
If Not LIBlock.EOF Then
    MsgBox "Customer blocked for IT tracking. Pls refer IT", vbCritical, "Message"
    Exit Sub
End If
  
'temporary closed because new PO not follow target device for marking 20141219
'Quah 20141103 marking check for FB, logic for other customers cannot be finalised yet.
'If Left(refno_TXT, 2) = "FB" Then
'    If InStr(TARGET_DEVICE_TXT, topx(1).TEXT) = 0 Then
'        MsgBox "Please check Marking Line2...", vbCritical, "Message"
'        Exit Sub
'    End If
'End If
  

If Trim(ww_txt) = "" Then   '2012-04-20
    MsgBox "Cannot save. Blank workweek!", vbCritical, "Message"
    Exit Sub
End If


REFNOx = Trim(refno_TXT)
If REFNOx = "" Then
    MsgBox "Cannot save. Invalid data!", vbCritical, "Message"
    Exit Sub
End If

'2012-02-08
If Trim(txtCusLot) = "" Then
    MsgBox "Cannot save. Invalid data!", vbCritical, "Message"
    Exit Sub
End If


If Left(REFNOx, 2) = "FS" Or Left(REFNOx, 2) = "FP" Then
        MsgBox "For FAIRCHILD FOM, pls ensure " & vbCrLf & "IQA has RECEIVED the wafers in FOM System!", vbInformation, "Message"

        '2012-01-12 check for XX marking (Planner forget to click on FAIRCHILD FOM button)
        If Mid(topx(1), 2, 2) = "XX" Then
            MsgBox "XX : Error in Marking?", vbCritical, "Message"
            Exit Sub
        End If

End If

'KO ADD FOR CHECK WAFERLOTNO AND TARGET DEVICE WITH IQA SYSTEM AS PER HAIGETE 20191104
If (Trim(custnameselect.TEXT) = "HAIGETE") Then
    Dim HAIiqa As ADODB.Recordset
    Set HAIiqa = New ADODB.Recordset
    Dim IQAChk
    IQAChk = "select *  FROM AIC_INVENTORY_MASTER WHERE CUSTOMERNAME='HAIGETE' and device_no ='" & Trim(TARGET_DEVICE_TXT.TEXT) & "' and waferlotno='" & Trim(WAFER) & "' "
    HAIiqa.Open IQAChk, wsDB
    If HAIiqa.EOF Then
       MsgBox " Targetdevice/Waferlotno not match with IQA System, Kindly verify again!!"
       Exit Sub
    End If
    HAIiqa.Close
End If
'20210302 Quah add for CHANGDIAN
If (Trim(custnameselect.TEXT) = "CHANGDIAN") Then
    Set HAIiqa = New ADODB.Recordset
    IQAChk = "select *  FROM AIC_INVENTORY_MASTER WHERE CUSTOMERNAME='CHANGDIAN' and device_no ='" & Trim(TARGET_DEVICE_TXT.TEXT) & "' and waferlotno='" & Trim(WAFER) & "' "
    Debug.Print IQAChk
    HAIiqa.Open IQAChk, wsDB
    If HAIiqa.EOF Then
       MsgBox " Targetdevice/Waferlotno not match with IQA System, Kindly verify again!!"
       Exit Sub
    End If
    HAIiqa.Close
End If
'KOko END FOR HAIGETE CHECK TARGET DEVICE


'check of po info changed.
If Trim(custnameselect) = liori_cust And Trim(lblPONo.TEXT) = liori_po And Trim(TARGET_DEVICE_TXT.TEXT) = liori_tgtdev Then
    'all key-info same.
Else
    'changes to keyinfo
    'if no more active LI, delete from baic_customer_po
    Dim chkcancel As ADODB.Recordset
    Set chkcancel = New ADODB.Recordset
'    SSQL = "select * from AIC_LOADING_INSTRUCTION where (STATUS='N' or STATUS='R') and CUSTOMER_NAME='" & Trim(liori_cust) & "' and TARGET_DEVICE='" & Trim(liori_tgtdev) & "' and PO_NO='" & Trim(liori_po) & "'"
    ssql = "select * from AIC_LOADING_INSTRUCTION where (STATUS='N' or STATUS='R') and  refno <> '" & Trim(refno_TXT) & "' and CUSTOMER_NAME='" & Trim(liori_cust) & "' and TARGET_DEVICE='" & Trim(liori_tgtdev) & "' and PO_NO='" & Trim(liori_po) & "'"
    Debug.Print ssql
    chkcancel.Open ssql, wsDB
    If chkcancel.EOF Then
        'no nore
        ssql = "delete from baic_customer_po wheRE CPO_CUST_SHORTNAME='" & Trim(liori_cust) & "' and CPO_PONO='" & Trim(liori_po) & "' and CPO_TARGETDEVICE='" & Trim(liori_tgtdev) & "'"
        Debug.Print ssql
        wsDB.Execute ssql
    End If
    chkcancel.Close
    Set chkcancel = Nothing
End If



If Trim(txt_pomode) = "" Then
        MsgBox "Cannot save. PO Mode cannot be blank.", vbCritical, "Message"
        Exit Sub
End If

'Quah 2010-08-18
'TMTECH if marking line 2got dot, block.
If Left(Trim(refno_TXT), 2) = "TH" Then
    If InStr(1, topx(2), ".") Then
        MsgBox "TMTECH marking cannot have DOT!", vbCritical, "Message"
        Exit Sub
    End If
End If
    
    
    
'Quah 2010-08-04 check for QFN device must register in AIMS
If InStr(package_lead_txt, "FN") > 0 Then
    Dim aimsdev As ADODB.Recordset
    Set aimsdev = New ADODB.Recordset
    ssql = "select * from baic_prodmast where pdm_deviceno='" & Trim(internal_device_no_txt) & "'"
    Debug.Print ssql
    aimsdev.Open ssql, wsDB
    If aimsdev.EOF = True Then
        MsgBox "Cannot save. Please register this device in AIMS System.", vbCritical, "Message"
        Exit Sub
    End If
    aimsdev.Close
    Set aimsdev = Nothing
End If
    
    
    If Trim(refno_TXT) = "" Then
        MsgBox " PLEASE INSERT REFERENCE NO!!!", vbCritical
        refno_TXT.SetFocus
    
    Else
        If Trim(stat) = "R" Then
            MsgBox "PARTICULAR LI CANNOT BE UPDATE AFTER RELEASE SUCCESSFULLY!", vbCritical, "ERROR"
            Exit Sub
        End If
        SQL$ = "SELECT STATUS FROM AIC_LOADING_INSTRUCTION WHERE REFNO  = '" & Trim(refno_TXT) & "'"
        Set RSQ = New ADODB.Recordset
        RSQ.Open SQL$, wsDB
        If RSQ.EOF = False Then
            If Trim(RSQ!Status) = "R" Then
                MsgBox "L.I. already released! Cannot update data!", vbCritical, "ERROR"
                Exit Sub
            End If
        End If
        
        'CHECK BD AND MARK SPEC
        'CHAR(1ST) = M : NO MARKINGSPEC CHECKING
        'If Left(Trim(refno_TXT), 2) = "OC" Or Left(Trim(refno_TXT), 2) = "SK" Or Left(Trim(refno_TXT), 2) = "AC" Or Left(Trim(refno_TXT), 2) = "HT" Then
        If Left(Right(Trim(cboCust.TEXT), 6), 1) <> "M" And Left(Trim(cboruntype), 11) <> "Engineering" Then
        'Else
            
            If Left(refno_TXT, 2) = "GM" Then
                SQL$ = "SELECT * FROM AIC_DEVICE_CONTROL WHERE CUSTOMER='CUSBD' AND TARGETDEVICE='MARKSPEC'" & _
                    " AND REMARK1='" & Trim(Me.bonding_diagram_txt) & "' AND REMARK2='" & Trim(Me.mark_spec_txt) & "'"
            Else
                SQL$ = "SELECT * FROM AIC_DEVICE_CONTROL WHERE CUSTOMER='CUSBD' AND TARGETDEVICE='MARKSPEC'" & _
                    " AND REMARK4='" & Trim(Me.internal_device_no_txt) & "' AND REMARK2='" & Trim(Me.mark_spec_txt) & "'"
            End If
            Debug.Print SQL$
            Set rs = New ADODB.Recordset
            rs.Open SQL$, wsDB
            If rs.EOF = False Then
                If Trim(rs!REMARK3) = "CLOSE" Then
                    MsgBox "PARTICULAR BONDING DIAGRAM WITH MARKING SPEC ALREADY CLOSED!PLEASE CHECK!", vbCritical, "ERROR"
                    Exit Sub
                End If
            Else
                If mark_spec_txt = "" Then  'Quah 20120912 add IF condition.
                    MsgBox "PARTICULAR BONDING DIAGRAM NOT INITIALIZE WITH MARKING SPEC!PLEASE CHECK!", vbCritical, "ERROR"
                    Exit Sub
                End If
            End If
        End If

        
        'CHECK TOPMARK WITH CUSLOTNO
        Select Case Trim(mark_spec_txt)
            Case "SG6841SZ":
                If Mid(Trim(Me.topx(2)), 2, 6) <> Left(Trim(txtCusLot), 6) Then
                    MsgBox "PARTICULAR TOPMARK3 NOT MATCH WITH 6 CHARACTER FROM CUSLOTNO!", vbCritical, "ERROR"
                    Me.topx(2).SetFocus
                    Exit Sub
                End If
            Case "SG6849":
                If Mid(Trim(Me.topx(2)), 2, 6) <> Left(Trim(txtCusLot), 6) Then
                    MsgBox "PARTICULAR TOPMARK3 NOT MATCH WITH 6 CHARACTER FROM CUSLOTNO!", vbCritical, "ERROR"
                    Me.topx(2).SetFocus
                    Exit Sub
                End If
            Case "SG6841DZ":
                If Mid(Trim(Me.topx(2)), 2, 6) <> Left(Trim(txtCusLot), 6) Then
                    MsgBox "PARTICULAR TOPMARK3 NOT MATCH WITH 6 CHARACTER FROM CUSLOTNO!", vbCritical, "ERROR"
                    Me.topx(2).SetFocus
                    Exit Sub
                End If
            Case "SG6848DZ1":
                If Mid(Trim(Me.topx(2)), 2, 6) <> Left(Trim(txtCusLot), 6) Then
                    MsgBox "PARTICULAR TOPMARK3 NOT MATCH WITH 6 CHARACTER FROM CUSLOTNO!", vbCritical, "ERROR"
                    Me.topx(2).SetFocus
                    Exit Sub
                End If
        End Select
        
        
'- Quah 20090402 Check if { matching with }, a few times planner incorrectly key in {T{
If InStr(topx(0), "{") > 0 Then
    If InStr(topx(0), "}") = 0 Then
            MsgBox "PLEASE CHECK MARKING. BRACKET NOT MATCH !", vbCritical, "ERROR"
            Exit Sub
    End If
End If
If InStr(topx(1), "{") > 0 Then
    If InStr(topx(1), "}") = 0 Then
            MsgBox "PLEASE CHECK MARKING. BRACKET NOT MATCH !", vbCritical, "ERROR"
            Exit Sub
    End If
End If
If InStr(topx(2), "{") > 0 Then
    If InStr(topx(2), "}") = 0 Then
            MsgBox "PLEASE CHECK MARKING ! BRACKET NOT MATCH !", vbCritical, "ERROR"
            Exit Sub
    End If
End If
If InStr(topx(3), "{") > 0 Then
    If InStr(topx(3), "}") = 0 Then
            MsgBox "PLEASE CHECK MARKING ! BRACKET NOT MATCH !", vbCritical, "ERROR"
            Exit Sub
    End If
End If
        
        
'- Quah 20081009 Check if MARKFILE registered by Engineer or not.
    Set rs = New ADODB.Recordset
    'If Trim(Me.package_lead_txt) Like "%FN%" Then
    
    '2010-05-14 Me.package_lead_txt change to Me.cbofullpackage
    If InStr(Me.cbofullpackage, "FN") > 0 Then
        SSQL1 = "SELECT * FROM AIC_MARKING_REF WHERE TARGETDEVICE='" & Trim(Me.TARGET_DEVICE_TXT) & "' AND PACKAGELEAD='" & Trim(Me.cbofullpackage) & "' AND CUSTNAME='" & Trim(custnameselect.TEXT) & "'"
    Else
        SSQL1 = "SELECT * FROM AIC_MARKING_REF WHERE DEVICENO='" & Trim(Me.internal_device_no_txt) & "' AND PACKAGELEAD='" & Trim(Me.cbofullpackage) & "' AND CUSTNAME='" & Trim(custnameselect.TEXT) & "'"
    End If
    Debug.Print SSQL1
    rs.Open SSQL1, wsDB
    markfileexist = "N"
    If rs.EOF = False Then
        If LenB(rs!MARKPLATE) <> 0 Then
            markfileexist = "Y"
        End If
    Else
        Set RS2 = New ADODB.Recordset
        SSQL1 = "SELECT * FROM AIC_MARKING_REF WHERE PACKAGELEAD='" & Trim(Me.cbofullpackage) & "' AND CUSTNAME='" & Trim(custnameselect.TEXT) & "' AND TARGETDEVICE='X'"
        Debug.Print SSQL1
        RS2.Open SSQL1, wsDB
            If RS2.EOF = False Then
                If LenB(RS2!MARKPLATE) <> 0 Then
                    markfileexist = "Y"
                End If
            End If
            RS2.Close
    End If
    rs.Close
    
    'Quah 20090916 compulsory markfile for Production lots. Agreed by Lingling
    'Quah 20090917 except for Smartcard (no marking)
    If SAVEREC.Caption <> "RELEASE LI" Then
        If Left(Trim(cboruntype), 11) = "Engineering" Or Left(Trim(Me.internal_device_no_txt), 4) = "SCRF" Or Left(Trim(Me.internal_device_no_txt), 3) = "SCM" Then
            If markfileexist = "N" Then
                s = MsgBox("MARKPLATE not registered for this Device. Proceed to save L.I.?", vbYesNo, "Message")
                If s <> 6 Then
                    Exit Sub
                End If
            End If
        Else    'Production Lots
            If markfileexist = "N" Then
    '             s = MsgBox("For Production lots, MARKPLATE must be registered for this Device.", vbCritical, "Message")
    '             Exit Sub
                s = MsgBox("MARKPLATE not registered for this Device. Proceed to save L.I.?", vbYesNo, "Message")
                If s <> 6 Then
                    Exit Sub
                End If
            End If
        End If
    
    End If
'- Quah 20081009
        
        
        s = MsgBox("CONFIRM UPDATE DATA?", vbYesNo, "Build Instruction")
        If s = 6 Then
            
            'Quah 20210705 insertLiBom at beginning of UPDATE.
            'Quah move INSERTBOM to MFGInventory Confirm Button
            'Call insertLIBOM(refno_TXT)
            
        '2011-12-29-1
        If Left(refno_TXT, 2) = "IJ" Then
            
            Dim ijrs As ADODB.Recordset
            Set ijrs = New ADODB.Recordset
            ssql = "select a.REFNO, WAFERNO, REMARKS1, REMARKS2, TEXT  from AIC_LI_LABELINFO a, AIC_LI_DUAL_DIE b, AIC_LABEL_REFERENCE where a.REFNO = '" & Trim(refno_TXT) & "' and ASSYLOTNO='IMPINJ IQA' and b.WAFERNO=WAFERLOTNO and a.REFNO=b.REFNO"
            Debug.Print ssql
            ijrs.Open ssql, wsDB
            If Not ijrs.EOF Then
                ijdef = Trim(ijrs!REMARKS2)
                ijwafer = Trim(ijrs!waferno)
            End If
            ijrs.Close
             If ijdef <> "" Then
                ijbal = InputBox("From this list, pls REMOVE wafers required for loading.", "IMPINJ BALANCE DIE", ijdef)
                If ijbal <> "" Then
                    If MsgBox("IMPINJ balance DIE-ID : " & ijbal & " ?", vbYesNo, "Message") = vbYes Then
                        Set ijrs = New ADODB.Recordset
                        ssql = "select * from AIC_LABEL_REFERENCE where ASSYLOTNO ='IMPINJ IQA' and WAFERLOTNO='" & ijwafer & "'"
                        Debug.Print ssql
                        ijrs.Open ssql, wsDB, adOpenDynamic, adLockOptimistic
                        If Not ijrs.EOF Then
                            ijrs!REMARKS2 = Trim(ijbal)
                            ijrs.Update
                        End If
                        ijrs.Close
                    End If
                End If
             End If
        End If
            
            
            
            'Check for Invoice UP
            TRAN_DATEx = Format(Now, "DD-MMM-YYYY")
            TRAN_TIMEx = Format(Time, "HH:MM:SS")
            Set rs = New ADODB.Recordset
            
'            SQL = "SELECT * FROM AIC_INVOICE_PRICE_DETAIL WHERE PRODUCT = '" & Trim(internal_device_no_txt) & "'AND TARGETDEVICE = '" & Trim(TARGET_DEVICE_TXT) & "'"
'2012-06-15 NEW INVOICE SYSTEM
'2012-06-20 skip below checking
            
'            SQL = "SELECT * FROM BAIC_INVOICE_PRICE_HEADER WHERE IPH_DEVICENO = '" & Trim(internal_device_no_txt) & "'"
'
'            rs.Open SQL, wsDB
'            If rs.EOF = True Then
''                If Weekday(Now) = "6" Or Weekday(Now) = "7" Or Weekday(Now) = "1" Then
''                    Set RS2 = New ADODB.Recordset
''                    SSQL2 = "SELECT * FROM AIC_INVOICE_UP_ALARM_TABLE WHERE PRODUCT = '" & Trim(internal_device_no_txt) & "'AND TARGETDEVICE = '" & Trim(TARGET_DEVICE_TXT) & "' "
''                    RS2.Open SSQL2, wsDB
''                    If RS2.EOF = True Then
''                        SSQL = "INSERT INTO AIC_INVOICE_UP_ALARM_TABLE (PRODUCT, TARGETDEVICE, REFNO, EMP_ID, TRAN_DATE, TRAN_TIME) " & _
''                               " VALUES('" & Trim(internal_device_no_txt) & "', '" & Trim(TARGET_DEVICE_TXT) & "', '" & Trim(refno_TXT) & "', '" & Trim(trim(login_id)) & "','" & TRAN_DATEX & "', '" & TRAN_TIMEX & "')"
''                        wsDB.Execute SSQL
''                    End If
''                    RS2.Close
''                Else
'                    MsgBox "ERROR: PRODUCT Unit Price Not Found, Please ask Finance to Setup First!!", vbCritical
'                    Exit Sub
''                End If
''            Else
''                Set RS1 = New ADODB.Recordset
''                SQL = "SELECT * FROM AIC_INVOICE_PRICE_HEADER WHERE DESC_ITEM5 = '" & Trim(internal_device_no_txt) & "'AND TARGETDEVICE = '" & Trim(TARGET_DEVICE_TXT) & "'"
''                RS1.Open SQL, wsDB
''                If RS1.EOF = True Then
'''                    If Weekday(Now) = "6" Or Weekday(Now) = "7" Or Weekday(Now) = "1" Then
'''                        Set RS2 = New ADODB.Recordset
'''                        SSQL2 = "SELECT * FROM AIC_INVOICE_UP_ALARM_TABLE WHERE PRODUCT = '" & Trim(internal_device_no_txt) & "'AND TARGETDEVICE = '" & Trim(TARGET_DEVICE_TXT) & "' "
'''                        RS2.Open SSQL2, wsDB
'''                        If RS2.EOF = True Then
'''                            SSQL = "INSERT INTO AIC_INVOICE_UP_ALARM_TABLE (PRODUCT, TARGETDEVICE, REFNO, EMP_ID, TRAN_DATE, TRAN_TIME) " & _
'''                                   " VALUES('" & Trim(internal_device_no_txt) & "', '" & Trim(TARGET_DEVICE_TXT) & "', '" & Trim(refno_TXT) & "', '" & Trim(trim(login_id)) & "','" & TRAN_DATEX & "', '" & TRAN_TIMEX & "')"
'''                            wsDB.Execute SSQL
'''                        End If
'''                        RS2.Close
'''                    Else
''                        MsgBox "ERROR: PRODUCT Unit Price Not Found, Please ask Finance to Setup First!!", vbCritical
''                        Exit Sub
'''                    End If
''                End If
''                RS1.Close
'            End If
'            rs.Close
            
        'CATHERINE 070514 - DELETE DIE INFO FROM AIC_DUAL_DIE
        wsSqlString = "DELETE AIC_LI_DUAL_DIE WHERE REFNO = '" & Trim(refno_TXT) & "'"
        wsDB.Execute wsSqlString
        
        Call SaveDie
            
            'Check location of DOT
            SpecChars = ""
            iCnt = 0
            Do While iCnt <= 4
                If chkdot(iCnt).Value = Checked Then
                    SpecChar(iCnt) = "."
                Else
                    SpecChar(iCnt) = " "
                End If
                SpecChars = SpecChars + SpecChar(iCnt)
                iCnt = iCnt + 1
            Loop
     
     
        'Quah 20090915 for SINOPOWER, need to add prefix to WIPPRD for linking to different test program in TEST ODBC.
'        If Left(Trim(refno_TXT), 2) = "SN" Then
'            Set rs = New ADODB.Recordset
'            SQL = "update wipprd set WPRD_USRDF_BGDAT_4='SN' WHERE WPRD_PROD ='" & Trim(internal_device_no_txt) & "'"
'            Debug.Print SQL
'            rs.Open SQL, wsDB
'        End If
     
        'Quah 2010-02-04 for Anpec,Sinopower update WIPPRD.WPRD_USRDF_BGDAT_3 = XAICBD_DEVICE (for Test-ODBC matching, requested by Ms Khoo)
        If Left(Trim(refno_TXT), 2) = "AN" Or Left(Trim(refno_TXT), 2) = "SN" Then
'            xbgdat3 = "NA"
'            If XAICBD_DEVICE <> "" Then
'                Set RS = New ADODB.Recordset
'                SQL = "select COUNT(distinct ZTRTE_MARKING) CNT from ztrte_tst_rte where ZTRTE_DEVICE_NO='" & TARGET_DEVICE_TXT & "'"
'                Debug.Print SQL
'                RS.Open SQL, wsDB
'                If RS!cnt > 1 Then
'                    xbgdat3 = XAICBD_DEVICE
'                End If
'                RS.Close
'
'                Set RS = New ADODB.Recordset
'                SQL = "update wipprd set WPRD_USRDF_BGDAT_3='" & xbgdat3 & "' where (WPRD_PRD_GRP_5='ANPEC' or  WPRD_PRD_GRP_5='SINOPOWER') AND WPRD_PROD ='" & Trim(internal_device_no_txt) & "'"
'                Debug.Print SQL
'                RS.Open SQL, wsDB
'            End If
        End If
     
        If Left(Trim(refno_TXT), 2) = "GM" Or Left(Trim(refno_TXT), 2) = "AN" Or Left(Trim(refno_TXT), 2) = "SN" Then
            '2010-11-03
            xbgdat3 = "NA"
            If XAICBD_DEVICE <> "" Then
                Set rs = New ADODB.Recordset
'                SQL = "select COUNT(distinct ZTRTE_MARKING) CNT from ztrte_tst_rte where ZTRTE_DEVICE_NO='" & TARGET_DEVICE_TXT & "'"
                SQL = "select COUNT(distinct TMH_MARKING) CNT from BAIC_TM_HEADER where TMH_TARGETDEVICE='" & TARGET_DEVICE_TXT & "'"
                Debug.Print SQL
                rs.Open SQL, wsDB
                If rs!cnt > 1 Then
                    xbgdat3 = XAICBD_DEVICE
                End If
                rs.Close

                Set rs = New ADODB.Recordset
'                SQL = "update wipprd set WPRD_USRDF_BGDAT_3='" & xbgdat3 & "' where (WPRD_PRD_GRP_5='ANPEC' or  WPRD_PRD_GRP_5='SINOPOWER') AND WPRD_PROD ='" & Trim(internal_device_no_txt) & "'"
                'aims
                SQL = " update baic_prodmast set pdm_data_1='" & xbgdat3 & "' where pdm_customer='GMT' and pdm_deviceno='" & Trim(internal_device_no_txt) & "'"
                Debug.Print SQL
                rs.Open SQL, wsDB
            End If
        End If
     
     
     
     
      'Quah.. 2010-10-15  update BAIC_CUSTOMER_PO
      '---------------------------------------------------------------
        If txt_pomode = "STANDARD" Then
            
            Call svrdatetime(xserverdate, xservertime, xshifttype, xproddate)
            
            If Trim(lblPONo) = "" Or Trim(lblPONo) = "NA" Or Trim(lblPONo) = "N/A" Then
                MsgBox "This is a STANDARD PO. Pls input the correct PO No.", vbCritical, "Message"
                Exit Sub
            End If
            If Trim(total_poqty) = "" Then
                MsgBox "This is a STANDARD PO. Pls input the TOTAL PO QTY.", vbCritical, "Message"
                Exit Sub
            End If
            Dim CpoRs As ADODB.Recordset
            Set CpoRs = New ADODB.Recordset
            sqltxt = "select * from baic_customer_po wheRE CPO_CUST_SHORTNAME='" & Trim(custnameselect.TEXT) & "' and CPO_PONO='" & lblPONo & "' and CPO_TARGETDEVICE='" & TARGET_DEVICE_TXT & "'"
            CpoRs.Open sqltxt, wsDB, adOpenDynamic, adLockOptimistic
            If CpoRs.EOF Then
                CpoRs.AddNew
                CpoRs!CPO_PO_MODE = "STANDARD"
                CpoRs!cpo_cust_shortname = Trim(custnameselect.TEXT)
                CpoRs!cpo_pono = Trim(lblPONo)
                CpoRs!cpo_targetdevice = Trim(TARGET_DEVICE_TXT)
                CpoRs!cpo_order_qty = total_poqty
                CpoRs!CPO_ORDER_YMD = xserverdate   'Now()
                CpoRs!cpo_prd_status = "OPN"
'                CpoRs!CPO_PRD_CLOSE_MODE = ""
'                CpoRs!CPO_PRD_CLOSE_YMD = ""
                CpoRs!CPO_FIN_STATUS = "OPN"
'                CpoRs!CPO_FIN_CLOSE_BY = ""
'                CpoRs!CPO_FIN_CLOSE_YMD = ""
'                CpoRs!CPO_REMARK = ""
                CpoRs!CPO_CREATED_BY = Trim(login_id)
                CpoRs!CPO_CREATION_YMD = Now()
                CpoRs!CPO_FIN_GOOD_QTY = 0
                CpoRs!CPO_FIN_REJ_QTY = 0
                CpoRs.Update
            Else
                CpoRs!cpo_order_qty = total_poqty
                CpoRs.Update
            End If
            CpoRs.Close
            Set CpoRs = Nothing
        End If
      '---------------------------------------------------------------
     
     
            Dim zadoLI As ADODB.Recordset
            Set zadoLI = New ADODB.Recordset
            Dim zadocounter As ADODB.Recordset
            Set zadocounter = New ADODB.Recordset
            wsSqlString = "select * from AIC_LOADING_INSTRUCTION" _
            & " where REFNO='" & Trim(refno_TXT) & "'"
            zadoLI.Open wsSqlString, wsDB, adOpenDynamic, adLockOptimistic
            If zadoLI.EOF = False Then
                zadoLI!DEVICE_NO = Trim(internal_device_no_txt)
                zadoLI!PACKAGE__LEAD = Trim(package_lead_txt)
                zadoLI!TARGET_DEVICE = Trim(TARGET_DEVICE_TXT)
                zadoLI!CUSLOTNO = Trim(txtCusLot)
                zadoLI!WAFER = Trim(WAFER)
                zadoLI!qty = Trim(qty)
                zadoLI!BD_NO = Trim(bonding_diagram_txt)
                zadoLI!CUSTOMER_NO = Left(Right(Trim(cboCust.TEXT), 4), 3)
                zadoLI!CUSTOMER_NAME = Trim(custnameselect.TEXT)
                zadoLI!MARKING_SPEC = Trim(mark_spec_txt)
                zadoLI!TOP1 = Trim((topx(0)))
                zadoLI!TOP2 = Trim((topx(1)))
                zadoLI!TOP3 = Trim((topx(2)))
                zadoLI!TOP4 = Trim((topx(3)))
                zadoLI!TOP5 = Trim((topx(4)))
                zadoLI!TOP6 = Trim((topx(5)))
                zadoLI!BOTTOM1 = Trim((bottom(0)))
                zadoLI!BOTTOM2 = Trim((bottom(1)))
                zadoLI!BOTTOM3 = Trim((bottom(2)))
                zadoLI!BOTTOM4 = Trim((bottom(3)))
                zadoLI!BOTTOM5 = Trim((bottom(4)))
                zadoLI!BOTTOM6 = Trim((bottom(5)))
                zadoLI!DATE_TRANX = DX
                zadoLI!TIME_TRANX = nowx
                zadoLI!EMP_ID = Trim(login_id)
                zadoLI!PO_NO = Trim(lblPONo)
                zadoLI!LOAD_TIME = Trim(cboruntype)     '2012-09-25
                zadoLI!Status = "N"
                zadoLI!Refno = Trim(refno_TXT)
                zadoLI!work_week = ww_txt
                zadoLI!PRINT_FLAG = "N"
                zadoLI!CATALOGNO = Trim(txtCatalogNo)
                zadoLI!SpecChar = SpecChars
                zadoLI!INTERNAL_BD = Trim(bd_no_txt.TEXT)      '2012-09-04
                zadoLI.Update
            End If
            zadoLI.Close
        
        
            'delete labelinfo, aic_loading_instruction_remark
            '-------------------------------------------------
            Dim donebefore
            Set UPDATEPERSON = New ADODB.Recordset
            wsSqlString = "SELECT * FROM AIC_LOADING_INSTRUCTION_REMARK WHERE REFNO='" & Trim(refno_TXT) & "'"
            UPDATEPERSON.Open wsSqlString, wsDB, adOpenDynamic, adLockOptimistic
            If Not UPDATEPERSON.EOF Then    'got record
                donebefore = "YES"
            Else
                donebefore = "NO"
            End If
            UPDATEPERSON.Close
            
            If donebefore = "YES" Then
                'Quah 20130509 exclude FAIRCHILD, req by Anita, due to PO# can be updated later. Dont want to effect the label-info data.
                'Quah 2021-03-04 exclude CHANGDIAN
                If Left(refno_TXT, 2) <> "CD" And Left(refno_TXT, 2) <> "FS" And Left(refno_TXT, 2) <> "FP" Then
                    ssql = "DELETE FROM AIC_LI_LABELINFO WHERE REFNO='" & Trim(refno_TXT) & "'"
                    wsDB.Execute ssql
                    ssql = "DELETE FROM AIC_LOADING_INSTRUCTION_REMARK WHERE REFNO='" & Trim(refno_TXT) & "'"
                    wsDB.Execute ssql
                End If
            
                'RE-SET SECOND PAGE, PLANNER NEED TO RE-KEY IN REVELANT INFO.
                Set UPDATEPERSON = New ADODB.Recordset
                wsSqlString = "SELECT top 1 * FROM BAIC_ALARM_TRANX"
                UPDATEPERSON.Open wsSqlString, wsDB, adOpenDynamic, adLockOptimistic
                If UPDATEPERSON.EOF = False Then
                    UPDATEPERSON!ALM_LOTNO = Trim(refno_TXT)
                    UPDATEPERSON!ALM_OPER = ""
                    UPDATEPERSON!ALM_REC_TYPE = "LI_U"
                    UPDATEPERSON!ALM_DATA_A1 = ""
                    UPDATEPERSON!ALM_DATA_A2 = ""
                    UPDATEPERSON!ALM_DATA_A3 = Trim(login_id)
                    UPDATEPERSON!ALM_DATA_A4 = ""
                    UPDATEPERSON!ALM_DATA_A5 = ""
                    UPDATEPERSON!ALM_DATA_N1 = 0
                    UPDATEPERSON!ALM_DATA_N2 = 0
                    UPDATEPERSON!TRANSACTION_YMD = Trim(xserverdate)
                    UPDATEPERSON.Update
                End If
                UPDATEPERSON.Close
            End If
            
            
            'Quah add 20210628
            'Quah move to beginning of Update 20210705
            'Call insertLIBOM(refno_TXT)
            
            MsgBox "LI Updated. Pls complete 2nd page.!", vbInformation, "Update"
            PULLDATA
            
            Call Addinfo_Click
        
        End If
    End If
End Sub


Private Sub cmdView_Click()
chkUpdate
'20170412 popup msg for RWK lots, req by OO
If InStr(internal_device_no_txt, "RWK") > 0 Then
    asn = MsgBox("Do you have 'Deviation Authorisation Form' " & Chr(13) & " for this rework lot?", vbYesNo, "Message")
    If asn = 7 Then
        Exit Sub
    End If
End If

'-------------------------CHONG ADD CHECKING ON UPDATED DEVICE WITHOUT VERIFY BOM 20230517----------------------------------------
'20211012
libom_ok = ""
Dim chkbom As ADODB.Recordset
Set chkbom = New ADODB.Recordset
'ssql = "select * from BAIC_LI_BOM where BOM_LIREF='" & Trim(refno_TXT) & "' and BOM_LI_RELEASE_YMD >= '2021-10-10'"
ssql = " select REFNO, DATE_TRANX, BOM_LIREF, BOM_DEVICENO from AIC_LOADING_INSTRUCTION " & _
       "    left outer join BAIC_LI_BOM on BOM_LIREF=refno where REFNO='" & Trim(refno_TXT) & "'"
Debug.Print ssql
chkbom.Open ssql, wsDB
If Not chkbom.EOF Then
   If InStr(internal_device_no_txt, "TST") = 0 And cboCust <> "STMICRO" Then 'KOKO ADD FOR PASSBY TST ST DEVICE BOM Temporary20220811
        If Format(chkbom!DATE_TRANX, "YYYYMMDD") >= "20211010" Then
            If IsNull(chkbom!BOM_LIREF) = True Or internal_device_no_txt <> chkbom!BOM_DEVICENO Then
                MsgBox "Please verify BOM first.", vbCritical, "Message"
                Exit Sub
            Else
                'ok'
                libom_ok = "YES"
            End If
        Else
            If IsNull(chkbom!BOM_LIREF) = False Then
                'ok'
                libom_ok = "YES"
            End If
        End If
   End If 'ko end 20220811
Else
    MsgBox "Error with LI", vbCritical, "Message"
    Exit Sub
End If
chkbom.Close
Set chkbom = Nothing
'==========================old code==============================
''20211012
'libom_ok = ""
'Dim chkbom As ADODB.Recordset
'Set chkbom = New ADODB.Recordset
''ssql = "select * from BAIC_LI_BOM where BOM_LIREF='" & Trim(refno_TXT) & "' and BOM_LI_RELEASE_YMD >= '2021-10-10'"
'ssql = " select REFNO, DATE_TRANX, BOM_LIREF from AIC_LOADING_INSTRUCTION " & _
'       "    left outer join BAIC_LI_BOM on BOM_LIREF=refno where REFNO='" & Trim(refno_TXT) & "'"
'Debug.Print ssql
'chkbom.Open ssql, wsDB
'If Not chkbom.EOF Then
'   If InStr(internal_device_no_txt, "TST") = 0 And cboCust <> "STMICRO" Then 'KOKO ADD FOR PASSBY TST ST DEVICE BOM Temporary20220811
'        If Format(chkbom!DATE_TRANX, "YYYYMMDD") >= "20211010" Then
'            If IsNull(chkbom!BOM_LIREF) = True Then
'                MsgBox "Please verify BOM first.", vbCritical, "Message"
'                Exit Sub
'            Else
'                'ok'
'                libom_ok = "YES"
'            End If
'        Else
'            If IsNull(chkbom!BOM_LIREF) = False Then
'                'ok'
'                libom_ok = "YES"
'            End If
'        End If
'   End If 'ko end 20220811
'Else
'    MsgBox "Error with LI", vbCritical, "Message"
'    Exit Sub
'End If
'chkbom.Close
'Set chkbom = Nothing


'==========================old code==============================
'-------------------------CHONG ADD CHECKING DEVICE WITHOUT VERIFY BOM 20230517----------------------------------------


'2012-02-13 ::::::::::: DIEBANK WAFER VALIDATION ::::::::::::::::::
'---------------------------------------------------------------------
'If NEW_OPT = True Then
'    Dim wfrchk As ADODB.Recordset
'    Dim wfrchk2 As ADODB.Recordset
'    Set wfrchk = New ADODB.Recordset
'    SSQL = "select WAFERNO, DIE_QTY from AIC_LI_DUAL_DIE where REFNO='" & Trim(refno_TXT) & "' and WAFER_QTY > 0"
'    Debug.Print SSQL
'    wfrchk.Open SSQL, wsDB
'    Do While Not wfrchk.EOF
'        dbwfr = Trim(wfrchk!waferno)
'        dbqty = wfrchk!DIE_QTY
'        Set wfrchk2 = New ADODB.Recordset
'        SSQL = "SELECT * FROM AIC_INVENTORY_MASTER WHERE WAFERLOTNO='" & dbwfr & "' AND STATUS = 'D-RECEIVED' and DIEBANKQTY >= " & dbqty
'        wfrchk2.Open SSQL, wsDB
'        If wfrchk2.EOF Then
'            MsgBox "Please check. Diebank qty not enough for : " & dbwfr & " ", vbCritical, "Message"
'        End If
'        wfrchk2.Close
'        wfrchk.MoveNext
'    Loop
'
'End If
'---------------------------------------------------------------------



Dim sql2 As String
Dim rsM As ADODB.Recordset
Dim jobno As String
Dim ORACODE As String

    If Trim(cboCust.TEXT) = "" Then
        MsgBox " PLEASE INSERT customer code!!!", vbCritical
        cboCust.SetFocus
        Exit Sub
    End If
    If Trim(refno_TXT) = "" Then
        MsgBox " PLEASE INSERT REFERENCE NO!!!", vbCritical
        refno_TXT.SetFocus
    
    Else


        'check li
        SQL$ = "select count(*) as ab from AIC_LOADING_INSTRUCTION where REFNO  = '" & Trim(refno_TXT) & "'"
        Set RSQ = New ADODB.Recordset
        RSQ.Open SQL$, wsDB
        If Val(RSQ!ab) < 1 Then
            MsgBox "LI Not Generated Yet!!", vbInformation, "Error"
            Exit Sub
        End If
        SQL$ = "select count(*) as ac from AIC_LOADING_INSTRUCTION_REMARK where REFNO  = '" & Trim(refno_TXT) & "'"
        Set RSQ = New ADODB.Recordset
        RSQ.Open SQL$, wsDB
        If Val(RSQ!ac) < 1 Then
            MsgBox "Please Key-In Additional Information Before Print!!", vbInformation, "Error"
            Exit Sub
        End If
        RSQ.Close
        
        SQL$ = " select count(*) as ad from aic_li_dual_die where refno = '" & Trim(refno_TXT) & "'"
        RSQ.Open SQL$, wsDB
        If Val(RSQ!ad) < 1 Then
          MsgBox " Please key in die information before print!!", vbInformation, "ERROR"
          Exit Sub
        End If
        RSQ.Close

        Dim liFile As ADODB.Recordset
        Dim AICX As ADODB.Recordset
        Set AICX = New ADODB.Recordset
        Set liFile = New ADODB.Recordset
        Dim aic As TableDef
        
        Set rsM = New ADODB.Recordset
     
        dtlx = App.Path & "\Report\liDB.MDB"
       '  dtlx = "c:\liDB.MDB"
        Set DTL = OpenDatabase(dtlx, False, False)
       
        DTL.Execute "DELETE FROM liTEMPX"
        DTL.Execute "DELETE FROM liADD"
        DTL.Execute "DELETE FROM MARK_SPEC"
        DTL.Execute "DELETE FROM TB_LABEL"
        DTL.Execute "DELETE FROM DUAL_DIE"
        DTL.Execute "DELETE FROM LABEL_INFO"
        'CATHERINE
        DTL.Execute "DELETE FROM SPEC_CHAR"
        DTL.Execute "DELETE FROM CUST_PO"
        
        
        sql2 = " SELECT ORA_CUSTCODE FROM AIC_LI_CUST_MASTER " & _
               " WHERE CUSTNO = '" & Left(Right(Trim(cboCust.TEXT), 4), 3) & "'"
        rsM.Open sql2, wsDB
        If Not rsM.EOF Then
          ORACODE = Trim(rsM("ora_custcode"))
        Else
          MsgBox " No oracle code found!"
        End If
        
        wsSqlString = "select * from AIC_LOADING_INSTRUCTION" _
          & " where REFNO  = '" & refno_TXT & "'"
        Debug.Print wsSqlString
        liFile.Open wsSqlString, wsDB
        If liFile.EOF = False Then
            'If IsNull(Trim(liFile!WAFER)) = True Then
            If LenB(Trim(liFile!WAFER)) = 0 Then
                MsgBox "DIE INFO IS NOT COMPLETE.", vbInformation
                liFile.Close
                Exit Sub
            End If
            
            'SPECIAL CHACRACTER CONTROL
            If LenB(Trim(liFile!SpecChar)) <> 0 Then
                STRSTRING = "INSERT INTO SPEC_CHAR(FIELD1, FIELD2, FIELD3, FIELD4, FIELD5) " & _
                            "values('" & Left(liFile!SpecChar, 1) & "', '" & Mid(liFile!SpecChar, 2, 1) & "', " & _
                            " '" & Mid(liFile!SpecChar, 3, 1) & "', '" & Mid(liFile!SpecChar, 4, 1) & "', " & _
                            "'" & Mid(liFile!SpecChar, 5, 1) & "')"
            DTL.Execute STRSTRING
            End If
        
            'green dot indicator for GM-WM
            GRDOT = ""
            If Left(refno_TXT, 2) = "GM" Then
                If Trim(liFile("SLNO")) = "GREENDOT" Then
                    GRDOT = "GREENDOT"
                End If
            End If
            STRSTRING = "INSERT INTO TB_LABEL(REFNO, REM4) " & _
                        "values('" & Trim(refno_TXT) & "','" & Trim(GRDOT) & "')"
            DTL.Execute STRSTRING
        
        
        
dbpackage = liFile("PACKAGE__LEAD")
dbfullpackage = dbpackage

'Quah 20090223 auto initialze long package name
'disable 20100923

'    Set RS = New ADODB.Recordset
'    sqltext = " select packagelead_full from aic_packagelead_master where PACKAGLEAD_SHORT='" & Trim(dbpackage) & "'"
'    Debug.Print sqltext
'    RS.Open sqltext, wsDB
'    If Not RS.EOF Then
'        dbfullpackage = Trim(RS!PACKAGELEAD_FULL)
'    End If
'    RS.Close
'    Set RS = Nothing
'end. Quah 20090223 auto initialize long package name
        
        
            With aic
                STRSTRING = "INSERT INTO liTEMPX VALUES ('" & liFile("DEVICE_NO") & "','" & _
                                          dbfullpackage & "','" & _
                                          (liFile("TARGET_DEVICE")) & "','" & _
                                          (liFile("WAFER")) & "'," & _
                                          (liFile("QTY")) & ",'" & _
                                          (liFile("BD_NO")) & "','" & _
                                          (liFile("INTERNAL_BD")) & "','" & _
                                          (liFile("CUSTOMER_NO")) & "','" & _
                                          (liFile("CUSTOMER_NAME")) & "','" & _
                                          (liFile("MARKING_SPEC")) & "','" & _
                                          (liFile("TOP1")) & "','" & _
                                          (liFile("TOP2")) & "','" & _
                                          (liFile("TOP3")) & "','" & _
                                          (liFile("TOP4")) & "','" & _
                                          (liFile("TOP5")) & "','" & _
                                          (liFile("TOP6")) & "','" & _
                                          (liFile("BOTTOM1")) & "','" & _
                                          (liFile("BOTTOM2")) & "','" & _
                                          (liFile("BOTTOM3")) & "','" & _
                                          (liFile("BOTTOM4")) & "','" & _
                                          (liFile("BOTTOM5")) & "','" & liFile("BOTTOM6") & "','" & _
                                          (liFile("DATE_TRANX")) & "'," & liFile("TIME_TRANX") & ",'" & _
                                          (liFile("EMP_ID")) & "','" & liFile("STATUS") & "','" & liFile("REFNO") & "','" & _
                                          (liFile("WORK_WEEK")) & "','" & liFile("CUSLOTNO") & "','" & liFile("PO_NO") & "','" & _
                                           ORACODE & "','')"
                DTL.Execute STRSTRING
            End With
            
            jobno = liFile("cuslotno")
            'Printing Daily Transfer Logsheet
        End If
        liFile.Close
        
        wsSqlString = "select * from AIC_LI_DUAL_DIE " _
                    & " where REFNO  = '" & refno_TXT & "'" _
                    & " order by seqno "
        liFile.Open wsSqlString, wsDB
        Dim LIDEX As String
        While Not liFile.EOF
          If IsNull(liFile!LIDIE_QTY) = True Then
            LIDIEX = "-"
          Else
            LIDIEX = Trim(liFile!LIDIE_QTY)
          End If
                                                            
        STRSTRING = " INSERT INTO DUAL_DIE " & _
                    " VALUES('" & liFile("refno") & "', " & _
                    " '" & liFile("partno") & "','" & liFile("waferno") & "', " & _
                    " " & liFile("wafer_qty") & "," & liFile("die_qty") & "," & liFile("seqno") & ", " & _
                    " '" & Trim(LIDIEX) & "')"
                                                            
          DTL.Execute STRSTRING
          
            If Left(Trim(refno_TXT), 2) = "BE" Then
                'DIANA 20150507 save die part for wafer scribe, bourns LI second page
                Set pg2 = New ADODB.Recordset
                SQL = "update AIC_LABEL_REFERENCE set remarks2='" & liFile("partno") & "' WHERE ASSYLOTNO = 'BOURNS IQA' AND remarks1 ='" & Trim(refno_TXT) & "' " & _
                        " and cuslotno='" & liFile("waferno") & "'"
                        Debug.Print SQL
                pg2.Open SQL, wsDB
            End If
                    
          liFile.MoveNext
        Wend
        liFile.Close
       
       'add additional instruction
        wsSqlString = "select * from AIC_LOADING_INSTRUCTION_REMARK" _
          & " where REFNO  = '" & refno_TXT & "'"
        liFile.Open wsSqlString, wsDB
        If liFile.EOF = False Then
            If Trim(liFile!URGENTLOT_FLAG) = "Y" Then
                urgflag = "URGENT LOT"
            Else
                urgflag = " "
            End If
            If Trim(liFile!DS_FLAG) = "Y" Then
                dsxflag = "DS"
            Else
                dsxflag = " "
            End If
            If Trim(liFile!ENGLOT_FLAG) = "E" Then
                englotflag = "E"
            Else
                englotflag = "N/A"
            End If
            If Trim(liFile!QTALOT_FLAG) = "Q" Then
                qtalotflag = "Q"
            Else
                qtalotflag = "N/A"
            End If
            If Trim(liFile!TEST_FLAG) = "Y" Then
                testxflag = "TESTED"
            ElseIf Trim(liFile!TEST_FLAG) = "N" Then
                testxflag = "UNTESTED"
            End If
            
            With aic
                'INSERT INTO ACCESS TABLE
                STRSTRING = "INSERT INTO liADD VALUES ('" & liFile("refno") & "','" & _
                (liFile("REM1")) & "','" & (liFile("REM2")) & "','" & (liFile("REM3")) & "','" & (liFile("REM4")) & "','" & (liFile("REM5")) & "','" & _
                (liFile("REM6")) & "','" & (liFile("REM7")) & "','" & (liFile("REM8")) & "','" & (liFile("REM9")) & "','" & (liFile("REM10")) & "','" & _
                (urgflag) & "','" & (englotflag) & "','" & (qtalotflag) & "','" & (testxflag) & "','" & (dsxflag) & "','" & _
                (liFile("SINGMARK_FLAG")) & "','" & (liFile("Form")) & "','" & (liFile("SL_NO")) & "','" & (liFile("TEST_PROCESS")) & "','" & (liFile("BIN1_TOP1")) & "','" & _
                (liFile("BIN1_TOP2")) & "','" & (liFile("BIN1_TOP3")) & "','" & (liFile("BIN1_TOP4")) & "','" & (liFile("BIN1_TOP5")) & "','" & (liFile("BIN1_TOP6")) & "','" & _
                (liFile("BIN3_TOP1")) & "','" & (liFile("BIN3_TOP2")) & "','" & (liFile("BIN3_TOP3")) & "','" & (liFile("BIN3_TOP4")) & "','" & _
                (liFile("BIN3_TOP5")) & "','" & (liFile("BIN3_TOP6")) & "','" & (liFile("BIN5_TOP1")) & "','" & (liFile("BIN5_TOP2")) & "','" & (liFile("BIN5_TOP3")) & "','" & _
                (liFile("BIN5_TOP4")) & "','" & (liFile("BIN5_TOP5")) & "','" & (liFile("BIN5_TOP6")) & "','" & _
                (liFile("loadingdate")) & "','" & (liFile("wsstartdate")) & "', '" & Trim(ROUTEx) & "')"
                DTL.Execute STRSTRING
            End With

            'Printing Daily Transfer Logsheet
        End If
        liFile.Close
        'Chk Mark Spec
        Dim MSpcRS As ADODB.Recordset
        Set MSpcRS = New ADODB.Recordset
        sqltext = "SELECT * FROM AIC_MARKING_SPEC_X WHERE SPECNO='" & Trim(mark_spec_txt) & "' AND APPROVAL = 'APPROVED'"
        MSpcRS.Open sqltext, wsDB
        If MSpcRS.EOF = False Then
            SPECNO = Trim(MSpcRS!SPECNO)
            TOP1 = Trim(MSpcRS!TOP1)
            TOP2 = Trim(MSpcRS!TOP2)
            TOP3 = Trim(MSpcRS!TOP3)
            TOP4 = Trim(MSpcRS!TOP4)
            TOP5 = Trim(MSpcRS!TOP5)
            TOP6 = Trim(MSpcRS!TOP6)
            BOTTOM1 = Trim(MSpcRS!BOTTOM1)
            BOTTOM2 = Trim(MSpcRS!BOTTOM2)
            BOTTOM3 = Trim(MSpcRS!BOTTOM3)
            BOTTOM4 = Trim(MSpcRS!BOTTOM4)
            BOTTOM5 = Trim(MSpcRS!BOTTOM5)
            BOTTOM6 = Trim(MSpcRS!BOTTOM6)
            STRSTRING = "INSERT INTO MARK_SPEC VALUES('" & Trim(refno_TXT) & "','" & SPECNO & "','" & TOP1 & "','" & TOP2 & "','" & TOP3 & "','" & TOP4 & "','" & TOP5 & "','" & TOP6 & "','" & BOTTOM1 & "','" & BOTTOM2 & "','" & BOTTOM3 & "','" & BOTTOM4 & "','" & BOTTOM5 & "','" & BOTTOM6 & "')"
            DTL.Execute STRSTRING
        End If
        MSpcRS.Close
        Set MSpcRS = Nothing
        'End Check Mark Spec
         
         
'insert customer po info
podatestr = Format(Me.txt_podate, "DD/MM/YYYY")
If total_poqty = "" Or total_poqty = "N/A" Then total_poqty = 0
ssql = "INSERT INTO CUST_PO (pomode,PONO,potgtdev,poqty,podate) " & _
        " VALUES('" & Trim(Me.txt_pomode) & "','" & Trim(Me.lblPONo) & "','" & Trim(Me.TARGET_DEVICE_TXT) & "'," & total_poqty & ", '" & podatestr & "')"
        Debug.Print ssql
DTL.Execute ssql
     
'CATHERINE 070606 - LABEL INFO CHANGE TABLE
'----------NEW INSERT LABEL INFO UPDATE 20070606-----------------------------------------------------------------
        SQL$ = "select * from AIC_LI_LABELINFO WHERE REFNO='" & Trim$(refno_TXT) & "' ORDER BY SEQNO"
        Set rsLABEL = New ADODB.Recordset
        rsLABEL.Open SQL$, wsDB, adOpenDynamic
        Do While rsLABEL.EOF = False
             ssql = "INSERT INTO LABEL_INFO(FIELD1, FIELD2, FIELD3) " & _
                    "VALUES( '" & Trim(rsLABEL!Label) & "', '" & Trim(rsLABEL!TEXT) & "',  " & _
                    " " & Trim(rsLABEL!seqno) & ")"
            DTL.Execute ssql
            rsLABEL.MoveNext
        Loop
        rsLABEL.Close

            
            MsgBox "PARTICULAR LI HEADER INFO ADD SUCCESSFULLY!", vbInformation, "ADD"
            SAVEREC.Enabled = False
            BOMLI_VIEW.Enabled = False 'temp 20221025
            Addinfo.Enabled = True


'Catherine 070607 - Label Info Format Change to General
'fetek use own format, label info too long kacau micrel 20150617 diana
        Me.CrystalReport2.WindowTitle = "LOADING INSTRUCTION"
        If Left(Trim(refno_TXT), 2) = "FK" Then
            
            ' add check box to select report with or without marking picture
            If Check2.Value = 1 Then
                '-------------------------------------------------------------------------------
                '20210526 - Update Marking_Patter.bmp with the actual image created by Planner (link by filename = LI refno).
                sourcefile = "\\aicwksvr2016\markingpattern\" & Trim(refno_TXT) & ".bmp"
                Debug.Print sourcefile
                Destinationfile = "C:\AICS SYSTEM\apos\Report\Marking_Pattern.bmp"
                FileCopy sourcefile, Destinationfile
                '-------------------------------------------------------------------------------
              
                CrystalReport2.ReportFileName = App.Path & "\Report\LoadingInstructionRpt_GenFK_M.RPT"
            
            Else
                CrystalReport2.ReportFileName = App.Path & "\Report\LoadingInstructionRpt_GenFK.RPT"
            End If
            
            
            
        '20171027 Quah add separate format for MICROCHIP.
        ElseIf Left(Trim(refno_TXT), 2) = "MC" Then
        If Check2.Value = 1 Then
                '-------------------------------------------------------------------------------
                '20210526 - Update Marking_Patter.bmp with the actual image created by Planner (link by filename = LI refno).
                sourcefile = "\\aicwksvr2016\markingpattern\" & Trim(refno_TXT) & ".bmp"
                Debug.Print sourcefile
                Destinationfile = "C:\AICS SYSTEM\apos\Report\Marking_Pattern.bmp"
                FileCopy sourcefile, Destinationfile
                '-------------------------------------------------------------------------------
              
                CrystalReport2.ReportFileName = App.Path & "\Report\loadinginstructionnewVF_MC_pic.RPT"
            
            Else
                CrystalReport2.ReportFileName = App.Path & "\Report\loadinginstructionnewVF_MC.RPT"
            End If
            ''CrystalReport2.ReportFileName = App.Path & "\Report\loadinginstructionnewVF_MC.RPT"
        '20200810 KO/Quah add separate format for M-CIRCUITS FOR TARGETDEVICE RIGHT 2 DIGIT A+ ,Requestor Choong WT
        
        ElseIf Left(Trim(refno_TXT), 2) = "MP" And Right(Trim(Me.TARGET_DEVICE_TXT), 2) = "A+" Then
           CrystalReport2.ReportFileName = App.Path & "\Report\loadinginstructionRpt_MP.RPT"
           
        ElseIf Left(Trim(refno_TXT), 2) = "FV" Then 'Ain add Map/Ink FORTUNE picture 20221028
            
            ' add check box to select report with or without marking picture
            If Check2.Value = 1 Then
                '-------------------------------------------------------------------------------
                '20210526 - Update Marking_Patter.bmp with the actual image created by Planner (link by filename = LI refno).
                sourcefile = "\\aicwksvr2016\markingpattern\" & Trim(refno_TXT) & ".bmp"
                Debug.Print sourcefile
                Destinationfile = "C:\AICS SYSTEM\apos\Report\Marking_Pattern.bmp"
                FileCopy sourcefile, Destinationfile
                '-------------------------------------------------------------------------------
              
                CrystalReport2.ReportFileName = App.Path & "\Report\LoadingInstructionRpt_GenFV_M.RPT"
            
            Else
                CrystalReport2.ReportFileName = App.Path & "\Report\LoadingInstructionRpt_FV.RPT"
            End If
            
        ElseIf Left(Trim(refno_TXT), 2) = "AP" Then 'Ain add Map/Ink FORTUNE picture 20221028
            
            ' add check box to select report with or without marking picture
            If Check2.Value = 1 Then
                '-------------------------------------------------------------------------------
                '20210526 - Update Marking_Patter.bmp with the actual image created by Planner (link by filename = LI refno).
                sourcefile = "\\aicwksvr2016\markingpattern\" & Trim(refno_TXT) & ".bmp"
                Debug.Print sourcefile
                Destinationfile = "C:\AICS SYSTEM\apos\Report\Marking_Pattern.bmp"
                FileCopy sourcefile, Destinationfile
                '-------------------------------------------------------------------------------
              
                CrystalReport2.ReportFileName = App.Path & "\Report\LoadingInstructionRpt_GenAP_M.RPT"
            
            Else
                CrystalReport2.ReportFileName = App.Path & "\Report\LoadingInstructionRpt_AP1.RPT"
            End If
            
        
        Else
            
            ' add check box to select report with or without marking picture
            If Check2.Value = 1 Then
                '-------------------------------------------------------------------------------
                '20210526 - Update Marking_Patter.bmp with the actual image created by Planner (link by filename = LI refno).
                sourcefile = "\\aicwksvr2016\markingpattern\" & Trim(refno_TXT) & ".bmp"
                Debug.Print sourcefile
                Destinationfile = "C:\AICS SYSTEM\apos\Report\Marking_Pattern.bmp"
                FileCopy sourcefile, Destinationfile
                '-------------------------------------------------------------------------------
                 
                'If Left(Trim(refno_TXT), 2) = "XXIJ" Then
                '    CrystalReport2.ReportFileName = App.Path & "\Report\LoadingInstructionRpt_General_M2.RPT"
                'Else
                    CrystalReport2.ReportFileName = App.Path & "\Report\LoadingInstructionRpt_General_M.RPT"
                'End If
            
'            ElseIf Left(Trim(refno_TXT), 2) = "AP" Then
'                CrystalReport2.ReportFileName = App.Path & "\Report\LoadingInstructionRpt_AP1.RPT"
'            ElseIf Left(Trim(refno_TXT), 2) = "FV" Then 'Ain add Map/Ink FORTUNE 20221027
'                CrystalReport2.ReportFileName = App.Path & "\Report\LoadingInstructionRpt_FV.RPT"
            Else
                CrystalReport2.ReportFileName = App.Path & "\Report\LoadingInstructionRpt_General.RPT"
            End If
        
        End If
        CrystalReport2.Destination = 0
        CrystalReport2.Action = 1
        Me.CrystalReport2.WindowState = crptMaximized
        
        '20210614 update specchar with M
        If Check2.Value = 1 Then
                STRSTRING = "update AIC_LOADING_INSTRUCTION set specchar='M' where REFNO='" & Trim(refno_TXT) & "'"
                wsDB.Execute STRSTRING
        End If
        
        'DIANA 20150507 BOURNS SECOND PAGE FOR WAFER SCRIBE
        'Quah 20181204 deactivate.
''''         If Left(Trim(refno_TXT), 2) = "BE" Then
''''            CrystalReportB.WindowTitle = "BOURNS WAFER SCRIBE"
''''            CrystalReportB.Connect = "PROVIDER=MSDASQL;dsn=AICSSQLDB;uid=ITAPP;pwd=APP;"
''''            CrystalReportB.ReportFileName = App.Path & "\Report\BOURNS SCRIBE.RPT"
''''            CrystalReportB.SelectionFormula = "{AIC_LABEL_REFERENCE.REMARKS1}='" & Trim(refno_TXT) & "' AND {AIC_LABEL_REFERENCE.ASSYLOTNO}='BOURNS IQA'"
''''            CrystalReportB.Destination = 0
''''            CrystalReportB.WindowState = crptMaximized
''''            CrystalReportB.Action = 1
''''         End If
        
        'CODES FOR PRINT MICROCHIP LOT ATTRIBUTES
        If Trim(Left(refno_TXT, 2)) = "MC" Then
            CrystalReportB.Connect = "PROVIDER=MSDASQL;dsn=AICSSQLDB;uid=ITAPP;pwd=APP;"
            CrystalReportB.ReportFileName = App.Path & "\report\AD_ATTRIBUTES_LI.RPT"
            CrystalReportB.SelectionFormula = "{CLS_LOT_INFO.CLS_LOTNO}='" & Trim(txtCusLot) & "' AND {AIC_LOADING_INSTRUCTION.STATUS}<>'C'"
            CrystalReportB.Destination = crptToWindow
            CrystalReportB.WindowState = crptMaximized
            CrystalReportB.Action = 1
        End If
    
    
    End If

If libom_ok = "YES" Then
  CrystalReportB.Connect = "PROVIDER=MSDASQL;dsn=AICSSQLDB;uid=ITAPP;pwd=APP;"
  CrystalReportB.ReportFileName = App.Path & "\report\BOMLI-A.RPT"
  CrystalReportB.SelectionFormula = "{BAIC_LI_BOM.BOM_LIREF}='" & Trim(refno_TXT) & "'"
  CrystalReportB.Destination = crptToWindow
  CrystalReportB.WindowState = crptMaximized
  CrystalReportB.Action = 1
End If

End Sub

Private Sub be_exit_Click()
    be_waferlotno.Clear
    be_listscribe.ListItems.Clear
    be_total = ""
    scribe = ""
    Check1.Value = Unchecked
    be_waferlotno.Locked = False
    Frame_BE.Visible = False
End Sub







Private Sub dd1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        ww_txt = "" '2014-11-18
        XX01 = dd1
        mm1.SetFocus
    End If
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
Me.top = 0
Me.Left = 0
 
 Dim WSFACILITY
 Dim sqltext As String
 Dim sqltext2 As String
 Dim chkAllpkg As ADODB.Recordset
 Dim rs As ADODB.Recordset
 Dim rsCnt As ADODB.Recordset
 Dim cust As String
 
    cDateTime = Date
    TIMEX = Format(Time, "HHMMSS")
    DX = cDateTime
    nowx = TIMEX
    ld_txt.Clear
    package_txt.Clear
    sqltext = ""
    sqltext2 = ""
    cust = ""
    
    'FOR ANPEC ONLY CONTROL LABEL INFO FIELD
    TARGET_DEVICEX = ""
       
'Quah 20090223 add long-package-name
    Set rs = New ADODB.Recordset
'    sqltext = " select * from aic_packagelead_master order by PACKAGELEAD_FULL "
sqltext = "select distinct pdm_packagelead PACKAGELEAD_FULL from baic_prodmast ORDER BY pdm_packagelead"
    rs.Open sqltext, wsDB
    If Not rs.EOF Then
      Do While Not rs.EOF
        cbofullpackage.AddItem Trim(rs!PACKAGELEAD_FULL)
        rs.MoveNext
      Loop
    End If
    rs.Close
    Set rs = Nothing
'end. Quah 20090223 add long-package-name
     
    'Get the record from customer master setting'
    Set rs = New ADODB.Recordset
    'sqltext = " SELECT DISTINCT CUSTNO, CUSTNAME FROM AIC_LI_CUST_MASTER WHERE PROG_MODULE = 'TNRBI_TWN' "
'Quah 20080617    sqltext = " SELECT DISTINCT GROUPING, CUSTNO, CUSTNAME FROM AIC_LI_CUST_MASTER"

'Quah 20080617 filter exclude Anpec, Apower, Atmel
'    SQLTEXT = " SELECT DISTINCT GROUPING, CUSTNO, CUSTNAME FROM AIC_LI_CUST_MASTER where (custno <> '016' and custno <> '018') and (custno <> '001' and custno <>'007' and custno <> '012' and custno <> '013' and custno <> '021') order by custname"

'REAL B2B include Atmel in GENERAL 007, 013, 001, 012
'    SQLTEXT = " SELECT DISTINCT GROUPING, CUSTNO, CUSTNAME FROM AIC_LI_CUST_MASTER where (custno <> '016' and custno <> '018') and  custno <> '021' order by custname"
sqltext = " SELECT DISTINCT GROUPING, CUSTNO, CUSTNAME FROM AIC_LI_CUST_MASTER where ( custno <>'001' or custno <>'007' or custno <>'012' or custno <>'013' ) order by custname"


'2012-11-12 move NIKO (021) to General screen
'  --  SQLTEXT = " SELECT DISTINCT GROUPING, CUSTNO, CUSTNAME FROM AIC_LI_CUST_MASTER where (custno <> '018') and (custno <> '001' and custno <>'007' and custno <> '012' and custno <> '013') order by custname"
'    SQLTEXT = " SELECT DISTINCT GROUPING, CUSTNO, CUSTNAME FROM AIC_LI_CUST_MASTER where (custno <> '018') and (custno <> '001' and custno <>'007' and custno <> '012' and custno <> '013' and custno <> '021') order by custname"
    
    
    rs.Open sqltext, wsDB
     
    If Not rs.EOF Then
      Do While Not rs.EOF
'        cust = Trim(RS("GROUPING")) & "-" & Trim(RS("custno")) & "-" & Trim(RS("custname"))
        cust = Trim(rs("custname")) & " - (" & Trim(rs("GROUPING")) & "|" & Trim(rs("custno")) & ")"
        
        cboCust.AddItem cust
        rs.MoveNext
      Loop
    Else
      MsgBox "No record found in customer master setting!"
    End If
    rs.Close
    Set rs = Nothing
    
    cboruntype.AddItem "Mass Production"
    cboruntype.AddItem "Engineering Lot"
    cboruntype.AddItem "Engineering E1"     '2012-08-15
    
    
    Set chkAllpkg = New ADODB.Recordset
   ' sqltext = "SELECT DISTINCT(WPRD_PRD_GRP_2) PACKAGE FROM WIPPRD WHERE WPRD_PRD_GRP_2 <> 'NA'"
'aims
    sqltext = "SELECT DISTINCT(pdm_package) PACKAGE FROM baic_prodmast WHERE pdm_package <> 'NA'"
    
    chkAllpkg.Open sqltext, wsDB
    Do While Not chkAllpkg.EOF
        package_txt.AddItem Trim(chkAllpkg!PACKAGE)
        chkAllpkg.MoveNext
    Loop
    chkAllpkg.Close
    Set chkAllpkg = Nothing
    Dim chkAllpkgWE As ADODB.Recordset
    Set chkAllpkgWE = New ADODB.Recordset
 '   sqltext = "SELECT DISTINCT(WPRD_PRD_GRP_3) LEAD FROM WIPPRD"
'aims
    sqltext = "SELECT DISTINCT(pdm_lead) LEAD FROM baic_prodmast"
    
    chkAllpkgWE.Open sqltext, wsDB
    Do While Not chkAllpkgWE.EOF
        ld_txt.AddItem Trim(chkAllpkgWE!LEAD)
        chkAllpkgWE.MoveNext
    Loop
    chkAllpkgWE.Close
    Set chkAllpkgWE = Nothing
End Sub

Private Sub Form_Activate()
chkUpdate
Me.top = 0
Me.Left = 0
End Sub



Private Sub fsfom_Click()
If Trim(TARGET_DEVICE_TXT.TEXT) <> "" And InStr(cboCust, "FAIRCHILD") > 0 Then
    s = MsgBox("Generate PID and Marking for FOM Lot?", vbYesNo, "Message")
    If s <> 6 Then
        Exit Sub
    End If
    
    wwfomcode = "??"
    xfompid = "??"
    Dim fomrs As ADODB.Recordset
    Set fomrs = New ADODB.Recordset
    SQL = "select GETDATE(),* from FOM_CALENDAR where FOM_STARTDT <= cast(GETDATE() as date) and FOM_ENDDT >= cast(GETDATE() as date)"
    fomrs.Open SQL, wsDB
    If Not fomrs.EOF Then
        xfomyear = Right(Trim(fomrs!fom_year), 1)
'        xfomyear2 = Right(Trim(fomrs!fom_year), 2)
        xfomweek = Format(Right(Trim(fomrs!fom_weekcode), 2), "00")
'        wwfomcode = Trim(fomrs!fom_year_code) & Trim(fomrs!fom_wws2)
        wwfomcode = Trim(fomrs!fom_year_code) & Trim(fomrs!fom_wwchar)       '20110728


'        Me.fslotcode = "Y" & xfomyear2 & xfomweek & wwfomcode
    Else
        MsgBox "Error! Current date not registered in FOM Calendar.", vbCritical, "Message"
        Exit Sub
    End If
    fomrs.Close
    Set fomrs = Nothing
    
    'default pidno.
    newpdino = "AC" & xfomyear & xfomweek & "XXXXN"
    
    'get last PID#, generate new PID.
'    pidyww = "AC" & xfomyear & xfomweek
'    Set fomrs = New ADODB.Recordset
'    SQL = "select * from aic_loading_instruction where (refno like 'FS%' or refno like 'FP%') and cuslotno like '" & pidyww & "%' order by cuslotno desc"
'    Debug.Print SQL
'    fomrs.Open SQL, wsDB
'    If Not fomrs.EOF Then
'        lastpidno = Int(Mid(Trim(fomrs!CUSLOTNO), 6, 4))
'        newpdino = pidyww & Format(Trim(Str(lastpidno + 1)), "0000") & "N"
'    Else
'        newpdino = pidyww & "0001N"
'    End If
'    fomrs.Close
'    Set fomrs = Nothing
    
    
    
    
    Set fomrs = New ADODB.Recordset
    SQL = "select * from fom_marking where fom_itemid='" & Trim(TARGET_DEVICE_TXT.TEXT) & "'"
    fomrs.Open SQL, wsDB
    If Not fomrs.EOF Then
        If fomrs!fom_dtcode_type <> "S2" Then
            MsgBox "Error! Only 'S2' is registered in FOM Datecode Calendar.", vbCritical, "Message"
            Exit Sub
        End If
        
        If Mid(Trim(fomrs!fom_topmark1), 1, 2) = "$Y" Then
            xfommark0 = "{FAIRCHILD LOGO}"
        Else
            MsgBox "Error! Invalid $Y marking data.", vbCritical, "Message"
            Exit Sub
        End If
        If Mid(Trim(fomrs!fom_topmark1), 3, 2) = "&Z" Then
            xfommark1a = "Y"
        Else
            MsgBox "Error! Invalid &Z marking data.", vbCritical, "Message"
            Exit Sub
        End If
        If Mid(Trim(fomrs!fom_topmark1), 5, 2) = "&2" Then
            xfommark1b = wwfomcode
        Else
            MsgBox "Error! Invalid &2 marking data.", vbCritical, "Message"
            Exit Sub
        End If
        If Mid(Trim(fomrs!fom_topmark1), 7, 2) = "&K" Then
            xfommark1c = "$$"
        Else
            MsgBox "Error! Invalid &K marking data.", vbCritical, "Message"
            Exit Sub
        End If
        xfommark2 = Trim(fomrs!fom_topmark2)
        xfommark3 = Trim(fomrs!fom_topmark3)
    Else
        MsgBox "Data Not Found in Fairchild Master List!", vbCritical, "Message"
        Exit Sub
    End If
    fomrs.Close
    Set fomrs = Nothing
        
    'get last lotcode, generate new lotcode
'    pidyww = "AC" & xfomyear & xfomweek
'    xfommark_partial = xfommark1a & xfommark1b
'    Set fomrs = New ADODB.Recordset
'    SQL = " select * from aic_loading_instruction where (refno like 'FS%' or refno like 'FP%') and " & _
'          " target_device = '" & Trim(TARGET_DEVICE_TXT.TEXT) & "' and top2 like '" & xfommark_partial & "%' and " & _
'          " cuslotno like '" & pidyww & "%' order by top2 desc"
'    Debug.Print SQL
'    fomrs.Open SQL, wsDB
'    If Not fomrs.EOF Then
'        lastpiddcode = Right(Trim(fomrs!TOP2), 2)
'        firstalpha = Left(lastpiddcode, 1)
'        lastalpha = Right(lastpiddcode, 1)
'        If lastalpha = "Z" Then
'            xfommark1c = Chr(Asc(firstalpha) + 1) & "A"
'        Else
'            xfommark1c = firstalpha & Chr(Asc(lastalpha) + 1)
'        End If
'    Else
'        xfommark1c = "AA"
'    End If
'    fomrs.Close
'    Set fomrs = Nothing
    
    xfommark1 = xfommark1a & xfommark1b & xfommark1c
    txtCusLot = newpdino
    topx(0) = xfommark0
    topx(1) = xfommark1
    topx(2) = xfommark2
    topx(3) = xfommark3
        
End If

End Sub

Private Sub homex_Click()
Call userPermission(login_id)
Unload Me
End Sub







Private Sub lblPONo_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    '20211206 Remove uppercase due to Chainpower need upper & small letters.
    KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii = 13 Then
        
        'RETRIEVE CUSTPO
        Dim CpoRs As ADODB.Recordset
        Set CpoRs = New ADODB.Recordset
        sqltxt = "select * from baic_customer where CUS_CODE='" & Left(Right(Trim(cboCust.TEXT), 4), 3) & "'"
        CpoRs.Open sqltxt, wsDB
        If Not CpoRs.EOF Then
            Debug.Print sqltxt
            txt_pomode = Trim(CpoRs!cus_po_mode)
        Else
            txt_pomode = "???"
        End If
        CpoRs.Close
        Set CpoRs = Nothing
        If txt_pomode = "STANDARD" Then
            Set CpoRs = New ADODB.Recordset
            sqltxt = "select * from baic_customer_po where CPO_CUST_SHORTNAME='" & Trim(custnameselect.TEXT) & "' and CPO_PONO='" & lblPONo & "' and CPO_TARGETDEVICE='" & TARGET_DEVICE_TXT & "'"
            CpoRs.Open sqltxt, wsDB
            If Not CpoRs.EOF Then
                total_poqty = CpoRs!cpo_order_qty
                txt_podate = CpoRs!CPO_ORDER_YMD
                total_poqty.Locked = True
                txt_podate.Locked = True
            Else
                txt_podate = Now()
                total_poqty.Locked = False
                txt_podate.Locked = False
            End If
            CpoRs.Close
            Set CpoRs = Nothing
        Else
            total_poqty = "N/A"
            txt_podate = "N/A"
            total_poqty.Locked = True
            txt_podate.Locked = True
        End If
      '---------------------------------------------------------------
      
      
      'Quah 20180718 disabled the autoconverion during ENTER-PO, only convert during ENTER-CUSLOT
        'diana 2015-06-29 NIKO YEAR TO ALPHABETS
'''''''         If Left(Trim(refno_TXT), 2) = "NK" Then
'''''''            If InStr(topx(2), "YWW") Then 'TOP3
'''''''
'''''''                'Quah add 2017-08-23, req by Mon.
'''''''                'Quah add 2 more devices 2017-11-30, req by Mon
'''''''                'NIKO PB521BX REV.AZ, PB606BX REV.BZ, PB5A2BX REV.AZ, PB606BA REV.CZ
'''''''                'Quah 20171229 add PB600BA REV.DZ, req by Mon
'''''''                'qUAH 20180402 ADD PB600BA REV.BZ, req by KC.
'''''''                    If InStr(TARGET_DEVICE_TXT, "PB521BX REV.AZ") > 0 Or InStr(TARGET_DEVICE_TXT, "PB606BX REV.BZ") > 0 _
'''''''                        Or InStr(TARGET_DEVICE_TXT, "PB5A2BX REV.AZ") > 0 Or InStr(TARGET_DEVICE_TXT, "PB606BA REV.CZ") > 0 _
'''''''                        Or InStr(TARGET_DEVICE_TXT, "PB600BA REV.DZ") > 0 Or InStr(TARGET_DEVICE_TXT, "PB600BA REV.BZ") > 0 Then
'''''''                        nkyy = Left(ww_txt, 2)
'''''''                        Select Case nkyy
'''''''                            Case "17"
'''''''                                xnky = "H"
'''''''                            Case "18"
'''''''                                xnky = "I"
'''''''                            Case "19"
'''''''                                xnky = "J"
'''''''                            Case "20"
'''''''                                xnky = "K"
'''''''                            Case "21"
'''''''                                xnky = "L"
'''''''                            Case "22"
'''''''                                xnky = "M"
'''''''                            Case "23"
'''''''                                xnky = "N"
'''''''                            Case "24"
'''''''                                xnky = "O"
'''''''                            Case "25"
'''''''                                xnky = "P"
'''''''                            Case "26"
'''''''                                xnky = "Q"
'''''''                            Case Else
'''''''                                MsgBox "NIKO calendar year not define for these device.", vbCritical, "Message"
'''''''                                Exit Sub
'''''''                        End Select
'''''''
'''''''                        xnkww = Right(ww_txt, 2)
'''''''
'''''''                        '20171218 Quah add condition for NIKO 2X2 PB606BA REV.CZ --> no need prefix A, req by Mon,KC.
'''''''                        '20171229 Quah add for PB600BA REV.DZ
'''''''                        '20180402 Quah add for PB600BA REV.BZ
'''''''                        If InStr(TARGET_DEVICE_TXT, "PB606BA REV.CZ") > 0 Or InStr(TARGET_DEVICE_TXT, "PB600BA REV.DZ") > 0 Or InStr(TARGET_DEVICE_TXT, "PB600BA REV.BZ") > 0 Then
'''''''                            topx(2) = xnky & xnkww
'''''''                        Else
'''''''                            topx(2) = "A" & xnky & xnkww
'''''''                        End If
'''''''                    End If
'''''''
'''''''
''''''''            'REPLACE WITH ALPHA-YEAR AND WORKWEEK FOR YWW
''''''''                Dim NK_Y
''''''''                Dim NKY As ADODB.Recordset
''''''''                Set NKY = New ADODB.Recordset
''''''''                THIS_YEAR = "20" & Mid(refno_TXT, 3, 2)
''''''''                SSQL = " select CUS_DATA_1 from BAIC_CUSTOMER_ADDSETUP where CUS_RECTYPE='NIKO_YEAR' and CUS_KEY_1='" & THIS_YEAR & "' "
''''''''                Debug.Print SSQL
''''''''                NKY.Open SSQL, wsDB
''''''''                    If NKY.EOF = False Then
''''''''                        NK_Y = NKY!cus_data_1 & Mid(ww_txt, 2, 2)
''''''''                    Else
''''''''                        MsgBox "NIKO MARKING NO YEAR FOUND!PLEASE CHECK!", vbCritical, "Message"
''''''''                        Exit Sub
''''''''                    End If
''''''''
''''''''                    topx(2).TEXT = Replace(topx(2).TEXT, "YWW", NK_Y)
''''''''                NKY.Close
''''''''                Set NKY = Nothing
'''''''            End If
'''''''         End If
         
        dd1.SetFocus
    End If
End Sub

Private Sub ld_txt_Click()
    Dim rsbd As ADODB.Recordset
    Set rsbd = New ADODB.Recordset
'    CSQLSTRING = "SELECT * FROM WIPPRD WHERE  WPRD_PRD_GRP_2 = '" & Trim(package_txt) & "' AND  WPRD_PRD_GRP_3 = '" & Trim(ld_txt) & "'"
    
'aims
    CSQLSTRING = "SELECT * FROM baic_prodmast where pdm_package = '" & Trim(package_txt) & "' AND  pdm_lead = '" & Trim(ld_txt) & "'"

    rsbd.Open CSQLSTRING, wsDB
    If rsbd.EOF Then
         MsgBox "Package & Lead is not setup in Database. Check with Planner!!!"
    Else
    End If
    rsbd.Close
    FG = Trim(package_txt) & Trim(ld_txt)
    If FG = "PDIP32" Or FG = "PDIP40" Then
        bottom(3).Enabled = False
        bottom(4).Enabled = False
        bottom(5).Enabled = False
     Else
        bottom(3).Enabled = True
        bottom(4).Enabled = True
        bottom(5).Enabled = True
     End If
End Sub

Private Sub ld_txt_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        Dim rsbd As ADODB.Recordset
        Set rsbd = New ADODB.Recordset
  '      CSQLSTRING = "SELECT * FROM WIPPRD WHERE  WPRD_PRD_GRP_2 = '" & Trim(package_txt) & "' AND  WPRD_PRD_GRP_3 = '" & Trim(ld_txt) & "'"
'AIMS
    CSQLSTRING = "SELECT * FROM baic_prodmast where pdm_package = '" & Trim(package_txt) & "' AND  pdm_lead = '" & Trim(ld_txt) & "'"

        rsbd.Open CSQLSTRING, wsDB
        If rsbd.EOF Then
            MsgBox " Package & Lead is not setup in Database. Check with Planner!!!"
        Else
            bonding_diagram_txt.Visible = True
            bonding_diagram_txt.SetFocus
            target_device_txt1.Visible = False
        End If
        rsbd.Close
        FG = Trim(package_txt) & Trim(ld_txt)
        If FG = "PDIP32" Or FG = "PDIP40" Then
            bottom(3).Enabled = False
            bottom(4).Enabled = False
            bottom(5).Enabled = False
        Else
            bottom(3).Enabled = True
            bottom(4).Enabled = True
            bottom(5).Enabled = True
        End If
    End If
End Sub

Private Sub INIT_MARK()
'    topx(0) = "":    topx(1) = ""
'    topx(2) = "":    topx(3) = ""
'    topx(4) = "":    topx(5) = ""
'    bottom(0) = "":    bottom(1) = ""
'    bottom(2) = "":    bottom(3) = ""
'    bottom(4) = "":    bottom(5) = ""
    
    For Xr = 0 To 5
        topInfo(Xr) = "":     botInfo(Xr) = ""
        topx(Xr) = vbNullString
        bottom(Xr) = vbnullstirng
        
        'Quah 20160517 enable Taggle for edit  marking.
        'Quah 20170619 enable Silterra for edit marking.
        'Quah 20180720 enable Taggle for edit marking
        'Quah 20181109 enabled MIKROELEKTRONIK for edit marking.
        'Quah 20190704 enabled for CHINA customer, req by Alice.
        If txtCustomer.TEXT = "TAIWAN" Or txtCustomer.TEXT = "CHINA" Or InStr(cboCust.TEXT, "TAGGLE") > 0 Or InStr(cboCust.TEXT, "MKR-IC") > 0 Or InStr(cboCust.TEXT, "SILTERRA") > 0 Or InStr(cboruntype, "Engineering") > 0 Then  '2014-11-18
            topx(Xr).Enabled = True
            bottom(Xr).Enabled = True
        'yana2022 open for 2 rework device req by mon
        '20230912 AIN ADD FOR KAK ROSE TO ENABLE MARKING EDIT FOR SOT23.AQ24COM.TGRPBF DEVICE
        ElseIf Trim(bdcombo.TEXT) = "SOIR14.NPA1C02142TNIP.RMAPBF" Or Trim(bdcombo.TEXT) = "SOIR14.NPA1C02142CNIP.RMAPBF" Or Trim(bdcombo.TEXT) = "SOT23.AQ24COM.TGRPBF" Then
            topx(Xr).Enabled = True
            bottom(Xr).Enabled = True
        'YANA ADD FOR MAHADHIR TO ENABLE MARKING EDIT FOR ENGINERING DEVICE
        ElseIf InStr(Trim(bdcombo.TEXT), "ENPA") Then
            topx(Xr).Enabled = True
            bottom(Xr).Enabled = True
        'YANA ADD FOR KAK ROSE TO ENABLE MARKING EDIT FOR REWORK DEVICE
        ElseIf InStr(Trim(bdcombo.TEXT), "RMA") Then
            topx(Xr).Enabled = True
            bottom(Xr).Enabled = True
        Else
            topx(Xr).Enabled = False
            bottom(Xr).Enabled = False
        End If
    Next Xr
End Sub

Private Sub mark_spec_txt_KeyPress(KeyAscii As Integer)
Dim rs As ADODB.Recordset
Dim sqlstr As String
Dim topMark As String
Dim cnt As Integer
Dim top, bottom As String
Dim Rs03 As ADODB.Recordset

KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    If Left(refno_TXT, 2) <> "NN" And Left(refno_TXT, 2) <> "NZ" Then   'Quah 20080613 add condition.
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If LenB(Trim(ww_txt)) = 0 Then
        MsgBox "ERROR - Please Enter Start Date Before Key in MarkSpec.", vbCritical
        Exit Sub
    End If
    
    FRU = InputBox("PLEASE INSERT MARK SPEC AGAIN!!", "AICSLI")
    If FRU <> "" Then
        If Trim(mark_spec_txt) <> FRU Then
            MsgBox "PLEASE RE-KEYIN MARK SPEC!!!", vbCritical
            mark_spec_txt = ""
            mark_spec_txt.SetFocus
            INIT_MARK
        Else
            '======================================================
            'CHECK BD AND MARK SPEC
'           If Left(Right(Trim(cboCust.TEXT), 6), 1) <> "M" And Left(Trim(cboruntype), 11) <> "Engineering" Then
           If Left(Trim(cboruntype), 11) <> "Engineering" Then
                
'                SQL$ = "SELECT * FROM AIC_DEVICE_CONTROL WHERE CUSTOMER='CUSBD' AND TARGETDEVICE='MARKSPEC'" & _
'                    " AND REMARK1='" & Trim(Me.bonding_diagram_txt) & "' AND REMARK2='" & Trim(Me.mark_spec_txt) & "'"
                
'Quah 20091221 add remark7 for targetdevice matching, exclude GMT and SGC from
'Quah 2010-05-18 for GMT, exlcude Targetdevice matching.

'Quah 2010-12-28 chg matching logic
'            If Left(refno_TXT, 2) = "GM" Then
'                SQL$ = "SELECT * FROM AIC_DEVICE_CONTROL WHERE CUSTOMER='CUSBD' AND TARGETDEVICE='MARKSPEC'" & _
'                    " AND REMARK1='" & Trim(Me.bonding_diagram_txt) & "' AND REMARK2='" & Trim(Me.mark_spec_txt) & "'"
'            Else
'                SQL$ = "SELECT * FROM AIC_DEVICE_CONTROL WHERE CUSTOMER='CUSBD' AND TARGETDEVICE='MARKSPEC'" & _
'                    " AND REMARK1='" & Trim(Me.bonding_diagram_txt) & "' AND REMARK2='" & Trim(Me.mark_spec_txt) & _
'                    "' AND REMARK7 = '" & Trim(Me.TARGET_DEVICE_TXT) & "'"
'            End If


           If Left(refno_TXT, 2) = "GM" Then
                SQL$ = "SELECT * FROM AIC_DEVICE_CONTROL WHERE CUSTOMER='CUSBD' AND TARGETDEVICE='MARKSPEC'" & _
                    " AND REMARK1='" & Trim(Me.bonding_diagram_txt) & "' AND REMARK2='" & Trim(Me.mark_spec_txt) & "'"
            Else
                SQL$ = "SELECT * FROM AIC_DEVICE_CONTROL WHERE CUSTOMER='CUSBD' AND TARGETDEVICE='MARKSPEC'" & _
                    " AND REMARK2='" & Trim(Me.mark_spec_txt) & _
                    "' AND REMARK4 = '" & Trim(Me.internal_device_no_txt) & "'"
            End If


                Set rs = New ADODB.Recordset
                Debug.Print SQL$
                rs.Open SQL$, wsDB
                If rs.EOF = False Then
                    If Trim(rs!REMARK3) = "CLOSE" Then
                        MsgBox "PARTICULAR BONDING DIAGRAM WITH MARKING SPEC ALREADY CLOSED!PLEASE CHECK!", vbCritical, "ERROR"
                        Exit Sub
                    End If
                Else
                    MsgBox "PARTICULAR BONDING DIAGRAM NOT INITIALIZE WITH MARKING SPEC!PLEASE CHECK!", vbCritical, "ERROR"
                    Exit Sub
                End If
            
                'MARKING SPEC CONTROL
                If Left(Trim(refno_TXT), 2) = "AS" Then
                    Call MarkSpec_Sanjose
                Else
                    Call MarkSpec
                    If Left(Trim(refno_TXT), 2) = "AD" Then
                        Call MarkSpec_AD
                    End If
                    'Quah 2013-01-09 AVT=NIKO
                    If Left(Trim(refno_TXT), 2) = "NK" Or Left(Trim(refno_TXT), 2) = "AO" Then
                        Call MarkSpec_NK
                    End If
                    If Left(Trim(refno_TXT), 2) = "II" Then
                        Call MarkSpec_II
                    End If
                    If Left(Trim(refno_TXT), 2) = "CT" Then
                        Call MarkSpec_CT
                    End If
                    If Left(Trim(refno_TXT), 2) = "GT" Then
                        Call MarkSpec_GT
                    End If
                    If Left(Trim(refno_TXT), 2) = "MP" Then  'Quah 20141215
'                        If InStr(topx(2).TEXT, "YYWW") > 0 Then
'                            topx(2).TEXT = Replace(topx(2).TEXT, "YYWW", ww_txt)
'                        End If
                        'Quah 20151223 new requirement
                        If InStr(topx(2).TEXT, "WYW") > 0 Then
                            Dim mpyy, mpww, mpayy, mp_wyw
                            mpyy = Left(ww_txt, 2)
                            mpww = Right(ww_txt, 2)
                            Select Case mpyy
                                Case "15"
                                    mpayy = "F"
                                Case "16"
                                    mpayy = "G"
                                Case "17"
                                    mpayy = "H"
                                Case "18"
                                    mpayy = "J"
                                Case "19"
                                    mpayy = "K"
                                Case "20"
                                    mpayy = "L"
                                Case "21"
                                    mpayy = "M"
                                Case "22"
                                    mpayy = "N"
                                Case "23"
                                    mpayy = "P"
                                Case "24"
                                    mpayy = "Q"
                                Case "25"
                                    mpayy = "R"
                                Case "26"
                                    mpayy = "S"
                                Case "27"
                                    mpayy = "T"
                                Case "28"
                                    mpayy = "U"
                                Case "29"
                                    mpayy = "V"
                                Case "30"
                                    mpayy = "W"
                                Case "31"
                                    mpayy = "X"
                                Case "32"
                                    mpayy = "Y"
                                Case "33"
                                    mpayy = "Z"
                                Case Otherwise
                                    MsgBox "Marking YEAR CODE not defined!", vbCritical, "Message"
                                    End
                            End Select
                            mp_wyw = Left(mpww, 1) & mpayy & Right(mpww, 1)
'                            topx(2).TEXT = Replace(topx(2).TEXT, "YYWW", ww_txt)
                            topx(2).TEXT = Replace(topx(2).TEXT, "WYW", mp_wyw)
                        End If
                    End If
                    
                    
                    If Left(Trim(refno_TXT), 2) = "BE" And LenB(txtCusLot) > 0 Then  'Quah 20090311 Bourns auto marking
                        bepoint = 0     'Quah 20091201 skip .1,.2 inserted by Planner to differentiate loading batch.
                        bepoint = InStr(txtCusLot, ".")
                        If bepoint > 0 Then
                            'KO ADD 20150909
                            'Changed as Per reqeust by Bourns ( engineer aic - Syamil 2015SEP09 )- SOIC PACKAGE ONLY
                            ' FROM 2 DIGI CHARACTER TO 3 DIGI EFFECTIVE 20150909
                            If Left(cbofullpackage, 4) = "SOIC" Then
                               topx(2).TEXT = Right(Mid(txtCusLot, 1, bepoint - 1), 3)
                            Else
                                topx(2).TEXT = Right(Mid(txtCusLot, 1, bepoint - 1), 2)
                            End If
                        Else
                        'Changed as Per reqeust by Bourns ( engineer aic - Syamil 2015SEP09 )-SOIC PACKAGE ONLY
                        ' FROM 2 DIGI CHARACTER TO 3 DIGI EFFECTIVE 20150909
                          'topx(2).TEXT = Right(txtCusLot, 2)
                            
                        '20181116 Quah addd, req by Anita
                        '20190118---> 3 LAST CUSLOTNO + MD
                        If InStr(topx(2).TEXT, "XXXXX") > 0 Then
                            BE_LAST5 = Right(WAFER.TEXT, 3) + "MD"
                            topx(2).TEXT = BE_LAST5
                        End If
                            
''''                            If Left(cbofullpackage, 4) = "SOIC" Then ' ko add for SOIC
''''                               topx(2).TEXT = Right(txtCusLot, 3)
''''                            Else
''''                               topx(2).TEXT = Right(txtCusLot, 2)
''''                            End If
                        End If
                        If InStr(topx(3).TEXT, "YWW") > 0 Then
                            y_ww = Right(ww_txt, 3)
                            topx(3).TEXT = Replace(topx(3).TEXT, "YWW", y_ww)
                        End If
                    End If
                    If Left(Trim(refno_TXT), 2) = "MN" Then 'Quah 20090629 Magnachip auto marking
'                        If Trim(topx(2).TEXT) = "YWLLLLGC" Then
                        If Left(Trim(topx(2).TEXT), 6) = "YWLLLL" Then 'Quah 2010-04-12 LL request to check only first 6 chars, last 2 is not fixed.
                            Set Rs03 = New ADODB.Recordset
                            CSQLSTRING = "select * from wwcal where ww_date='" & Format(Date, "DD-MMM-YYYY") & "'"
                            Debug.Print CSQLSTRING
                            Rs03.Open CSQLSTRING, wsDB
                            If Not Rs03.EOF Then
'                                topx(2).TEXT = Trim(Rs03!ww_magnachip) & "(ASSY#L4#)GC"
                                topx(2).TEXT = Trim(Rs03!ww_magnachip) & "(ASSY#L4#)" & Right(Trim(topx(2).TEXT), 2)
                            Else
                                MsgBox "ERROR IN MAGNACHIP MARKING-DATECODE. PLEASE CHECK !!!!", vbCritical
                                topx(2).TEXT = ""
                                mark_spec_txt.TEXT = ""
                                Exit Sub
                            End If
                            Rs03.Close
                        End If
                    End If
                    If Left(Trim(refno_TXT), 2) = "UL" Then 'Quah 20091008 Ultrachip auto marking
                        If Trim(topx(2).TEXT) = "YYMMWW" Then
                            Set Rs03 = New ADODB.Recordset
                            CSQLSTRING = "select * from wwcal where ww_date='" & Format(Date, "DD-MMM-YYYY") & "'"
                            Debug.Print CSQLSTRING
                            Rs03.Open CSQLSTRING, wsDB
                            If Not Rs03.EOF Then
                                topx(2).TEXT = Trim(Rs03!ww_ultrachip)
                            Else
                                MsgBox "ERROR IN ULTRACHIP MARKING-DATECODE. PLEASE CHECK !!!!", vbCritical
                                topx(2).TEXT = ""
                                mark_spec_txt.TEXT = ""
                                Exit Sub
                            End If
                            Rs03.Close
                        End If
                    End If
                    
                    If Left(Trim(refno_TXT), 2) = "EN" Then
                        If Trim(topx(2).TEXT) = "F-XXXXXX" Then
                            topx(2).TEXT = Replace(topx(2).TEXT, "XXXXXX", txtCusLot)
                        End If
                    End If
                    
                                     
                    'Quah 20090211 add Fairchild-Pg (FP)
                    If Left(Trim(refno_TXT), 2) = "FS" Or Left(Trim(refno_TXT), 2) = "FP" Then     'Different prefix setting.
                        Set Rs03 = New ADODB.Recordset
                        CSQLSTRING = "SELECT * FROM AIC_SETUP_DATA WHERE TABLE_NAME='FSMARKING' and PROG_ITEM1='" & Trim(Me.TARGET_DEVICE_TXT) & "'"
                        Rs03.Open CSQLSTRING, wsDB
                        If Not Rs03.EOF Then
                            X1 = Trim(Rs03!prog_item2)
                            X2 = Trim(Rs03!prog_item3)
                        Else
                            X1 = ""
                            X2 = ""
                        End If
                        Rs03.Close
                        
                        'Quah 2010-05-28 add condition. If not in the above device list, then topx(2) = mark spec
                        If X1 <> "" Then
                            Me.topx(2) = X1 & Left(Trim(WAFER.TEXT), 6) & X2
                        Else
                            Me.topx(2) = Me.mark_spec_txt
                        End If
                        
'
'                        If Trim(Me.TARGET_DEVICE_TXT) = "LTA504SGZ" Then
'                            Me.topx(2) = "H1" & Left(Trim(WAFER.Text), 6) & "YYWWQ"
'                        ElseIf Trim(Me.TARGET_DEVICE_TXT) = "SG6961SZ" Then
'                            Me.topx(2) = "B1" & Left(Trim(WAFER.Text), 6) & "YYWWQ"
'                        ElseIf Trim(Me.TARGET_DEVICE_TXT) = "LTA504SGZF" Then
'                            Me.topx(2) = "F1" & Left(Trim(WAFER.Text), 6) & "YYWWQ"
'                        ElseIf Trim(Me.TARGET_DEVICE_TXT) = "SG5841DZ" Then
'                            Me.topx(2) = "D3" & Left(Trim(WAFER.Text), 6) & "YYWWQ"
'                        ElseIf Trim(Me.TARGET_DEVICE_TXT) = "SG5841JDZ" Then
'                            Me.topx(2) = "D3" & Left(Trim(WAFER.Text), 6) & "YYWWQ"
'                        ElseIf Trim(Me.TARGET_DEVICE_TXT) = "SG6848DZ1" Then
'                            Me.topx(2) = "B" & Left(Trim(WAFER.Text), 6) & "YYWWQ"
'                        ElseIf Trim(Me.TARGET_DEVICE_TXT) = "SG6846LSZ" Then
'                            Me.topx(2) = "F" & Left(Trim(WAFER.Text), 6) & "YYWWQ"
'                        ElseIf Trim(Me.TARGET_DEVICE_TXT) = "SG6846LDZ" Then
'                            Me.topx(2) = "F" & Left(Trim(WAFER.Text), 6) & "YYWWQ"
'                        ElseIf Trim(Me.TARGET_DEVICE_TXT) = "SG684965DZ" Then
'                            Me.topx(2) = "B" & Left(Trim(WAFER.Text), 6) & "YYWWQ"
'                        ElseIf Trim(Me.TARGET_DEVICE_TXT) = "SG6841DZ" Then
'                            Me.topx(2) = "K" & Left(Trim(WAFER.Text), 6) & "YYWWQ"
'                        ElseIf Trim(Me.TARGET_DEVICE_TXT) = "SG6841SZ" Then
'                            Me.topx(2) = "K" & Left(Trim(WAFER.Text), 6) & "YYWWQ"
'                        Else
'                            Me.topx(2) = "F" & Left(Trim(WAFER.Text), 6) & "YYWWQ"
'                        End If
                    End If
                
                End If
           End If       'skip Marking
        End If
    Else
        mark_spec_txt = ""
        mark_spec_txt.SetFocus
        INIT_MARK
    End If
    End If
End If
End Sub
Private Sub MarkSpec_Sanjose()
INIT_MARK

Dim rsbd As ADODB.Recordset
Set rsbd = New ADODB.Recordset

CSQLSTRING = "SELECT * FROM AIC_MARKING_SPEC_X WHERE  SPECNO = '" & Trim(mark_spec_txt) & "'"
rsbd.Open CSQLSTRING, wsDB
If Not rsbd.EOF Then
   XWORKW = Trim(ww_txt)
   XWORKW_YWW = Right(Trim(ww_txt), 3)
   XWORKW_WW = Right(Trim(ww_txt), 2)
   XWAFERC = Trim(wafer_lot_txt)
   FRNTFER = Trim(Left(wafer_lot_txt, 2))
   
   YOSHTRA1 = "N"
   YOSHTRA2 = "N"
   YOSHTRA3 = "N"
   YOSHTRA4 = "N"
   YOSHTRA5 = "N"
   YOSHTRA6 = "N"
   
   YOSHTRI1 = "N"
   YOSHTRI2 = "N"
   YOSHTRI3 = "N"
   YOSHTRI4 = "N"
   YOSHTRI5 = "N"
   YOSHTRI6 = "N"
   
   topInfo(0) = Trim((rsbd("top1")))
   topInfo(1) = Trim((rsbd("top2")))
   topInfo(2) = Trim((rsbd("top3")))
   topInfo(3) = Trim((rsbd("top4")))
   topInfo(4) = Trim((rsbd("top5")))
   topInfo(5) = Trim((rsbd("top6")))
   '-------------------------------------
   TEMPINFO1 = Trim((rsbd("top1")))
   TEMPINFO2 = Trim((rsbd("top2")))
   TEMPINFO3 = Trim((rsbd("top3")))
   TEMPINFO4 = Trim((rsbd("top4")))
   TEMPINFO5 = Trim((rsbd("top5")))
   TEMPINFO6 = Trim((rsbd("top6")))
   
   'TOPM1
   If Trim((rsbd("top1"))) = "{}" Then
    topInfo(0) = Trim((rsbd("top1")))
    topx(0) = topInfo(0)
   End If
   
   'TOPM2-------------------------------------------------------
   
   If InStr(1, Trim((rsbd("top2"))), "ATMEL") > 0 Then
    TEMPINFO2 = TEMPINFO2
   
   YOSHTRA2 = "Y"
   End If
   If InStr(1, Trim((rsbd("top2"))), "ATMEL LOGO") > 0 Then
    TEMPINFO2 = TEMPINFO2
    
    YOSHTRA2 = "Y"
   End If
   If InStr(1, Trim((rsbd("top2"))), "ATMEL SJ") > 0 Then
    TEMPINFO2 = TEMPINFO2
    
    YOSHTRA2 = "Y"
   End If
   
   If InStr(1, Trim((rsbd("top2"))), "YYWW") > 0 Then
    CHKTOP2 = XWORKW
    TEMPINFO2 = Replace(TEMPINFO2, "YYWW", CHKTOP2)
    
    YOSHTRA2 = "Y"
   ElseIf InStr(1, Trim((rsbd("top2"))), "YWW") > 0 Then
    CHKTOP2 = XWORKW_YWW
    TEMPINFO2 = Replace(TEMPINFO2, "YWW", CHKTOP2)
    
    YOSHTRA2 = "Y"
   End If
                  
    If YOSHTRA2 = "Y" Then
        topx(1) = TEMPINFO2
    topInfo(1).BackColor = &HC0FFFF
    topInfo(1).ForeColor = RED
    topInfo(1).FontBold = True
    End If
    
   'END TOPM2----------------------------------------------------
   
   'TOPM3
   If InStr(1, Trim((rsbd("top3"))), "YYWW") > 0 Then
    CHKTOP3 = XWORKW
    TEMPINFO3 = Replace(TEMPINFO3, "YYWW", CHKTOP3)
   
    YOSHTRA3 = "Y"
   ElseIf InStr(1, Trim((rsbd("top3"))), "YWW") > 0 Then
    CHKTOP3 = XWORKW_YWW
    TEMPINFO3 = Replace(TEMPINFO3, "YWW", CHKTOP3)
    
    YOSHTRA3 = "Y"
   End If
    
    If YOSHTRA3 = "Y" Then
        topx(2) = TEMPINFO3
        topInfo(2).BackColor = &HC0FFFF
        topInfo(2).ForeColor = RED
        topInfo(2).FontBold = True
    End If
    
   
   'TOPM4
   If InStr(1, Trim((rsbd("top4"))), "YYWW") > 0 Then
    CHKTOP4 = XWORKW
    TEMPINFO4 = Replace(TEMPINFO4, "YYWW", CHKTOP4)
    
    YOSHTRA4 = "Y"
   ElseIf InStr(1, Trim((rsbd("top4"))), "YWW") > 0 Then
    CHKTOP4 = XWORKW_YWW
    TEMPINFO4 = Replace(TEMPINFO4, "YWW", CHKTOP4)
    
    YOSHTRA4 = "Y"
   End If
   
   If YOSHTRA4 = "Y" Then
    topx(3) = TEMPINFO4
    topInfo(3).BackColor = &HC0FFFF
    topInfo(3).ForeColor = RED
    topInfo(3).FontBold = True
   End If
   
   'TOPM5
   If InStr(1, Trim((rsbd("top5"))), "YYWW") > 0 Then
    CHKTOP5 = XWORKW
    TEMPINFO5 = Replace(TEMPINFO5, "YYWW", CHKTOP5)
    
    YOSHTRA5 = "Y"
   ElseIf InStr(1, Trim((rsbd("top5"))), "YWW") > 0 Then
    CHKTOP5 = XWORKW_YWW
    TEMPINFO5 = Replace(TEMPINFO5, "YWW", CHKTOP5)
    
    YOSHTRA5 = "Y"
   End If
   
    If YOSHTRA5 = "Y" Then
    topx(4) = TEMPINFO5
    topInfo(4).BackColor = &HC0FFFF
    topInfo(4).ForeColor = RED
    topInfo(4).FontBold = True
    End If
   
   'TOPM6
   If InStr(1, Trim((rsbd("top6"))), "YYWW") > 0 Then
    CHKTOP6 = XWORKW
    TEMPINFO6 = Replace(TEMPINFO6, "YYWW", CHKTOP6)
    
    YOSHTRA6 = "Y"
   ElseIf InStr(1, Trim((rsbd("top6"))), "YWW") > 0 Then
    CHKTOP6 = XWORKW
    TEMPINFO6 = Replace(TEMPINFO6, "YWW", CHKTOP6)
    
    YOSHTRA6 = "Y"
   End If
   
   If InStr(1, Trim((rsbd("top6"))), "SL449") > 0 Then
    TEMPINFO6 = Replace(TEMPINFO6, "SL449", "SL449")
    
    YOSHTRA6 = "Y"
   ElseIf InStr(1, Trim((rsbd("top6"))), "SL473") > 0 Then
    TEMPINFO6 = Replace(TEMPINFO6, "SL473", "SL473")
    
    YOSHTRA6 = "Y"
   End If
   
    If YOSHTRA6 = "Y" Then
        topx(5) = TEMPINFO6
    topInfo(5).BackColor = &HC0FFFF
    topInfo(5).ForeColor = RED
    topInfo(5).FontBold = True
    End If

    '-----BOTTOM MARK---->
    botInfo(0) = Trim((rsbd("bottom1")))
    botInfo(1) = Trim((rsbd("bottom2")))
    botInfo(2) = Trim((rsbd("bottom3")))
    botInfo(3) = Trim((rsbd("bottom4")))
    botInfo(4) = Trim((rsbd("bottom5")))
    botInfo(5) = Trim((rsbd("bottom6")))
    
    BEMPINFO1 = Trim((rsbd("bottom1")))
    BEMPINFO2 = Trim((rsbd("bottom2")))
    BEMPINFO3 = Trim((rsbd("bottom3")))
    BEMPINFO4 = Trim((rsbd("bottom4")))
    BEMPINFO5 = Trim((rsbd("bottom5")))
    BEMPINFO6 = Trim((rsbd("bottom6")))
    
   'bottom1---->
   If Trim((rsbd("bottom1"))) = "{}" Then
    botInfo(0) = Trim((rsbd("bottom1")))
    bottom(0) = botInfo(0)
   End If
    
   If InStr(1, Trim((rsbd("bottom1"))), "CCCCCC-F") > 0 Then
    bottom(0) = XWAFERC
    botInfo(0).BackColor = &HC0FFFF
    botInfo(0).ForeColor = RED
    botInfo(0).FontBold = True
   ElseIf InStr(1, Trim((rsbd("bottom1"))), "YYWW") > 0 Then
        bottom(0) = XWORKW
        botInfo(0).BackColor = &HC0FFFF
        botInfo(0).ForeColor = RED
        botInfo(0).FontBold = True
   ElseIf InStr(1, Trim((rsbd("bottom1"))), "YWW") > 0 Then
        bottom(0) = Right(XWORKW, 3)
        botInfo(0).BackColor = &HC0FFFF
        botInfo(0).ForeColor = RED
        botInfo(0).FontBold = True
   End If
   
   'bottom2--->
   If InStr(1, Trim((rsbd("bottom2"))), "CCCCCC-F") > 0 Then
        'bottom(1) = XWAFERC
        BEMPINFO2 = Replace(BEMPINFO2, "CCCCCC-F", XWAFERC)
        botInfo(1).BackColor = &HC0FFFF
        botInfo(1).ForeColor = RED
        botInfo(1).FontBold = True
        
        YOSHTRI2 = "Y"
   ElseIf InStr(1, Trim((rsbd("bottom2"))), "F") > 0 Then
        'bottom(1) = XWAFERC
        
        'additional for soik
        If Left(Trim(bdcombo), 4) = "SOIK" Then
        
            If InStr(1, XWAFERC, "-") > 0 Then
            BEMPINFO2 = Replace(BEMPINFO2, "F", Right(XWAFERC, 1))
            botInfo(1).BackColor = &HC0FFFF
            botInfo(1).ForeColor = RED
            botInfo(1).FontBold = True
            End If
        
        Else
        BEMPINFO2 = Replace(BEMPINFO2, "F", Right(XWAFERC, 1))
        botInfo(1).BackColor = &HC0FFFF
        botInfo(1).ForeColor = RED
        botInfo(1).FontBold = True
        End If
        YOSHTRI2 = "Y"
   End If
   
   If InStr(1, Trim((rsbd("bottom2"))), "YYWW") > 0 Then
        'bottom(1) = XWORKW
        BEMPINFO2 = Replace(BEMPINFO2, "YYWW", XWORKW)
        botInfo(1).BackColor = &HC0FFFF
        botInfo(1).ForeColor = RED
        botInfo(1).FontBold = True
        
        YOSHTRI2 = "Y"
   ElseIf InStr(1, Trim((rsbd("bottom2"))), "YWW") > 0 Then
        'bottom(1) = Right(XWORKW, 3)
        BEMPINFO2 = Replace(BEMPINFO2, "YWW", XWORKW_YWW)
        botInfo(1).BackColor = &HC0FFFF
        botInfo(1).ForeColor = RED
        botInfo(1).FontBold = True
        
        YOSHTRI2 = "Y"
   End If
   
    If YOSHTRI2 = "Y" Then
        bottom(1) = BEMPINFO2
    End If
   
   'bottom3----->
   If InStr(1, Trim((rsbd("bottom3"))), "yQ") > 0 Then
        
        BEMPINFO3 = Replace(BEMPINFO3, "yQ", FRNTFER)
        botInfo(2).BackColor = &HC0FFFF
        botInfo(2).ForeColor = RED
        botInfo(2).FontBold = True
        
        YOSHTRI3 = "Y"
   End If
   
   If InStr(1, Trim((rsbd("bottom3"))), "CCCCCC") > 0 Then
        
        If Left(Trim(bdcombo), 4) = "SOIK" Then
        
            If InStr(1, XWAFERC, "-") > 0 Then
            KOPI = InStr(1, XWAFERC, "-")

            BEMPINFO3 = Replace(BEMPINFO3, "CCCCCC", Left(XWAFERC, KOPI - 1))
            botInfo(2).BackColor = &HC0FFFF
            botInfo(2).ForeColor = RED
            botInfo(2).FontBold = True
           
            Else
            
            BEMPINFO3 = Replace(BEMPINFO3, "CCCCCC", XWAFERC)
            botInfo(2).BackColor = &HC0FFFF
            botInfo(2).ForeColor = RED
            botInfo(2).FontBold = True
            'End If
            
            End If
        
        Else
        
            BEMPINFO3 = Replace(BEMPINFO3, "CCCCCC", XWAFERC)
            botInfo(2).BackColor = &HC0FFFF
            botInfo(2).ForeColor = RED
            botInfo(2).FontBold = True
        
        End If
        
        YOSHTRI3 = "Y"
   End If
   
   If InStr(1, Trim((rsbd("bottom3"))), "YYWW") > 0 Then
        'bottom(2) = XWORKW
        BEMPINFO3 = Replace(BEMPINFO3, "YYWW", XWORKW)
        botInfo(2).BackColor = &HC0FFFF
        botInfo(2).ForeColor = RED
        botInfo(2).FontBold = True
        
        YOSHTRI3 = "Y"
   ElseIf InStr(1, Trim((rsbd("bottom3"))), "YWW") > 0 Then
        'bottom(2) = XWORKW
        BEMPINFO3 = Replace(BEMPINFO3, "YWW", XWORKW_YWW)
        botInfo(2).BackColor = &HC0FFFF
        botInfo(2).ForeColor = RED
        botInfo(2).FontBold = True
        
        YOSHTRI3 = "Y"
   End If
           
    If YOSHTRI3 = "Y" Then
        bottom(2) = BEMPINFO3
    End If
   
   'bottom4----->
   If InStr(1, Trim((rsbd("bottom4"))), "yQ") > 0 Then
        'bottom(3) = FRNTFER & XWORKW
        BEMPINFO4 = Replace(BEMPINFO4, "yQ", FRNTFER)
        botInfo(3).BackColor = &HC0FFFF
        botInfo(3).ForeColor = RED
        botInfo(3).FontBold = True
        
        YOSHTRI4 = "Y"
   End If
   
   If InStr(1, Trim((rsbd("bottom4"))), "YYWW") > 0 Then
        'bottom(3) = XWORKW
        BEMPINFO4 = Replace(BEMPINFO4, "YYWW", XWORKW)
        botInfo(3).BackColor = &HC0FFFF
        botInfo(3).ForeColor = RED
        botInfo(3).FontBold = True
        
        YOSHTRI4 = "Y"
   End If
   
    If YOSHTRI4 = "Y" Then
        bottom(3) = BEMPINFO4
    End If
    
    'CATHERINE 070522 - BOTTOM 3 IS CUSTLOT
    bottom(1).TEXT = Trim(txtCusLot)
    
   'bottom5----->
   If InStr(1, Trim((rsbd("bottom5"))), "yQ") > 0 Then
        'bottom(4) = FRNTFER & XWORKW
        BEMPINFO5 = Replace(BEMPINFO5, "yQ", FRNTFER)
        botInfo(4).BackColor = &HC0FFFF
        botInfo(4).ForeColor = RED
        botInfo(4).FontBold = True
        
        YOSHTRI5 = "Y"
   End If
   
   If InStr(1, Trim((rsbd("bottom5"))), "YYWW") > 0 Then
        'bottom(4) = XWORKW
        BEMPINFO5 = Replace(BEMPINFO5, "YYWW", XWORKW)
        botInfo(4).BackColor = &HC0FFFF
        botInfo(4).ForeColor = RED
        botInfo(4).FontBold = True
        
        YOSHTRI5 = "Y"
   End If
   
    If YOSHTRI5 = "Y" Then
        bottom(4) = BEMPINFO5
    End If
   
   'bottom6----->
   If InStr(1, Trim((rsbd("bottom6"))), "yQYYWW") > 0 Then
        bottom(5) = FRNTFER & XWORKW
        botInfo(5).BackColor = &HC0FFFF
        botInfo(5).ForeColor = RED
        botInfo(5).FontBold = True
   ElseIf InStr(1, Trim((rsbd("bottom6"))), "YYWW") > 0 Then
        bottom(5) = XWORKW
        botInfo(5).BackColor = &HC0FFFF
        botInfo(5).ForeColor = RED
        botInfo(5).FontBold = True
   End If
      
Else
    MsgBox "Marking Spec is not setup in Database. Check with Planner!!!"
    ld_txt.SetFocus
     
End If
rsbd.Close
  
FG = Trim(package_txt) & Trim(ld_txt)
 
If FG = "PDIP32" Or FG = "PDIP40" Then

   bottom(3) = ""
   bottom(4) = ""
   bottom(5) = ""

End If
End Sub

Private Sub MarkSpec()
Dim rs As ADODB.Recordset
Dim sqlstr As String
Dim topMark, botMark As String
Dim cnt As Integer
Dim top, bottom As String

INIT_MARK
Dim rsbd As ADODB.Recordset
Set rsbd = New ADODB.Recordset
CSQLSTRING = "SELECT * FROM AIC_MARKING_SPEC_X WHERE SPECNO = '" & Trim(mark_spec_txt) & "' AND APPROVAL = 'APPROVED'"
Debug.Print CSQLSTRING
rsbd.Open CSQLSTRING, wsDB
If Not rsbd.EOF Then
    XWORKW = Trim(ww_txt)
    XWORKW_YWW = Right(Trim(ww_txt), 3)
    XWORKW_WW = Right(Trim(ww_txt), 2)
    XWAFERC = Trim(wafer_lot_txt)
    FRNTFER = Trim(Left(wafer_lot_txt, 2))

    YOSHTRA1 = "N"
    YOSHTRA2 = "N"
    YOSHTRA3 = "N"
    YOSHTRA4 = "N"
    YOSHTRA5 = "N"
    YOSHTRA6 = "N"

    YOSHTRI1 = "N"
    YOSHTRI2 = "N"
    YOSHTRI3 = "N"
    YOSHTRI4 = "N"
    YOSHTRI5 = "N"
    YOSHTRI6 = "N"
    
    
    Set rs = New ADODB.Recordset
     
   sqlstr = " SELECT TOPMARK_CUST, BOTMARK_CUST FROM AIC_LI_CUST_MASTER " _
          & " WHERE CUSTNO = '" & Left(Right(Trim(cboCust.TEXT), 4), 3) & "' "
       '   & " AND CUST_GROUP = 'TESTING2' "
          
   rs.Open sqlstr, wsDB
                        
    If Not rs.EOF Then
      topMark = rs("topmark_cust")
      botMark = rs("BOTMARK_CUST")
    Else
      MsgBox " No record found for topmark_cust!"
    End If
    
    'TOP MARK
    cnt = 0
    Do Until cnt > 5
      top = "top" & cnt + 1
      bottom = "BOTTOM" & cnt + 1
      If (cnt + 1) = topMark Then
        Me.topx(cnt) = Left(Trim(txtCusLot), 6)
        Me.topInfo(cnt) = Trim(rsbd(top))
      Else
        Me.topx(cnt) = Trim(rsbd(top))
        Me.topInfo(cnt) = Trim(rsbd(top))
      End If
      cnt = cnt + 1
    Loop
    
    

    
    'Quah 2010-05-18 for GMT, topx1 retrieve back from topx(1).tag (stored during pulldata)
    If Left(Trim(refno_TXT), 2) = "GM" Then
        topx(1) = Trim(topx(1).Tag)
    End If
    'Quah 2014-06-17 for FITIPOWER, topx retrieve back from topx().tag (stored during pulldata)
    If InStr(cboCust, "FITIPOWER") > 0 Then
        topx(0) = Trim(topx(1).Tag)
        topx(1) = Trim(topx(2).Tag)
        topx(2) = Trim(topx(3).Tag)
        topx(3) = Trim(topx(4).Tag)
    End If
    
    'Quah 2012-03-30 for Hexa take the Marking from the imported data.
    If Left(Trim(refno_TXT), 2) = "HT" Then
        Dim htmark As ADODB.Recordset
        Set htmark = New ADODB.Recordset
        ssql = "select * from aic_loading_instruction where refno='" & Trim(refno_TXT) & "'"
        htmark.Open ssql, wsDB
        If Not htmark.EOF Then
            topx(0) = "{}"
            topx(1) = Trim(htmark!TOP2)
            topx(2) = Trim(htmark!TOP3)
        End If
        htmark.Close
        Set htmark = Nothing
    End If
    
    
    'BOTTOM MARK
    cnt = 0
    Do Until cnt > 5
      top = "top" & cnt + 1
      bottom = "BOTTOM" & cnt + 1
      If (cnt + 1) = botMark Then
        Me.bottom(cnt) = Left(Trim(txtCusLot), 6)
        Me.botInfo(cnt) = Trim(rsbd(bottom))
      Else
        Me.bottom(cnt) = Trim(rsbd(bottom))
        Me.botInfo(cnt) = Trim(rsbd(bottom))
      End If
      cnt = cnt + 1
    Loop
    
Else
    If InStr(bdcombo, "RMA") > 0 Or InStr(bdcombo, "RWK") > 0 Or InStr(bdcombo, "TTR") > 0 Then     'Quah 20100105
        'can edit marking.
        mark_spec_txt = "N/A"
        mark_spec_txt.Locked = False
    Else
        MsgBox "Marking Spec is not setup in Database. Pls check.!!!", vbCritical, "Message"
        mark_spec_txt = "N/A" 'QUAH 20120912 DEFAULT
        
'        If Left(Trim(refno_TXT), 2) = "ST" Then
'            End
'        End If
        
        '''''End '20170728 req by KC & Mon, to block if no markfile.
    End If
   ' ld_txt.SetFocus
End If
rsbd.Close

FG = Trim(package_txt) & Trim(ld_txt)
If FG = "PDIP32" Or FG = "PDIP40" Then
    Me.bottom(3) = ""
    Me.bottom(4) = ""
    Me.bottom(5) = ""
End If
End Sub

Private Sub MarkSpec_NK()
cnt = 0
Do Until cnt > 5
    'CHECK CUSTLOTNO/PRODUCT LOT
    If InStr(Trim(topInfo(cnt)), "PRO LOT(ASSY#L2#)") > 0 Then
        Me.topx(cnt) = Replace(Me.topInfo(cnt), "PRO LOT(ASSY#L2#)", Trim(txtCusLot))
    End If
    If InStr(Trim(botInfo(cnt)), "PRO LOT(ASSY#L2#)") > 0 Then
        Me.bottom(cnt) = Replace(Me.botInfo(cnt), "PRO LOT(ASSY#L2#)", Trim(txtCusLot))
    End If
    cnt = cnt + 1
Loop

End Sub

Private Sub MarkSpec_AD()
cnt = 0
Do Until cnt > 5
    'CHECK WW
    If InStr(Trim(topInfo(cnt)), "YWW") Then
        Me.topx(cnt) = Replace(Me.topx(cnt), "YWW", Right(Trim(ww_txt), 3))
    End If
    If InStr(Trim(botInfo(cnt)), "YWW") Then
        Me.bottom(cnt) = Replace(Me.bottom(cnt), "YWW", Right(Trim(ww_txt), 3))
    End If
    
    'CHECK CUSLOTNO
    If InStr(Trim(topInfo(cnt)), "CUSTLOTNO") Then
        Me.topx(cnt) = Left(Trim(txtCusLot), 1) & Trim(Mid(Trim(txtCusLot), 3, 20))
    End If
    If InStr(Trim(botInfo(cnt)), "CUSTLOTNO") Then
        Me.bottom(cnt) = (Mid(Trim(txtCusLot), 2, 20))
    End If
    cnt = cnt + 1
Loop
End Sub

Private Sub MarkSpec_CT()
cnt = 0
Do Until cnt > 5
    'CHECK CUSTLOTNO/DATE CODE
    If InStr(Trim(topInfo(cnt)), "DATE CODE") Then
        Me.topx(cnt) = Replace(Me.topx(cnt), "DATE CODE", Trim(Mid(Trim(txtCusLot), 2, 25)))
    End If
    If InStr(Trim(botInfo(cnt)), "DATE CODE") Then
        Me.bottom(cnt) = Replace(Me.bottom(cnt), "DATE CODE", Trim(Mid(Trim(txtCusLot), 2, 25)))
    End If
    cnt = cnt + 1
Loop
End Sub
Private Sub MarkSpec_GT()
cnt = 0
Do Until cnt > 5
    'Check Customer REFNO and Replace Work Week
    If InStr(Trim(topInfo(cnt)), "YYWW") Then
        Me.topx(cnt) = Replace(Me.topx(cnt), "YYWW", Right(Trim(ww_txt), 4))
    End If
    cnt = cnt + 1
Loop
End Sub

Private Sub MarkSpec_II()
cnt = 0

'CHECK VALID MARK SPEC
If Trim(topx(0)) <> "{IR LOGO}" Then
    MsgBox "ERROR - TOPM1 must be <<{IRC LOGO}>>. Please Check your MarkSpec.", vbCritical
    INIT_MARK
    Exit Sub
End If

If Right(Trim(TARGET_DEVICE_TXT), 3) = "PBF" And topx(1) = "PYWWM" Then     'TARGET DEVICE LAST 3 DIGITS
ElseIf Right(Trim(TARGET_DEVICE_TXT), 3) <> "PBF" And topx(1) = "YWWM" Then
Else
    MsgBox "ERROR - TOPM1 must be <<TargetDevicePBF: PYWWM else YWWM>>. Please Check ur MarkSpec.", vbCritical
    INIT_MARK
    Exit Sub
End If

If (topx(2) <> "XXXX (last 4 char N/C)") And (topx(2) <> "XXXX_ (last 4 char N/C)") Then
    MsgBox "ERROR - TOPM3 must be <<XXXX (last 4 char N/C)>>. Please Check your MarkSpec.", vbCritical
    INIT_MARK
    Exit Sub
End If

Do Until cnt > 5
    'CHECK XXXX = NANACODE/ CUSLOTNO
    '2011-10-04 IR nanacode change to ProdOrder eg 10025UM.1
    pointpos = InStr(Trim(txtCusLot), ".")
    If pointpos > 4 Then
        irprodorder = Mid(Trim(txtCusLot), pointpos - 4, 4)
    Else
        irprodorder = "????"
        MsgBox "Pls check marking data : ????", vbCritical, "Message"
    End If
    If InStr(Trim(topInfo(cnt)), "XXXX (last 4 char N/C)") Then
'        Me.topx(cnt) = Replace(Me.topx(cnt), "XXXX (last 4 char N/C)", Right(Trim(txtCusLot), 4))
        '2011-10-04
        Me.topx(cnt) = Replace(Me.topx(cnt), "XXXX (last 4 char N/C)", irprodorder)
    End If
    If InStr(Trim(topInfo(cnt)), "XXXX_ (last 4 char N/C)") Then
        Me.topx(cnt) = Replace(Me.topx(cnt), "XXXX_ (last 4 char N/C)", irprodorder) & "_"
    End If
    If InStr(Trim(botInfo(cnt)), "XXXX (last 4 char N/C)") Then
        Me.bottom(cnt) = Replace(Me.bottom(cnt), "XXXX (last 4 char N/C)", irprodorder)
    End If
    cnt = cnt + 1
Loop

Me.topx(0).Enabled = False
Me.topx(1).Enabled = False
Me.topx(2).Enabled = False
Me.topx(3).Enabled = False

End Sub

Private Sub mm1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        ww_txt = "" '2014-11-18
        
        XX02 = mm1
        yy1.SetFocus
    End If
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub NEW_OPT_Click()
    If NEW_OPT.Value = True Then SAVEREC.Caption = "SAVE NEW"
End Sub

Private Sub OpenAppComm_Click()
Call Shell("C:\WINDOWS\system32\mspaint.exe", vbNormalNoFocus)

End Sub

Private Sub optBD_Click()
    bdcombo.Clear
    cbobonding_diagram.Visible = False
    cbobonding_diagram.Clear
    target_device_txt1.Visible = False
    target_device_txt1 = vbNullString
    bonding_diagram_txt.Visible = True
    bonding_diagram_txt.SetFocus
End Sub

Private Sub optTargetDevice_Click()
bdcombo.Clear
target_device_txt1.Visible = True
bonding_diagram_txt.Visible = False
'bonding_diagram_txt.Text = vbNullString
cbobonding_diagram.Visible = False
cbobonding_diagram.Clear
target_device_txt1.SetFocus
If Left(refno_TXT, 2) = "GM" Then
    target_device_txt1 = Trim(TARGET_DEVICE_TXT)
End If
End Sub

Private Sub package_txt_Click()
    Dim rsbd As ADODB.Recordset
    Set rsbd = New ADODB.Recordset
 '   CSQLSTRING = "SELECT * FROM WIPPRD WHERE  WPRD_PRD_GRP_2 = '" & Trim(package_txt) & "'"
'AIMS
    CSQLSTRING = "SELECT * FROM baic_prodmast WHERE  PDM_PACKAGE = '" & Trim(package_txt) & "'"

    rsbd.Open CSQLSTRING, wsDB
    If rsbd.EOF Then
        MsgBox " Package  is not setup in Database. Check with Planner!!!"
    Else
        ld_txt.SetFocus
    End If
    rsbd.Close
End Sub

Private Sub package_txt_KeyPress(KeyAscii As Integer)
    Dim rsbd As ADODB.Recordset
    Set rsbd = New ADODB.Recordset
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
'        CSQLSTRING = "SELECT * FROM WIPPRD WHERE  WPRD_PRD_GRP_2 = '" & Trim(package_txt) & "'"
        
'AIMS
    CSQLSTRING = "SELECT * FROM baic_prodmast WHERE  PDM_PACKAGE = '" & Trim(package_txt) & "'"
        
        rsbd.Open CSQLSTRING, wsDB
        If rsbd.EOF Then
            MsgBox " Package  is not setup in Database. Check with Planner!!!"
        Else
            ld_txt.SetFocus
        End If
        rsbd.Close
    End If
End Sub

Private Sub quantity_txt_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        Me.lblPONo.SetFocus
        qty = Trim(quantity_txt)
        'If mark_spec_txt.Enabled = True Then
        '     mark_spec_txt.SetFocus
        'Else
        'dd1.SetFocus
        'End If
   End If
End Sub

Private Sub init_header()
    Dim cnt  As Integer
    
    internal_device_no_txt = "":    package_lead_txt = ""
    TARGET_DEVICE_TXT = "":    WAFER = ""
    qty = "":    quantity_txt = ""
    total_poqty = "": txt_podate = "": txt_pomode = ""
    mark_spec_txt = "":    stat = ""
    NEW_OPT.Value = False:    RELEASE_OPT.Value = False:    CANCEL_OPT.Value = False
    bonding_diagram_txt = "":    wafer_lot_txt = ""
    bdcombo.Clear:    package_txt = ""
    ld_txt = "":    ww_txt = ""
    bd_no_txt = "":    bd_no_txt1 = ""
    txtCusLot = "":    lblPONo = ""
    NOWAFERLOTx = 0
 
End Sub
Private Sub refno_TXT_GotFocus()
    lblU.Visible = False
    lbltestonly.Visible = False
    lblFroute = "": lblLroute = ""
    lblOpr1 = "": lblOpr2 = "":  lblTUT = ""
    
End Sub
' add at 11/6/2021 kokkeong

Private Sub show1()
refno_TXT.TEXT = UCase(refno_TXT.TEXT)
Label52.Caption = refno_TXT.TEXT & ".bmp"
Image1.Picture = LoadPicture("\\aicwksvr2016\MarkingPattern\" & Label52.Caption)
End Sub

' add at 11/6/2021 kokkeong
Sub checkFile()

If Dir$("\\aicwksvr2016\MarkingPattern\" & Label52.Caption) <> "" Then
    show1
    Check2.Value = 1
Else
    Image1.Picture = LoadPicture("\\aicwksvr2016\MarkingPattern\NoMarking.bmp")
    Check2.Value = 0
End If
      
End Sub

Private Sub refno_TXT_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then

'    bom_verify = ""

' 11/6/2021 kokkeong
' add label 52 and checkFile

    Label52.Caption = (refno_TXT.TEXT & ".bmp")
    checkFile
      
      lirefpre = Left(Trim(refno_TXT), 2)
      If lirefpre <> Me.custprefix Then
           MsgBox "Customer and Refno not match!", vbCritical, "Message"
           Call RESET
           refno_TXT.TEXT = vbNullString
           Exit Sub
      End If
      
    If Left(refno_TXT, 2) = "AP" Then
        AP = Right(refno_TXT, 7)
        If AP <= 1405012 Then
            MsgBox "FOR REFNO BEFORE AP1405013 PLEASE USE DUAL DIE SCREEN!", vbCritical, "Message"
            Call RESET
            refno_TXT.TEXT = vbNullString
            Exit Sub
        End If
    End If
      
      
      
'      If lirefpre = "AP" Then
'        MsgBox "Not ready!", vbCritical, "Message"
'        Exit Sub
'      End If
      
'      If lirefpre = "MT" Or lirefpre = "SG" Or lirefpre = "SL" Or lirefpre = "AP" Or lirefpre = "AN" Or lirefpre = "AD" Or lirefpre = "AS" Or lirefpre = "AC" Or lirefpre = "NK" Then


'20160811 Block for FITIPOWER-BJ, req  by LY, due to no forecast and label spec W1-C025 not ready.
If lirefpre = "FJ" Then
      MsgBox "Customer Spec W1-C025 not ready. Refer Test Engr. ", vbCritical, "Message"
      Call RESET
      refno_TXT.TEXT = vbNullString
      Exit Sub
End If
  
'20160817 Quah temp block for UBIQ, due to not yet implement for TT to show lots < 84 days for combine.
If prefix = "UB" Then
'      MsgBox "Pls refer IT, for 84 days combine lot control.", vbCritical, "Message"
'      Call RESET
'      refno_TXT.TEXT = vbNullString
'      Exit Sub
End If



'Quah 2012-11-02 unblock NIKO
'      If lirefpre = "MT" Or lirefpre = "SG" Or lirefpre = "SL" Or lirefpre = "AN" Or lirefpre = "AD" Or lirefpre = "AS" Or lirefpre = "AC" Or lirefpre = "NK" Then
      If lirefpre = "MT" Or lirefpre = "SG" Or lirefpre = "SL" Or lirefpre = "AN" Or lirefpre = "AD" Or lirefpre = "AS" Or lirefpre = "AC" Then
        
          MsgBox "This screen is not for this customer. Please select the correct screen from the top menu."
          Call RESET
          refno_TXT.TEXT = vbNullString
          Exit Sub
      End If
    
    'Quah 2014-11-24
    txtMarkingType.TEXT = ""
    If Left(refno_TXT, 2) = "BE" Then
        txtMarkingType.TEXT = "INPUT"
    End If
    
    
    Dim AIC_LIDB As ADODB.Recordset
    Dim SQL As String
    Set AIC_LIDB = New ADODB.Recordset
    SQL = "SELECT * FROM AIC_LOADING_INSTRUCTION WHERE REFNO = '" & refno_TXT & "'"
    AIC_LIDB.Open SQL, wsDB
    If AIC_LIDB.EOF = False Then
        bonding_diagram_txt.Visible = True
        PULLDATA
        Call PassParameter
        NEWFLAG = False
        If lirefpre = "GM" Or lirefpre = "FS" Or lirefpre = "FP" Or lirefpre = "TF" Or lirefpre = "AA" Then       '2010-10-29 - gmt, 20110713 FS,FP, 20120201-AAT
            '20110913 - add TF
            '20120201 - add AAT
            If internal_device_no_txt = "" Then
                optTargetDevice = True
                target_device_txt1 = TARGET_DEVICE_TXT
                TARGET_DEVICE_TXT.Locked = True
                If lirefpre = "AA" Then
                    total_poqty = Me.qty
                End If
            End If
        End If
        
        '20110913
        If lirefpre = "TF" Then 'Telefunken blanket po
            lblPONo = "N/A"
        End If
        
        
        Call cboCust_LostFocus
        
    Else
        MsgBox "NO RECORD FOUND!", vbCritical, "Message"
        Call RESET
        NEWFLAG = True
    End If
End If
End Sub
Private Sub xdatar()

    Dim XDATAV As ADODB.Recordset
    Set XDATAV = New ADODB.Recordset
    
    sqltext = "SELECT REM9, WSSTARTDATE FROM AIC_LOADING_INSTRUCTION_REMARK WHERE REFNO='" & Trim(refno_TXT) & "'"
              
              XDATAV.Open sqltext, wsDB, adOpenDynamic, adLockOptimistic
              
              If XDATAV.EOF = False Then
                    If InStr(1, XDATAV!REM9, "UPDATE ON") > 0 Then
                        lblU.Visible = True
                    Else
                        lblU.Visible = False
                    End If
                    
                    'QUAH 20080729
                    If Not IsNull(XDATAV!wsstartdate) Then
                    'Quah 20090327 improve on formatting.
                        
                        xws = Format(XDATAV!wsstartdate, "MM/DD/YYYY")  'Quah 20140929 date formatting
                        pos1 = InStr(1, xws, "/")
                        pos2 = InStr(pos1 + 1, xws, "/")
                        mm1 = Format(Mid(xws, 1, pos1 - 1), "00")
                        dd1 = Format(Mid(xws, pos1 + 1, pos2 - pos1 - 1), "00")
                        yy1 = Right(xws, 4)
'                        mm1 = Mid(XDATAV!wsstartdate, 1, 2)
 '                       dd1 = Mid(XDATAV!wsstartdate, 4, 2)
  '                      yy1 = Mid(XDATAV!wsstartdate, 7, 4)
                    End If
              End If
              XDATAV.Close
              Set XDATAV = Nothing

End Sub

Private Sub RELEASE_OPT_Click()
    If RELEASE_OPT.Value = True Then
    
        '2011-11-16
'        If Left(refno_TXT, 2) = "IJ" Or ((Left(refno_TXT, 2) = "MX" Or Left(refno_TXT, 2) = "MQ") And Trim(Me.TARGET_DEVICE_TXT) = "GW2150NT") Then
'            MsgBox "Temporary block for tracking purpose. Pls refer IT.", vbCritical, "Message"
'            RELEASE_OPT.Value = False
'            Exit Sub
'        End If
    
    'CATHERINE 070601
        SAVEREC.Caption = "RELEASE LI"
        cmdUpdate.Enabled = False
        SAVEREC.Visible = True
        SAVEREC.Enabled = True
       ' cmdLiRelease.Visible = True
       ' cmdLiRelease.Enabled = True
    End If
End Sub

Private Sub SaveDie()
Set adoMD = New ADODB.Recordset
wsSqlString = " SELECT * FROM AIC_LI_DUAL_DIE" _
& " WHERE REFNO = '" & Trim(refno_TXT) & "'  ORDER BY SEQNO"
adoMD.Open wsSqlString, wsDB, adOpenDynamic, adLockOptimistic
If adoMD.EOF = True Then
    iCnt = 1
    Do While iCnt <= lvwDie.ListItems.Count
        Set itmx = lvwDie.ListItems(iCnt)
        'If adoMD.EOF = True Then
           adoMD.AddNew
           adoMD!Refno = Trim(refno_TXT)
           adoMD!PartNo = Trim(itmx.SubItems(1))
           adoMD!waferno = Trim(itmx.SubItems(2))
           adoMD!WAFER_QTY = Trim(itmx.SubItems(5))
           adoMD!DIE_QTY = Trim(itmx.SubItems(4))
           adoMD!LIDIE_QTY = Trim(itmx.SubItems(3))
           adoMD!seqno = Trim(itmx.TEXT)
           adoMD.Update
       'End If
       iCnt = iCnt + 1
    Loop
Else
   MsgBox "DIE INFO ALREADY EXIST! PLEASE USE UPDATE!"
End If
adoMD.Close
End Sub

Private Sub insertLIBOM(refno_TXT)
       '20210602 INSERT INTO BAIC_LI_BOM 'KOKO 20210627
       ssql = " delete from BAIC_LI_BOM where Bom_LIREF = '" & Trim(refno_TXT) & "'"
       wsDB.Execute ssql
       
'''       ssql = " insert into BAIC_LI_BOM select LI.REFNO, LI.DEVICE_NO,BOM.SERIAL_NO,BOM.OPERATION_CODE,BOM.PART_NO,BOM.PART_SHORT_DESC,BOM.USAGE, " & _
'''              "  LI.QTY, (LI.QTY*BOM.USAGE) PARTQTY , BOM.UOM,BOM.YIELD_FACTOR,BOM.APT_PRINT,BOM.BD_NUMBER,LI.DATE_TRANX " & _
'''              "   from AIC_LOADING_INSTRUCTION LI, AIC_BOM_COMPONENT BOM " & _
'''              "   where REFNO='" & Trim(refno_TXT) & "' and LI.DEVICE_NO=BOM.DEVICE_NO " & _
'''              "  and BOM.REMARK='A' " & _
'''              "  order by BOM.SERIAL_NO "
       
'20210630 add in 4 extra cols
'BOM_STARTWIP  20,7
'BOM_COMPLETEWIP 20.7
'BOM_BALANCEWIP 20, 7
'BOM_LISTATUS nvarchar(10)

'Quah 2021-09-14 insert only selected parts by Planner (link to BAIC_LOAD_PART)
'-------------------------------------------------------
       ssql = " insert into BAIC_LI_BOM select LI.REFNO, LI.DEVICE_NO,BOM.SERIAL_NO,BOM.OPERATION_CODE,BOM.PART_NO,BOM.PART_SHORT_DESC,BOM.USAGE, " & _
              "  LI.QTY, (LI.QTY*BOM.USAGE) PARTQTY , 0,0,0,'', BOM.UOM,BOM.YIELD_FACTOR,BOM.APT_PRINT,BOM.BD_NUMBER,LI.DATE_TRANX " & _
              "   from AIC_LOADING_INSTRUCTION LI, AIC_BOM_COMPONENT BOM, BAIC_LOAD_PART " & _
              "   where REFNO='" & Trim(refno_TXT) & "' and LI.DEVICE_NO=BOM.DEVICE_NO " & _
              "  and LOAD_LIREF=LI.REFNO and LOAD_PARTNO=BOM.PART_NO " & _
              "  and BOM.REMARK='A' " & _
              "  order by BOM.SERIAL_NO "

       wsDB.Execute ssql
'-------------------------------------------------------

'''''       ssql = " insert into BAIC_LI_BOM select LI.REFNO, LI.DEVICE_NO,BOM.SERIAL_NO,BOM.OPERATION_CODE,BOM.PART_NO,BOM.PART_SHORT_DESC,BOM.USAGE, " & _
'''''              "  LI.QTY, (LI.QTY*BOM.USAGE) PARTQTY , 0,0,0,'', BOM.UOM,BOM.YIELD_FACTOR,BOM.APT_PRINT,BOM.BD_NUMBER,LI.DATE_TRANX " & _
'''''              "   from AIC_LOADING_INSTRUCTION LI, AIC_BOM_COMPONENT BOM " & _
'''''              "   where REFNO='" & Trim(refno_TXT) & "' and LI.DEVICE_NO=BOM.DEVICE_NO " & _
'''''              "  and BOM.REMARK='A' " & _
'''''              "  order by BOM.SERIAL_NO "
End Sub

Private Sub RL_ASSY_Click()
Call RL_Assy_inv.Show
End Sub

Private Sub saverec_Click()
chkUpdate

'20230221 Chong add ALLSENSORS cus lot check
'If Trim(custnameselect.TEXT) = "ALL SENSORS" Then
If Left(Trim(refno_TXT), 2) = "LS" Then
    If Trim(topx(2).TEXT) = "CUST LOT NO" Then
        MsgBox "Invalid CUST LOT NO, Please reenter.", vbCritical, "Message"
        Exit Sub '20230328 AIN CHANGE POSITION DUE TO USER CANNOT SAVE
    End If
    
End If
'End If

        
'20211008
'If Trim(bom_verify) = "" Then
'    MsgBox "Please check and confirm BOM parts.", vbCritical, "Message"
'    Exit Sub
'End If

If Trim(lblPONo) = "QUAH" Then      'should be QUAL ?
    MsgBox "Invalid PO No.", vbCritical, "Message"
    Exit Sub
End If

'20201001
If (Trim(custnameselect.TEXT) = "HCLY") Then
    MsgBox "HCLY label not ready in LabelMatrix.", vbCritical, "Message"
    Exit Sub
End If

'20201228 EK request, if device not running > 3 years, get ENGR confirmation first, for correct TEST PROGRAM.
Dim chklastload As ADODB.Recordset
Set chklastload = New ADODB.Recordset
ssql = "select startdate, DateDiff(Day, startdate, GETDATE()) diff from AIC_WIP_HEADER, BAIC_LOTMAST WHERE ASSYLOTNO=LTM_LOTNO AND LTM_STATUS <> 'TRLT' AND LTM_TARGETDEVICE ='" & Trim(TARGET_DEVICE_TXT.TEXT) & "' ORDER BY STARTDATE DESC"
Debug.Print ssql
chklastload.Open ssql, wsDB
If Not chklastload.EOF Then
    If chklastload!diff > 1095 Then
       MsgBox "Targetdevice last load > 3 yrs ago." & vbCrLf & "Pls confirm any changes with Engrs before loading.", vbCritical, "Message"
       If MsgBox("Targetdevice > 3 years." & vbCrLf & "Confirm to Proceed???", vbYesNo, "Message") = vbNo Then
           Exit Sub
       End If
    End If
End If
chklastload.Close


'KOKO ADD FOR CHECK WAFERLOTNO AND TARGET DEVICE WITH IQA SYSTEM AS PER HAIGETE 20191104 MIXED DEVICE
If (Trim(custnameselect.TEXT) = "HAIGETE") Then
    Dim HAIiqa As ADODB.Recordset
    Set HAIiqa = New ADODB.Recordset
    Dim IQAChk
    IQAChk = "select *  FROM AIC_INVENTORY_MASTER WHERE CUSTOMERNAME='HAIGETE' and device_no ='" & Trim(TARGET_DEVICE_TXT.TEXT) & "' and waferlotno='" & Trim(WAFER) & "' "
    Debug.Print IQAChk
    HAIiqa.Open IQAChk, wsDB
    If HAIiqa.EOF Then
       MsgBox " Targetdevice/Waferlotno not match with IQA System, Kindly verify again!!"
       Exit Sub
    End If
    HAIiqa.Close
End If
'20210302 Quah add for CHANGDIAN
If (Trim(custnameselect.TEXT) = "CHANGDIAN") Then
    Set HAIiqa = New ADODB.Recordset
    IQAChk = "select *  FROM AIC_INVENTORY_MASTER WHERE CUSTOMERNAME='CHANGDIAN' and device_no ='" & Trim(TARGET_DEVICE_TXT.TEXT) & "' and waferlotno='" & Trim(WAFER) & "' "
    Debug.Print IQAChk
    HAIiqa.Open IQAChk, wsDB
    If HAIiqa.EOF Then
       MsgBox " Targetdevice/Waferlotno not match with IQA System, Kindly verify again!!"
       Exit Sub
    End If
    HAIiqa.Close
End If
'KOko END FOR HAIGETE CHECK TARGET DEVICE



'Quah 20171127, for NIKO, check the length of Cuslotno & Marking.
If (Trim(custnameselect.TEXT) = "NIKO") Then
    'Quah 2018081 skip RWK, simplify logic.
    If InStr(package_lead_txt, "2X2") > 0 Or InStr(internal_device_no_txt, "RMA") > 0 Or InStr(internal_device_no_txt, "RWK") > 0 Then
        'OK SKIP CHECKING.
        '2X2 use YWW conversion formula.
    Else
        If Len(txtCusLot) <> 17 And Len(txtCusLot) <> 9 Then
            MsgBox "Invalid length. Please check Custlotno !!", vbCritical, "Message"
            Exit Sub
        End If
    End If
End If


'Quah add 20210204, check for same cuslotno for NIKO, GSTEK
If (Trim(custnameselect.TEXT) = "NIKO") Then
    Set chklastload = New ADODB.Recordset
    ssql = "select REFNO, CUSLOTNO from AIC_LOADING_INSTRUCTION where REFNO like 'NK%' and REFNO <> '" & refno_TXT.TEXT & "' and CUSLOTNO='" & txtCusLot.TEXT & "'"
    Debug.Print ssql
    chklastload.Open ssql, wsDB
    If Not chklastload.EOF Then
        MsgBox "Duplicate Custlotno !!", vbCritical, "Message"
        Exit Sub
    End If
    chklastload.Close
End If


'Quah add 20220523, same waferlot (cuslotno) cannot use in different LI.
If (Trim(custnameselect.TEXT) = "STMICRO") And CANCEL_OPT.Value = False Then
    Set chklastload = New ADODB.Recordset
    ssql = "select REFNO, CUSLOTNO from AIC_LOADING_INSTRUCTION where REFNO like 'ST%' and REFNO <> '" & refno_TXT.TEXT & "' and CUSLOTNO='" & txtCusLot.TEXT & "' and status <> 'C'"
    Debug.Print ssql
    chklastload.Open ssql, wsDB
    If Not chklastload.EOF Then
        MsgBox "Duplicate Custlotno !!", vbCritical, "Message"
        Exit Sub
    End If
    chklastload.Close

    If Len(txtCusLot.TEXT) >= 10 Then
        MsgBox "Cuslotno too long. Recommended max 9 chars !!", vbCritical, "Message"
'        Exit Sub '20220524
    End If
 If InStr(internal_device_no_txt, "TST") > 0 Then 'KOKO ADD TO CHECK THE TST DEVICE 20220807
  'CustAICLOT = Left(txtCusLot, 2)
     If RL_LOTNO_LBL = "" And Left(txtCusLot, 2) <> "ST" Then
        MsgBox "Please click Rawline button "
        Exit Sub
     Else
          'KOKO ADD 20220808 'KOKOKOKO
        ' XX = refno_TXT
      Set ChkLIQTY = New ADODB.Recordset
      SQLLI = "select SUBD_LI_REFNO,CUSLOTNO,QTY,SUBD_RL_QTY_ISSUE from AIC_LOADING_INSTRUCTION," & _
              " BAIC_RL_ASSY_DETAIL where REFNO = SUBD_LI_REFNO AND REFNO = '" & refno_TXT.TEXT & "'"
              ChkLIQTY.Open SQLLI, wsDB
              If Not ChkLIQTY.EOF Then
                  
                  If txtCusLot <> Trim(ChkLIQTY!CUSLOTNO) Then
                    If Val(qty) <> ChkLIQTY!qty Then
                       Set rs = New ADODB.Recordset
                      SQL$ = "update AIC_LOADING_INSTRUCTION set QTY=" & Val(qty) & ", CUSLOTNO='" & txtCusLot & "' where REFNO='" & refno_TXT.TEXT & "'"
                            Debug.Print SQL$
                            rs.Open SQL, wsDB
                    End If
                    
                    
                  End If
              
              End If
              
     
     End If
  End If 'koko add for ST RL LOAD
  
  'add Ain 20221206 Rawline' cuslotno is same with device no
  'add Asyraf 20221209 Rawline' cuslotno is same with device no (check for ST)
  If Left(txtCusLot, 2) = "ST" Then
  
  Set chklastload = New ADODB.Recordset
  SSQL1 = "SELECT *  FROM BAIC_LOTMAST LEFT JOIN ST_LOADING_MAST " & _
        "ON LTM_TARGETDEVICE = ST_Rawline WHERE LTM_LOTNO='" & txtCusLot & "' and ST_FG_Code = '" & TARGET_DEVICE_TXT & "' "
    Debug.Print SSQL1
    chklastload.Open SSQL1, wsDB
    If chklastload.EOF Then
    'If (Trim(TARGET_DEVICE_TXT.TEXT) <> chklastload!ST_FG_code) Then
        MsgBox "Custlot Does Not Tally with Device No!!", vbCritical, "Message"
        Exit Sub
    End If
    chklastload.Close
  
  End If
  
  Set chklastload = New ADODB.Recordset
  ssql = "select * from AIC_INVENTORY_MASTER where CUSTOMERNAME = 'STMICRO' AND WAFERLOTNO ='" & txtCusLot.TEXT & "' AND DIEQTY > 0"
  
  
'    Debug.Print ssql
'    chklastload.Open ssql, wsDB
'    If chklastload.EOF Then
'        MsgBox "Custlot Does Not Exist!!", vbCritical, "Message"
'        Exit Sub
'    End If
'    chklastload.Close
'  End If
  
'  '20220525 Checking for wafer exist in iqa
'  If Left(txtCusLot, 5) <> "DUMMY" Then   'skip for dummy lot
'    Set chklastload = New ADODB.Recordset
'    ssql = "select * from AIC_INVENTORY_MASTER where CUSTOMERNAME = 'STMICRO' AND WAFERLOTNO ='" & txtCusLot.TEXT & "' AND DIEQTY > 0"
'    Debug.Print ssql
'    chklastload.Open ssql, wsDB
'    If chklastload.EOF Then
'        MsgBox "Custlot Does Not Exist!!", vbCritical, "Message"
'        Exit Sub
'    End If
'    chklastload.Close
'  End If
End If


'If (Trim(custnameselect.TEXT) = "GSTEK") Then
'    Set chklastload = New ADODB.Recordset
'    ssql = "select REFNO, CUSLOTNO from AIC_LOADING_INSTRUCTION where REFNO like 'GS%' and REFNO <> '" & refno_TXT.TEXT & "' and CUSLOTNO='" & txtCusLot.TEXT & "'"
'    Debug.Print ssql
'    chklastload.Open ssql, wsDB
'    If Not chklastload.EOF Then
'        MsgBox "Duplicate Custlotno !!", vbCritical, "Message"
'        Exit Sub
'    End If
'    chklastload.Close
'End If




'''''If (Trim(custnameselect.TEXT) = "NIKO") Then
'''''    If Len(txtCusLot) <> 17 Then
'''''        MsgBox "Invalid length. Please check Custlotno !!", vbCritical, "Message"
'''''        Exit Sub
'''''    End If
'''''
'''''    'for marking....
'''''    If InStr(package_lead_txt, "2X2") > 0 Then
'''''        'skip marking check.
'''''    ElseIf InStr(package_lead_txt, "SOIC") > 0 Then
'''''        If Len(topx(3)) <> 17 Then
'''''            MsgBox "Invalid length. Please check Marking !!", vbCritical, "Message"
'''''            Exit Sub
'''''        End If
'''''    Else
'''''        If Len(topx(2)) <> 14 Then
'''''            MsgBox "Invalid length. Please check Marking !!", vbCritical, "Message"
'''''            Exit Sub
'''''        End If
'''''    End If
'''''End If



'Quah 20170426 temporary block new FETEK (to link to Axelite label format).
If (Trim(custnameselect.TEXT) = "FETEK") And SAVEREC.Caption = "&SAVE" Then
    MsgBox "Please inform IT, for AXELITE label format !!", vbCritical, "Message"
End If
'Quah 20170620 temporary block PANJIT (to prepare datamatrix label & yield report).
If Trim(custnameselect.TEXT) = "PANJIT" And SAVEREC.Caption = "&SAVE" Then
    MsgBox "Please inform IT, Datamatrix label ready? ", vbCritical, "Message"
    Exit Sub
End If

'Quah 20170620 check for AMPHENOL unregistered APT for Port Type and Dispensing Type (for MASS PROD lots).
If Trim(custnameselect.TEXT) = "AMPHENOL" And InStr(cboruntype, "Mass Production") > 0 Then
    'Quah 20170718 skip checking for NPX & HSE, requested by Anita.
    If InStr(internal_device_no_txt, "HSE") > 0 Or InStr(internal_device_no_txt, "NPX") > 0 Or InStr(internal_device_no_txt, "CAP") > 0 Then
        'skip checking.
    Else
        Dim RS2 As ADODB.Recordset
        
        '20181101 add checking in PDM_CATEGORY
        Set RS2 = New ADODB.Recordset
        ssql = "SELECT * FROM BAIC_PRODMAST WHERE pdm_targetdevice='" & internal_device_no_txt & "' and PDM_CATEGORY='NPA'"
        Debug.Print ssql
        RS2.Open ssql, wsDB
        If Not RS2.EOF Then
            NPALOT = True
        Else
            NPALOT = False
        End If
        RS2.Close
        
        If NPALOT = True Then
            Set RS2 = New ADODB.Recordset
            ssql = "SELECT * FROM BAIC_COMTBL WHERE TBL_REC_TYPE='GEPT'" & _
                   " AND TBL_KEY_A20= '" & Trim(TARGET_DEVICE_TXT) & "'"
            Debug.Print ssql
            RS2.Open ssql, wsDB
            If RS2.EOF = True Then
                MsgBox "AMPHENOL Device Not Registered For APT." & vbCrLf & "Pls inform CUSTOMER SERVICE.", vbCritical, "Message"
                Exit Sub
            End If
            RS2.Close
        End If
    End If
End If


'''''''ko add for micrel validate the Customer bd keyin
'''''''KO add 20151103 Included Microchip
''''''If (Trim(custnameselect.TEXT) = "MICREL" Or Trim(custnameselect.TEXT) = "MICROCHIP") And saverec.Caption = "&SAVE" Then
''''''    BD_FOUND = "N"
''''''    rekey_bd.Show vbModal
''''''    If BD_FOUND = "N" Then
''''''        MsgBox "Customer BD Unmatched !!", vbCritical, "Message"
''''''        Exit Sub
''''''    End If
''''''End If

'Quah 2017-01-12
'Mary request block NS if assy price is zero.
Dim LIBlock As ADODB.Recordset
If custnameselect.TEXT = "NS" Then
    Set LIBlock = New ADODB.Recordset
    ssql = "select * from BAIC_INVOICE_PRICE_DETAIL where IPD_DEVICENO='" & Trim(internal_device_no_txt) & "' and IPD_EXP_YMD > GETDATE() and IPD_ASSY_PRICE > 0"
    Debug.Print ssql
    LIBlock.Open ssql, wsDB
    If LIBlock.EOF Then
        'CHECK WITH MARY
        '20190108 Anita request disabled.
        'MsgBox "Assy Price is zero. Pls inform Finance.", vbCritical, "Message"
        'Exit Sub
    End If
End If

'Quah 2012-07-24
Set LIBlock = New ADODB.Recordset
ssql = "select * from BAIC_CUSTOMER_ADDSETUP where CUS_RECTYPE='BLOCK_LI' and CUS_SHORTNAME='" & Trim(custnameselect.TEXT) & "'"
Debug.Print ssql
LIBlock.Open ssql, wsDB
If Not LIBlock.EOF Then
    MsgBox "Customer blocked for IT tracking. Pls refer IT", vbCritical, "Message"
    Exit Sub
End If
  
'Quah 20141103 marking check for FB, logic for other customers cannot be finalised yet.
'If Left(refno_TXT, 2) = "FB" Then
'    If InStr(TARGET_DEVICE_TXT, topx(1).TEXT) = 0 Then
'        MsgBox "Please check Marking Line2...", vbCritical, "Message"
'        Exit Sub
'    End If
'End If
'
  
'diana 20140912 LY want to block some new devices at loading
If Left(refno_TXT, 2) = "UB" Then
Dim new_dev As ADODB.Recordset
Set new_dev = New ADODB.Recordset
nSSQL = "select * from BAIC_CUSTOMER_ADDSETUP where CUS_RECTYPE='NEW_DEVICE' and CUS_SHORTNAME='UBIQ' AND CUS_KEY_1 = '" & Trim(Me.TARGET_DEVICE_TXT) & "'"
Debug.Print nSSQL
new_dev.Open nSSQL, wsDB
    If Not new_dev.EOF Then
        MsgBox "UBIQ NEW DEVICE. REFER ENGR BEFORE LOADING!", vbCritical, "Message"
        Exit Sub
    End If
End If

If Trim(ww_txt) = "" Then   '2012-04-20
    MsgBox "Cannot save. Blank workweek!", vbCritical, "Message"
    Exit Sub
End If

REFNOx = Trim(refno_TXT)
If REFNOx = "" Then
    MsgBox "Cannot save. Invalid data!", vbCritical, "Message"
    Exit Sub
End If

'2012-02-08
If Trim(txtCusLot) = "" Then
    MsgBox "Cannot save. Invalid data!", vbCritical, "Message"
    Exit Sub
End If


If Left(REFNOx, 2) = "FS" Or Left(REFNOx, 2) = "FP" Then
        MsgBox "For FAIRCHILD FOM, pls ensure " & vbCrLf & "IQA has RECEIVED the wafers in FOM System!", vbInformation, "Message"
        
        'for RMA lots, also need to received.
        If Left(Trim(txtCusLot), 1) = "R" Then
            MsgBox "For RMA lots, please remember to do Receipt first." & vbCrLf & _
                   "(Go to FOM, FG Receipt ==> RMA, Click the save button)", vbInformation, "Message"
        End If
        
        '2012-01-12 check for XX marking (Planner forget to click on FAIRCHILD FOM button)
        If Mid(topx(1), 2, 2) = "XX" Then
            MsgBox "XX : Error in Marking?", vbCritical, "Message"
            Exit Sub
        End If

End If


If Trim(txt_pomode) = "" Then
        MsgBox "Cannot save. PO Mode cannot be blank.", vbCritical, "Message"
        Exit Sub
End If

If RELEASE_OPT.Value = True Then    'IF RELEASED.


'******
'CZZ Added 2012-08-29 LI Material Shortage Prevention
If Left(refno_TXT, 2) = "ST" And InStr(internal_device_no_txt, "TST") > 0 Then  'ko add 20220815
 'ignore check materialbalance
 ' Exit Sub
Else
    If Not CheckMaterialBalance(Trim(refno_TXT.TEXT)) Then
        Exit Sub
    End If
End If
'******
    
    PRODUCTX = Trim(internal_device_no_txt)
    
    
    'quah 2013-04-25 no bom if route dont go thru process 1300/1400/2100
    Dim chkpro As ADODB.Recordset
    Set chkpro = New ADODB.Recordset
    ssql = "select * from BAIC_ROUTING, BAIC_PRODMAST where PDM_ASSY_ROUTE=RTG_ROUTE and PDM_DEVICENO='" & Trim(PRODUCTX) & "' and RTG_OPER in (1300,1400,2100)"
    chkpro.Open ssql, wsDB
    If Not chkpro.EOF Then
        bompro = "BOM REQUIRED"
    Else
        bompro = "BOM NOT REQUIRED"
    End If
    chkpro.Close
    Set chkpro = Nothing
    ''' add bompro to below checking.
    
    
    '-----check for BOM
    '2021-12-10 disable bom check, due to already have Verify BOM function.
    '2012-01-19 add to exclude RMA
    If (Right(Trim(PRODUCTX), 6) <> "TTRPBF" And Right(Trim(PRODUCTX), 3) <> "TTR") And (Left(Trim(PRODUCTX), 4) <> "TROJ" _
        And Right(Trim(PRODUCTX), 3) <> "RWK") And (Left(Trim(PRODUCTX), 4) <> "SODR") And Left(Trim(PRODUCTX), 5) <> "TSODR" And Right(Trim(PRODUCTX), 6) <> "RWKPBF" And _
        Right(Trim(PRODUCTX), 6) <> "PBFRWK" And Right(Trim(PRODUCTX), 6) <> "RMAPBF" And Right(Trim(PRODUCTX), 6) <> "RESPBF" And Right(Trim(PRODUCTX), 3) <> "TST" And _
        (Right(Trim(PRODUCTX), 6) <> "RTTPBF") And (Right(Trim(PRODUCTX), 6) <> "RTRPBF") And (Right(Trim(PRODUCTX), 3) <> "RMA") And _
        Right(Trim(PRODUCTX), 6) <> "TSTPBF" And Left(Trim(REFNOx), 2) <> "OS" And _
        bompro = "BOM REQUIRED" Then
        Set orRS = New Recordset
            orSQL = "select B.device_no, B.part_no, B.part_short_desc " & _
                 "from aic_bom_header A, aic_bom_component B " & _
                 "Where a.bill_sequence_id = b.bill_sequence_id " & _
                 "and A.active_flag = 'Y' " & _
                 "and a.device_no = '" & Trim(PRODUCTX) & "' " & _
                 "and b.apt_print = 'Y' " & _
                 "and b.remark = 'A' " & _
                 "and b.part_no not like 'DIE%' " & _
                 "and b.part_no not like '052%' " & _
                 "order by b.serial_no "
                 Debug.Print orSQL
        orRS.Open orSQL, wsDB
        If orRS.EOF = True Then
            If InStr(Trim(Me.TARGET_DEVICE_TXT), "NPX") > 0 And Left(Trim(refno_TXT), 2) = "GE" Then  'Quah temporary for NPX ENGR 2013-01-15
'                MsgBox REFNOx & " : GE NPX, temporary bypassed BOM-Check ! ", vbCritical, "Message"
            Else
'                MsgBox REFNOx & "-BOM/DIESOURCE not found. Refer to Material Planner!!", vbCritical, "Message"
'                Exit Sub
            End If
        
        End If
        orRS.Close
    
    ' ko add 08-Jul-2020 for NPA LOADING BY PASS THE DIE CHECKING AT BOM , THE REST CUSTOMER NEED TO CHECK DIE
        If InStr(Trim(Me.TARGET_DEVICE_TXT), "NPA") > 0 And Left(Trim(refno_TXT), 2) = "NV" Then  'Quah temporary for NPX ENGR 2013-01-15
    
         orSQL = "select B.device_no, B.part_no, B.part_short_desc " & _
              "from aic_bom_header A, aic_bom_component B " & _
              "Where a.bill_sequence_id = b.bill_sequence_id " & _
              "and A.active_flag = 'Y' " & _
              "and a.device_no = '" & Trim(PRODUCTX) & "' " & _
              "and b.remark = 'A' order by b.serial_no "
       Else 'KO ADD FOR DIE CHECKING NPA 2020-07-08
       orSQL = "select B.device_no, B.part_no, B.part_short_desc " & _
              "from aic_bom_header A, aic_bom_component B " & _
              "Where a.bill_sequence_id = b.bill_sequence_id " & _
              "and A.active_flag = 'Y' " & _
              "and a.device_no = '" & Trim(PRODUCTX) & "' " & _
              "and b.remark = 'A' " & _
              "and b.part_no like 'DIE%' " & _
              "order by b.serial_no "
       
       End If 'KO ADD FOR DIE CHECKING NPA 2020-07-08
              
              
        orRS.Open orSQL, wsDB
        Debug.Print orSQL
        If orRS.EOF = True Then
'            MsgBox REFNOx & "-BOM/DIESOURCE not found. Refer to Material Planner!!", vbCritical, "Message"
'            Exit Sub
        End If
        orRS.Close
    End If
End If
'----check for BOM




'Quah 2010-08-18
'TMTECH if marking line 2got dot, block.
If Left(Trim(refno_TXT), 2) = "TH" Then
    If InStr(1, topx(2), ".") Then
        MsgBox "TMTECH marking cannot have DOT!", vbCritical, "Message"
        Exit Sub
    End If
End If



'Quah 2010-08-04 check for QFN device must register in AIMS
If InStr(package_lead_txt, "FN") > 0 Then
    Dim aimsdev As ADODB.Recordset
    Set aimsdev = New ADODB.Recordset
    ssql = "select * from baic_prodmast where pdm_deviceno='" & Trim(internal_device_no_txt) & "'"
    Debug.Print ssql
    aimsdev.Open ssql, wsDB
    If aimsdev.EOF = True Then
        MsgBox "Cannot save. Please register this device in AIMS System.", vbCritical, "Message"
        Exit Sub
    End If
    aimsdev.Close
    Set aimsdev = Nothing
End If


Dim adoMD As Recordset
Dim waferstr As String
Dim waferq As String
Dim dieq As String
Dim cnt As Integer
        
If Me.mm1 <> "" Then
        SDate = CDate(Me.mm1 + "/" + Me.dd1 + "/" + Me.yy1)
        
        'Quah 20080605 check for StartDate > 5 days
        'Quah 20140723 open for 10 days due to Hari Raya advance loading, req by Anita
        If (SDate - Date > 5) Or (SDate - Date < 0) Then
            MsgBox ("SYSTEM DOES NOT ALLOW START-DATE MORE THAN 5 DAYS !! Please Check Your Input.")
            MsgBox ("System will now set the Start-Date to the default tomorrow's date.")
            'Quah 20080605 set default date and WW
            wwdate = Format(Date + 1, "DDMMYYYY")
            dd1 = Mid(wwdate, 1, 2)
            mm1 = Mid(wwdate, 3, 2)
            yy1 = Mid(wwdate, 5, 4)
            
            MsgBox "Please double check WORKWEEK & MARKING.", vbInformation, "Message"
            Exit Sub
        End If
End If
        
    'Check location of DOT
    SpecChars = ""
    iCnt = 0
    Do While iCnt <= 4
        If chkdot(iCnt).Value = Checked Then
            SpecChar(iCnt) = "."
        Else
            SpecChar(iCnt) = " "
        End If
        SpecChars = SpecChars + SpecChar(iCnt)
        iCnt = iCnt + 1
    Loop
       
    If cboCust.TEXT = "" Then
       MsgBox "PLEASE INSERT CUSTOMER NO !!", vbCritical
       cboCust.SetFocus
       Exit Sub
    End If
    If Trim(refno_TXT) = "" Then
        MsgBox " PLEASE INSERT REFERENCE NO!!!", vbCritical
        refno_TXT.SetFocus
        Exit Sub
    End If
'   DIANA 20150508 REMOVE BECAUSE ld_txt does not exists
'    If Left(Trim(refno_TXT), 2) <> "OS" Then    'Quah 20091116
'        If package_txt = "" Then
'            MsgBox " PLEASE SELECT PACKAGE!!!", vbCritical
'            Exit Sub
'        End If

'        If ld_txt = "" Then
'            MsgBox " PLEASE INSERT LEAD!!!", vbCritical
'            ld_txt.SetFocus
'            Exit Sub
'        End If
'    End If
    If bonding_diagram_txt = "" Then
        MsgBox " PLEASE INSERT AIC BD NO!!!", vbCritical
        bonding_diagram_txt.SetFocus
        Exit Sub
    End If
    If bdcombo = "" Then
        If CANCEL_OPT.Value = False And RELEASE_OPT.Value = False Then
            MsgBox " PLEASE INSERT CUSTOMER BD NO!!!", vbCritical
            bdcombo.SetFocus
            Exit Sub
        End If
    End If
    'Quah 20180222 add lbltestonly condition.
    If Left(Right(Trim(cboCust.TEXT), 6), 1) <> "M" And Left(Trim(cboruntype), 11) <> "Engineering" And lbltestonly <> "TEST ONLY" Then
        If mark_spec_txt = "" Then
            MsgBox " PLEASE INSERT MARK SPEC!!!", vbCritical, "Message"
            mark_spec_txt.SetFocus
            Exit Sub
        End If
    End If
    If ww_txt = "" Then
        MsgBox " PLEASE INSERT WORK WEEK!!!", vbCritical
        ww_txt.SetFocus
        Exit Sub
    End If
    
    
    'CHECK BD AND MARK SPEC
    'CHAR(1ST) = M : NO MARKINGSPEC CHECKING
    'If Left(Trim(refno_TXT), 2) = "OC" Or Left(Trim(refno_TXT), 2) = "SK" Or Left(Trim(refno_TXT), 2) = "AC" Then
    If Left(Right(Trim(cboCust.TEXT), 6), 1) <> "M" And Left(Trim(cboruntype), 11) <> "Engineering" Then
        If Left(refno_TXT, 2) = "GM" Then
             SQL$ = "SELECT * FROM AIC_DEVICE_CONTROL WHERE CUSTOMER='CUSBD' AND TARGETDEVICE='MARKSPEC'" & _
                 " AND REMARK1='" & Trim(Me.bonding_diagram_txt) & "' AND REMARK2='" & Trim(Me.mark_spec_txt) & "'"
         Else
             SQL$ = "SELECT * FROM AIC_DEVICE_CONTROL WHERE CUSTOMER='CUSBD' AND TARGETDEVICE='MARKSPEC'" & _
                 " AND REMARK2='" & Trim(Me.mark_spec_txt) & _
                 "' AND REMARK4 = '" & Trim(Me.internal_device_no_txt) & "'"
         End If
        
        Set rs = New ADODB.Recordset
        Debug.Print SQL$
        rs.Open SQL$, wsDB
        If rs.EOF = False Then
            If Trim(rs!REMARK3) = "CLOSE" Then
                MsgBox "PARTICULAR BONDING DIAGRAM WITH MARKING SPEC ALREADY CLOSED!PLEASE CHECK!", vbCritical, "ERROR"
                Exit Sub
            End If
        Else
            If mark_spec_txt = "" And lbltestonly <> "TEST ONLY" Then  'Quah 20120912 add IF condition.
                MsgBox "PARTICULAR BONDING DIAGRAM NOT INITIALIZE WITH MARKING SPEC!PLEASE CHECK!", vbCritical, "ERROR"
                Exit Sub
            End If
        End If
        rs.Close
    End If


'- Quah 20090402 Check if { matching with }, a few times planner incorrectly key in {T{
If InStr(topx(0), "{") > 0 Then
    If InStr(topx(0), "}") = 0 Then
            MsgBox "PLEASE CHECK MARKING. BRACKET NOT MATCH !", vbCritical, "ERROR"
            Exit Sub
    End If
End If
If InStr(topx(1), "{") > 0 Then
    If InStr(topx(1), "}") = 0 Then
            MsgBox "PLEASE CHECK MARKING. BRACKET NOT MATCH !", vbCritical, "ERROR"
            Exit Sub
    End If
End If
If InStr(topx(2), "{") > 0 Then
    If InStr(topx(2), "}") = 0 Then
            MsgBox "PLEASE CHECK MARKING ! BRACKET NOT MATCH !", vbCritical, "ERROR"
            Exit Sub
    End If
End If
If InStr(topx(3), "{") > 0 Then
    If InStr(topx(3), "}") = 0 Then
            MsgBox "PLEASE CHECK MARKING ! BRACKET NOT MATCH !", vbCritical, "ERROR"
            Exit Sub
    End If
End If


'- Quah 20081009 Check if MARKFILE registered by Engineer or not.
    Set rs = New ADODB.Recordset
    'If Trim(Me.package_lead_txt) Like "%FN%" Then
    
    '2010-05-14 Me.package_lead_txt change to Me.cbofullpackage
    If InStr(Me.cbofullpackage, "FN") > 0 Then
        SSQL1 = "SELECT * FROM AIC_MARKING_REF WHERE TARGETDEVICE='" & Trim(Me.TARGET_DEVICE_TXT) & "' AND PACKAGELEAD='" & Trim(Me.cbofullpackage) & "' AND CUSTNAME='" & Trim(custnameselect.TEXT) & "'"
    Else
        SSQL1 = "SELECT * FROM AIC_MARKING_REF WHERE DEVICENO='" & Trim(Me.internal_device_no_txt) & "' AND PACKAGELEAD='" & Trim(Me.cbofullpackage) & "' AND CUSTNAME='" & Trim(custnameselect.TEXT) & "'"
    End If
    
    Debug.Print SSQL1
    rs.Open SSQL1, wsDB
    markfileexist = "N"
    If rs.EOF = False Then
        If LenB(rs!MARKPLATE) <> 0 Then
            markfileexist = "Y"
        End If
    Else
        Set RS2 = New ADODB.Recordset
        '2010-05-14 Me.package_lead_txt change to Me.cbofullpackage
        SSQL1 = "SELECT * FROM AIC_MARKING_REF WHERE PACKAGELEAD='" & Trim(Me.cbofullpackage) & "' AND CUSTNAME='" & Trim(custnameselect.TEXT) & "' AND TARGETDEVICE='X'"
        Debug.Print SSQL1
        RS2.Open SSQL1, wsDB
            If RS2.EOF = False Then
                If LenB(RS2!MARKPLATE) <> 0 Then
                    markfileexist = "Y"
                End If
            End If
            RS2.Close
    End If
    rs.Close
    
    'Quah 20090916 compulsory markfile for Production lots. Agreed by Lingling
    'Quah 20090917 except for Smartcard (no marking)
    If SAVEREC.Caption <> "RELEASE LI" Then
        If Left(Trim(cboruntype), 11) = "Engineering" Or Left(Trim(Me.internal_device_no_txt), 4) = "SCRF" Or Left(Trim(Me.internal_device_no_txt), 3) = "SCM" Then
            If markfileexist = "N" Then
                s = MsgBox("MARKPLATE not registered for this Device. Proceed to save L.I.?", vbYesNo, "Message")
                If s <> 6 Then
                    Exit Sub
                End If
            End If
        Else    'Production Lots
            If markfileexist = "N" Then
                s = MsgBox("MARKPLATE not registered for this Device. Proceed to save L.I.?", vbYesNo, "Message")
                If s <> 6 Then
                    Exit Sub
                End If
            End If
        End If
    End If
'- Quah 20081009
    
    
    
    s = MsgBox("CONFIRM SAVE DATA?", vbYesNo, "Build Instruction")
    If s = 6 Then
    
        'Quah 20210705 insertLiBom at beginning of SAVE.
        'Quah move INSERTBOM to MFGInventory Confirm Button
        'Call insertLIBOM(refno_TXT)
    
        '2011-12-29 - 2
        If Left(refno_TXT, 2) = "IJ" Then
            
            Dim ijrs As ADODB.Recordset
            Set ijrs = New ADODB.Recordset
            ssql = "select a.REFNO, WAFERNO, REMARKS1, REMARKS2, TEXT  from AIC_LI_LABELINFO a, AIC_LI_DUAL_DIE b, AIC_LABEL_REFERENCE where a.REFNO = '" & Trim(refno_TXT) & "' and ASSYLOTNO='IMPINJ IQA' and b.WAFERNO=WAFERLOTNO and a.REFNO=b.REFNO"
            ijrs.Open ssql, wsDB
            If Not ijrs.EOF Then
                ijdef = Trim(ijrs!REMARKS2)
                ijwafer = Trim(ijrs!waferno)
            End If
            ijrs.Close
             If ijdef <> "" Then
                ijbal = InputBox("From this list, pls REMOVE wafers required for loading.", "IMPINJ BALANCE DIE", ijdef)
                If ijbal <> "" Then
                    If MsgBox("IMPINJ balance DIE-ID : " & ijbal & " ?", vbYesNo, "Message") = vbYes Then
                        Set ijrs = New ADODB.Recordset
                        ssql = "select * from AIC_LABEL_REFERENCE where ASSYLOTNO ='IMPINJ IQA' and WAFERLOTNO='" & ijwafer & "'"
                        Debug.Print ssql
                        ijrs.Open ssql, wsDB, adOpenDynamic, adLockOptimistic
                        If Not ijrs.EOF Then
                            ijrs!REMARKS2 = Trim(ijbal)
                            ijrs.Update
                        End If
                        ijrs.Close
                    End If
                End If
             End If
        End If
    
    
    
        'Check for Invoice UP
        TRAN_DATEx = Format(Now, "DD-MMM-YYYY")
        TRAN_TIMEx = Format(Time, "HH:MM:SS")
        Set rs = New ADODB.Recordset
'        SQL = "SELECT * FROM AIC_INVOICE_PRICE_DETAIL WHERE PRODUCT = '" & Trim(internal_device_no_txt) & "'AND TARGETDEVICE = '" & Trim(TARGET_DEVICE_TXT) & "'"
'Quah new invoice system 2012-06-15
        
'2012-06-20 skip below checking.
'        SQL = "SELECT * FROM BAIC_INVOICE_PRICE_HEADER where iph_deviceno= '" & Trim(internal_device_no_txt) & "'"
'        rs.Open SQL, wsDB
'        If rs.EOF = True Then
'            If (Weekday(Now) = "6" And Format(Time, "HHMM") > "0830") Or Weekday(Now) = "7" Or Weekday(Now) = "1" Then 'Fri, Sat, Sun
'                s = MsgBox("ERROR: Price not setup by Finance yet. Proceed to save? ", vbYesNo, "Message")
'                If s = 7 Then   'NO proceed.
'                    Exit Sub
'                Else
'                    Set RS2 = New ADODB.Recordset
'                    SSQL2 = "SELECT top 2 * FROM AIC_INVOICE_ALARM_TABLE"
'                    RS2.Open SSQL2, wsDB
'                    If RS2.EOF = True Then
'                        SSQL = "INSERT INTO AIC_INVOICE_ALARM_TABLE (PRODUCT, TARGETDEVICE, REFNO, EMP_ID, TRAN_DATE, TRAN_TIME) " & _
'                               " VALUES('" & Trim(internal_device_no_txt) & "', '" & Trim(TARGET_DEVICE_TXT) & "', '" & Trim(refno_TXT) & "', '" & Trim(login_id) & "','" & TRAN_DATEx & "', '" & TRAN_TIMEx & "')"
'                        wsDB.Execute SSQL
'                    End If
'                    RS2.Close
'                End If
'            End If
'        End If
'        rs.Close
    
    
    
        If Left(Trim(refno_TXT), 2) = "GM" Or Left(Trim(refno_TXT), 2) = "AN" Or Left(Trim(refno_TXT), 2) = "SN" Then
            '2010-11-03
            xbgdat3 = "NA"
            If XAICBD_DEVICE <> "" Then
                Set rs = New ADODB.Recordset
                SQL = "select COUNT(distinct TMH_MARKING) CNT from BAIC_TM_HEADER where TMH_TARGETDEVICE='" & TARGET_DEVICE_TXT & "'"
                Debug.Print SQL
                rs.Open SQL, wsDB
                If rs!cnt > 1 Then
                    xbgdat3 = XAICBD_DEVICE
                End If
                rs.Close

                Set rs = New ADODB.Recordset
                SQL = " update baic_prodmast set pdm_data_1='" & xbgdat3 & "' where pdm_customer='GMT' and pdm_deviceno='" & Trim(internal_device_no_txt) & "'"
                Debug.Print SQL
                rs.Open SQL, wsDB
            End If
        End If
    
    
    
      'Quah.. 2010-10-15  update BAIC_CUSTOMER_PO
      '---------------------------------------------------------------
        If txt_pomode = "STANDARD" Then
            Call svrdatetime(xserverdate, xservertime, xshifttype, xproddate)
            
            If Trim(lblPONo) = "" Or Trim(lblPONo) = "NA" Or Trim(lblPONo) = "N/A" Then
                MsgBox "This is a STANDARD PO. Pls input the correct PO No.", vbCritical, "Message"
                Exit Sub
            End If
            If Trim(total_poqty) = "" Or Trim(txt_podate) = "" Then
                MsgBox "This is a STANDARD PO. Pls input the TOTAL PO QTY and PO DATE.", vbCritical, "Message"
                Exit Sub
            End If
            Dim CpoRs As ADODB.Recordset
            Set CpoRs = New ADODB.Recordset
            sqltxt = "select * from baic_customer_po wheRE CPO_CUST_SHORTNAME='" & Trim(custnameselect.TEXT) & "' and CPO_PONO='" & lblPONo & "' and CPO_TARGETDEVICE='" & TARGET_DEVICE_TXT & "'"
            CpoRs.Open sqltxt, wsDB, adOpenDynamic, adLockOptimistic
            If CpoRs.EOF Then
                CpoRs.AddNew
                CpoRs!CPO_PO_MODE = "STANDARD"
                CpoRs!cpo_cust_shortname = Trim(custnameselect.TEXT)
                CpoRs!cpo_pono = Trim(lblPONo)
                CpoRs!cpo_targetdevice = Trim(TARGET_DEVICE_TXT)
                CpoRs!cpo_order_qty = total_poqty
                CpoRs!CPO_ORDER_YMD = xserverdate 'txt_podate
                CpoRs!cpo_prd_status = "OPN"
'                CpoRs!CPO_PRD_CLOSE_MODE = ""
'                CpoRs!CPO_PRD_CLOSE_YMD = ""
                CpoRs!CPO_FIN_STATUS = "OPN"
'                CpoRs!CPO_FIN_CLOSE_BY = ""
'                CpoRs!CPO_FIN_CLOSE_YMD = ""
'                CpoRs!CPO_REMARK = ""
                CpoRs!CPO_CREATED_BY = Trim(login_id)
                CpoRs!CPO_CREATION_YMD = Now()
                CpoRs!CPO_FIN_GOOD_QTY = 0
                CpoRs!CPO_FIN_REJ_QTY = 0
                CpoRs.Update
            Else
                CpoRs!cpo_order_qty = total_poqty
                CpoRs.Update
            End If
            CpoRs.Close
            Set CpoRs = Nothing
        End If
      '---------------------------------------------------------------
    
    
        'INSERT NEW LI
        Dim adoLI As ADODB.Recordset
        Set adoLI = New ADODB.Recordset
        wsSqlString = "select * from AIC_LOADING_INSTRUCTION" _
           & " where REFNO='" & Trim(refno_TXT) & "'"
        adoLI.Open wsSqlString, wsDB, adOpenDynamic, adLockOptimistic
        If adoLI.EOF Then
            adoLI.AddNew
            adoLI!DEVICE_NO = Trim(internal_device_no_txt)
            adoLI!PACKAGE__LEAD = Trim(package_lead_txt)
            adoLI!TARGET_DEVICE = Trim(TARGET_DEVICE_TXT)
            adoLI!WAFER = Trim(WAFER)
            adoLI!qty = Trim(qty)
            adoLI!LOAD_TIME = Trim(cboruntype)
            adoLI!BD_NO = Trim(bonding_diagram_txt)
            adoLI!CUSTOMER_NO = Left(Right(Trim(cboCust.TEXT), 4), 3)
            adoLI!CUSTOMER_NAME = Trim(custnameselect.TEXT)
            adoLI!MARKING_SPEC = Trim(mark_spec_txt)
            If InStr(internal_device_no_txt, "TST") > 0 And adoLI!CUSTOMER_NAME = "STMICRO" Then 'KOKO 20221107
            adoLI!SpecChar = "PTL" 'KOKOKO20221107
            Else
            adoLI!SpecChar = SpecChars
            End If 'KOKOKO 20221107
            Top1DAT = Trim(topx(0))
            Top2DAT = Trim(topx(1))
            Top3DAT = Trim(topx(2))
            Top4DAT = Trim(topx(3))
            Top5DAT = Trim(topx(4))
            Top6DAT = Trim(topx(5))
            Bot1DAT = Trim(bottom(0))
            Bot2DAT = Trim(bottom(1))
            Bot3DAT = Trim(bottom(2))
            Bot4DAT = Trim(bottom(3))
            Bot5DAT = Trim(bottom(4))
            Bot6DAT = Trim(bottom(5))
            adoLI!TOP1 = Top1DAT
            adoLI!TOP2 = Top2DAT
            adoLI!TOP3 = Top3DAT
            adoLI!TOP4 = Top4DAT
            adoLI!TOP5 = Top5DAT
            adoLI!TOP6 = Top6DAT
            adoLI!BOTTOM1 = Bot1DAT
            adoLI!BOTTOM2 = Bot2DAT
            adoLI!BOTTOM3 = Bot3DAT
            adoLI!BOTTOM4 = Bot4DAT
            adoLI!BOTTOM5 = Bot5DAT
            adoLI!BOTTOM6 = Bot6DAT
            adoLI!DATE_TRANX = DX
            adoLI!TIME_TRANX = nowx
            adoLI!EMP_ID = Trim(login_id)
            'adoLI!EMP_ID = Trim(lblPONo)
            adoLI!Status = "N"
            adoLI!Refno = Trim(refno_TXT)
            adoLI!work_week = ww_txt
            adoLI!PRINT_FLAG = "N"
            adoLI!CUSLOTNO = Trim(txtCusLot)
            'adoLI!PO_NO = lblPONo
            'adoLI!PO_NO = Left(trim(login_id), 5)
            adoLI!PO_NO = Trim(lblPONo)
            adoLI!CATALOGNO = Trim(txtCatalogNo)
            adoLI!INTERNAL_BD = Trim(bd_no_txt.TEXT)      '2012-09-04
            adoLI.Update
            
            Call SaveDie
            '*****KOKO
            'KOKO ADD 20210627 ' FOR ADD LI BOM
            '20210705 move to beginning of SAVE
            'Call insertLIBOM(refno_TXT)
            'KOKO END
            Dim adocounter As ADODB.Recordset
            Set adocounter = New ADODB.Recordset
  
            YDCT = Format(Date, "YY")
            MDCT = Format(Date, "MM")
            DDCT = Format(Date, "DD")

            wsSqlString = "select * from AIC_LI_COUNTER" _
                 & " where PREFIX = '" & Trim(prefix) & "' AND customer_no ='" & Left(Right(Trim(cboCust.TEXT), 4), 3) & "' AND M_COUNTER='" & MDCT & "' AND Y_COUNTER='" & YDCT & "'"
            adocounter.Open wsSqlString, wsDB, adOpenDynamic, adLockOptimistic
            If Not adocounter.EOF Then
                adocounter!Counter = adocounter!Counter + 1
                
                adocounter.Update
            Else
                adocounter.AddNew
                adocounter!Y_Counter = YDCT
                adocounter!M_Counter = MDCT
                adocounter!Counter = 1
                adocounter!CUSTOMER_NO = Left(Right(Trim(cboCust.TEXT), 4), 3)
                adocounter!prefix = prefix
                adocounter.Update
            End If
            adocounter.Close

             
             'if exist & status = "N" & OPTION2 = release
        ElseIf Trim(stat) = "N" And RELEASE_OPT.Value = True Then
                
                adoLI!Status = "R"
                adoLI!DATE_TRANX = DX
                adoLI!TIME_TRANX = nowx
                adoLI.Update
                
'******
'CZZ Added 2012-08-29 LI Material Shortage Prevention
                AddLIReserved Trim(refno_TXT.TEXT)
'******
                
                ' === AUTO INTERFACE start 2011-04-14 =================================
'                    REFNOX = Trim(Me.lvwToLoad.ListItems(iCnt).Text)
'""Q                    WAFER_LOTNOX = Trim(Me.lvwToLoad.ListItems(iCnt).SubItems(1))
'""Q                    CUSLOTNOX = Trim(Me.lvwToLoad.ListItems(iCnt).SubItems(2))
 '                   PRODUCTX = Trim(Me.lvwToLoad.ListItems(iCnt).SubItems(3))
'""Q                    TARGET_DEVICEX = Trim(Me.lvwToLoad.ListItems(iCnt).SubItems(4))
'""Q                    CUSTOMERX = cboCust.Text
'""Q                    WAFER_LOT_QTYX = Trim(Me.lvwToLoad.ListItems(iCnt).SubItems(5))
            
                    Set countRs = New ADODB.Recordset
                    ssql = "select * FROM AIC_SYNC_COUNTER"
                    countRs.Open ssql, wsDB
                    If countRs.EOF = False Then ORACLEWKNOx = countRs!Counter
                    countRs.Close
                        
                        REFNOx = Trim(adoLI!Refno)
                        WAFERX = Trim(adoLI!WAFER)
                        CUSTLOTNOX = Trim(adoLI!CUSLOTNO)
                        PRODUCTX = Trim(adoLI!DEVICE_NO)
                        
                        'rechck ENGR status 2010-09-14
                        'Disabled 2012-08-29 due to E1
'                        SSQL = "select refno, rem7 from aic_loading_instruction_remark where refno like '" & REFNOx & "'"
'                        Dim engrrs As ADODB.Recordset
'                        Set engrrs = New ADODB.Recordset
'                        engrrs.Open SSQL, wsDB
'                        xengrlot = ""
'                        If Not engrrs.EOF Then
'                            If Trim(engrrs!REM7) = "RED" Then
'                                xengrlot = "Engineering Lot"
'                            End If
'                        End If
'                        engrrs.Close
'                        SSQL = "update aic_loading_instruction set load_time='" & xengrlot & "' where refno like '" & REFNOx & "'"
'                        wsDB.Execute SSQL
                        
                        ssql = "UPDATE AIC_SYNC_COUNTER SET COUNTER = '" & Trim(ORACLEWKNOx + 1) & "'"
                        wsDB.Execute ssql
                       
                        ssql = "DELETE FROM AIC_WK_ORDER_ALL where ORDER_NO='" & REFNOx & "'"
                        wsDB.Execute ssql
                       
                        SSQL1 = "INSERT INTO aic_wk_order_all (ORACLE_WORK_ORDER) VALUES ('" & Trim(ORACLEWKNOx) & "')"
                        wsDB.Execute SSQL1
                       
                       ssql = "UPDATE AIC_LOADING_INSTRUCTION SET SYNC='Y', SYNC_DATE = '" & Format(Date, "DD-MMM-YYYY") & "', SYNC_TIME = '" & Format(Time, "HH:MM:SS") & "' WHERE REFNO = '" & REFNOx & "'"
                       wsDB.Execute ssql
                        
                        ssql = "UPDATE aic_wk_order_all " & _
                             " SET WAFER_LOTNO = '" & Trim(adoLI!WAFER) & "', CUST_LOTNO = '" & Trim(adoLI!CUSLOTNO) & "', " & _
                             " ORDER_NO = '" & Trim(REFNOx) & "', PRODUCT =  '" & Trim(adoLI!DEVICE_NO) & "', " & _
                             " TARGET_DEVICE = '" & Trim(adoLI!TARGET_DEVICE) & "', CUSTOMER = '" & Trim(adoLI!CUSTOMER_NAME) & "', " & _
                             " CUSTOMER_CODE = '" & Trim(adoLI!CUSTOMER_NO) & "', WAFER_LOT_QTY = '" & Trim(adoLI!qty) & "', " & _
                             " TOP_MARK_1 = '" & Trim(adoLI!TOP1) & "',TOP_MARK_2 = '" & Trim(adoLI!TOP2) & "', " & _
                             " TOP_MARK_3 = '" & Trim(adoLI!TOP3) & "', TOP_MARK_4 = '" & Trim(adoLI!TOP4) & "', " & _
                             " TOP_MARK_5 = '" & Trim(adoLI!TOP5) & "', TOP_MARK_6 = '" & Trim(adoLI!TOP6) & "', " & _
                             " BOTTOM_MARK_1 = '" & Trim(adoLI!BOTTOM1) & "', BOTTOM_MARK_2 = '" & Trim(adoLI!BOTTOM2) & "', " & _
                             " BOTTOM_MARK_3 = '" & Trim(adoLI!BOTTOM3) & "', BOTTOM_MARK_4 = '" & Trim(adoLI!BOTTOM4) & "', " & _
                             " BOTTOM_MARK_5 = '" & Trim(adoLI!BOTTOM5) & "', BOTTOM_MARK_6 =  '" & Trim(adoLI!BOTTOM6) & "', " & _
                             " BOM_ID = '296546', DIE_WAFER_LOT_1 = '" & Trim(adoLI!WAFER) & "', " & _
                             " DIE_QTY_1 = '" & Trim(adoLI!qty) & "', LOT_FLAG = 'N', " & _
                             " CATALOG_NO = '" & Trim(adoLI!CATALOGNO) & "' " & _
                             " WHERE ORACLE_WORK_ORDER = '" & Trim(ORACLEWKNOx) & "'"
                        wsDB.Execute ssql
                       
                       
                       
                       Set rs = New Recordset
                       SSQL1 = "SELECT cus_destination FROM baic_customer WHERE CUS_CODE = '" & Trim(adoLI!CUSTOMER_NO) & "' "
                       rs.Open SSQL1, wsDB
                                      
                            SSQL1 = "UPDATE aic_wk_order_all SET DESTINATION = '" & Trim(rs!CUS_DESTINATION) & "' " & _
                                     " WHERE ORACLE_WORK_ORDER = '" & Trim(ORACLEWKNOx) & "'"
                            wsDB.Execute SSQL1
                            rs.Close
            
                       Set orRS = New Recordset
                       'Quah 2013-05-27 Order by Oper
                       'Quah 2018-10-18 Order by Oper, Sno.... req by EK TAN.
'''                         orSQL = "select B.device_no, B.part_no, B.part_short_desc " & _
'''                                  "from aic_bom_header A, aic_bom_component B " & _
'''                                  "Where a.bill_sequence_id = b.bill_sequence_id " & _
'''                                  "and A.active_flag = 'Y' " & _
'''                                  "and a.device_no = '" & Trim(PRODUCTX) & "' " & _
'''                                  "and b.apt_print = 'Y' " & _
'''                                  "and b.remark = 'A' " & _
'''                                  "and b.part_no not like 'DIE%' " & _
'''                                  "and b.part_no not like '052%' " & _
'''                                  "order by b.operation_code, b.serial_no  "
'                                  "order by b.serial_no "
                        
'                        ' Quah/Liyana 2021-12-16 link to BAIC_LI_BOM to get parts selected by Planner.
                            orSQL = "select B.device_no, B.part_no, B.part_short_desc " & _
                                    "from BAIC_LI_BOM A, AIC_BOM_COMPONENT B " & _
                                    "Where a.BOM_PARTNO = b.PART_NO " & _
                                    "and a.BOM_DEVICENO = b.DEVICE_NO " & _
                                    "and a.BOM_LIREF = '" & REFNOx & "' " & _
                                    "and b.remark = 'A' " & _
                                    "and b.part_no not like 'DIE%' " & _
                                    "and b.part_no not like '052%' " & _
                                    "order by b.operation_code, b.serial_no  "

                        
                        orRS.Open orSQL, wsDB
                               
                        'UPDATE TO aic_wk_order_all - BOM_PART_NO(ICNT), BOM_ITEM_ID_DESC(ICNT), BOM_ITEM_ID_SDESC(ICNT) (KEY:PRODUCT)
                        iCnt = 1
                        Do While Not orRS.EOF
                            If iCnt < 16 Then 'KOKO ADD 20191112 FOR SKIP PART NO MORE THAN 16
                            BOM_PART_NOX = Trim(orRS!PART_NO)                                                                     'Part No.
                            BOM_ITEM_IDZ_DESCX = IIf(IsNull(orRS!PART_SHORT_DESC), "-", Left(Trim(orRS!PART_SHORT_DESC), 50))     'Short Desc
                            BOM_ITEM_ID_SDESCX = IIf(IsNull(orRS!PART_SHORT_DESC), "-", Left(Trim(orRS!PART_SHORT_DESC), 50))     'Short Desc
                            
                                 wsSQL = " UPDATE aic_wk_order_all SET BOM_PART_NO_" & Trim(Str(iCnt)) & " ='" & BOM_PART_NOX & "', " & _
                                         " BOM_ITEM_ID_DESC_" & Trim(Str(iCnt)) & " = '" & BOM_ITEM_IDZ_DESCX & "', " & _
                                         " BOM_ITEM_ID_SDESC_" & Trim(Str(iCnt)) & " = '" & BOM_ITEM_ID_SDESCX & "'" & _
                                         " WHERE ORACLE_WORK_ORDER = '" & Trim(ORACLEWKNOx) & "'"
                                         Debug.Print wsSQL
                                 wsDB.Execute wsSQL
                            

                          End If
                          
                          iCnt = iCnt + 1
                            orRS.MoveNext
                        Loop
                        orRS.Close
                           
                       'RETRIEVE DIE_SOURCE
                       Set orRS = New ADODB.Recordset
                         orSQL = "select B.device_no, B.part_no, B.part_short_desc " & _
                                  "from aic_bom_header A, aic_bom_component B " & _
                                  "Where a.bill_sequence_id = b.bill_sequence_id " & _
                                  "and A.active_flag = 'Y' " & _
                                  "and a.device_no = '" & Trim(PRODUCTX) & "' " & _
                                  "and b.remark = 'A' " & _
                                  "and b.part_no like 'DIE%' " & _
                                  "order by b.operation_code "
'                                  "order by b.serial_no "
                       
                       orRS.Open orSQL, wsDB
                       
                       If orRS.EOF = False Then
                            wsSQL = " UPDATE aic_wk_order_all SET DIE_SOURCE_1 = '" & Trim(orRS!PART_NO) & "' " & _
                                " WHERE ORACLE_WORK_ORDER = '" & Trim(ORACLEWKNOx) & "'"
                            wsDB.Execute wsSQL
                       
                       End If
                       orRS.Close

'                    End If
'                    tempRS.Close
                '=== AUTO INTERFACE end ===============================================
                MsgBox " PARTICULAR LI RELEASED. PLEASE PROCEED WITH LOT CREATION.", vbInformation, "Message"
                Unload Me
                LI_General.Show
                Exit Sub
        ElseIf Trim(stat) = "N" And CANCEL_OPT.Value = True Then
                'check for loaded lots.
                Set adocounter = New ADODB.Recordset
                wsSqlString = "select * from AIC_WIP_HEADER, baic_lotmast " _
                     & " where assylotno=ltm_lotno_9 and ltm_deleted='N' and orderno='" & Trim(refno_TXT) & "'"
                Debug.Print wsSqlString
                adocounter.Open wsSqlString, wsDB
                If Not adocounter.EOF Then
                    MsgBox "Cannot CANCEL. Lots already loaded. Please check.", vbCritical, "Message"
                    Exit Sub
                Else
                    adoLI!Status = "C"
                    adoLI!DATE_TRANX = DX
                    adoLI!TIME_TRANX = nowx
                    adoLI.Update
                    'delete if exist in aic_wk_order_all  '2011-04-14
                    '-------------------
                    ssql = "DELETE FROM AIC_WK_ORDER_ALL where ORDER_NO='" & Trim(refno_TXT) & "'"
                    wsDB.Execute ssql
                    '-------------------
                    
                    'if no more active LI, delete from baic_customer_po
                    Dim chkcancel As ADODB.Recordset
                    Set chkcancel = New ADODB.Recordset
                    ssql = "select * from AIC_LOADING_INSTRUCTION where (STATUS='N' or STATUS='R') and CUSTOMER_NAME='" & Trim(custnameselect.TEXT) & "' and TARGET_DEVICE='" & Trim(TARGET_DEVICE_TXT) & "' and PO_NO='" & Trim(lblPONo) & "'"
                    Debug.Print ssql
                    chkcancel.Open ssql, wsDB
                    If chkcancel.EOF Then
                        'no nore
                        ssql = "delete from baic_customer_po wheRE CPO_CUST_SHORTNAME='" & Trim(custnameselect.TEXT) & "' and CPO_PONO='" & lblPONo & "' and CPO_TARGETDEVICE='" & TARGET_DEVICE_TXT & "'"
                        Debug.Print ssql
                        wsDB.Execute ssql
                    End If
                    chkcancel.Close
                    Set chkcancel = Nothing
                    
                    MsgBox "PARTICULAR LI CANCELLED SUCCESSFULLY ! ", , "Message"
                    Unload Me
                    LI_General.Show
                    Exit Sub
                End If
                adocounter.Close
        Else
                MsgBox "DATA ALREADY EXIST! PLEASE USE UPDATE!"
        End If
        adoLI.Close
        PULLDATA
    End If
End Sub
Private Sub target_device_txt1old_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
            bonding_diagram_txt.TEXT = vbNullString
            bd_no_txt.TEXT = vbNullString
            bd_no_txt1.TEXT = vbNullString
            TARGET_DEVICE_TXT = vbNullString
            cbobonding_diagram.Visible = True
            target_device_txt1.Visible = False
            
            cbobonding_diagram.SetFocus
            Set rsbd = New ADODB.Recordset
            CSQLSTRING = "SELECT * FROM aic_bd_no WHERE AICBD_TARGET_DEVICE = '" & Trim(target_device_txt1) & "' AND AICBD_STATUS='INUSE'"
            rsbd.Open CSQLSTRING, wsDB
            Do While Not rsbd.EOF
                cbobonding_diagram.AddItem Trim(rsbd!AICBD_CUSBD_NUMBER)
                rsbd.MoveNext
            Loop
            rsbd.Close
            Set rsbd = Nothing
            TARGET_DEVICE_TXT.TEXT = Trim(target_device_txt1.TEXT)
    End If
End Sub
Private Sub target_device_txt1_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
            
            If package_txt.TEXT = "" Then
                MsgBox "Please select Package !"
                Exit Sub
            End If
            
            If Left(Trim(refno_TXT), 2) = "GM" Then
                If Trim(TARGET_DEVICE_TXT) = "" Then
                    MsgBox "GMT L.I. must be import from softcopy...", vbCritical, "Message"
                    Exit Sub
                End If
                If TARGET_DEVICE_TXT <> "" And target_device_txt1 <> "" Then
                    If TARGET_DEVICE_TXT <> target_device_txt1 Then
                        MsgBox "Cannot change Target Device", vbCritical, "Message"
                        Exit Sub
                    End If
                End If
            End If
            
            
            
            bonding_diagram_txt.TEXT = vbNullString
            bd_no_txt.TEXT = vbNullString
            bd_no_txt1.TEXT = vbNullString
            cbobonding_diagram.Visible = True
            cbobonding_diagram.Clear
            target_device_txt1.Visible = False
            TARGET_DEVICE_TXT = vbNullString
            TARGET_DEVICE_TXT = Trim(target_device_txt1)
            Set rsbd = New ADODB.Recordset
            
            'Quah 20080825 do not convert target_device_txt1 to UpperCase, because Customer CORERIVER uses some small letters (e.g. MiDAS, TouchCore) in the targetdevice.
            CSQLSTRING = "SELECT * FROM aic_bd_no WHERE AICBD_TARGET_DEVICE = '" & Trim(target_device_txt1) & "' AND AICBD_STATUS='INUSE'"
            Debug.Print CSQLSTRING
            rsbd.Open CSQLSTRING, wsDB
            If rsbd.EOF = False Then
                Do While Not rsbd.EOF
                    cbobonding_diagram.AddItem Trim(rsbd!AICBD_CUSBD_NUMBER)
                    rsbd.MoveNext
                Loop
            Else
                Dim REV, REV1, REV2 As String
                Dim WIPPRDx, RSBD1 As ADODB.Recordset
                Set WIPPRDx = New ADODB.Recordset
                Dim ssql, BDSQL As String
                
                REV = vbNullString
    '            SSQL = "SELECT DISTINCT WPRD_USRDF_SMDAT_2 FROM WIPPRD WHERE  WPRD_PRD_GRP_2 = '" & Trim(package_txt.TEXT) & "' AND  WPRD_PRD_GRP_3 = '" & Trim(ld_txt.TEXT) & "' AND WPRD_DESC = '" & Trim(target_device_txt1) & "'"
'AIMS
                ssql = "SELECT DISTINCT PDM_INTERNAL_BD WPRD_USRDF_SMDAT_2 FROM baic_prodmast WHERE  PDM_PACKAGE = '" & Trim(package_txt.TEXT) & "' AND  PDM_LEAD = '" & Trim(ld_txt.TEXT) & "' AND PDM_TARGETDEVICE = '" & Trim(target_device_txt1) & "'"
                
                WIPPRDx.Open ssql, wsDB
                If WIPPRDx.EOF = False Then
                    Do While Not WIPPRDx.EOF
                        If Not IsNull(WIPPRDx!WPRD_USRDF_SMDAT_2) Then
                            REV = Trim(WIPPRDx!WPRD_USRDF_SMDAT_2)
                            
                            bdlen = Len(Trim(REV))
                            REV1 = Left(Trim(REV), bdlen - 2)
'                            REV1 = Left(Trim(REV), 9)
                            
                            REV2 = Right(Trim(REV), 1)
                            Set RSBD1 = New ADODB.Recordset
                            'BDSQL = "SELECT * FROM AIC_BD_NO WHERE AICBD_BD_NUMBER = '" & Trim(REV1) & "' AND AICBD_REVISION = '" & Trim(REV2) & "'"
                            BDSQL = "SELECT * FROM AIC_BD_NO WHERE AICBD_BD_NUMBER = '" & Trim(REV1) & "' AND AICBD_REVISION = '" & Trim(REV2) & "' AND AICBD_STATUS = 'INUSE'"
                            RSBD1.Open BDSQL, wsDB
                            
                            If RSBD1.EOF = False Then
                                Do While Not RSBD1.EOF
                                    cbobonding_diagram.AddItem Trim(RSBD1!AICBD_CUSBD_NUMBER)
                                    RSBD1.MoveNext
                                Loop
                            End If
                            RSBD1.Close
                        Else
                            MsgBox "INTERNAL BD NO NOT FOUND.", vbInformation
                            Exit Sub
                        End If
                    WIPPRDx.MoveNext
                    DoEvents
                    Loop
                Else
                
 '                  MsgBox Trim(package_txt.TEXT) & " (WPRD_PRD_GRP_2)" & Chr(13) & Trim(ld_txt.TEXT) & " (WPRD_PRD_GRP_3)" & Chr(13) & Trim(target_device_txt1) & " (WPRD_DESC)" & Chr(13) & Chr(13) & "Record Not Match In WIPPRD. Please check ORACLE/WORKSTREAM Device Registration."
                    MsgBox Trim(package_txt.TEXT) & " (PACKAGE)" & Chr(13) & Trim(ld_txt.TEXT) & " (LEAD)" & Chr(13) & Trim(target_device_txt1) & " (TARGETDEVICE)" & Chr(13) & Chr(13) & "Record Not Match In AIMS. Please check AIMS Device Registration."

'                    MsgBox "BD NOT REGISTERED / ALREADY OBSOLETED.", vbInformation
                    Exit Sub
                End If
                WIPPRDx.Close
            End If
            rsbd.Close
            Set rsbd = Nothing
            cbobonding_diagram.SetFocus
    End If
End Sub

Private Sub target_device_txt1WITHOUTCUSTBD_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
            bonding_diagram_txt.TEXT = vbNullString
            TARGET_DEVICE_TXT = vbNullString
            bd_no_txt.TEXT = vbNullString
            bd_no_txt1.TEXT = vbNullString
            bdcombo.Clear
            'cbobonding_diagram.Visible = True
            'target_device_txt1.Visible = False
            'cbobonding_diagram.SetFocus
            Set rsbd = New ADODB.Recordset
'            CSQLSTRING = "SELECT * FROM BAIC_PRODMAST WHERE  PDM_PACKAGE = '" & Trim(package_txt) & "' AND  PDM_LEAD = '" & Trim(ld_txt) & "' AND PDM_TARGETDEVICE = '" & Trim(target_device_txt1) & "'"
            'Quah 2012-12-12 exclude inactive devices.
            CSQLSTRING = "SELECT * FROM BAIC_PRODMAST WHERE  PDM_PACKAGE = '" & Trim(package_txt) & "' AND  PDM_LEAD = '" & Trim(ld_txt) & "' AND PDM_TARGETDEVICE = '" & Trim(target_device_txt1) & "' and (pdm_inactive_date ='' or pdm_inactive_date is null)"
            rsbd.Open CSQLSTRING, wsDB
            Do While Not rsbd.EOF
                bdcombo.AddItem Trim(rsbd!WPRD_PROD)
                rsbd.MoveNext
            Loop
            rsbd.Close
            Set rsbd = Nothing
            bdcombo.Visible = True
            bdcombo.SetFocus
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text1 = lvwDie.ListItems.Count + 1
    Text2 = vbnulsltring
    Text3 = vbNullString
    Text4 = vbNullString
    Text5 = vbNullString
    Text6 = vbNullString
    Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text3.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    If Left(Me.refno_TXT, 2) = "AP" And Text3.TEXT <> "" Then
        If Text3.TEXT = Me.txtCusLot Then
            MsgBox "Apower Cuslot and Waferlot cannot be same.", vbCritical, "Message"
            Text3.TEXT = ""
        End If
    End If
    
                   '20230323 Chong add to check st wafer expired
                '--------------------------------------------
                If Left(Me.refno_TXT, 2) = "ST" And Text3.TEXT <> "" Then
                    If Left(Trim(Text3.TEXT), 1) = "U" Then
                        Dim chkexpirewafer As ADODB.Recordset
                        Set chkexpirewafer = New ADODB.Recordset
                        ssql = " SELECT CONVERT(VARCHAR,[STW_LOCKDATE],112)AS LOCK_DATE, CONVERT (VARCHAR,[STW_EXP_DATE],112) AS EXP_DATE, CONVERT(VARCHAR,GETDATE(),112) AS CUR_DATE, BYPASS_IND, CONVERT(VARCHAR,[STW_BYPASSDATE],112) AS BYPASS_DATE" & _
                                " FROM BAIC_ST_WAFER_MCODE " & _
                                " WHERE STW_WLOT ='" & Trim(Text3.TEXT) & "' "
                        Debug.Print ssql
                        chkexpirewafer.Open ssql, wsDB
                        If Not chkexpirewafer.EOF Then
                            If chkexpirewafer!CUR_DATE >= chkexpirewafer!lock_date Then
                                expiredWaferBox = MsgBox("" & Trim(Text3.TEXT) & " will Expired on: " & Trim(chkexpirewafer!exp_date) & "", vbCritical, "Message")
                                If chkexpirewafer!BYPASS_IND = 1 Then
                                    If chkexpirewafer!CUR_DATE >= chkexpirewafer!BYPASS_DATE Then
                                        MsgBox ("WAFERNO = " & Trim(Text3.TEXT) & " only BYPASS until: " & Trim(chkexpirewafer!BYPASS_DATE) & " .Unable to proceed.")
                                        MsgBox ("CANCELLED for WAFERNO = " & Trim(Text3.TEXT) & " Enter another WAFER!")
                                        Text3.TEXT = ""
                                    Else
                                        MsgBox ("" & Trim(Text3.TEXT) & " BYPASSED until: " & Trim(chkexpirewafer!BYPASS_DATE) & "")
                                    End If
                                Else
                                
                                MsgBox ("WAFER LOCKED!!CANCELLED for WAFERNO = " & Trim(Text3.TEXT) & " .Enter another WAFER!")
                                Text3.TEXT = ""
                                End If
                            End If
                         End If
                         chkexpirewafer.Close
                    End If
                End If
                '--------------------------------------------
                
    Text4.SetFocus
End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
Call CheckNumeric(KeyAscii)
If KeyAscii = 13 Then
    Text5.SetFocus
End If
End Sub

Private Sub CheckNumeric(KeyAscii As Integer)
If IsNumeric(Chr$(KeyAscii)) = False And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> 13 Then
KeyAscii = 0
End If
'If IsNumeric(Chr$(KeyAscii)) = False And KeyAscii <> 13 Then
'KeyAscii = 0
'End If
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
Call CheckNumeric(KeyAscii)
If KeyAscii = 13 Then
    Text6.SetFocus
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
Call CheckNumeric(KeyAscii)
Dim lvwDieCnt As Integer
lvwDieCnt = lvwDie.ListItems.Count + 1
If KeyAscii = 13 Then
    If LenB(Text1) = 0 Or LenB(Text2) = 0 Or LenB(Text3) = 0 Or LenB(Text4) = 0 Or LenB(Text5) = 0 Or LenB(Text6) = 0 Then
        MsgBox "Cannnot have Empty Field.", vbInformation
        Exit Sub
    End If
    
                
                '20230323 Chong add to check st wafer expired
                '--------------------------------------------
                If Left(Me.refno_TXT, 2) = "ST" And Text3.TEXT <> "" Then
                    If Left(Trim(Text3.TEXT), 1) = "U" Then
                        Dim chkexpirewafer As ADODB.Recordset
                        Set chkexpirewafer = New ADODB.Recordset
                        ssql = " SELECT CONVERT(VARCHAR,[STW_LOCKDATE],112)AS LOCK_DATE, CONVERT (VARCHAR,[STW_EXP_DATE],112) AS EXP_DATE, CONVERT(VARCHAR,GETDATE(),112) AS CUR_DATE, BYPASS_IND, CONVERT(VARCHAR,[STW_BYPASSDATE],112) AS BYPASS_DATE" & _
                                " FROM BAIC_ST_WAFER_MCODE " & _
                                " WHERE STW_WLOT ='" & Trim(Text3.TEXT) & "' "
                        Debug.Print ssql
                        chkexpirewafer.Open ssql, wsDB
                        If Not chkexpirewafer.EOF Then
                            If chkexpirewafer!CUR_DATE >= chkexpirewafer!lock_date Then
                                expiredWaferBox = MsgBox("" & Trim(Text3.TEXT) & " will Expired on: " & Trim(chkexpirewafer!exp_date) & "", vbCritical, "Message")
                                If chkexpirewafer!BYPASS_IND = 1 Then
                                    If chkexpirewafer!CUR_DATE >= chkexpirewafer!BYPASS_DATE Then
                                        MsgBox ("WAFERNO = " & Trim(Text3.TEXT) & " only BYPASS until: " & Trim(chkexpirewafer!BYPASS_DATE) & " .Unable to proceed.")
                                        MsgBox ("CANCELLED for WAFERNO = " & Trim(Text3.TEXT) & " Enter another WAFER!")
                                        Text3.TEXT = ""
                                    Else
                                        MsgBox ("" & Trim(Text3.TEXT) & " BYPASSED until: " & Trim(chkexpirewafer!BYPASS_DATE) & "")
                                    End If
                                Else
                                
                                MsgBox ("WAFER LOCKED!!CANCELLED for WAFERNO = " & Trim(Text3.TEXT) & " .Enter another WAFER!")
                                Text3.TEXT = ""
                                End If
                            End If
                         End If
                         chkexpirewafer.Close
                    End If
                End If
                '--------------------------------------------
                
    
    
    
    
    If Trim(Text1) = "1" Then
        WAFER = Trim(Text3)
        qty = Trim(Text4)
    End If
    
    
    'Quah 2014-01-16 DIODES cuslotno= YYWW-WAFER.AS
    If Left(refno_TXT, 2) = "DT" Then
        dotx = InStr(Trim(Text3), ".")
        If dotx > 0 Then
            txtCusLot = Trim(ww_txt) & "-" & Mid(Trim(Text3), 1, dotx - 1) & ".AS"
            
            'Quah 20200423 Mon req convert marking for DIODES LLLYWW
            If topx(2) = "LLLYWW" Then
                topx(2) = Right(Mid(Trim(Text3), 1, dotx - 1), 3) & Right(ww_txt, 3)
            End If
        Else
            txtCusLot = "?????"
        End If
    End If
    
    'Quah 2012-07-24 check for Fairchild masterlist
    If Left(Trim(refno_TXT), 2) = "FS" Or Left(Trim(refno_TXT), 2) = "FP" Then
        fairpart = Trim(Text2.TEXT)
        fairpart = Replace(fairpart, "MOS : ", "")
        fairpart = Replace(fairpart, "IC : ", "")
        fairpart = Replace(fairpart, "DIE 1 : ", "")
        fairpart = Replace(fairpart, "DIE 2 : ", "")
        
        
        Dim checkFSRs As ADODB.Recordset
        Set checkFSRs = New ADODB.Recordset
        If Left(Trim(Text2.TEXT), 2) = "IC" Or Left(Trim(Text2.TEXT), 5) = "DIE 1" Then
            ssql = "select * from fom_marking where FOM_ITEMID='" & Trim(TARGET_DEVICE_TXT) & "' and FOM_COMPID like '" & Trim(fairpart) & "%'"
        Else
            ssql = "select * from fom_marking where FOM_ITEMID='" & Trim(TARGET_DEVICE_TXT) & "' and FOM_COMPID='" & Trim(fairpart) & "'"
        End If
        Debug.Print ssql
        checkFSRs.Open ssql, wsDB
        If checkFSRs.EOF Then
            MsgBox "Partid not match with targetdevice in FOM Masterlist.", vbCritical, "Message"
            Exit Sub
        End If
        checkFSRs.Close
    End If
    
    
    
''''''''    'Quah 20110701 check MICREL wafer duplicate loading.
''''''''    'KO add 20151103 Included Microchip
''''''''    If (Left(Trim(refno_TXT), 2) = "MU" Or Left(Trim(refno_TXT), 2) = "MC") Then
''''''''        If WAFER <> "NA" And WAFER <> "N/A" Then
''''''''            Dim checkMURs As ADODB.Recordset
''''''''            Set checkMURs = New ADODB.Recordset
'''''''''            SSQL = "select * from AIC_LOADING_INSTRUCTION where STATUS in ('R','N') and REFNO like 'MU%' AND REFNO <> '" & Trim(refno_TXT) & "' and wafer='" & Trim(WAFER) & "'"
'''''''''2012-06-22 check by Cuslot
''''''''            SSQL = "select * from AIC_LOADING_INSTRUCTION where STATUS in ('R','N') and (REFNO like 'MU%' OR  REFNO like 'MC%') AND REFNO <> '" & Trim(refno_TXT) & "' and CUSLOTNO='" & Trim(txtCusLot) & "'"
''''''''            Debug.Print SSQL
''''''''            checkMURs.Open SSQL, wsDB
''''''''            If Not checkMURs.EOF Then
''''''''                MsgBox "Error! Wafer already used in L.I. : " & Trim(checkMURs!Refno), vbCritical, "Message"
''''''''                Exit Sub
''''''''            End If
''''''''        End If
''''''''    End If
    
'=========================
    'Quah 20140416 Anita request Waferlot must match with DeviceType in IQA table.
    'for BE, FS, FP, GE, TA
    If Left(Trim(refno_TXT), 2) = "BE" Or Left(Trim(refno_TXT), 2) = "FS" Or Left(Trim(refno_TXT), 2) = "FP" Or Left(Trim(refno_TXT), 2) = "GE" Or Left(Trim(refno_TXT), 2) = "TA" Or Left(Trim(refno_TXT), 2) = "NV" Then
        'Quah 20160923 also skip checking for TST, RWK
        If InStr(internal_device_no_txt, "-CAP") > 0 Or InStr(internal_device_no_txt, "TST") > 0 Or InStr(internal_device_no_txt, "RWK") > 0 Then
            '20140507 req by Anita, for GE CAP loading.
            'no need to check
        Else
            Dim devchk As ADODB.Recordset
            Set devchk = New ADODB.Recordset
'            ssql = " select * from AIC_INVENTORY_MASTER where WAFERLOTNO='" & Trim(Text3) & "' and DEVICE_NO='" & Trim(Text2) & "' "
            
            'Quah 20200505 for Amphenol, link by device first 5 chars, exclude HSE.
            If (Left(Trim(refno_TXT), 2) = "NV" Or Left(Trim(refno_TXT), 2) = "GE") And InStr(internal_device_no_txt, "HSE") = 0 Then
                ssql = " select * from AIC_INVENTORY_MASTER where WAFERLOTNO='" & Trim(Text3) & "' and DEVICE_NO like '" & Left(Trim(Text2), 5) & "%' "
            Else
                ssql = " select * from AIC_INVENTORY_MASTER where WAFERLOTNO='" & Trim(Text3) & "' and DEVICE_NO='" & Trim(Text2) & "' "
            End If
            
            Debug.Print ssql
            devchk.Open ssql, wsDB
            If devchk.EOF Then
            
                'Quah 20141003 temporary bypass (req by Anita) for Hari Raya advance loading...
                'Quah 20141111 re-open back the control.
                '20170414 temporary bypass. (Nabila)
                '20170417 activate back.
                MsgBox "DiePart and Waferlot not match in IQA record....", vbCritical, "Message"
                Exit Sub
            
            End If
            devchk.Close
            Set devchk = Nothing
        End If
    End If
'=========================
    
    If Val(Text1) > lvwDieCnt - 1 Then
        Set itmx = lvwDie.ListItems.Add(1, , lvwDieCnt)
        itmx.SubItems(1) = Trim(Text2)
        itmx.SubItems(2) = Trim(Text3)
        itmx.SubItems(3) = Trim(Text4)
        itmx.SubItems(4) = Trim(Text5)
        itmx.SubItems(5) = Trim(Text6)
    Else
        Set itmx = lvwDie.SelectedItem
        itmx.TEXT = Trim(Text1)
        itmx.SubItems(1) = Trim(Text2)
        itmx.SubItems(2) = Trim(Text3)
        itmx.SubItems(3) = Trim(Text4)
        itmx.SubItems(4) = Trim(Text5)
        itmx.SubItems(5) = Trim(Text6)
    End If
    Call ResetlvwDie
End If
End Sub

Private Sub ResetlvwDie()
Text1 = vbNullString
Text2 = vbnulsltring
Text3 = vbNullString
Text4 = vbNullString
Text5 = vbNullString
Text6 = vbNullString
Text1.SetFocus
End Sub

Private Sub lvwDie_Click()
If lvwDie.ListItems.Count <> 0 Then
Set itmx = lvwDie.SelectedItem
Text1 = itmx.TEXT
Text2 = itmx.SubItems(1)    'DIE PART NO
Text3 = itmx.SubItems(2)    'WAFER LOT
Text4 = itmx.SubItems(3)    'BUILD DIE QTY
Text5 = itmx.SubItems(4)    'DIE QTY
Text6 = itmx.SubItems(5)    'NO OF WAFER
End If
End Sub
Private Sub lvwDie_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwDie.SortOrder = lvwAscending
    lvwDie.SortKey = ColumnHeader.Index - 1
    lvwDie.Sorted = True
    lvwDie.Sorted = False
End Sub
Private Sub lvwDie_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If MsgBox("Do you want to Delete '" & Trim$(lvwDie.SelectedItem.SubItems(1)) & "'?", vbQuestion + vbYesNo, "Delete") = vbYes Then
            lvwDie.ListItems.Remove (lvwDie.SelectedItem.Index)
            For iCnt = 1 To lvwDie.ListItems.Count
                lvwDie.ListItems(iCnt).TEXT = iCnt
            Next
            Call ResetlvwDie
            Call lvwDie_Click
        End If
    End If
End Sub

Private Sub topx_KeyPress(Index As Integer, KeyAscii As Integer)
    'KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        c = Index + 1
        If c < 6 Then
            topx(c).SetFocus
        ElseIf c = 6 Then
            bottom(0).SetFocus
        End If
    End If
End Sub


'******
'CZZ Added 2012-08-29 LI Material Shortage Prevention
Private Sub txtCheckBalance_Click()
CheckMaterialBalance Trim(refno_TXT.TEXT)
End Sub
'******


Private Sub txtCusLot_Change()
'20170613 Anita req, NS clear marking
If Left(Trim(refno_TXT), 2) = "NN" Then
    topx(0).TEXT = ""
    topx(1).TEXT = ""
    topx(2).TEXT = ""
    topx(3).TEXT = ""
    topx(4).TEXT = ""
    topx(5).TEXT = ""
  '  RL_Assy_inv.lbl_sublot = LI_General.RL_LOTNO_LBL
  '  RL_Assy_inv.txtqty = LI_General.total_poqty
  '  RL_Assy_inv.txtqty = LI_General.qty
End If
End Sub

Private Sub txtCusLot_GotFocus()
oricuslotno = txtCusLot
End Sub

Private Sub txtCusLot_KeyPress(KeyAscii As Integer)
'  KeyAscii = Asc(UCase(Chr(KeyAscii)))
  KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii = 13 Then
        'Quah 2010-05-27 for GMT, auto prefix B.
        If Left(Trim(refno_TXT), 2) = "GM" Then
            If Left(Trim(txtCusLot), 1) <> "" And Left(Trim(txtCusLot), 1) <> "B" Then
                txtCusLot = "B" & Trim(txtCusLot)
            End If
        End If
        
        'QUAH 20181204 BOURNS AUTO LINE 3
        If Left(Trim(refno_TXT), 2) = "BE" And Len(Trim(txtCusLot)) > 5 Then
            
                ''' Conversion change to during APT
                
''''            Dim BeRs As ADODB.Recordset
''''            'match month
''''            Set BeRs = New ADODB.Recordset
''''            SSQL = "SELECT TBL_DATA_A1 FROM BAIC_COMTBL  WHERE TBL_REC_TYPE='BEMK' AND TBL_KEY_A12='MONTH' AND TBL_KEY_9=" & Month(xserverdate)
''''            Debug.Print SSQL
''''            BeRs.Open SSQL, wsDB, adOpenDynamic, adLockOptimistic
''''            If Not BeRs.EOF Then
''''                mthcode = BeRs!TBL_DATA_A1
''''            End If
''''            BeRs.Close
''''
''''            'match day
''''            Set BeRs = New ADODB.Recordset
''''            SSQL = "SELECT TBL_DATA_A1 FROM BAIC_COMTBL  WHERE TBL_REC_TYPE='BEMK' AND TBL_KEY_A12='DAY' AND TBL_KEY_9=" & Day(xserverdate)
''''            Debug.Print SSQL
''''            BeRs.Open SSQL, wsDB, adOpenDynamic, adLockOptimistic
''''            If Not BeRs.EOF Then
''''                daycode = BeRs!TBL_DATA_A1
''''            End If
''''            BeRs.Close
            
            
            Dim newcode
'            newcode = Mid(txtCusLot, 5, 3) & mthcode & daycode
            newcode = Mid(txtCusLot, 5, 3) & "XX"
            topx(2).TEXT = newcode
        End If
        
        'Quah 2013-04-08 for cust ENE
        If Left(Trim(refno_TXT), 2) = "EN" Then
            If Trim(topx(2).TEXT) = "F-XXXXXX" Then
                topx(2).TEXT = Replace(topx(2).TEXT, "XXXXXX", txtCusLot)
            End If
        End If
        
        
        'Quah 2021-08-26 for cust ALL SENSORS
        If Left(Trim(refno_TXT), 2) = "LS" Then
            If Trim(topx(2).TEXT) = "CUST LOT NO" Then
                topx(2).TEXT = txtCusLot
            End If
        End If
        
        
        '20180718 add for RAFFAR
        If Left(Trim(refno_TXT), 2) = "RF" Then
            If InStr(topInfo(2), "YWW") Then
                RFYWW = Right(Trim(ww_txt), 3)
                topx(2).TEXT = Replace(topInfo(2).TEXT, "YWW", RFYWW)
            End If
            If InStr(package_lead_txt, "2X2") > 0 Then
                'Quah add 20200813, req by Mon.
                topx(2).TEXT = Trim(txtCusLot)
            End If
            If InStr(package_lead_txt, "4X4") > 0 Then
                'Quah add 20200814, req by Mon.
                topx(2).TEXT = RFYWW & Left(Trim(txtCusLot), 3) & "A#"
            End If
        
        End If
        
        
        '20210204 add for SILTERRA
        If Left(Trim(refno_TXT), 2) = "GS" Then
            topx(1).TEXT = Replace(topInfo(1).TEXT, "XXX", Left(txtCusLot, 3))
            topx(2).TEXT = Replace(topInfo(2).TEXT, "XXXXXX", Right(txtCusLot, 6))
        End If
        
        
        '20170308 add for SILTERRA
        If Left(Trim(refno_TXT), 2) = "SI" Then
            topx(2).TEXT = Trim(txtCusLot)
            topx(3).TEXT = Trim(ww_txt)
        End If
        
        
        'diana 2014-12-22 for customer IMPINJ
        'Quah disabled 20210922, refer to coding at WORKWEEK.
'        If Left(Trim(refno_TXT), 2) = "IJ" And Len(txtCusLot) = 6 Then
'            X_Mark = Right(Trim(txtCusLot), 5)
'            yz = Right(Trim(yy1), 2)
'            topx(3).TEXT = X_Mark & yz & mm1
'        End If
        
                'diana 2015-02-05 jaeyoung auto-marking
         If Left(Trim(refno_TXT), 2) = "JC" Then
            If InStr(topx(3), "(Cust lot)") Then
               topx(3).TEXT = Trim(txtCusLot)
            End If
            If InStr(topx(4), "YYWW") Then
               topx(4).TEXT = Trim(ww_txt)
            End If
        End If
        
        
        '2020-07-09 Apower front-6 cuslot (before dash) populate to marking line2.
         If Left(Trim(refno_TXT), 2) = "AP" Then
'            If topInfo(2) = "CUST LOT" Then
                clotdash = InStr(Trim(txtCusLot), "-")
               topx(2).TEXT = Mid(Trim(txtCusLot), 1, clotdash - 1)
'            End If
        End If
        
        
         '20180814 Quah add for Microchip PLCC.
         If Left(Trim(refno_TXT), 2) = "MC" And InStr(cbofullpackage, "PLCC") > 0 Then
            Dim aiplcc As ADODB.Recordset
            
            If InStr(topx(1).TEXT, "SLNO") > 0 Then
                Set aiplcc = New ADODB.Recordset
                ssql = "select CLS_DATA_ALPHA from CLS_LOT_INFO where cls_lotno='" & Trim(txtCusLot) & "' and CLS_PARAMETER='Marking Line 1'"
                aiplcc.Open ssql, wsDB, adOpenDynamic, adLockOptimistic
                Debug.Print ssql
                If Not aiplcc.EOF Then
                    fulldat1A = Trim(aiplcc!cls_data_alpha)
                    topx(1).TEXT = Right(fulldat1A, 5)
                End If
                aiplcc.Close
            End If
            
            
            If InStr(topx(2).TEXT, "PART NO") > 0 Then
                Set aiplcc = New ADODB.Recordset
                ssql = "select CLS_DATA_ALPHA from CLS_LOT_INFO where cls_lotno='" & Trim(txtCusLot) & "' and CLS_PARAMETER='Marking Line 2'"
                aiplcc.Open ssql, wsDB, adOpenDynamic, adLockOptimistic
                Debug.Print ssql
                If Not aiplcc.EOF Then
                    fulldat1 = Trim(aiplcc!cls_data_alpha)
                    topx(2).TEXT = Mid(fulldat1, 5)
                End If
                aiplcc.Close
            End If
'            If InStr(topx(2).TEXT, "SPEED, TYPE, DIE_REVISION") > 0 Then
'                Set aiplcc = New ADODB.Recordset
'                ssql = "select CLS_DATA_ALPHA from CLS_LOT_INFO where cls_lotno='" & Trim(txtCusLot) & "' and CLS_PARAMETER='Marking Line 3'"
'                aiplcc.Open ssql, wsDB, adOpenDynamic, adLockOptimistic
'                Debug.Print ssql
'                If Not aiplcc.EOF Then
'                    fulldat2 = Trim(aiplcc!cls_data_alpha)
'                    topx(2).TEXT = Mid(fulldat2, 5, 9) + "V"
'                End If
'                aiplcc.Close
'            End If
            If InStr(topx(3).TEXT, "TRACE CODE") > 0 Then
                Set aiplcc = New ADODB.Recordset
                '20200218 Quah - change line 4 to 3, due to customer revise column.
                ssql = "select CLS_DATA_ALPHA from CLS_LOT_INFO where cls_lotno='" & Trim(txtCusLot) & "' and CLS_PARAMETER='Marking Line 3'"
                Debug.Print ssql
                aiplcc.Open ssql, wsDB, adOpenDynamic, adLockOptimistic
                Debug.Print ssql
                If Not aiplcc.EOF Then
                    fulldat3 = Trim(aiplcc!cls_data_alpha)
                End If
                aiplcc.Close
                If fulldat3 = "<jc>YYWWNNN" Then
                    Set aiplcc = New ADODB.Recordset
                    ssql = "select CLS_DATA_ALPHA from CLS_LOT_INFO where cls_lotno='" & Trim(txtCusLot) & "' and CLS_PARAMETER='YYWWNNN'"
                    aiplcc.Open ssql, wsDB, adOpenDynamic, adLockOptimistic
                    Debug.Print ssql
                    If Not aiplcc.EOF Then
                        topx(3).TEXT = Trim(aiplcc!cls_data_alpha)
                    End If
                    aiplcc.Close
                Else
                    topx(3).TEXT = "?????"
                End If
            End If
         
         End If
        
                'diana 2015-03-25 mini-circuits auto-marking (markspec DVGA1-242A+)
         If Left(Trim(refno_TXT), 2) = "MP" Then
'            If InStr(topx(3), "YYWW") Then
'               topx(3).TEXT = Trim(ww_txt)
'            End If
         
            'Quah 20151223 new requirement
            If InStr(topx(3).TEXT, "WYW") > 0 Then
                Dim mpyy, mpww, mpayy, mp_wyw
                mpyy = Left(ww_txt, 2)
                mpww = Right(ww_txt, 2)
                Select Case mpyy
                    Case "15"
                        mpayy = "F"
                    Case "16"
                        mpayy = "G"
                    Case "17"
                        mpayy = "H"
                    Case "18"
                        mpayy = "J"
                    Case "19"
                        mpayy = "K"
                    Case "20"
                        mpayy = "L"
                    Case "21"
                        mpayy = "M"
                    Case "22"
                        mpayy = "N"
                    Case "23"
                        mpayy = "P"
                    Case "24"
                        mpayy = "Q"
                    Case "25"
                        mpayy = "R"
                    Case "26"
                        mpayy = "S"
                    Case "27"
                        mpayy = "T"
                    Case "28"
                        mpayy = "U"
                    Case "29"
                        mpayy = "V"
                    Case "30"
                        mpayy = "W"
                    Case "31"
                        mpayy = "X"
                    Case "32"
                        mpayy = "Y"
                    Case "33"
                        mpayy = "Z"
                    Case Otherwise
                        MsgBox "Marking YEAR CODE not defined!", vbCritical, "Message"
                        End
                End Select
                mp_wyw = Left(mpww, 1) & mpayy & Right(mpww, 1)
'                            topx(2).TEXT = Replace(topx(2).TEXT, "YYWW", ww_txt)
                topx(3).TEXT = Replace(topx(3).TEXT, "WYW", mp_wyw)
            End If
         End If
         
                'diana 2015-06-29 NIKO YEAR TO ALPHABETS
         If Left(Trim(refno_TXT), 2) = "NK" Then
            'Quah 20180801 add for SOIC and 3X3
            If InStr(cbofullpackage, "3X3") > 0 Then
                Me.topx(1) = Replace(Me.topInfo(1), "XXX", Left(Trim(txtCusLot), 3))
                Me.topx(2) = Replace(Me.topInfo(2), "XXXX", Mid(Trim(txtCusLot), 4, 4))
            End If
            If InStr(cbofullpackage, "SOIC") > 0 Then
                Me.topx(3) = Trim(txtCusLot)
            End If


'            If InStr(topx(2), "YWW") Then 'TOP3
'20210521 change to second page.
'''''''            If InStr(topInfo(2), "YWW") And InStr(cbofullpackage, "2X2") > 0 Then '20180718 base on TOP3 template.
'''''''
'''''''            'REPLACE WITH ALPHA-YEAR AND WORKWEEK FOR YWW
'''''''                Dim NK_Y
'''''''                Dim NKY As ADODB.Recordset
'''''''                Set NKY = New ADODB.Recordset
''''''''                THIS_YEAR = "20" & Mid(refno_TXT, 3, 2)
'''''''                THIS_YEAR = "20" & Mid(ww_txt, 1, 2)     '20180718
'''''''                ssql = " select CUS_DATA_1 from BAIC_CUSTOMER_ADDSETUP where CUS_RECTYPE='NIKO_YEAR' and CUS_KEY_1='" & THIS_YEAR & "' "
'''''''                NKY.Open ssql, wsDB
'''''''                    If NKY.EOF = False Then
''''''''''                        NK_Y = NKY!cus_data_1 & Mid(ww_txt, 2, 2)
'''''''
'''''''                        'Quah 20180718 correction for WW data.
'''''''                        NK_Y = NKY!cus_data_1 & Mid(ww_txt, 3, 2)
'''''''
'''''''                        If Int(Right(NK_Y, 2)) < 1 Or Int(Right(NK_Y, 2)) > 53 Then
'''''''                            MsgBox "INVALID WORKWEEK. PLS CHECK !!!!", vbCritical, "Message"
'''''''                            End
'''''''                        End If
'''''''                    Else
'''''''                        MsgBox "NIKO MARKING YEAR NOT FOUND!PLEASE CHECK!", vbCritical, "Message"
'''''''                        End
'''''''                    End If
'''''''
''''''''                    topx(2).TEXT = Replace(topx(2).TEXT, "YWW", NK_Y)
'''''''                    topx(2).TEXT = Replace(topInfo(2).TEXT, "YWW", NK_Y)
'''''''                NKY.Close
'''''''                Set NKY = Nothing
'''''''            End If
         
'20210521 change to second page.
            '20210322 add for SOT23
'''''            If InStr(topInfo(1), "YWW") And InStr(cbofullpackage, "SOT23") > 0 Then
''''''                Dim NK_Y
''''''                Dim NKY As ADODB.Recordset
'''''                Set NKY = New ADODB.Recordset
'''''                THIS_YEAR = "20" & Mid(ww_txt, 1, 2)     '20180718
'''''                ssql = " select CUS_DATA_1 from BAIC_CUSTOMER_ADDSETUP where CUS_RECTYPE='NIKO_YEAR' and CUS_KEY_1='" & THIS_YEAR & "' "
'''''                NKY.Open ssql, wsDB
'''''                    If NKY.EOF = False Then
'''''                        NK_Y = NKY!cus_data_1 & Mid(ww_txt, 3, 2)
'''''                        If Int(Right(NK_Y, 2)) < 1 Or Int(Right(NK_Y, 2)) > 53 Then
'''''                            MsgBox "INVALID WORKWEEK. PLS CHECK !!!!", vbCritical, "Message"
'''''                            End
'''''                        End If
'''''                    Else
'''''                        MsgBox "NIKO MARKING YEAR NOT FOUND!PLEASE CHECK!", vbCritical, "Message"
'''''                        End
'''''                    End If
'''''                    topx(1).TEXT = Replace(topInfo(1).TEXT, "YWW", NK_Y)
'''''                NKY.Close
'''''                Set NKY = Nothing
'''''            End If
'''''

         End If
                
                 'diana 2015-02-12 mazet auto-marking (markspec SIM12)
                 'diana 2015-02-27 add mazet top3 marking (markspec LC1.4)
                 '20170727 MAZET change name to AMS (AY)
         If Left(Trim(refno_TXT), 2) = "MG" Or Left(Trim(refno_TXT), 2) = "AY" Then
            If InStr(topx(4), "Lotnumber") Then 'mark_spec_txt = "SIM12"
               topx(4).TEXT = Trim(txtCusLot)
            End If
            If InStr(topx(5), "YYWW") Then
               topx(5).TEXT = Trim(ww_txt)
            End If
            If InStr(topx(3), "YYWW") Then 'markspec LC1.4
               topx(3).TEXT = Trim(ww_txt)
            End If
        End If

        
        Call ns_marking
        chkShow.SetFocus
        
    End If
End Sub

Private Sub txtCusLot_LostFocus()

If txtMarkingType.TEXT = "INPUT" And oricuslotno <> "" And txtCusLot <> oricuslotno Then   'Quah 2014-11-24    'initialize device, markspec, marking
    INIT_MARK
    mark_spec_txt = ""
    lbldeviceno = ""
    bdcombo = ""
End If



If Left(Trim(refno_TXT), 2) = "AN" Or Left(Trim(refno_TXT), 2) = "SN" Then
    If cboruntype.TEXT = "Mass Production" Then
        
'            If yy1 = "2010" Or yy1 = "2011" Then
'                If Left(Trim(txtCusLot), 2) = "XM" Or Left(Trim(txtCusLot), 2) = "XN" Then
'                    ' ok
'                Else
'                    MsgBox ("ANPEC/SINOPOWER Customer Lot No. for year 2010=XM, 2011=XN !" & Chr(13))
'                    txtCusLot.TEXT = ""
'                    Exit Sub
'                End If
'            Else
'                    MsgBox ("ANPEC/SINOPOWER Customer Lot No. for year 2010=XM, 2011=XN !" & Chr(13))
'                    txtCusLot.TEXT = ""
'                    Exit Sub
'            End If
'
        
            '2011-12-22
            If yy1 = "2011" Or yy1 = "2012" Then
                If Left(Trim(txtCusLot), 2) = "XN" Or Left(Trim(txtCusLot), 2) = "XO" Then
                    ' ok
                Else
                    MsgBox ("ANPEC/SINOPOWER Customer Lot No. for year 2011=XN, 2012=XO !" & Chr(13))
                    txtCusLot.TEXT = ""
                    Exit Sub
                End If
            Else
                    MsgBox ("ANPEC/SINOPOWER Customer Lot No. for year 2011=XN, 2012=XO !" & Chr(13))
                    txtCusLot.TEXT = ""
                    Exit Sub
            End If
        
        
        
    End If
End If


If Left(Trim(refno_TXT), 2) = "GM" Then 'Auto prefix B for GMT  2010-05-19
    If Left(Trim(txtCusLot), 1) = "B" Then
        'ok
    Else
        Me.txtCusLot = "B" & Trim(txtCusLot)
    End If
End If

End Sub

Private Sub txtNoWaferLot1_Change()
Dim cnt As Integer

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    waferFlag = False
    If KeyAscii = 13 Then
      If IsNumeric(Trim(txtNoWaferLot)) Then
         If CInt(txtNoWaferLot) > 50 Then
           MsgBox "Maximum no of wafer lot is 50! "
           Exit Sub
         End If
      Else
        MsgBox " Please key in No of wafer lot in number!"
        Exit Sub
      End If
   
      If cboCust.TEXT <> "" Then
        waferFlag = True
        FRM_MULTIWFR.Show
      Else
        MsgBox "Please select customer !"
        cboCust.SetFocus
      End If
    End If

End Sub



Private Sub UploadComm_Click()
CDL.Filter = "Picture File | *.bmp"
CDL.ShowOpen
If CDL.FileName <> "" Then
    patternfilename = Left(Right(CDL.FileName, 13), 9)
    If patternfilename <> Trim(refno_TXT) Then
        MsgBox "LI Refno not match with selected Marking Pattern", vbCritical, "Message"
        Exit Sub
    Else
        Image1.Picture = LoadPicture(CDL.FileName)
    End If
End If



End Sub

Private Sub ww_txt_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If Left(Trim(refno_TXT), 2) = "BE" And LenB(ww_txt) > 0 Then
            y_ww = Right(ww_txt, 3)
            topx(3).TEXT = Replace(topx(3).TEXT, "YWW", y_ww)
        End If
        '20110708
        If Left(Trim(refno_TXT), 2) = "CM" Then
            '20110714 AMICCOM
            xyymmdd = Right(yy1, 2) & mm1 & dd1
            topx(2).TEXT = Trim(topInfo(2).TEXT)
            topx(3).TEXT = Trim(topInfo(3).TEXT)
            topx(2).TEXT = Replace(topx(2).TEXT, "YYWW", ww_txt)
            topx(2).TEXT = Replace(topx(2).TEXT, "YYMMDD", xyymmdd)
            topx(2).TEXT = Replace(topx(2).TEXT, "(CustLot#)", "")   '2011-08-11 refer Beh
            topx(3).TEXT = Replace(topx(3).TEXT, "YYWW", ww_txt)
            topx(3).TEXT = Replace(topx(3).TEXT, "YYMMDD", xyymmdd)
            topx(3).TEXT = Replace(topx(3).TEXT, "CustLot#)", "")
        End If
        
        'Quah 20210922, for Impinj... populate Cuslotn+Workweek to Marking.
        If Left(Trim(refno_TXT), 2) = "IJ" Then
            X_Mark = Right(Trim(txtCusLot), 5)
            topx(3).TEXT = X_Mark & ww_txt
        End If

        
        
    
    End If
End Sub

Private Sub yy1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        
        SDate = CDate(Me.mm1 + "/" + Me.dd1 + "/" + Me.yy1)
        
        'Quah 20080605 check for StartDate > 5 days
        'Quah 20140723 open for 10 days due to Hari Raya advance loading, req by Anita
        If (SDate - Date > 5) Or (SDate - Date < 0) Then
            MsgBox ("SYSTEM DOES NOT ALLOW START-DATE MORE THAN 5 DAYS !! Please Check Your Input.")
            MsgBox ("System will now set the Start-Date to the default tomorrow's date.")
            'Quah 20080605 set default date and WW
            wwdate = Format(Date + 1, "DDMMYYYY")
            dd1 = Mid(wwdate, 1, 2)
            mm1 = Mid(wwdate, 3, 2)
            yy1 = Mid(wwdate, 5, 4)
            Exit Sub
        End If
        
        Dim wwcalrs As ADODB.Recordset
        
        If Left(Trim(refno_TXT), 2) = "AS" Then
            MsgBox "??????"
            Exit Sub
            Call SanjoseWW
        
        'Quah 2013-01-09 AVT=NIKO
        ElseIf Left(Trim(refno_TXT), 2) = "NK" Or Left(Trim(refno_TXT), 2) = "AO" Then
            Call MarkSpec_NK
            
            '2017-08-23 for PDFN2X2 (PB521BX REV.AZ, PB606BX REV.BZ) automated marking line 3, req by Mon.
            'Quah add 2 more devices 2017-11-30, req by Mon
            'NIKO PB521BX REV.AZ, PB606BX REV.BZ, PB5A2BX REV.AZ, PB606BA REV.CZ
            'Quah 20171229 add for PB600BA REV.DZ
            'qUAH 20180726 CLOSE BELOW CODES, TRIGGER ON CUSLOTNO ENTER.
            
'            If InStr(TARGET_DEVICE_TXT, "PB521BX REV.AZ") > 0 Or InStr(TARGET_DEVICE_TXT, "PB606BX REV.BZ") > 0 _
'                Or InStr(TARGET_DEVICE_TXT, "PB5A2BX REV.AZ") > 0 Or InStr(TARGET_DEVICE_TXT, "PB606BA REV.CZ") > 0 _
'                Or InStr(TARGET_DEVICE_TXT, "PB600BA REV.DZ") > 0 Or InStr(TARGET_DEVICE_TXT, "PB600BA REV.BZ") Then
''''            If InStr(TARGET_DEVICE_TXT, "PB521BX REV.AZ") > 0 Or InStr(TARGET_DEVICE_TXT, "PB606BX REV.BZ") > 0 Then
'                nkyy = Left(ww_txt, 2)
'                Select Case nkyy
'                    Case "17"
'                        xnky = "H"
'                    Case "18"
'                        xnky = "I"
'                    Case "19"
'                        xnky = "J"
'                    Case "20"
'                        xnky = "K"
'                    Case "21"
'                        xnky = "L"
'                    Case "22"
'                        xnky = "M"
'                    Case "23"
'                        xnky = "N"
'                    Case "24"
'                        xnky = "O"
'                    Case "25"
'                        xnky = "P"
'                    Case "26"
'                        xnky = "Q"
'                    Case Else
'                        MsgBox "NIKO calendar year not define for these device.", vbCritical, "Message"
'                        Exit Sub
'                End Select
'
'                xnkww = Right(ww_txt, 2)
'
'                '20171218 Quah add condition for NIKO 2X2 PB606BA REV.CZ --> no need prefix A, req by Mon,KC.
'                '20171229 Quah add for PB600BA REV.DZ , req by Mon.
'                '20180402 Quah add for PB600BA REV.BZ , req by KC.
'                If InStr(TARGET_DEVICE_TXT, "PB606BA REV.CZ") > 0 Or InStr(TARGET_DEVICE_TXT, "PB600BA REV.DZ") > 0 Or InStr(TARGET_DEVICE_TXT, "PB600BA REV.BZ") > 0 Then
'                    topx(2) = xnky & xnkww
'                Else
'                    topx(2) = "A" & xnky & xnkww
'                End If
'            End If
        
        ElseIf Left(Trim(refno_TXT), 2) = "AD" Then
'            Dim WWxx, datecodexx As String
'            XX03 = yy1
'            SDate = CDate(Me.mm1 + "/" + Me.dd1 + "/" + Me.yy1)
'            datecodexx = Format(Format(SDate), "YY", vbSunday, vbFirstJan1) + Format(Format(Format(SDate), "WW", vbSunday, vbFirstJan1), "00")
'            If Right(Trim(datecodexx), 1) = "1" Or Right(Trim(datecodexx), 1) = "3" Or Right(Trim(datecodexx), 1) = "5" Or Right(Trim(datecodexx), 1) = "7" Or Right(Trim(datecodexx), 1) = "9" Then
'                 WWxx = Right(Trim(datecodexx), 2) + 1
'                Me.ww_txt = Format(Format(SDate), "YY", vbSunday, vbFirstJan1) + Format(Trim(WWxx), "00")
'            Else
'                Me.ww_txt = Format(Format(SDate), "YY", vbSunday, vbFirstJan1) + Format(Format(Format(SDate), "WW", vbSunday, vbFirstJan1), "00")
'            'If mark_spec_txt.Enabled = True Then mark_spec_txt.SetFocus
'            End If
        
            '2011-12-29 new logic, ww pull from WIPCAL
            Set wwcalrs = New ADODB.Recordset
            wwcalsql = "SELECT * FROM WIPCAL WHERE WCAL_YEAR='" & Me.yy1 & "' AND WCAL_MONTH ='" & Me.mm1 & "' AND WCAL_DAY ='" & Me.dd1 & "'AND WCAL_FACILITY ='AICS'"
            wwcalrs.Open wwcalsql, wsDB, adOpenDynamic, adLockOptimistic
            Debug.Print wwcalsql
            If Not wwcalrs.EOF Then
                   LOTYY = Trim(wwcalrs!WCAL_PLNG_YEAR)
                   LOTWW = Format(Trim(wwcalrs!WCAL_PLNG_WEEK), "00")
                   Me.ww_txt = Right(LOTYY & LOTWW, 4)
            End If
            wwcalrs.Close
            Set wwcalrs = Nothing
        
        Else
'            XX03 = yy1
'            SDate = CDate(Me.mm1 + "/" + Me.dd1 + "/" + Me.yy1)
            'ko add 20101224
            
            'Me.ww_txt = Format(Format(SDate), "YY", vbSunday, vbFirstJan1) + Format(Format(Format(SDate), "WW", vbSunday, vbFirstJan1), "00")
            'If mark_spec_txt.Enabled = True Then mark_spec_txt.SetFocus
        
            '2011-12-29 new logic, ww pull from WIPCAL
            Set wwcalrs = New ADODB.Recordset
            wwcalsql = "SELECT * FROM WIPCAL WHERE WCAL_YEAR='" & Me.yy1 & "' AND WCAL_MONTH ='" & Me.mm1 & "' AND WCAL_DAY ='" & Me.dd1 & "'AND WCAL_FACILITY ='AICS'"
            wwcalrs.Open wwcalsql, wsDB, adOpenDynamic, adLockOptimistic
            Debug.Print wwcalsql
            If Not wwcalrs.EOF Then
                   LOTYY = Trim(wwcalrs!WCAL_PLNG_YEAR)
                   LOTWW = Format(Trim(wwcalrs!WCAL_PLNG_WEEK), "00")
                   Me.ww_txt = Right(LOTYY & LOTWW, 4)
            End If
            wwcalrs.Close
            Set wwcalrs = Nothing
        
        
        End If
                    
        'Quah 2014-01-16 DIODES cuslotno= YYWW-WAFER.AS
        If Left(refno_TXT, 2) = "DT" Then
            dotx = InStr(Trim(WAFER), ".")
            If dotx > 0 Then
                txtCusLot = Trim(ww_txt) & "-" & Mid(Trim(WAFER), 1, dotx - 1) & ".AS"
            Else
                txtCusLot = "?????"
            End If
        End If
        
         'diana 2015-02-05 jaeyoung auto-marking
         If Left(Trim(refno_TXT), 2) = "JC" Then
            If InStr(topx(3), "(Cust lot)") Then
               topx(3).TEXT = Trim(txtCusLot)
            End If
            If InStr(topx(4), "YYWW") Then
               topx(4).TEXT = Trim(ww_txt)
            End If
        End If
        
         'diana 2015-02-12 mazet auto-marking (markspec SIM12)
         'diana 2015-02-27 add mazet top3 marking (markspec LC1.4)
         '20170727 MAZET change name to AMS (AY)
         If Left(Trim(refno_TXT), 2) = "MG" Or Left(Trim(refno_TXT), 2) = "AY" Then
            If InStr(topx(4), "Lotnumber") Then 'mark_spec_txt = "SIM12"
               topx(4).TEXT = Trim(txtCusLot)
            End If
            If InStr(topx(5), "YYWW") Then
               topx(5).TEXT = Trim(ww_txt)
            End If
            If InStr(topx(3), "YYWW") Then 'markspec LC1.4
               topx(3).TEXT = Trim(ww_txt)
            End If
        End If
            
        
        'Quah 20091222 alert message to change Bourns workweek during last-week of December.
        'LingLing requested Ko to set to workweek ww01 for ww53 2009.
        'For Bourns, still need to have ww53.
'        If Left(Trim(refno_TXT), 2) = "BE" Then
'            MsgBox "Please check WORKWEEK & MARKING for Bourns, if incorrect, pls amend accordingly.", vbInformation, "Message"
'        End If
        
        '20110708
        If Left(Trim(refno_TXT), 2) = "CM" Then
            '20110714 AMICCOM
            xyymmdd = Right(yy1, 2) & mm1 & dd1
            topx(2).TEXT = Trim(topInfo(2).TEXT)
            topx(3).TEXT = Trim(topInfo(3).TEXT)
            topx(2).TEXT = Replace(topx(2).TEXT, "YYWW", ww_txt)
            topx(2).TEXT = Replace(topx(2).TEXT, "YYMMDD", xyymmdd)
            topx(2).TEXT = Replace(topx(2).TEXT, "(CustLot#)", "")
            topx(3).TEXT = Replace(topx(3).TEXT, "YYWW", ww_txt)
            topx(3).TEXT = Replace(topx(3).TEXT, "YYMMDD", xyymmdd)
            topx(3).TEXT = Replace(topx(3).TEXT, "(CustLot#)", "")
        End If
        
        
        
        mark_spec_txt.SetFocus
    End If
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Private Sub SanjoseWW()
XX03 = yy1
   
If mm1 = "" Or dd1 = "" Or yy1 = "" Then
MsgBox "PLEASE KEY IN DATE!!!"
Exit Sub
End If

DATEMB = mm1 & "/" & dd1 & "/" & yy1
DATEMB = Format(DATEMB, "DD-MMM-YY")
DATEYY = Format(Date, "YY")
'-----------CHECK WW------

Dim ImpWW As ADODB.Recordset
Set ImpWW = New ADODB.Recordset

sqltext = "SELECT * FROM WWCAL_SJ WHERE WW_DATE='" & DATEMB & "'"
            ImpWW.Open sqltext, wsDB
            
            If ImpWW.EOF = False Then
            
            WWKK = ImpWW!ww_workweek
            DDKK = ImpWW!WW_DAY
            
            End If
            ImpWW.Close
            Set ImpWW = Nothing
'-------------------------
If Right(Trim(yy1), 2) = "04" Then
DATEYY = "04"
ElseIf Right(Trim(yy1), 2) = "05" Then
DATEYY = "05"
ElseIf Right(Trim(yy1), 2) = "06" Then
DATEYY = "06"
ElseIf Right(Trim(yy1), 2) = "07" Then
DATEYY = "07"
ElseIf Right(Trim(yy1), 2) = "08" Then
DATEYY = "08"
ElseIf Right(Trim(yy1), 2) = "09" Then
DATEYY = "09"
End If

If Len(WWKK) = 1 Then
WWKK = "0" & WWKK
End If
ww_txt = DATEYY & WWKK
   
If mark_spec_txt.Enabled = True Then
mark_spec_txt.SetFocus
End If
End Sub


Private Sub PULLDIE()
Set rsWAFER = New ADODB.Recordset
SQL = "SELECT * FROM AIC_LI_DUAL_DIE WHERE REFNO = '" & Trim(refno_TXT) & "' ORDER BY SEQNO desc"
rsWAFER.Open SQL, wsDB
Do While Not rsWAFER.EOF
    Set itmx = lvwDie.ListItems.Add(1, , "")
    itmx.TEXT = rsWAFER!seqno
    If Not IsNull(rsWAFER!PartNo) Then itmx.SubItems(1) = rsWAFER!PartNo
    If Not IsNull(rsWAFER!waferno) Then itmx.SubItems(2) = rsWAFER!waferno
    If Not IsNull(rsWAFER!WAFER_QTY) Then itmx.SubItems(5) = rsWAFER!WAFER_QTY
    If Not IsNull(rsWAFER!LIDIE_QTY) Then itmx.SubItems(3) = rsWAFER!LIDIE_QTY
    If Not IsNull(rsWAFER!DIE_QTY) Then itmx.SubItems(4) = rsWAFER!DIE_QTY
    rsWAFER.MoveNext
Loop
rsWAFER.Close
End Sub
Private Sub PULLDATA()
Call RESET
Call xdatar

Set rsLI = New ADODB.Recordset
SQL = "select * from AIC_LOADING_INSTRUCTION where REFNO='" & Trim(refno_TXT) & "'"
rsLI.Open SQL, wsDB
If rsLI.EOF = False Then
    xsd = Trim(rsLI!Status)
    'lblPONo = "5500000028"
    
    'Quah 20080729
    If Not IsNull(rsLI!LOAD_TIME) Then
        cboruntype = Trim(rsLI!LOAD_TIME)
    End If
    If InStr(cboruntype, "Engineering") = 0 Then
        cboruntype = "Mass Production"
    End If
    
    If xsd = "N" Then
        cmdUpdate.Enabled = True
        SAVEREC.Enabled = False
        BOMLI_VIEW.Enabled = True
        NEW_OPT.Enabled = False
        CANCEL_OPT.Enabled = True
        RELEASE_OPT.Enabled = True
        'cmdUpdData.Visible = False
    ElseIf xsd = "C" Then
        MsgBox "REF NO ALREADY CANCEL!", vbCritical, "ERROR"
        cmdUpdate.Enabled = False
        SAVEREC.Enabled = False
        BOMLI_VIEW.Enabled = False 'temp 20221025
        NEW_OPT.Enabled = False
        CANCEL_OPT.Enabled = False
        RELEASE_OPT.Enabled = False
        'cmdUpdData.Visible = False
    ElseIf xsd = "R" Then
        MsgBox "L.I. already released! Cannot update data!", vbCritical, "ERROR"
        cmdUpdate.Enabled = False
        SAVEREC.Enabled = False
        BOMLI_VIEW.Enabled = False 'temp 20221025
        NEW_OPT.Enabled = False
        CANCEL_OPT.Enabled = False
        RELEASE_OPT.Enabled = False
        'cmdUpdData.Visible = True
    End If
      
      Call PULLDIE
      
      'Check location of DOT
      If Not IsNull(rsLI!SpecChar) Then
        If Left(rsLI!SpecChar, 1) = "." Then
            chkdot(0).Value = Checked
        End If
        iCnt = 1
        Do While iCnt <= 4
            If Mid(rsLI!SpecChar, iCnt + 1, 1) = "." Then
                chkdot(iCnt).Value = Checked
            End If
            iCnt = iCnt + 1
        Loop
     End If

        
      If Not IsNull(Trim(rsLI!DEVICE_NO)) Then internal_device_no_txt = Trim(rsLI!DEVICE_NO)
      If Not IsNull(Trim(rsLI!PACKAGE__LEAD)) Then package_lead_txt = Trim(rsLI!PACKAGE__LEAD)
      bdcombo.Visible = False
      lbldeviceno.Visible = True
      If Not IsNull(Trim(rsLI!DEVICE_NO)) Then lbldeviceno = Trim(rsLI!DEVICE_NO)
      If Not IsNull(Trim(rsLI!TARGET_DEVICE)) Then TARGET_DEVICE_TXT = Trim(rsLI!TARGET_DEVICE)
      
      'Quah 20170928
      If target_device_txt1 = "" Then
          If Not IsNull(Trim(rsLI!TARGET_DEVICE)) Then TARGET_DEVICE_TXT = Trim(rsLI!TARGET_DEVICE)
      End If
      
'      If lbldeviceno = "" And TARGET_DEVICE_TXT <> "" Then  '2010-05-06
'            optTargetDevice.Value = True
'            target_device_txt1 = TARGET_DEVICE_TXT
'      End If
      
      
      TARGET_DEVICEX = Trim(TARGET_DEVICE_TXT)
      If Not IsNull(Trim(rsLI!WAFER)) Then WAFER = Trim(rsLI!WAFER)
      If Not IsNull(Trim(rsLI!WAFER)) Then wafer_lot_txt = Trim(rsLI!WAFER)
      If Not IsNull(Trim(rsLI!qty)) Then qty = Trim(rsLI!qty)
      If Not IsNull(Trim(rsLI!Status)) Then stat = Trim(rsLI!Status)
      If Not IsNull(Trim(rsLI!qty)) Then quantity_txt = Trim(rsLI!qty)
      If Not IsNull(Trim(rsLI!BD_NO)) Then
        bonding_diagram_txt = Trim(rsLI!BD_NO)
        BONDING_DIAGRAMx = Trim(rsLI!BD_NO)
      End If
      If Not IsNull(Trim(rsLI!CUSTOMER_NO)) Then customer_no_txt = Trim(rsLI!CUSTOMER_NO)
      If Not IsNull(Trim(rsLI!CUSTOMER_NAME)) Then customer_name_txt = Trim(rsLI!CUSTOMER_NAME)
      If Not IsNull(Trim(rsLI!MARKING_SPEC)) Then mark_spec_txt = Trim(rsLI!MARKING_SPEC)
      If Not IsNull(Trim(rsLI!work_week)) Then ww_txt = Trim(rsLI!work_week)
      If Not IsNull(Trim(rsLI!CUSLOTNO)) Then txtCusLot = Trim(rsLI!CUSLOTNO)
      If Not IsNull(Trim(rsLI!PO_NO)) Then lblPONo = Trim(rsLI!PO_NO)
      If Not IsNull(Trim(rsLI!CATALOGNO)) Then txtCatalogNo = Trim(rsLI!CATALOGNO)
                 
      'Quah.. 2010-10-13  check for CUSTOMER_PO base on TgtDevice, PO
      '---------------------------------------------------------------
        Dim CpoRs As ADODB.Recordset
        Set CpoRs = New ADODB.Recordset
        sqltxt = "select * from baic_customer where CUS_CODE='" & Left(Right(Trim(cboCust.TEXT), 4), 3) & "'"
        CpoRs.Open sqltxt, wsDB
        If Not CpoRs.EOF Then
            txt_pomode = Trim(CpoRs!cus_po_mode)
        Else
            txt_pomode = "???"
        End If
        CpoRs.Close
        Set CpoRs = Nothing
        If txt_pomode = "STANDARD" Then
            Set CpoRs = New ADODB.Recordset
            sqltxt = "select * from baic_customer_po where CPO_CUST_SHORTNAME='" & customer_name_txt & "' and CPO_PONO='" & lblPONo & "' and CPO_TARGETDEVICE='" & TARGET_DEVICE_TXT & "'"
            CpoRs.Open sqltxt, wsDB
            If Not CpoRs.EOF Then
                total_poqty = CpoRs!cpo_order_qty
                txt_podate = CpoRs!CPO_ORDER_YMD
                total_poqty.Locked = True
                txt_podate.Locked = True
            Else
                total_poqty.Locked = False
                txt_podate.Locked = False
            End If
            CpoRs.Close
            Set CpoRs = Nothing
        Else
            total_poqty = "N/A"
            txt_podate = "N/A"
            total_poqty.Locked = True
            txt_podate.Locked = True
        End If
      '---------------------------------------------------------------
                 
                 
                 
                 
                 
                 
      If Left(Trim(refno_TXT), 2) = "AD" Or Left(Trim(refno_TXT), 2) = "HT" Then
        target_device_txt1.Visible = True
        target_device_txt1.SetFocus
        If Not IsNull(Trim(rsLI!TARGET_DEVICE)) Then target_device_txt1 = Trim(rsLI!TARGET_DEVICE)
        If Not IsNull(Trim(rsLI!CATALOGNO)) Then CATALOGNOx = Trim(rsLI!CATALOGNO)
        optTargetDevice.Value = True
        optBD.Value = False
        'bonding_diagram_txt.Visible = False
      End If

'      XTEMP = IIf(IsNull(rsLI!PACKAGE__LEAD), "", Trim(rsLI!PACKAGE__LEAD))
      
    If Not IsNull(rsLI!PACKAGE__LEAD) Then
        XTEMP = Trim(rsLI!PACKAGE__LEAD)
    Else
        XTEMP = ""
    End If
      
      
      XTEMPLEN = Len(XTEMP)
      If XTEMPLEN > 5 Then
        dashpos = InStr(1, XTEMP, "-")      'QUAH 20090623
        ld_txt = Mid(XTEMP, dashpos + 1)    'QUAH 20090623
         package_txt = Left(XTEMP, XTEMPLEN - (XTEMPLEN - dashpos))
         If package_txt = "" Then
            package_txt = XTEMP
         End If
      ElseIf XTEMPLEN = 5 Then
          If XTEMP = "BGA14" Then
              ld_txt = Right(XTEMP, 2)
              package_txt = Left(XTEMP, XTEMPLEN - 2)
          Else
              ld_txt = Right(XTEMP, 1)
              package_txt = Left(XTEMP, XTEMPLEN - 1)
          End If
       Else
              ld_txt = ""
              package_txt = XTEMP
      End If
      cbofullpackage.TEXT = Trim(XTEMP)
      
      If Not IsNull(rsLI!TOP1) Then topx(0) = Trim(rsLI!TOP1)
      If Not IsNull(rsLI!TOP2) Then topx(1) = Trim(rsLI!TOP2)
      If Not IsNull(rsLI!TOP3) Then topx(2) = Trim(rsLI!TOP3)
      If Not IsNull(rsLI!TOP4) Then topx(3) = Trim(rsLI!TOP4)
      If Not IsNull(rsLI!TOP5) Then topx(4) = Trim(rsLI!TOP5)
      
      If Not IsNull(rsLI!TOP1) Then topx(1).Tag = Trim(rsLI!TOP1)   'Quah 20140617 to store xls-imported data.
      If Not IsNull(rsLI!TOP2) Then topx(2).Tag = Trim(rsLI!TOP2)   'Quah 2010-05-18 to store GMT xls-imported data.
      If Not IsNull(rsLI!TOP3) Then topx(3).Tag = Trim(rsLI!TOP3)   'Quah 20140617 to store xls-imported data.
      If Not IsNull(rsLI!TOP4) Then topx(4).Tag = Trim(rsLI!TOP4)   'Quah 20140617 to store xls-imported data.
      
      
      If Not IsNull(rsLI!TOP6) Then
          topx(5) = Trim(rsLI!TOP6)
          If Trim(rsLI!TOP6) = "SL449" Then
             txtSLUT = Right(Trim(rsLI!TOP6), 3)
          End If
          FoundSl = txtSLUT
      End If
      If Not IsNull(rsLI!BOTTOM1) Then bottom(0) = Trim(rsLI!BOTTOM1)
      If Not IsNull(rsLI!BOTTOM2) Then bottom(1) = Trim(rsLI!BOTTOM2)
      If Not IsNull(rsLI!BOTTOM3) Then bottom(2) = Trim(rsLI!BOTTOM3)
      If Not IsNull(rsLI!BOTTOM4) Then bottom(3) = Trim(rsLI!BOTTOM4)
      If Not IsNull(rsLI!BOTTOM5) Then bottom(4) = Trim(rsLI!BOTTOM5)
      If Not IsNull(rsLI!BOTTOM6) Then bottom(5) = Trim(rsLI!BOTTOM6)
      stat = Trim(rsLI!Status)
     ' mark_spec_txt.Enabled = False
     
    Set rsbd = New ADODB.Recordset
 '   CSQLSTRING = "SELECT * FROM WIPPRD WHERE WPRD_PROD = '" & Trim(lbldeviceno.Caption) & "' "
'AIMS
    CSQLSTRING = "SELECT * FROM BAIC_PRODMAST WHERE PDM_DEVICENO = '" & Trim(lbldeviceno.Caption) & "' "
    
    rsbd.Open CSQLSTRING, wsDB
    'Do While Not rsbd.EOF
    If rsbd.EOF = False Then
'        bd_no_txt.TEXT = Trim(rsbd!WPRD_USRDF_SMDAT_2)
        bd_no_txt.TEXT = Trim(rsbd!PDM_INTERNAL_BD)
    End If
    rsbd.Close
    Set rsbd = Nothing
    
    '2012-09-04
    If IsNull(rsLI!INTERNAL_BD) = False And Trim(rsLI!INTERNAL_BD) <> "" Then
        bd_no_txt.TEXT = Trim(rsLI!INTERNAL_BD)
    End If
    
Else
    Call xdatar
    init_header
    lbldeviceno.Visible = False
    bdcombo.Visible = True
    bdcombo.Clear
    mark_spec_txt.Enabled = True
    SAVEREC.Enabled = True
    BOMLI_VIEW.Enabled = True
    cmdUpdate.Enabled = False
    cmdUpdData.Visible = False
    NEW_OPT.Enabled = True
    CANCEL_OPT.Enabled = False
    RELEASE_OPT.Enabled = False
End If
rsLI.Close

If Left(refno_TXT, 2) = "FS" Or Left(refno_TXT, 2) = "FP" Then
    fsfom.Enabled = True
Else
    fsfom.Enabled = False
End If


'populate ori info, for baic_customer_po matching during save-click.
liori_cust = Trim(custnameselect.TEXT)
liori_po = Trim(lblPONo.TEXT)
liori_tgtdev = Trim(TARGET_DEVICE_TXT.TEXT)

End Sub


Private Sub RESET()
ROUTEx = "-"
optBD.Value = True
optTargetDevice.Value = False

bonding_diagram_txt.TEXT = vbNullString
target_device_txt1.TEXT = vbNullString
cbobonding_diagram.Clear

txtCusLot.TEXT = vbNullString
txtNoWaferLot.TEXT = vbNullString
lblPONo.TEXT = vbnulstring
dd1.TEXT = vbNullString
mm1.TEXT = vbNullString
yy1.TEXT = vbNullString
mark_spec_txt.TEXT = vbNullString
stat.Caption = vbNullString

package_txt.TEXT = vbNullString
ld_txt.TEXT = vbNullString
lbldeviceno.Caption = ""
ww_txt.TEXT = vbNullString
lblFroute.Caption = vbNullString
lblLroute.Caption = vbNullString
lblOpr1.Caption = vbNullString
lblTUT.Caption = vbNullString

bdcombo.Clear
cboruntype.TEXT = ""

txt_pomode = ""
total_poqty = ""
txt_podate = ""


NEW_OPT.Value = False
CANCEL_OPT.Value = False
RELEASE_OPT.Value = False

mark_spec_txt.Enabled = True
For iCnt = 0 To 5
    topx(iCnt).TEXT = vbNullString
    topInfo(iCnt).TEXT = vbNullString
    bottom(iCnt).TEXT = vbNullString
    botInfo(iCnt).TEXT = vbNullString
Next iCnt

internal_device_no_txt.TEXT = vbNullString
package_lead_txt.TEXT = vbNullString
WAFER.TEXT = vbNullString
qty.TEXT = vbNullString

TARGET_DEVICE_TXT.TEXT = vbNullString
bd_no_txt.TEXT = vbNullString
bd_no_txt1.TEXT = vbNullString

'Unload FRM_MULTIWFR

Text1.TEXT = vbNullString
Text2.TEXT = vbNullString
Text3.TEXT = vbNullString
Text4.TEXT = vbNullString
Text5.TEXT = vbNullString
Text6.TEXT = vbNullString

txtCatalogNo = vbNullString
CATALOGNOx = vbNullString

iCnt = 0
Do While iCnt <= 4
    chkdot(iCnt).Value = Unchecked
    iCnt = iCnt + 1
Loop

lvwDie.ListItems.Clear

SAVEREC.Enabled = True
BOMLI_VIEW.Enabled = True
cmdUpdate.Enabled = False
      
    wwdate = Format(Date + 1, "DDMMYYYY")
    dd1 = Mid(wwdate, 1, 2)
    mm1 = Mid(wwdate, 3, 2)
    yy1 = Mid(wwdate, 5, 4)
    SDate = CDate(Me.mm1 + "/" + Me.dd1 + "/" + Me.yy1)

'    Me.ww_txt = Format(Format(SDate), "YY", vbSunday, vbFirstJan1) + Format(Format(Format(SDate), "WW", vbSunday, vbFirstJan1), "00")

    '2011-12-29 new logic, ww pull from WIPCAL
    Dim wwcalrs As ADODB.Recordset
    Set wwcalrs = New ADODB.Recordset
    wwcalsql = "SELECT * FROM WIPCAL WHERE WCAL_YEAR='" & Me.yy1 & "' AND WCAL_MONTH ='" & Me.mm1 & "' AND WCAL_DAY ='" & Me.dd1 & "'AND WCAL_FACILITY ='AICS'"
    wwcalrs.Open wwcalsql, wsDB, adOpenDynamic, adLockOptimistic
    Debug.Print wwcalsql
    If Not wwcalrs.EOF Then
           LOTYY = Trim(wwcalrs!WCAL_PLNG_YEAR)
           LOTWW = Format(Trim(wwcalrs!WCAL_PLNG_WEEK), "00")
           Me.ww_txt = Right(LOTYY & LOTWW, 4)
    End If
    wwcalrs.Close
    Set wwcalrs = Nothing
        
        
        'Quah 20091222 alert message to change Bourns workweek during last-week of December.
        'LingLing requested Ko to set to workweek ww01 for ww53 2009.
        'For Bourns, still need to have ww53.
'        If Left(Trim(refno_TXT), 2) = "BE" Then
'            MsgBox "Please check WORKWEEK & MARKING for Bourns, if incorrect, pls amend accordingly.", vbInformation, "Message"
'        End If
    
    'Quah 20090225 Bourns auto marking
    If Left(Trim(refno_TXT), 2) = "BE" And LenB(ww_txt) > 0 And InStr(topx(3).TEXT, "YWW") > 0 Then
'        topx(3).TEXT = Right(ww_txt, 3)
        
        'Quah 2014-05-12
         y_ww = Right(ww_txt, 3)
         topx(3).TEXT = Replace(topx(3).TEXT, "YWW", y_ww)
    End If
     
End Sub

Private Sub ns_marking()

'Quah 20080606 for NS, the marking is auto-pulled from NS text file.
die_run_txt.TEXT = ""
If Left(Trim(refno_TXT), 2) = "NN" And LenB(txtCusLot) > 0 Then
    topx(0).TEXT = ""
    topx(1).TEXT = ""
    topx(2).TEXT = ""
    topx(3).TEXT = ""
    topx(4).TEXT = ""
    topx(5).TEXT = ""
    
    Dim ns_file As String
'    ns_file = "\\aicwksvr2\NSEM ePT\PT\PT_" & Trim(txtCusLot.TEXT) & ".TXT"
    '20180503
    ns_file = "\\aicwksvr2016\NSEM_ePT\PT_" & Trim(txtCusLot.TEXT) & ".TXT"
    
    Dim DoesFileExist As Boolean
    DoesFileExist = FileExists(ns_file)

    If DoesFileExist = False Then
'        MsgBox "Error ePT File : '" & ns_file & "' !!!" & Chr(13) & "Please check Custlotno."
'        Exit Sub
        '2013-03-05 STLIM req to block LI if no softcopy marking file is found.
'        '2013-03-15 ZUL ADD TO SEARCH IN BAIC_ALARM_TRANX

        Set rs2x = New ADODB.Recordset
        SSQL2 = "SELECT * FROM BAIC_ALARM_TRANX WHERE ALM_LOTNO= '" & Trim(txtCusLot) & "' AND ALM_REC_TYPE='NSMK'"
        Debug.Print SSQL2
        rs2x.Open SSQL2, wsDB
        If Not rs2x.EOF Then
            topx(0).Enabled = True
            topx(1).Enabled = True
            topx(2).Enabled = True
            topx(3).Enabled = True
            topx(4).Enabled = True
            topx(5).Enabled = True
            MsgBox "ePT file not found, Please Proceed for Manual Marking.", vbInformation, "MESSAGE"
            Exit Sub
        Else
            MsgBox "Error ePT File : '" & ns_file & "' !!!" & vbCrLf & "Please check Custlotno.", vbCritical, "ERROR"
            End
        End If
        
    End If
    
    Dim xfound
    Dim templine
    xfound = 0
    Open ns_file For Input As 1
     Do While Not EOF(1)
         Line Input #1, templine
         
         'Quah 20090416 get DIERUN#
         If Left(templine, 9) = "DIE RUN #" Then
             die_run_txt = Trim(Mid(templine, 18, 28))
         End If
         
        If xfound > 0 Then
            xfound = xfound + 1
         End If
         If xfound > 3 Then
            If Trim(templine) = "" Then
                xfound = 0
            End If
         End If
         If Mid(templine, 6, 7) = "TOPMARK" Then
            xfound = 1
         End If
         
         If xfound = 3 And Trim(templine) = "" Then     'Quah 20090805 need to read additional line because of suspected carriage-return.
                xfound = 2  'reset counter to 2
         End If
         
         If xfound = 3 Then
            'Quah 20090115 NS Logo must base on $ on first line, If not $ sign, then blank logo. Refer to LingLing/Anita/Leong (wrong marking issue).
            'Quah 20090121 revised condition: $N = NS Logo, $X = Special Logo, requested by LingLing
            'Quah 20090206 revised condition: If $ but not X or N, then block.
            If Left(Trim(templine), 2) = "$N" Then
                topx(0).TEXT = "{NS LOGO}"
            ElseIf Left(Trim(templine), 2) = "$X" Then
                topx(0).TEXT = "{SPECIAL LOGO}"
                MsgBox ("Please check: $X = SPECIAL LOGO")
            ElseIf Left(Trim(templine), 1) = "$" Then           'not $N, not $X
                topx(0).TEXT = "{UNKNOWN}"
                MsgBox ("Unable to proceed. Unknown $ Marking. Kindly verify with NS !!!")
                Trim(txtCusLot.TEXT) = ""
                Exit Sub
            Else
                topx(0).TEXT = "{}"
                MsgBox ("Please check: NO NS LOGO")
            End If
            'Quah 20090115 - end modification.
            
            If Left(Trim(templine), 1) = "$" Then   'Quah 20080213 if first letter is $, then start from 3rd letter.
                topx(1).TEXT = Mid(Trim(templine), 3)
            Else
                topx(1).TEXT = Trim(templine)
            End If
         End If
         If xfound = 4 Then
            topx(2).TEXT = Trim(templine)
         End If
         If xfound = 5 Then
            topx(3).TEXT = Trim(templine)
         End If
         If xfound = 6 Then
            topx(4).TEXT = Trim(templine)
         End If
         If xfound = 7 Then
'            topx(5).TEXT = Trim(templine)
             topx(4).TEXT = topx(4).TEXT & ", " & Trim(templine)    '2011-11-11 for IPS
         End If
    Loop
    Close #1
    
    'Quah 20080613 Pull Package, Start Qty and DeviceID
    xpullqty = 0
    xpulldev = ""
    Open ns_file For Input As 1
     Do While Not EOF(1)
        Line Input #1, templine
         If Mid(templine, 1, 10) = "PACKAGE ID" Then
            xpullpackage = Trim(Mid(templine, 19, 4))
            If xpullpackage = "PDIP" Then xpullpackage = "IPS"
            xpull_lead = Trim(Str(CDbl(Trim(Mid(templine, 23, 3)))))
         End If
         
         'Quah 2014-02-11 change from position 46 to 36, 63 to 53.
         If Mid(templine, 36, 9) = "START QTY" Then
            xpullqty = CDbl(Trim(Mid(templine, 53, 6)))
            Me.Text3 = Me.txtCusLot
            Me.Text4 = xpullqty
            Me.Text5 = xpullqty
         End If
         
         'Quah 2014-02-11 new PT format change position from 46 to 36, 63 to 53.
         If Mid(templine, 36, 9) = "DEVICE ID" Then
            lastpos = InStr(53, templine, " ") - 53
            
            xpulldev = Trim(Mid(templine, 53, lastpos))
            'Quah 2012-11-09 NS-TI Integration: Targetdevice need to exclude /NOPB
            'Quah 2013-01-29 LL req to remove below logic, due to BD already updated.
''            xnopb = InStr(xpulldev, "/")
''            If xnopb > 0 Then
''                xpulldev = Left(xpulldev, xnopb - 1)
''            End If
            
         End If
    Loop
    Close #1
    'oct2022
            topx(0).Enabled = False
            topx(1).Enabled = False
            topx(2).Enabled = False
            topx(3).Enabled = False
            topx(4).Enabled = False
            topx(5).Enabled = False
    
    If lbltestonly <> "TEST ONLY" Then  '2010-05-17
        Me.package_txt = xpullpackage
        Me.ld_txt = xpull_lead
        
        Me.optTargetDevice = True
        Me.target_device_txt1 = xpulldev
        
        Dim chkAllpkg As ADODB.Recordset
        Set chkAllpkg = New ADODB.Recordset
        sqltext = "SELECT * FROM BAIC_PRODMAST WHERE PDM_PACKAGELEAD = '" & Trim(package_txt) & Trim(ld_txt) & "'"
        
        chkAllpkg.Open sqltext, wsDB
        If chkAllpkg.EOF = True Or package_txt = "NA" Then
            package_txt = vbNullString
            ld_txt = vbNullString
            MsgBox "PackageLead is not setup in Database. Check with Planner!!!"
        End If
        chkAllpkg.Close
        Set chkAllpkg = Nothing
    End If
End If
End Sub

Private Function FileExists(FileName As String) As Boolean
    On Error GoTo ErrorHandler
    ' get the attributes and ensure that it isn't a directory
    FileExists = (GetAttr(FileName) And vbDirectory) = 0
ErrorHandler:
    ' if an error occurs, this function returns False
End Function