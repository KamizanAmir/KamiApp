Public PubSeqHeader
Public PubTranx
Public Poll_SEQ(50)
Public Poll_OPER(50)
Public Poll_OPERNAME(50)
Public Poll_TRANS_STATION(50)
Public Poll_APT_PRINT(50)
Public Poll_TESTPT_PRINT(50)
Public Poll_AREA(50)
Public Poll_AREA_STATUS(50)
Public Poll_PREV_OPER(50)
Public Poll_NEXT_OPER(50)
Public Poll_EDCATTACH(50)
Public PubLong
Public RecTmp1(500)
Public RecTmp2(500)
Public RecTmp3(500)
Public RecTmp4(500)
Public RecTmp5(500)
Public RecTmp6(500)
Public RecTmp7(500)
Public MAXVAL
Public PubTargetdevice
Public HoldFlag
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public login_id, pchostname, user_dept, MesSysName, MesSysVer As String
Public xserverdate, xservertime, xshifttype, xproddate
Public wsDB As ADODB.Connection
Public wsDB_CLS As ADODB.Connection
Public txtQTY(60), cmdMrlt(60), txtLOT(60), runnumber, txtSEQ(60)
Public Rej_combo(50), lblParmXLastseq(35), edcrej_combo(16), lblSeqX(34)
Public Oper_lbl(40), operdesc_lbl(40), operseq_lbl(40), opertype_lbl(40), opergrp_lbl(40)
Public oper_T_Station(40), oper_Grp_Station(40), oper_input_Station(40)
Public atmeloricuslot
Public MARKSTRING, bMARKSTRING As String
Public Trteseq(20), Trteoper(20), Trtedesc(20)
Public PubRejCode(20)
Public PubRejLoss(20)
Public gAPT_reprint_flag
Public CoreVersion As Double
Public Sec_AuthArea As String 
Public Sec_AuthCat As String  
Public Sec_AuthFunc As String    
Public Sec_EmpClass As String 
Public Sec_EmpGroup As String 
Public Login_Name As String
Public MsgBoxResponse As Integer
Public MsgBoxDefaultButton As Integer
Private Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetParent& Lib "user32" (ByVal hwnd As Long)
Dim sPattern As String, hFind As Long
Public errTransaction As String
Private Function EnumWinProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
  Dim k As Long, sName As String
  If GetParent(hwnd) = 0 Then
     sName = Space$(128)
     k = GetWindowText(hwnd, sName, 128)
     If k > 0 Then
        sName = Left$(sName, k)
        If lParam = 0 Then sName = UCase(sName)
        If sName Like sPattern Then
           hFind = hwnd
           EnumWinProc = 0
           Exit Function
        End If
     End If
  End If
  EnumWinProc = 1
End Function
Public Function FindWindow(sWild As String, Optional bMatchCase As Boolean = True) As Long
  sPattern = sWild
  If Not bMatchCase Then sPattern = UCase(sPattern)
  EnumWindows AddressOf EnumWinProc, bMatchCase
  FindWindow = hFind
End Function
Public Sub svrdatetime(xserverdate, xservertime, xshifttype, xproddate)
    Dim svr_dateRs As ADODB.Recordset
    Dim sqldate As String
    Set svr_dateRs = New ADODB.Recordset
    sqldate = "SELECT GETDATE() svrdate FROM DUAL"
    Debug.Print sqldate
    svr_dateRs.Open sqldate, wsDB, adOpenForwardOnly, adLockReadOnly
    xserverdate = svr_dateRs.Fields("svrdate").Value
    Set svr_dateRs = New ADODB.Recordset
    sqldate = "select CONVERT(varchar(20),getdate(),114) svrtime from DUAL"
    svr_dateRs.Open sqldate, wsDB, adOpenForwardOnly, adLockReadOnly
    xservertime = svr_dateRs.Fields("svrtime").Value 
    Dim timehh As Integer
    timehh = Val(Format(xservertime, "HH"))
   If timehh >= 19 Or timehh < 7 Then
      xshifttype = "1"
      If timehh >= 19 Then
         xproddate = Format(xserverdate + 1, "YYYYMMDD")
      Else
         xproddate = Format(xserverdate, "YYYYMMDD")
      End If
   Else
      xshifttype = "2"
      xproddate = Format(xserverdate, "YYYYMMDD")
   End If
End Sub

Public Sub userPermission(login_id)
        Dim i, j
        If Sec_EmpClass = "ADM" Then
            For Each i In Main_Menu.Toolbar1.Buttons
                For Each j In i.ButtonMenus
                    j.Enabled = True
                Next
            Next
        Else
            For Each i In Main_Menu.Toolbar1.Buttons
                For Each j In i.ButtonMenus
                    If InStr(Sec_AuthFunc, j.Key) Then
                        j.Enabled = True
                    End If
                Next
            Next
        End If
        Main_Menu.Toolbar1.Buttons(8).Enabled = True
End Sub
Public Sub exitlogic(exittype As Integer)
Call svrdatetime(xserverdate, xservertime, xshifttype, xproddate)
Dim hostexitRS As ADODB.Recordset
Set hostexitRS = New ADODB.Recordset
Dim hostexitRS2 As ADODB.Recordset
Set hostexitRS2 = New ADODB.Recordset
Dim sqlstring, lastid As String
If exittype = 1 Then
    sqlstring = "select * from baic_comtbl where TBL_REC_TYPE='LOGN' and TBL_KEY_A20='" & pchostname & "' and TBL_KEY_A12='AIMS'"
    hostexitRS.Open sqlstring, wsDB
    If Not hostexitRS.EOF Then
        sqlstring = "update baic_comtbl set TBL_DATA_A2='" & Format(xserverdate, "YYYYMMDD-HH:MM") & "', TBL_REMARK='LOGOFF' where TBL_REC_TYPE='LOGN' and TBL_KEY_A20='" & pchostname & "' and TBL_KEY_A12='AIMS'"
        Debug.Print sqlstring
        hostexitRS2.Open sqlstring, wsDB
    End If
    hostexitRS.close
    End
ElseIf exittype = 2 Then
        sqlstring = "select * from baic_comtbl where TBL_REC_TYPE='LOGN' and TBL_KEY_A20='" & pchostname & "' and TBL_KEY_A12='AIMS'"
        hostexitRS.Open sqlstring, wsDB
        If Not hostexitRS.EOF Then
            sqlstring = "update baic_comtbl set TBL_DATA_A2='" & Format(xserverdate, "YYYYMMDD-HH:MM") & "', TBL_REMARK='LOGOFF' where TBL_REC_TYPE='LOGN' and TBL_KEY_A20='" & pchostname & "' and TBL_KEY_A12='AIMS'"
            Debug.Print sqlstring
            hostexitRS2.Open sqlstring, wsDB
        End If
        hostexitRS.close
        End
Else    
    If MsgBox("Logoff System?", vbYesNo, "Message") = 6 Then
        sqlstring = "select * from baic_comtbl where TBL_REC_TYPE='LOGN' and TBL_KEY_A20='" & pchostname & "' and TBL_KEY_A12='AIMS'"
        hostexitRS.Open sqlstring, wsDB
        If Not hostexitRS.EOF Then
            sqlstring = "update baic_comtbl set TBL_DATA_A2='" & Format(xserverdate, "YYYYMMDD-HH:MM") & "', TBL_REMARK='LOGOFF' where TBL_REC_TYPE='LOGN' and TBL_KEY_A20='" & pchostname & "' and TBL_KEY_A12='AIMS'"
            Debug.Print sqlstring
            hostexitRS2.Open sqlstring, wsDB
        End If
        hostexitRS.close
        Main_Menu.Caption = MesSysName
        Call DisableAllMenu
        User_Login.Show
    End If
End If
End Sub
Public Sub DisableAllMenu()
        Dim i, j 
        For Each i In Main_Menu.Toolbar1.Buttons
            For Each j In i.ButtonMenus
                j.Enabled = False
            Next
        Next
        Unload Main_BackGround
End Sub
Public Function Sec_AuthOper(ByVal strOper As String) As Boolean
    If Sec_EmpClass = "ADM" Then
        Sec_AuthOper = True
    Else
        Dim rsSelect As New ADODB.Recordset
        Dim sqlSelect As String
        sqlSelect = "select * from BAIC_COMTBL " & _
                    "where TBL_REC_TYPE = 'OPR' " & _
                        "and TBL_KEY_9 = " & Trim(strOper) & " " & _
                    "order by TBL_KEY_9"
                    Debug.Print sqlSelect
        With rsSelect
            .Open sqlSelect, wsDB
                If .EOF Then
                MsgBox "Please request Technical Support.", _
                    vbCritical, "Unknown operation"
                Sec_AuthOper = False
            ElseIf Sec_EmpClass = "ENG" Then
                If InStr(Sec_AuthArea, Trim(.Fields("TBL_KEY_A20").Value)) > 0 Then
                    Sec_AuthOper = True
                Else
                    Sec_AuthOper = False
                    MsgBox "You are not authorized to perform " & vbCrLf & "operation [" & _
                    strOper & "] in area [" & Trim(.Fields("TBL_KEY_A20").Value) & "]", _
                    vbCritical, "Invalid Authorization"
                End If
            ElseIf Trim(.Fields("TBL_KEY_A12").Value) <> Sec_AuthCat Or _
                    InStr(Sec_AuthArea, Trim(.Fields("TBL_KEY_A20").Value)) < 1 Then
                    MsgBox "Your user category [" & Sec_AuthCat & "] " & vbCrLf & _
                    "are not authorized to perform " & vbCrLf & "operation [" & _
                    strOper & "] in area [" & Trim(.Fields("TBL_KEY_A20").Value) & "]", _
                    vbCritical, "Invalid user category and/or area"
                Sec_AuthOper = False

            Else
                Sec_AuthOper = True
            End If
        End With
    End If
End Function
Public Function funcSubTotal(excelApp As Excel.Application, rngTotal As Range, strStop As String) As Double
rngTotal.Worksheet.Activate
rngTotal.Offset(-1, 0).Activate
Do While Trim(excelApp.ActiveCell.Value) <> strStop
    rngTotal.Value = Val(rngTotal.Value) + Val(excelApp.ActiveCell.Value)
    excelApp.ActiveCell.Offset(-1, 0).Activate
Loop
funcSubTotal = Val(rngTotal.Value)
End Function
Public Function chkupdate()
    Call svrdatetime(xserverdate, xservertime, xshifttype, xproddate)
    Dim rsData As New ADODB.Recordset
    Dim sqlData As String
    sqlData = "select cast(tbl_data_a1 as datetime) TIME_BEGIN, cast(tbl_data_a2 as datetime) TIME_END from BAIC_COMTBL where TBL_REC_TYPE = 'AIC' and TBL_KEY_A20 = 'SHUTDOWN'"
    rsData.Open sqlData, wsDB
    If xserverdate > rsData!TIME_BEGIN And xserverdate < rsData!TIME_END Then
        MsgBox "AIMS is under maintenance." & vbCrLf & _
                "This program will close now.", vbExclamation, "System Locked"
        End
    End If
    rsData.close
    sqlData = "select TBL_KEY_9 VERSION, TBL_DATA_A3 FILEPATH from BAIC_COMTBL where TBL_REC_TYPE = 'DPLY' and TBL_KEY_A20 = 'AIMS'"
    rsData.Open sqlData, wsDB
    If rsData.EOF Then
        MsgBox "Unable to retrieve update information, please request technical support.", vbCritical, "AIMS Updater"
        rsData.close
        End
    End If
    If Val(Trim(rsData!Version)) > CoreVersion Then
        Shell Left(Trim(rsData!filepath), InStrRev(Trim(rsData!filepath), "\") - 1) & "\Updater.exe", vbNormalFocus
        rsData.close
        End
    End If
    rsData.close
End Function
Public Function funcNextTabOrder(KeyAscii As Integer)
If KeyAscii = 13 Then
   Sendkeys "{TAB}"
End If
End Function
Public Function funcHighlightText(objControl As Control)
On Error GoTo ErrHandler
    objControl.SelStart = 0
    objControl.SelLength = Len(objControl.Text)
On Error GoTo 0
Exit Function
ErrHandler:
    MsgBox "Error occured when control passed into [funcHighlightText], please check control's properties.", vbCritical, "Code Error"
    Resume Next
End Function
Public Function funcUCaseText(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Function
Public Function funcComboDropDown()
    Sendkeys "{F4}"
End Function
Private Function Sendkeys(Text$, Optional wait As Boolean)
    CreateObject("WScript.Shell").Sendkeys Text$, wait
End Function
Public Function funcFillList(strSQL As String, ctlTarget As Control)
Dim rsData As New ADODB.Recordset
Dim strTemp As String
Dim objField As Field
rsData.Open strSQL, wsDB
If Not rsData.EOF Then
    Do While Not rsData.EOF
        strTemp = ""
        For Each objField In rsData.Fields
            strTemp = strTemp & Trim(objField.Value) & " | "
        Next
        strTemp = Left(strTemp, Len(strTemp) - 3)
                
        On Error GoTo ErrHandler
            ctlTarget.AddItem strTemp
        On Error GoTo 0
        rsData.MoveNext
    Loop
Else
End If
rsData.close
Exit Function
ErrHandler:
    MsgBox "Error occured when control passed into [funcFillList], please check control's properties.", vbCritical, "Code Error"
    Exit Function
End Function
Public Function funcBuildField(wsTarget As Worksheet, colTarget As Collection, colKey As Collection, StartRow As Long, StartCol As Long) As Range
Dim curRow As Long 
Dim curCol As Long 
Dim curInd As Long 
Dim maxcol As Long 
Dim mergeStart, mergeEnd As Integer
With wsTarget
curRow = StartRow
curCol = StartCol
curInd = 0
maxcol = StartCol
mergeStart = -1 
mergeEnd = 0
For curInd = 1 To colTarget.Count
    If colTarget(curInd) = "#NEW ROW#" Then
        If curCol > maxcol Then
            maxcol = curCol - 1
        End If
        curCol = StartCol - 1
        curRow = curRow + 1
    ElseIf colTarget(curInd) = "#MERGE#" Then
        If mergeStart = -1 Then 
            mergeStart = curCol - 1
            mergeEnd = curCol
        Else 
            mergeEnd = curCol
        End If
        .Cells(curRow, curCol).Value = ""
    Else
        If mergeStart <> -1 Then 
            .Range(.Cells(curRow, mergeStart), .Cells(curRow, mergeEnd)).Merge
            mergeStart = -1
        End If
        .Cells(curRow, curCol).Value = Trim(colTarget(curInd))
    End If
    If colKey(curInd) <> " " Then
        On Error GoTo ErrHandler
            colTarget.Add curCol, colKey(curInd)
        On Error GoTo 0
    End If
    curCol = curCol + 1
Next
End With
If curCol > maxcol Then
    maxcol = curCol - 1
End If
Set funcBuildField = wsTarget.Cells(curRow, maxcol)
Exit Function
ErrHandler:
    MsgBox "Internal data error, please request technical support.", vbCritical, "Duplicate key"
End Function
Public Function funcAddField(strCaption As String, colTarget As Collection, colKey As Collection, Optional strKey As String)
    If Trim(strKey) = "" Then
        colKey.Add " "
    ElseIf strKey = "##SAME##" Then
        colKey.Add strCaption
    Else
        colKey.Add strKey
    End If
    colTarget.Add strCaption
End Function
Public Function MsgBox(strMsg As String, Optional intParam As Integer, Optional strTitle As String) As Integer
If IsMissing(intParam) Then
    intParam = 0
End If
If IsMissing(strTitle) Then
    strTitle = ""
End If
frmMsg.lblMsg.Caption = strMsg
frmMsg.txtMsg.Height = frmMsg.lblMsg.Height
frmMsg.txtMsg.Width = frmMsg.lblMsg.Width
If frmMsg.txtMsg.Width < 3500 Then
    frmMsg.txtMsg.Width = 3500
End If
frmMsg.txtMsg.Text = strMsg
frmMsg.Caption = strTitle
If (intParam And CInt(15)) = 0 Then
    frmMsg.cmdNo.Visible = False
    MsgBoxDefaultButton = 0
    frmMsg.cmdYes.Caption = "OK"
Else
    frmMsg.cmdNo.Visible = True
    frmMsg.cmdYes.Caption = "YES"
    If (intParam And CInt(256)) = 0 Then
        MsgBoxDefaultButton = 0
    Else
        MsgBoxDefaultButton = 1
    End If
End If
If (intParam And CInt(64)) > 0 Then
    Set frmMsg.imgMsg.Picture = frmMsg.imgLstMsg.ListImages("Information").Picture
ElseIf (intParam And CInt(48)) = 48 Then
    Set frmMsg.imgMsg.Picture = frmMsg.imgLstMsg.ListImages("Exclamation").Picture
ElseIf (intParam And CInt(32)) = 32 Then
    Set frmMsg.imgMsg.Picture = frmMsg.imgLstMsg.ListImages("Question").Picture
ElseIf (intParam And CInt(16)) = 16 Then
    Set frmMsg.imgMsg.Picture = frmMsg.imgLstMsg.ListImages("Critical").Picture
Else
    Set frmMsg.imgMsg.Picture = Nothing
End If
frmMsg.Show vbModal
MsgBox = MsgBoxResponse
End Function
Public Function funcVerifyCombo(cboInput As Control)
Dim lngIndex As Long
On Error GoTo ErrHandler
If cboInput.ListIndex > -1 Then
    Exit Function
End If
For lngIndex = 0 To cboInput.ListCount - 1
    If Trim(cboInput.List(lngIndex)) = Trim(cboInput.Text) Then
        cboInput.ListIndex = lngIndex
        Exit For
    End If
Next
On Error GoTo 0
Exit Function
ErrHandler:
    MsgBox "The control passed may not be a combo box.", vbCritical, "Error at funcVerifyCombo"
    Exit Function
End Function
Public Sub UpdateInvoiceSummary()
Dim sqltext As String
Dim InsSum As ADODB.Recordset
Set InsSum = New ADODB.Recordset
wsDB.BeginTrans
wsDB.Execute "update T1 Set T1.ITEM_NUM_3 = Null from BAIC_INVOICE_SUMMARY T0, AIC_INVOICE_HEADER_B T1 Where T0.INS_INVOICENO = T1.INVOICENO and T1.STATUS = 'CANCEL'"
wsDB.Execute "delete T0 from BAIC_INVOICE_SUMMARY T0, AIC_INVOICE_HEADER_B T1 Where T0.INS_INVOICENO = T1.INVOICENO and (T1.STATUS = 'CANCEL' or T1.ITEM_NUM_3 is null)"
sqltext = "INSERT INTO BAIC_INVOICE_SUMMARY " & _
"SELECT " & _
"YY.INVOICENO,YY.DTL_NO,YY.LOT_NO,YY.TARGETDEVICE,YY.CUSTOMER_NAME,RTRIM(XX.SHIP_TYPE),YY.BILLING_DATE,'X' XX,SUBSTRING(XX.CHARGE_STATUS,1,3),YY.CUSTOMER_PO, " & _
"YY.SHIP_QTY,YY.AMOUNT,YY.CREATEDATE FROM " & _
"(SELECT INVOICENO,CHARGE_STATUS,SUBSTRING(SHIP_TYPE,1,10) SHIP_TYPE FROM AIC_INVOICE_HEADER_B WHERE ITEM_NUM_3 IS NULL AND (STATUS='C' OR STATUS='CANCEL')) XX, " & _
"(SELECT A.INVOICENO,A.DTL_NO,A.LOT_NO,B.CUSTOMER_NAME,B.BILLING_DATE,'X' XX,'C' CC,A.CUSTOMER_PO, " & _
"A.SHIP_QTY,ASSY_AMOUNT+TEST_AMOUNT+TNR_AMOUNT+INDEX_TEST_TIME_AMOUNT1+INDEX_TEST_TIME_AMOUNT2 AMOUNT, GETDATE() CREATEDATE,B.TARGETDEVICE " & _
"FROM AIC_INVOICE_DETAIL_DATA A, AIC_INVOICE_DETAIL_HEADER B " & _
"Where a.LOT_NO = B.LOT_NO " & _
"AND A.INVOICENO=B.INVOICENO " & _
"AND A.CANCEL_FLAG='N' " & _
"AND B.BILLING_DATE > 20100101 AND DTL_NO NOT LIKE 'GM%X%') YY " & _
"Where " & _
"XX.INVOICENO = YY.INVOICENO"
Debug.Print sqltext
wsDB.Execute sqltext
Dim InsSum2 As ADODB.Recordset
Set InsSum2 = New ADODB.Recordset
sqltext = "UPDATE BAIC_INVOICE_SUMMARY SET INS_INVOICE_TYPE='GOOD' WHERE INS_INVOICENO NOT LIKE '%REJ%' AND INS_INVOICE_TYPE='X'"
wsDB.Execute sqltext
Dim InsSum3 As ADODB.Recordset
Set InsSum3 = New ADODB.Recordset
sqltext = "UPDATE BAIC_INVOICE_SUMMARY SET INS_INVOICE_TYPE='REJ' WHERE INS_INVOICENO LIKE '%REJ%' AND INS_INVOICE_TYPE='X'"
wsDB.Execute sqltext
Dim InsSum4 As ADODB.Recordset
Set InsSum4 = New ADODB.Recordset
sqltext = "UPDATE AIC_INVOICE_HEADER_B SET ITEM_NUM_3=CONVERT(varchar(8), getdate(), 112) WHERE ITEM_NUM_3 IS NULL AND (STATUS='C' OR STATUS='CANCEL')"
wsDB.Execute sqltext
sqltext = "delete T0 from baic_invoice_summary T0, baic_invoice_header T1 where T0.ins_invoiceno = T1.inh_invoiceno and T1.inh_summarized_ymd is null"
wsDB.Execute sqltext
sqltext = "insert into baic_invoice_summary " & _
            " select INH_INVOICENO, PLD_DTLNO, PLD_LOTNO, LTM_TARGETDEVICE, INH_CUSNAME, INH_MFG_FLOW, CAST(convert(varchar(10),INH_INVORI_YMD, 112) as int), case when INH_LOT_TYPE = 'R' then 'REJ' else 'GOOD' end, case when INH_LOT_TYPE = 'R' then 'REJ' else case when INH_FOC = 'N' then 'CHG' else 'FOC' end end, PLD_CUSTOMER_PONO, PLD_QTY, null, GETDATE()" & _
            " From BAIC_INVOICE_HEADER, BAIC_PL_BUNDLE_HEADER, BAIC_PL_BUNDLE_DATA, BAIC_LOTMAST" & _
            " Where INH_INVOICENO = PLB_INVOICENO" & _
                    " and PLB_BUNDLENO = PLD_BUNDLENO" & _
                    " and PLD_LOTNO = LTM_LOTNO" & _
                    " and INH_STATUS = 'SHIPPED'" & _
                    " and inh_summarized_ymd is null"
wsDB.Execute sqltext                   
sqltext = "update baic_invoice_header set inh_summarized_ymd = getdate() where inh_summarized_ymd is null and inh_status = 'SHIPPED'"
wsDB.Execute sqltext
wsDB.CommitTrans                   
End Sub
Public Function funcRollback()
wsDB.RollbackTrans
MsgBox "Error #" & Err.Number & " @ " & errTransaction & vbCrLf & _
        "Source: " & Err.Source & vbCrLf & _
        Err.Description, vbCritical, "Please inform IT Department!"
End
End Function
Public Function funcRollback_CLS()
wsDB_CLS.RollbackTrans
MsgBox "Error #" & Err.Number & " @ " & errTransaction & vbCrLf & _
        "Source: " & Err.Source & vbCrLf & _
        Err.Description, vbCritical, "Please inform IT Department!"
End
End Function