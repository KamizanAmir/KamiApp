Private Sub Form_Load()

'********************************
If Day(Date) = 28 Or Day(Date) = 29 Or Day(Date) = 30 Or Day(Date) = 31 Then
    planmonth = Month(Date) + 1
    fromdate = Format(Date, "28-MMM-YYYY 07:00")
Else
    planmonth = Month(Date)
    lastmth = Date - Day(Date) - 1
    fromdate = Format(lastmth, "28-MMM-YYYY 07:00")
End If
'planmonth = 3
'fromdate = "28-Feb-2022 07:00"

todate = Format(Date, "DD-MMM-YYYY 06:59")
rptdate = Format(Now(), "YYYYMMDD-HHMM")
'********************************


Dim wsDB As ADODB.Connection
Set wsDB = New ADODB.Connection
wsDB.ConnectionString = "PROVIDER=MSDASQL;dsn=AICSSQLDB;uid=ITAPP;pwd=APP;"
wsDB.CursorLocation = adUseServer
wsDB.Open

qryline = "select distinct rev_deviceno 'LTM_DEVICENO' from BAIC_REVENUE_PLAN where REV_MONTH=" & planmonth & "  and REV_PLANQTY > 0 union " & _
" select distinct LTM_DEVICENO FROM BAIC_LOTMAST WHERE LTM_STATUS='ACT' union " & _
" select distinct LTM_DEVICENO FROM BAIC_LOTMAST WHERE LTM_STATUS='ACT' union " & _
" select distinct LTM_DEVICENO FROM BAIC_FG_HEADER, BAIC_LOTMAST WHERE FGH_STATUS IN ('PKGRCV','QAACPT','PCKDTL') AND FGH_LOTNO=LTM_LOTNO union " & _
" select distinct LTM_DEVICENO FROM BAIC_FG_HEADER, BAIC_LOTMAST WHERE FGH_STATUS IN ('FGI','INVOICED') AND FGH_LOTNO=LTM_LOTNO union " & _
" select distinct LTM_DEVICENO FROM BAIC_FG_HEADER, BAIC_LOTMAST WHERE FGH_STATUS='SHIP' AND FGH_LAST_TRANS_YMD BETWEEN '" & fromdate & "' AND '" & todate & "' AND FGH_LOTNO=LTM_LOTNO union " & _
" select distinct LTM_DEVICENO FROM BAIC_DS_INVENTORY, BAIC_LOTMAST WHERE DSI_STATUS='ACCEPT' AND DSI_LOTNO=LTM_LOTNO union " & _
" select distinct LTM_DEVICENO from BAIC_DS_HEADER, BAIC_DS_DETAIL_DATA, BAIC_LOTMAST where DSH_STATUS='SHIPPED' and DSH_SHIP_YMD BETWEEN '" & fromdate & "' AND '" & todate & "' and DSH_DROPSHIPNO=DDD_DROPSHIPNO and DDD_ORI_LOTNO=LTM_LOTNO "

'qryline = " select distinct LTM_DEVICENO FROM BAIC_LOTMAST WHERE LTM_STATUS='ACT' union " & _
'" select distinct LTM_DEVICENO FROM BAIC_LOTMAST WHERE LTM_STATUS='ACT' union " & _
'" select distinct LTM_DEVICENO FROM BAIC_FG_HEADER, BAIC_LOTMAST WHERE FGH_STATUS IN ('PKGRCV','QAACPT','PCKDTL') AND FGH_LOTNO=LTM_LOTNO union " & _
'" select distinct LTM_DEVICENO FROM BAIC_FG_HEADER, BAIC_LOTMAST WHERE FGH_STATUS IN ('FGI','INVOICED') AND FGH_LOTNO=LTM_LOTNO union " & _
'" select distinct LTM_DEVICENO FROM BAIC_FG_HEADER, BAIC_LOTMAST WHERE FGH_STATUS='SHIP' AND FGH_LAST_TRANS_YMD BETWEEN '" & fromdate & "' AND '" & todate & "' AND FGH_LOTNO=LTM_LOTNO union " & _
'" select distinct LTM_DEVICENO FROM BAIC_DS_INVENTORY, BAIC_LOTMAST WHERE DSI_STATUS='ACCEPT' AND DSI_LOTNO=LTM_LOTNO union " & _
'" select distinct LTM_DEVICENO from BAIC_DS_HEADER, BAIC_DS_DETAIL_DATA, BAIC_LOTMAST where DSH_STATUS='SHIPPED' and DSH_SHIP_YMD BETWEEN '" & fromdate & "' AND '" & todate & "' and DSH_DROPSHIPNO=DDD_DROPSHIPNO and DDD_ORI_LOTNO=LTM_LOTNO "


        sqlReport = qryline
        Debug.Print sqlReport
        Dim rsReport As ADODB.Recordset
        Dim rs2 As ADODB.Recordset
        
        Set rsReport = New ADODB.Recordset
        rsReport.Open sqlReport, wsDB

        If rsReport.EOF Then 'no record
            xrecord = "N"
        
        Else 'record exists
            xrecord = "Y"

            'Declaration for XLS
            Dim excelApp As New Excel.Application
            Dim wbReport As Workbook
            Dim wsReport As Worksheet
            Dim rejRs As ADODB.Recordset

            excelApp.Visible = False
            excelApp.ScreenUpdating = False
            excelApp.Interactive = False
            excelApp.IgnoreRemoteRequests = True
            
            Set wbReport = excelApp.Workbooks.Add
            Set wsReport = wbReport.ActiveSheet
            
            With wsReport

                .Cells(1, 1).Value = "REVENUE REPORT"
                .Cells(2, 1).Value = "FROM : " & fromdate & " - " & todate
                .Rows(1).Font.Bold = True
                .Rows(2).Font.Bold = True
                
                .Cells(4, 1).Value = "PDM_REPORT_CAT2"
                .Cells(4, 2).Value = "PDM_CUSTOMER"
                .Cells(4, 3).Value = "PDM_REVENUE_CAT2"
                .Cells(4, 4).Value = "PDM_PROCESS_REF"
                .Cells(4, 5).Value = "PDM_PACKAGELEAD"
                .Cells(4, 6).Value = "PDM_DEVICENO"
                .Cells(4, 7).Value = "PDM_TARGETDEVICE"
                .Cells(4, 8).Value = "DEVICE_PRICE (USD)"
                .Cells(4, 9).Value = "FORECAST_QTY"
                .Cells(4, 10).Value = "FORECAST_AMT"
                .Cells(4, 11).Value = "DELTA_QTY"
                .Cells(4, 12).Value = "MTD_SHIP (AIC INVOICE)"
                .Cells(4, 13).Value = "WIP_FOL"
                .Cells(4, 14).Value = "WIP_EOL"
                .Cells(4, 15).Value = "WIP_TEST"
                .Cells(4, 16).Value = "FG"
                .Cells(4, 17).Value = "DC_WIP"
                .Cells(4, 18).Value = "DC_SHIP"
                .Cells(4, 19).Value = "MTD_FG"
                
'''''                .Cells(4, 11).Value = "WIP_SPLITINV"
''''                .Cells(4, 12).Value = "WIP_AWAITING_PACK"
                
                '.Cells(4, 13).Value = "FG"
            
                
''''                .Cells(4, 13).Value = "WIP_PRODPACK"
''''                .Cells(4, 14).Value = "WIP_FGPRIME"
''''                .Cells(4, 15).Value = "WIP_FGPOST"
                
'                .Cells(4, 16).Value = "MTD_SHIP (AIC INVOICE)"
               ' .Cells(4, 17).Value = "DC_WIP"
                '.Cells(4, 18).Value = "DC_SHIP"
                
'                .Cells(4, 17).Value = "MTD_SHIP (AIC INVOICE)"
                '.Cells(4, 18).Value = "DC_WIP"
                '.Cells(4, 19).Value = "DC_SHIP"
                
                '.Cells(4, 19).Value = "FORECAST_QTY"
                '.Cells(4, 20).Value = "DEVICE_PRICE"
                '.Cells(4, 21).Value = "FORECAST_AMT"
                '.Cells(4, 22).Value = "DELTA_QTY"
                
                .Rows(4).Font.Bold = True
                .Rows(4).Font.Color = vbBlue
    
                Dim strCategory, strLossCat, strLossQty As String
    
                intRow = 5 'Starting row# for data insertion
        
                Do While Not rsReport.EOF
                    xdevice = Trim(rsReport!LTM_DEVICENO)
                    Debug.Print xdevice
                    'MAIN INFO
                    Set rs2 = New ADODB.Recordset
                    ssql = "select PDM_REPORT_CAT2, PDM_CUSTOMER, PDM_REVENUE_CAT2,PDM_PROCESS_REF, " & _
                           " PDM_PACKAGELEAD, PDM_DEVICENO, PDM_TARGETDEVICE " & _
                           " From baic_prodmast where pdm_deviceno='" & xdevice & "'"
                    rs2.Open ssql, wsDB
                    If Not rs2.EOF Then
                        .Cells(intRow, 1).Value = rs2!PDM_REPORT_CAT2
                        .Cells(intRow, 2).Value = rs2!PDM_CUSTOMER
                        .Cells(intRow, 3).Value = rs2!PDM_REVENUE_CAT2
                        .Cells(intRow, 4).Value = rs2!PDM_PROCESS_REF
                        .Cells(intRow, 5).Value = rs2!PDM_PACKAGELEAD
                        .Cells(intRow, 6).Value = rs2!PDM_DEVICENO
                        .Cells(intRow, 7).Value = rs2!PDM_TARGETDEVICE
                    End If
                    rs2.Close
                    
                    '----------------------------------------------------------------------------------
                    'WIP
                    '----------------------------------------------------------------------------------
                    Set rs2 = New ADODB.Recordset
                    ssql = "select PDM_REPORT_CAT2, PDM_CUSTOMER, PDM_REVENUE_CAT2,PDM_PROCESS_REF, " & _
                           " PDM_PACKAGELEAD, LTM_DEVICENO, LTM_TARGETDEVICE, " & _
                           " sum(case when LTM_OPER >= 901 and LTM_OPER < 2100 then LTM_QTY else 0 end) TOTAL_FOL," & _
                           " sum(case when LTM_OPER >= 2100 and LTM_OPER < 3500 then LTM_QTY else 0 end) TOTAL_EOL, " & _
                           " sum(case when LTM_OPER >= 3500 and LTM_OPER < 9000 then LTM_QTY else 0 end) TOTAL_TEST " & _
                           " From baic_prodmast , baic_lotmast " & _
                           "    where ltm_status='ACT' and ltm_deviceno=pdm_deviceno and pdm_deviceno='" & xdevice & "' " & _
                           "    GROUP BY PDM_REPORT_CAT2, PDM_CUSTOMER, PDM_REVENUE_CAT2,PDM_PROCESS_REF, PDM_PACKAGELEAD, LTM_DEVICENO, LTM_TARGETDEVICE"

''''''                           " sum(case when LTM_OPER >= 8000 and LTM_OPER <= 8600 then LTM_QTY else 0 end) TOTAL_SPLITINV, "
''''''                           " sum(case when LTM_OPER = 7000 then LTM_QTY else 0 end) TOTAL_AWAITING_PACK, "

                    
                    rs2.Open ssql, wsDB
                    If Not rs2.EOF Then
                        .Cells(intRow, 13).Value = Format(rs2!TOTAL_FOL, "###,###,###,###")
                        .Cells(intRow, 14).Value = Format(rs2!TOTAL_EOL, "###,###,###,###")
                        .Cells(intRow, 15).Value = Format(rs2!TOTAL_TEST, "###,###,###,###")
'''                        .Cells(intRow, 11).Value = rs2!TOTAL_SPLITINV
'''                        .Cells(intRow, 12).Value = rs2!TOTAL_AWAITING_PACK
                    End If
                    rs2.Close
                    
                    '----------------------------------------------------------------------------------
                    'FG ('PKGRCV','QAACPT','PCKDTL','FGI','INVOICED')
                    '----------------------------------------------------------------------------------
                    Set rs2 = New ADODB.Recordset
                    ssql = " select " & _
                           " sum(FGH_QTY) FGQTY " & _
                           " From BAIC_FG_HEADER, BAIC_LOTMAST, BAIC_PRODMAST " & _
                           " Where FGH_STATUS IN ('PKGRCV','QAACPT','PCKDTL','FGI','INVOICED') and pdm_deviceno = '" & xdevice & "' " & _
                           " AND FGH_LOTNO=LTM_LOTNO AND LTM_DEVICENO=PDM_DEVICENO and pdm_deviceno='" & xdevice & "'" & _
                           " GROUP BY PDM_REPORT_CAT2, PDM_CUSTOMER, PDM_REVENUE_CAT2,PDM_PROCESS_REF, PDM_PACKAGELEAD, LTM_DEVICENO, LTM_TARGETDEVICE"

                    rs2.Open ssql, wsDB
                    If Not rs2.EOF Then
                        .Cells(intRow, 16).Value = Format(rs2!FGQTY, "###,###,###,###")
                
                    End If
                    rs2.Close
                    
                    
                    '----------------------------------------------------------------------------------
                    'MTD_FG   (Fgh_Creation_Ymd)
                    '----------------------------------------------------------------------------------
                    Set rs2 = New ADODB.Recordset
                    ssql = " select " & _
                           " sum(FGH_QTY) FGMTD " & _
                           " From BAIC_FG_HEADER, BAIC_LOTMAST, BAIC_PRODMAST " & _
                           " Where FGH_STATUS NOT LIKE 'CANCEL%' and pdm_deviceno = '" & xdevice & "' " & _
                           " AND FGH_CREATION_YMD BETWEEN '" & fromdate & "' AND '" & todate & "'  " & _
                           " AND FGH_LOTNO=LTM_LOTNO AND LTM_DEVICENO=PDM_DEVICENO and pdm_deviceno='" & xdevice & "'" & _
                           " GROUP BY PDM_REPORT_CAT2, PDM_CUSTOMER, PDM_REVENUE_CAT2,PDM_PROCESS_REF, PDM_PACKAGELEAD, LTM_DEVICENO, LTM_TARGETDEVICE"

                    rs2.Open ssql, wsDB
                    If Not rs2.EOF Then
                        .Cells(intRow, 19).Value = Format(rs2!FGMTD, "###,###,###,###")
                
                    End If
                    rs2.Close
                    
                    
                    '----------------------------------------------------------------------------------
                    'FG PACK ('FG PRIME, FG POST')
                    '----------------------------------------------------------------------------------
'''''''                    Set rs2 = New ADODB.Recordset
'''''''                    ssql = " select " & _
'''''''                           " sum(case when FGH_STATUS='FGI' then FGH_QTY else 0 end) FG_PRIME, " & _
'''''''                           " sum(case when FGH_STATUS='INVOICED' then FGH_QTY else 0 end) FG_POST " & _
'''''''                           " From BAIC_FG_HEADER, BAIC_LOTMAST, BAIC_PRODMAST " & _
'''''''                           " Where FGH_STATUS IN ('FGI','INVOICED') " & _
'''''''                           " AND FGH_LOTNO=LTM_LOTNO AND LTM_DEVICENO=PDM_DEVICENO and pdm_deviceno='" & xdevice & "'" & _
'''''''                           " GROUP BY PDM_REPORT_CAT2, PDM_CUSTOMER, PDM_REVENUE_CAT2,PDM_PROCESS_REF, PDM_PACKAGELEAD, LTM_DEVICENO, LTM_TARGETDEVICE"
'''''''
'''''''                    rs2.Open ssql, wsDB
'''''''                    If Not rs2.EOF Then
'''''''                        .Cells(intRow, 14).Value = rs2!FG_PRIME
'''''''                        .Cells(intRow, 15).Value = rs2!FG_POST
'''''''                    End If
'''''''                    rs2.Close

                    '----------------------------------------------------------------------------------
                    'MTD SHIP (AIC INVOICES)
                    '----------------------------------------------------------------------------------
                    Set rs2 = New ADODB.Recordset
                    ssql = " select " & _
                           " SUM(FGH_QTY) SHIPQTY " & _
                           " From BAIC_FG_HEADER, BAIC_LOTMAST, BAIC_PRODMAST " & _
                           " Where FGH_STATUS='SHIP' AND FGH_LAST_TRANS_YMD BETWEEN '" & fromdate & "' AND '" & todate & "' and LTM_DEVICENO='" & xdevice & "'" & _
                           " and FGH_LOTNO=LTM_LOTNO and LTM_DEVICENO=PDM_DEVICENO " & _
                           " GROUP BY PDM_REPORT_CAT2, PDM_CUSTOMER, PDM_REVENUE_CAT2,PDM_PROCESS_REF, PDM_PACKAGELEAD, LTM_DEVICENO, LTM_TARGETDEVICE"
                    rs2.Open ssql, wsDB
                    If Not rs2.EOF Then
                    
                        .Cells(intRow, 12).Value = Format(rs2!SHIPQTY, "###,###,###,###")
                    End If
                    rs2.Close

                    '----------------------------------------------------------------------------------
                    'DC WIP
                    '----------------------------------------------------------------------------------
                    Set rs2 = New ADODB.Recordset
                    ssql = " select " & _
                           " SUM(DSI_BAL_QTY) DCBALQTY " & _
                           " From BAIC_DS_INVENTORY, BAIC_LOTMAST, BAIC_PRODMAST " & _
                           " Where DSI_STATUS='ACCEPT' " & _
                           " AND DSI_LOTNO=LTM_LOTNO and LTM_DEVICENO=PDM_DEVICENO and PDM_DEVICENO='" & xdevice & "'" & _
                           " GROUP BY PDM_REPORT_CAT2, PDM_CUSTOMER, PDM_REVENUE_CAT2,PDM_PROCESS_REF, PDM_PACKAGELEAD, LTM_DEVICENO, LTM_TARGETDEVICE"
                    rs2.Open ssql, wsDB
                    If Not rs2.EOF Then
                        .Cells(intRow, 17).Value = Format(rs2!DCBALQTY, "###,###,###,###")
                  
                    End If
                    rs2.Close


                    '----------------------------------------------------------------------------------
                    'DC SHIP
                    '----------------------------------------------------------------------------------
                    Set rs2 = New ADODB.Recordset
                    ssql = " select LTM_DEVICENO, sum(DDD_QTY) DCSHIP from BAIC_DS_HEADER, BAIC_DS_DETAIL_DATA, BAIC_LOTMAST " & _
                           " where DSH_STATUS='SHIPPED' and DSH_SHIP_YMD BETWEEN '" & fromdate & "' AND '" & todate & "' " & _
                           " and DSH_DROPSHIPNO=DDD_DROPSHIPNO and DDD_ORI_LOTNO=LTM_LOTNO and LTM_DEVICENO='" & xdevice & "' " & _
                           " group by LTM_DEVICENO "
                    rs2.Open ssql, wsDB
                    If Not rs2.EOF Then
                        .Cells(intRow, 18).Value = Format(rs2!DCSHIP, "###,###,###,###")
                  
                    End If
                    rs2.Close




                    '----------------------------------------------------------------------------------
                    'FORECAST PLAN
                    '----------------------------------------------------------------------------------
                    Set rs2 = New ADODB.Recordset
                    ssql = " select * from BAIC_REVENUE_PLAN where REV_DEVICENO='" & xdevice & "' and REV_MONTH=" & planmonth
                    rs2.Open ssql, wsDB
                    If Not rs2.EOF Then
                    'ko add 20210414
                       If rs2!rev_planqty = 0 Then
                          .Cells(intRow, 9).Value = 0
                          .Cells(intRow, 11).Value = Format(0 - (.Cells(intRow, 12).Value), "###,###,###,###")
                       Else
                          .Cells(intRow, 9).Value = Format(rs2!rev_planqty, "###,###,###,###")
                          .Cells(intRow, 11).Value = Format(rs2!rev_planqty - (.Cells(intRow, 12).Value), "###,###,###,###")
                       End If
                    
                    Else
                       .Cells(intRow, 9).Value = 0
                       .Cells(intRow, 11).Value = Format(0 - .Cells(intRow, 12).Value, "###,###,###,###")
                       
                       If .Cells(intRow, 11).Value < 0 Then
                          .Cells(intRow, 11).Font.Color = vbRed
                          
                       End If
                    End If
                    
                    rs2.Close
                    


                    '----------------------------------------------------------------------------------
                    'DEVICE PRICE
                    '----------------------------------------------------------------------------------
                    Set rs2 = New ADODB.Recordset
                    ssql = " select IPD_ASSY_PRICE+IPD_TEST_PRICE+IPD_TNR_PRICE 'DEVPRICE' from BAIC_INVOICE_PRICE_DETAIL where IPD_DEVICENO='" & xdevice & "' and IPD_EXP_YMD > GETDATE()  order by IPD_EXP_YMD desc"
                    rs2.Open ssql, wsDB
                    If Not rs2.EOF Then
                        
                        'Convertion RMB-to-USD Rate 6.5 for ChangDian Only 20210517 disable
                      '  If Cells(intRow, 2).Value = "CHANGDIAN" Or .Cells(intRow, 2).Value = "HCLY" Or .Cells(intRow, 2).Value = "SILIKRON" Or .Cells(intRow, 2).Value = "GREEN POWER" Then
                      'Previous included Haigete and changdian ( no conversion -- from RMB TO USD', Effective 15-05-2021 used USD Currency
                      
                        If .Cells(intRow, 2).Value = "HAIGETE" Or .Cells(intRow, 2).Value = "HCLY" Or .Cells(intRow, 2).Value = "SILIKRON" Or .Cells(intRow, 2).Value = "GREEN POWER" Then
                            .Cells(intRow, 8).Value = Format(rs2!DEVPRICE / 6.5, "###.00000")
                            .Cells(intRow, 10).Value = rs2!DEVPRICE / 6.5 * .Cells(intRow, 9).Value
                            .Cells(intRow, 10).Value = Format(rs2!DEVPRICE / 6.5 * .Cells(intRow, 9).Value, "###,###,###.##")
                        
                        '20211013 For AAS NPA ASIC.. zero price.. requested by AC Chuah
                        ElseIf .Cells(intRow, 3).Value = "AAS NPA ASIC" Then
                            .Cells(intRow, 8).Value = "0.00"
                            .Cells(intRow, 10).Value = "0.00"
                        
                        Else
                            .Cells(intRow, 8).Value = Format(rs2!DEVPRICE, "###.00000")
                            .Cells(intRow, 10).Value = Format(rs2!DEVPRICE * .Cells(intRow, 9).Value, "###,###,###.##")
                        End If
                        
                    End If
                    rs2.Close
                                                            
                    intRow = intRow + 1
                    rsReport.MoveNext
                Loop
                
               .Range("A1").Select
           
           End With

            wsReport.UsedRange.Columns.AutoFit
      '    wbReport.SaveAs "c:\aicreport\REVENUE_RPT_" & rptdate & ".xls"
        '    wbReport.Close
    
    
        xfilename = "REVENUE_RPT_" & rptdate
        wbReport.SaveAs "C:\AICREPORT\" & xfilename & ".xls"
        wbReport.Close
        Set wsReport = Nothing
    
            '--clear excel memory
            excelApp.Visible = True
            excelApp.Interactive = True
            excelApp.ScreenUpdating = True
            excelApp.IgnoreRemoteRequests = False
            excelApp.Quit
            Set excelApp = Nothing
      
        'excelApp.Visible = True ' ko others module statement ok 202104
        'excelApp.ScreenUpdating = True
        'excelApp.Interactive = True
        'excelApp.IgnoreRemoteRequests = False

        End If
        '---------------------------
        rsReport.Close
        Set rsReport = Nothing


'----- auto send by email

            Dim olApp As Outlook.Application
           Set olApp = CreateObject("Outlook.Application")

           Dim olNs As Outlook.Namespace
           Set olNs = olApp.GetNamespace("MAPI")
           olNs.Logon

           Dim olMail As Outlook.MailItem
           Dim myAttachments As Outlook.Attachments
           Set olMail = olApp.CreateItem(olMailItem)
           Set myAttachments = olMail.Attachments
           
           olMail.To = "INT_REPORT2021"
           'olMail.To = "py_sim@aicsemicon.com"
           olMail.Subject = xfilename
           myAttachments.Add "C:\AICREPORT\" & xfilename & ".XLS", olByValue, 1, xfilename
           olMail.Body = vbCrLf & vbCrLf & "AUTO-GENERATED REPORT FROM " & vbCrLf & " AIC SEMICONDUCTOR SDN. BHD."
           olMail.Send

           olNs.Logoff
           Set olNs = Nothing
           Set olMail = Nothing
           Set olAppt = Nothing
           Set olItem = Nothing
           Set olApp = Nothing
           Set myAttachments = Nothing


Unload Me
End Sub
