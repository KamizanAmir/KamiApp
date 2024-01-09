    
'***CZZ***
'July 06, 2010
'+ Icons added to Main Menu GUI
'
'July 07, 2010
'o Modified Main Menu GUI for Enquiry > Reports
'
'July 08, 2010
'+ Menu added: Enquiry > Customer Report
'
'August 02, 2010
'+ Added function to prevent multiple instance
'
'August 03, 2010
'+ Menu and code added: OTHER > RESCREEN
'*********
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
'----------------------------------------
'*** pls include modification date ****
'----------------------------------------
'4000998-2023-12-27 20231227 PY Stock Count December 28 and Kamizan add Cycle Time Process in CS Report
'4000997-2023-12-15 20231215 cs report - stmicro add column
'4000996-2023-11-30 20231130 Stockcount Update
'4000995-2023-11-29 UPDATE TEST BIN YIELD * FOR ST DEVICES

'4000994-2023-11-28FITAD 20231128 STMICRO BIN TEST SCRAP, STMICRO MVOU SCRAP (NOT YET RELEASE)
'4000993-2023-11-21 20231124 blank update to force prod pc update
'4000992-2023-11-21 20231121 Unlocked Release Lot for SE LIM as QA Backup 20231121 without blocking different owner'
'4000991-2023-11-17 20231117 AIN JY ASK TO REMOVE THE MESSAGE
'                   20231117 AIN JY ASK TO PREVIEW FIRST BEFORE PRINT
'4000990-2023-11-15 SEPARATE TST AND T000 CRYSTAL REPORT
'4000989-2023-11-15 FIXED BUGS
'4000988-2023-11-15 20231115 AIN ADD DATA BARCODECHIP
'4000987-2023-11-09 PACKMOV CHECK - CHONG reorganize the packmov condition at 1307 operation with 3 hardcoded dual die device
'4000986-2023-10-31 20231020 CHONG add a checking to block Future-Hold lot from ST-BinConvert 20231031
'4000985-2023-10-27 20231027 PY Stock Count October change to 28 Oct
'4000984-2023-10-27 20231027 PY Stock Count October
'4000983-2023-10-25 Fixed bugs
'4000982-2023-10-20 20231020 CHONG add a checking to block Future-Hold lot from split 20231020
'4000981-2023-10-20 20231020 AIN DEVICE BTA41,BSE
'4000980-2023-10-19 20231019 AIN FOR DEVICE BTA41-BSE/Q ONLY
'4000979-2023-10-13 20231013 Ain adjust marking for device BTA41 BSE
'4000978-2023-10-10 231010 Sim Pei Ying Add Mother Lot
'4000977-2023-10-06 request by mahadhir from 56 days to 14 days (2 weeks) '20231006 'Chong
'4000976-2023-10-05 AIN 20231005 AIN Plant CT Query
'                   20230914 AIN LITTELFUSE AND DEVICE
'4000975-2023-09-04 ASYRAF UPDATE ISSUE
'4000974-2023-08-30 AIN 20230828 ADD 2 COLUMN TUN DUN AT TEST REPORT
'4000973-2023-08-28 20230828 Stockcount Update
'4000972-2023-08-10 20230810 AIN CHANGE DATABASE SAPSVR TO SAPB1SVR
'4000971-2023-08-07 Chong Modify Amphenol Housing : Oper 2460 low yield 99% at COMTBL
'4000970-2023-07-25 2023-07-25 ASYRAF Fixed Error
'4000969-2023-07-25 2023-07-25 ASYRAF BLOCK LOT THAT HAVE MORE THAN 800 QTY(R91) AND MORE THAN 2400 QTY(NOT RD91)
'4000968-2023-07-25 2023-07-25 ASYRAF BLOCK LOT THAT HAVE MORE THAN 800 QTY(R91) AND MORE THAN 600 QTY(NOT RD91)
'4000967-2023-07-13 2023-07-13 AIN SKIP FOR ROUTE TDFN8.FA7607OTR.GCUPBF
'4000966-2023-07-10 20230707 ENHANCEMENT FOR TEST REPORT
'4000965-2023-06-26 Stockcount 27/6/2023
'4000964-2023-06-22 UPDATE SYSTEM VERSION AS LAST COMPILE SET WRONG VERSION
'4000963-2023-06-22 CHONG FIX BUG WHERE EDC FOR ST RETRIEVE COLUMN NAME (REMARK5) data which will be overite by mc maintenance
'4000962-2023-06-12 CHONG FIX BUG WHERE ST-DB-MVOU RETRIEVE ALL CUSTOMER LOTNO
'4000961-2023-06-06 CHONG 20230606 ST-3000 LOWYIELD DOES NOT REQUIRE QA
'4000960-2023-06-02 AIN 20230602 SPLIT AT FOL FOR PLANNER ONLY
'4000959-2023-06-01 Chong EDC update, fix bug (eproxy cure1/2 as category 1350)
'4000958-2023-05-31 Chong 20230531 EDC update, fix bug
'4000957-2023-05-31 Chong 20230531 EDC update, add operation to catch EDC --1350-EPOXY CURE, 1305-DIE BOND2, 1355-EPOXY CURE2, 1405-WIRE BOND2, 1450-WBOND CU
'4000956-2023-05-29  Stockcount 29th May, unlock for keyin - Siva request
'4000955-2023-05-26  99% default low yield (FOI) as WTCHOONG request | Stockcount 29th May
'4000954-2023-05-18 AIN 20230518 STD QTY FOR RD91 25, ELSE 600
'4000953-2023-05-17 AIN 20230517 ADD OPER 3500 FOR QFN MERGE
'4000952-2023-05-15 AIN 20230515 CHANGE QFN MERGE START AT 5500
'4000951-2023-05-15 AIN 20230515 CATEG NPA SPLIT AT VMI, SOT NOT SETUP YET AND AUTHORIZE PERSON CAN SPLIT AT VMI ONLY
'4000950-2023-05-12 AIN 20230512 CATEG NPA SPLIT AT FINAL TEST
'4000949-2023-05-12 AIN 20230512 FIXED BUGS
'4000948-2023-05-12 AIN 20230512 FIX BUG AT PROD MVOU
'                   AIN 20230508 SPLIT AT VMI
'4000947-2023-04-28 AIN 20230428 UPDATE MARKING RD91 AT TEST ROUTE
'4000946-2023-04-27 Chong comment out stockcount time lock for user key in
'4000945-2023-04-25 ASYRAF FIXED TRACE CODE CREATION AT 1230 PROD MVOU (Exclude TOP3WIRE)
'4000944-2023-04-17 CHONG STMICRO LOWYIELD CHECK AT OPERATION 3000
'4000943-2023-04-11 ASYRAF REDO TRACE CODE CREATION AT DB MVOU for Discrepancies - Cutoff 6 Apr/Previously at PROD MVOU
'4000942-2023-04-05 CHONG FIX BUG FOR RE-PRINT TT
'4000941-2023-04-04 AIN APT ADD ON DOP3 + COLUMN RAILCODING CHECK
'4000940-2023-04-03 CHONG FIX REPRINT CHECK TO SEPERATE TT/VISION
'4000939-2023-04-03 AIN ADD 20230304 BAIC_SUBINV_MASTER MOVEOUT EDC
'4000938-2023-03-29 20230328 CHONG add function to lock Test Traveller REPRINT
'4000937-2023-03-28 ASYRAF : Revise Test Report Issue
'4000936-2023-03-28 ASYRAF : TRACE CODE CREATION FOR TOP3WIRE AT MVOU OPER 1000 & 1155
'4000935-2023-03-27 Fix Test Report Issue
'4000934-2023-03-27 Stockcount update
'4000933-2023-03-24 ASYRAF : COMMENT ALL SENSOR NEW CODE @ PRODMVOU DUE TO USER CANNOT MVOU
'4000932-2023-03-24 AIN : ADJUST APT FOR TOP3 PRODUCT
'4000931-2023-03-13 Fix RD91 Standard Pack
'4000930-2023-03-03 UPDATE ON ST_DB_MVOU
'4000929-2023-03-03 Fix on Top3wire product deduct die qty instead of wafer qty
'4000928-2023-03-01 Release For DOP3 Reject code YAS KOKO Yield Analyis sheet
'4000927-2023-02-28 Release for stockcount key in day 27,28
'4000926-2023-02-27 Release for stockcount key in day 27,28
'4000925-2023-02-22 Fix bug from last update
'4000924-2023-02-22 diebank split,mvou,printapt to not show STMICRO,TST device|20230222 Chong add to lock ST tst device from loading to Diebank function
'4000923-2023-02-20 chong update merge lot for STMICRO, rd91 max size =25 ''20230220 CHONG ADD FOR RD91 MAX QTY = 25
'                   Stockcount update(no changes, JAN stockcount on 27th)
'4000922-2023-02-14 chong update diebank create lot to not show STMICRO,TST device
'4000921-2023-01-27 Stockcount update 'closed (correct date and time)
'4000920-2023-01-27 Stockcount update 'release backto allow key in
'4000919-2023-01-27 Stockcount update 'release backto allow key in' (set to day 29)
'4000918-2023-01-27 Stockcount update 'release backto allow key in' (set to day 28)
'4000917-2023-01-27 Stockcount update
'4000916-2023-01-13 Fiexed bugs
'4000915-2023-01-12 QA must release all hold lot for all customer
'4000914-2023-01-10  CS Report - adds on STMIRCO new format
'4000913-2023-01-09  Yield add & change for TOP3 & DOP3 for BIN7, RD91 BIN6,7
'4000912-2023-01-06  Fixed 2 digit WW for Tracecode
'4000911-2023-01-05  Remove insert Yield Percent by Asyraf
'4000910-2023-01-05  Remove insert Yield Percent by Asyraf
'4000909-2023-01-05  Yield change for TOP3 & DOP3 for BIN6 from 0.5 to 5.7
'4000908-2023-01-04  add yield percent input and output
'                    add category SOT23, QFN, SOICMF, TSOP, TOP3, DOP3 for low yield
'4000907-2022-12-29  Stockcount month - Fixed bugs
'4000906-2022-12-28  Stockcount month 20221229
'4000905-2022-12-28  Stockcount month 12
'4000904-2022-12-26 Capture the 2023 assembly PO Condition 'not release yet by KOKOKO
'4000903-2022-12-13 Fix bugs on Release Lot
'4000902-2022-12-09 Stockcount ST Preload LI | Create Lot (Fix 0 Qty Isse)
'4000901-2022-12-08 Stockcount ST Preload LI | Create Lot (Fix 0 Qty Isse)
'4000900-2022-11-29 Stockcount month 11- Fixed bugs
'4000899-2022-11-25 Stockcount month 11
'4000898-2022-11-23 Fixed bugs
'4000897-2022-11-23 Fixed bugs (MICROSHIP / CREATE LOT)
'4000896-2022-11-22 Fixed bugs
'4000895-2022-11-22 Fixed bugs
'4000894-2022-11-17 Fixed bugs
'4000893-2022-11-17 Fixed bugs
'4000892-2022-11-17 Fixed bugs
'4000891-2022-11-17 Fixed bugs
'4000890-2022-11-17 Fixed bugs
'4000889-2022-11-17 Fixed old Genrate Lot, Split Lot : Block STMICRO TST Device from old process
'4000888-2022-11-16 Fixed bugs Merge Lot (BIN)
'4000887-2022-11-16 Fixed bugs Create Lot
'4000884-2022-11-14 Fixed bugs Create Lot
'4000883-2022-11-11 Fixed bugs Test Traveller RD91
'4000882-2022-11-11 Create list LI APT
'4000881-2022-11-10 Fixed bugs : Update WIP Header for ST Preload LI (Create Lot)
'4000880-2022-11-09 Fixed bugs
'4000879 2022-11-08 Fixed bugs
'4000878 2022-11-08 Fixed bugs
'4000877 2022-11-08 Change position DIEBANK after ENQUIRY
'4000876 2022-11-08 Fixed bugs
'4000875 2022-11-07 Create for ST PRELOAD TEST (Create, Split, MVOU, Print and Reprint)
'4000874 2022-11-01 Fixed bugs
'4000872 2022-11-01 Fixed bugs
'4000871 2022-10-31 Add new query to fetch latest railcode seq
'4000870 2022-10-28 Yana - Bypass planner JY login to allow revrse for STMICRO & Add condition to bypass movout 0 quantity for STMICRO
'4000869 2022-10-14 - Fix BUG
'4000868 2022-10-03 Ain - Add RD91 Test Traveller
'4000867 2022-09-28 Yana - Open Stock COunt
'4000866 2022-09-27 Yana - Adjusment on filtering of cuslotno for update ST marking by 9Digit.
'                        - Adjustment on stock count date
'                        - Adjustment on formula of goldwire calculation
'4000865 2022-09-26 - Fix bugs
'4000864 2022-09-06 - Fix bugs
'4000863 2022-09-06 - Fix coding for updating markfile for STMICRO
'4000862 2022-09-05 - Fix bugs
'4000861 2022-09-02 - Remove terminate lot blocking for ST
'4000860 2022-08-30 - Release new APT to be included '5 PSI LEAK TEST' for device NPA-2-C02215 (AIC) req by TG OO
'4000859 2022-08-29 - Fix bugs
'4000858 2022-08-29 - Open StockCount for August
'4000857 2022-08-26 - Fix Moveout ST 3010 to 3400 bugs error
'4000856 2022-08-25 - Fix WIP Report column
'4000855 2022-08-25 - Adjust ST WIP Report
'4000854 2022-08-15 - KO - Adjustment on moveout error Message 3010 ~ 3400 for STMICRO
'4000853 2022-08-13 - Fix bugs
'4000849 2022-08-12 - ST MICRO, open split during AS
'4000848 2022-08-11 - Adjustment for STMICRO (marking,validation)
'4000847 2022-07-28 - Yana - Open Stock Count untill 4pm
'4000846 2022-07-18 - Yana - Fix bugs
'4000843 2022-07-14 - Yana - Fix bugs
'4000840 2022-07-14 - KO - Change "QA BULK ID" text to "QA NUMBER" during 5520 moveout
'4000839 2022-07-05 - Yana - Fix bugs - Logic for convert tracecode MFYWWXXX during moveout 1307
'4000838 2022-07-05 - Yana - Fix bugs - Test Traveller STMICRO file not found
'4000837 2022-06-30 - Fix bugs
'4000836 2022-06-28 - StockCount open until 17 for FOL.
'4000835 2022-06-28 - Add Spool input for StockCount goldwire.
'4000834 2022-06-01 - Yana - Fix bugs
'4000833 2022-05-30 - Quah/Yana - Stock Count Date Adjustment
'4000832 2022-05-26 - Quah/Yana - Set up for ST APT, set up Wafer deduction during MOVOU 1220
'4000831 2022-05-10 - Yana - Adjustment of MICROCHIP COMBINE WAFER
'4000830 2022-05-10 - Quah - BINCONVERT, SPLIT LOT for STMICRO.
'                   - Yana - Add export data to Excell for Microchip Combine Wafer
'4000829 2022-05-09 - Quah/Yana - Release MICROCHIP COMBINE WAFER
'4000828 2022-04-28 - Quah - Setting for auto StockCount date on 28th.
'4000827 2022-04-19 - KO/Yana - FIX bugs for validate emp id for STMICRO process
'4000825 2022-04-19 - KO - FIX lowyield bugs
'4000824 2022-04-18 - KO/Yana - Setup blocking message for unvalid employee ID for STMICRO process
'4000823 2022-04-18 - Yana - Release ST Test Traveller
'4000822 2022-04-04 - Quah - EDC STMICRO validate Emp ID.
'4000820 2022-04-04 - Quah - Fix for STMICRO auto generate QA BULK ID.
'4000819 2022-03-25 - Quah - STMICRO auto generate QA BULK ID.
'4000818 2022-03-24 - Quah - Stkcount 2022-03-28 set date.
'4000817 2022-03-22 - Quah - Remove ST checking on Process Transaction authorization, pending for further instruction.
'4000816 2022-03-22 - Quah/Yana - Release ST APT
'4000815 2022-03-21 - Quah/Yana - Fix bugs
'4000814 2022-03-21 - Quah - Machine for ST
'4000812 2022-03-18 - Quah/Yana - Block ST MICRO for split, merge and reverse
'4000811 2022-03-15 - Quah - ST APT
'4000810 2022-03-01 - Yana - Add WAFER_QTY in ssql query to match column with xwaferqty
'4000809 2022-02-28 - Quah - ST logic for marking and railcode
'4000808 2022-01-27 - YANA - StockCount on 28th Feb 2022. CHANGE TO 1600
'4000806 2022-01-27 - Quah - StockCount on 28th Jan 2022.
'4000804 2021-12-29 - Quah - StockCount on 30th Dec 2021.
'4000803 2021-12-29 - Quah - StockCount on 30th Dec 2021.
'4000802 2021-12-29 - Quah - StockCount on 30th Dec 2021.
'4000800 2021-12-20 - Quah - TempCycle must be kept 5 days before TEST (check during 3500 PROC)
'4000799 2021-12-16 - Quah - Release for ST APT.
'4000798 2021-12-14 - Quah - STMicro skip 'Remove Remnant' at 3010 APT Print.
'4000797 2021-11-29 - Yana/Quah - Release TT-STD-FTSOIC2020.rpt to include retest in one page
'4000795 2021-11-17 - Yana - Add 'LIKE' condition for reject in View Transaction
'4000794 2021-10-28 - Quah - Material Count set between 8am-3pm on 28th Oct (first month using system).
'4000793 2021-10-26 - Quah - Fix bugs.
'4000792 2021-10-26 - Quah - Material StockCount for every 28th.
'4000791 2021-10-05 - Quah - Impinj customise marking data base on cuslot last 5 char.
'4000790 2021-10-04 - Quah - TempCycle must be kept 7 days before TEST (check during 3500 PROC)
'4000789 2021-09-07 - Quah - Auto update for Impinj Cuslot and Topm4 during APT print (follow last5).
'4000788 2021-09-03 - Quah - Assembly Yield Report for SOT23 EOL, req by Eddie. Yield = (TNF_Out/MOLD_In)*100
'4000786 2021-09-01 - Quah - Fix yield calc error for SOT23 EOL Yield (from Mold to TNF)
'4000785 2021-08-25 - Quah - Eddie request SOT EOL lowyield control at 99.5% (TnF-Out/Mold-In) --- Testrun for ChangDian first.
'4000784 2021-07-29 - Quah - Lowyield trigger set at default 99.5 for QFN package at FOI-3000 (if no specific setting at COMTBL), req by Eddie.
'4000783 2021-07-29 - Quah - Bart Port Setting for B Devices allow for Parts 025.0006.0000 (old) & 025.0020.0000 (new)
'4000782 2021-07-28 - Quah - CS Yield Report (SOT) : TestOut add Split Qty (Cyndi request).
'4000781 2021-07-26 - Quah - Fix the ELSE=NA to solve the NPA lowyield block (feedback by OO, Eddie)
'4000780 2021-06-28 - Ko - Release for BOM flush-out during MVOU. (BOM Project)
'4000779 2021-06-22 - Test traveller for TT-SCMT2 ,TT-SCM REQUESTED BY SE LIM, FIZA
'4000778 2021-06-22 - remove the use square punch for smartcard
'4000777 2021-06-11 - SOT Yield Rpt (TestOut + Split).
'4000776 2021-06-03 - Insert into BAIC_LI_BOM during Create Lot.
'4000775 2021-06-02 - Release the WASSYPT-SOT/WASSYPT to include the marking pattern refer to LI RECOMPILE
'4000774 2021-06-02 - Release the WASSYPT-SOT/WASSYPT to include the marking pattern refer to LI
'4000773 2021-05-28 - Fix added NEW Customer all sensors at 96 cycles and fix test traveller
'4000772 2021-05-24 - Fix bugs for Changdian Merge.
'4000771 2021-05-21 - Fix bugs.
'4000770 2021-05-20 - add SOT23 PRODUCT LOW Yield 99.5 as per SB Heng Req 20210520
'4000769 2021-05-18 - Add in Port Attached column req by siti QA /WT Choong 20210518 release
'4000768 2021-05-17 - Fix bug for Merge SQA rules.
'4000767 2021-05-10 - Fix bug -remove dot for CJ,CD marking.
'4000766 2021-05-05 - CS Report reset back setting for SOT Yield Report.
'4000765 2021-05-05 - Remove dot for CJ,CD marking.
'4000764 2021-04-20 - CJ,CD merging can allow SQA lots to be mother, but don't allow to merge into other lots (due to SQA must carry till end).
'4000763 2021-04-20 - Markfile for QFN lots, previously match by Targetdevice, Mon req change to Deviceno.
'4000762 2021-03-24 - Setup APT for ALLSENSORS.
'---- BANJO ATTACK on 10thMarch 8pm -------------------
'4000758 2021-03-02 - Quah Change to QFNSVR/D folder for auto TESTIN TESTOUT.
'4000757 2021-03-01 - Quah/Ko Temporary bypass QFNSERVER for auto TESTIN TESTOUT, due to QFNSERVER down.
'4000756 2021-02-23 - Quah - APT print preview if login = ITADM. Others print to printer.
'4000755 2021-02-23 - KO - add range at wassypt-sot23 requested by koay, included fix PO Closed report
'4000754 2021-02-17 - KO/Quah - PO Close add INS_PONO Matching with AIC_WIP_HEADER REMARKS PO ,SHIP PO DOUBLE
'4000753 2021-02-16 - Quah - LotSearch by Oper, add RANGE data.
'4000752 2021-02-02 - KO-CJ TEST TRAVELLERT THE TARGET DEVICE CHANGE TO IDAUTOMATION FONT INSTEAD OF CODE39 DUE TO ()
'4000751 2020-12-24 - Bourn TT without QR CODE, STANDARD REPORT
'4000750 2020-12-04 - Zuraini Amphenol Test Rpt.. change 1620MVOU to Lth_Qty_New instead of Lth_Qty_Old (bugs??).
'4000749 2020-12-01 - Apower MERGE block if different HF or DRYPACK.
'4000748 2020-11-12 - Fix bug.
'4000747 2020-11-12 - Fix bug.
'4000746 2020-11-12 - Quah/Ko Add FTBIN_QTY and FTBIN_INV to Baic_Loss_qty to track BIN3,4,5 FT REJECT qty and RMS Shipment (request by ERIC/EK for CHAINPOWER)
'4000745 2020-11-03 - Quah/Ko for check release lot reverse back to old method due to comma, comma onhold CJ base on \\Qfnsvr\d2\AIMS Test Out' as per EK TAN Requested
'4000744 2020-11-03 - Quah/Ko for check release lot without onhold CJ base on \\Qfnsvr\d2\AIMS Test Out' as per EK TAN Requested
'4000743 2020-11-02 - Quah/Ko for auto on hold and auto fill up reject qty for customer CJ base on \\Qfnsvr\d2\AIMS Test Out' as per EK TAN Requested
'4000742 2020-10-22 - Quah/Ko  During print APT, validate BOM BD revision with LI BD revision.
'4000741 2020-10-21 - ko add FOR npa vision page 2 to final test. ENABLE the vision 2page print for NPA, Requester Test Fiza
'4000740 2020-10-13 - ko for add npa vision page 2 to final test. disable the vision 2page print for NPA, Requester Test Fiza
'4000739 2020-09-24 - Quah - WriteToComtbl for EK Test Lots during print TT.
'4000738 2020-08-11 - KO - ADD THE M-CIRCUIT FORMAT FOR APT . Requested by Choong WT ON Targetdevice A+ is need to underline apt name MP FORMAT
'4000737 2020-07-14 - KO - change the .Connect = wsDB to .Connect = "PROVIDER=MSDASQL;dsn=AICSSQLDB;uid=ITAPP;pwd=APP" to support windows 10 verson
'4000736 2020-06-01 - KO - DB Created lot -change from ADODC1 & ADODC2 TO listview format
'4000735 2020-06-01 - KO - Test Traveller format -FOR CHAINPOWER & FETEK (SOT23) ONLY APPLIED FOR NEW CONDITION FOLLOW HAIGETE, AS PER KOAY FEEDBACK
'4000734 2020-06-01 - KO - Test Traveller format -TT-QFN-VISON change to TT-QFN-VISION2020 added the orentation as per hidthir requested
'4000733 2020-05-22 - Quah - CV check, skip for TAGGLE, request by Anita,Cyndi.
'4000732 2020-05-19 - KO CHANGE SOIC New FORMAT AS PER KOAY HB REQUEST for some admendment
'4000731 2020-05-12 - KO CHANGE SOIC FORMAT TO NEW FORMAT AS PER KOAY HB REQUEST tt-std-ft to TT-STD-FTSOIC2020
'4000730 2020-04-30 - KO/Quah - Deactivate CRYSTAL in TEST MASTER, to prevent OCX error on new Windows10 laptops.
'4000729 2020-04-30 - Quah - Deactivate CRYSTAL in TEST MASTER, to prevent OCX error on new Windows10 laptops.
'4000728 2020-04-30 - Quah - Fix logic for TEST-IN qty in SOT Yeild Report.
'4000727 2020-04-29 - Quah - GA-USA remove Brushing in APT process,req by LS,KP.
'4000726 2020-04-22 - Quah - Fix bug, for NPA barb port setting.
'4000725 2020-04-22 - Quah - Fix bug, for NPA barb port setting.
'4000724 2020-04-22 - Quah - User interface for NPA barb port setting.
'4000723 2020-04-20 - Quah - CS Yield Report for SOT.
'4000722 2020-04-16 - Quah - APT printed 2 pages, change back to double-sided.
'4000721 2020-04-02 - Quah - Apower show DRYPACK on APT Remark for SOT23S.AP2305GNS.CUPBF
'4000720 2020-03-26 - Quah - Allow all SOT23 to MRLT at 4400.
'4000719 2020-03-13 - Quah - Allow merge at 4400 for FETEK SOT23.
'4000718 2020-03-10 - ko/quah expanded the cuslotno len in crystal report
'4000717 2020-03-04 - ko/quah ADD counting for need to perform SQA
'4000716 2020-03-03 - ko/quah for sot23s package same category
'4000715 2020-02-29 - ko for add TEST PROGRAM EXTENDED LEN
'4000714 2020-02-27 - ko for add QFN NEW FORMAT RECOMPILE
'4000713 2020-02-26 - ko for add QFN NEW FORMAT RECOMPILE
'4000712 2020-02-25 -ko for add QFN NEW FORMAT
'4000711 2020-01-29 - Quah/Alif - For Haigete SPLIT lot, update blank for LTM_BIN_INDICATOR.
'4000710 2020-01-20 - Quah - TST excluded for Port printing in APT.
'4000709 2019-12-27 - Quah - BYPASS Haigete Bin3,4,5 checking if already RELEASE by ENGR before.
'4000708 2019-12-24 - Quah - Release for TT-STD-FTCJNEW.
'4000706 2019-12-16 - Quah - Haigete FT BIN03,04,05 lowyield control.
'4000705 2019-12-10 - Quah - Haigete - allow Merging at 4400/5500.
'4000704 2019-12-09 - Quah - Haigete certain PO allow merge without checking cuslotno.
'4000703 2019-11-29 - KO - REOPEN LOCK FUNCTION FOR MERGE FUNCTION FOR CJ AS PER KOAY REQUESTED
'4000702 2019-11-29 - KO - TEMPORARY UNMERGE FUNCTION OPEN FOR CJ AS PER KOAY REQUESTED
'4000701 2019-11-26 - KO - FIX New format for test traveller on sot23 as per koay requested
'4000700 2019-11-26 - KO - New format for test traveller on sot23 as per koay requested
'4000699 2019-11-21 - Quah - SubAssy report include Create Date.
'4000698 2019-11-13 - Quah - Raffar APT 4x4 no need "To Clear All Die"
'4000697 2019-11-01 - ko  -  FIXED BUG FOR MERGE LEN CJ LOTS
'4000696 2019-10-08 - ko  -  RANGE BY HAITGETE FOLLOW BY PACKAGE SOT23
'4000695 2019-09-26 - ko  -  residual Merge SOT23 without checking cuslotno as per koay requested
'4000694 2019-09-11 - ko  -  Merge SOT23 AT 2000 QTY 138240 , 2500 QTY - 276480 REQ KANG/SIVA/KP,Buyoff Koay
'4000693 2019-08-27 - Quah  -  Skip Yield checking for Split lot, req by Koay.
'4000692 2019-08-23 - KO    -  Test program too long , extend
'4000691 2019-08-22 - KO    -  Fixed bourns need to perform sqa traveller
'4000690 2019-08-21 - Quah    - Haigete MRLT open to 3500, 4400
'4000689 2019-08-21 - Quah    - Add back Crystal component in VIEW TRANS form.... Got Error.
'4000688 2019-08-21 - Quah    - Remove Crystal component in VIEW TRANS form.
'4000687 2019-08-16 - KO/Quah - HAIGETE FIX THE WRONG TRIGGER MESSAGE TEST NOT SETUP
'4000686 2019-08-15 - KO/Quah - HAIGETE NEW pull from the range baic_lot_info
'4000685 2019-08-15 - KO/Quah - HAIGETE NEW pull from the range baic_lot_info
'4000684 2019-08-15 - KO/Quah - HAIGETE NEW TEST TRAVELLER FORMAT AND NEW ROUTE T-SOIC-HC1 FOR CUTDOWN TRANSACTION
'4000683 2019-08-07 - Quah - Fix bug for HOLD RELEASE report.
'4000682 2019-07-11 - Quah - Bourns SW setting in MVOU-EDC.
'4000681 2019-07-10 - Quah - Wip Summary include Range.
'4000680 2019-07-10 - Quah - Onhold Wip include Range.
'4000679 2019-07-09 - Quah - Summary Output include Range.
'4000678 2019-07-05 - Quah - APT insert CLEANING after 3010, if package is less than 2x2.
'4000677 2019-07-03 - Quah - AgSn on APT Remarks for Greenpower 2N7002K-AGSN
'4000676 2019-07-03 - Quah - AgSn on APT Remarks for Greenpower 2N7002K2-HF
'4000675 2019-06-13 - Quah - AgSn on APT Remarks for Greenpower 2N7002Y-G, BSS84-G, 2N7002K-G
'4000674 2019-05-20 - Quah - TOP-BEST allow Merging of different cuslot, for PO (107091902AI,107091903 AI,107120501 AI)
'4000673 2019-04-01 - Quah - Enforce MICROCHIP lot merge at 3000 if have > 1 APT.
'4000672 2019-03-27 - KO/QUAH- ADD BOURNS AT GE,BE- ASSY,TEST, REJ AND ADD SKP LM06
'4000671 2019-02-25 - KO- WASSY-BE APT Format for bourns
'4000670 2019-01-24 - Quah - QRcode TT format for Bourns.
'4000669 2019-01-03 - Quah - QWAVE replaced MAXTEK for SAMPLE SPLIT.
'4000668 2018-12-07 - Quah - WASSYPT-SOT.rpt format link to CHANGJIANG.
'4000667 2018-12-05 - Quah - REQUEST SHIP DATE entry for DIODES.
'4000666 2018-12-04 - Quah - Bourns convert marking & cuslot during APT print.
'4000665 2018-11-16 - Quah - Bourns Custlot convert YMMDD during APT print.
'4000664 2018-11-01 - Quah - Clear bugs.
'4000663 2018-11-01 - Quah - Amphenol GEPT check BarbPort info only for PDM_CATEOGRY=NPA.
'4000662 2018-10-18 - Quah - APLUSTEK CS Yield Report shows B1,B2,B3 TEST-OUT qty.
'4000661 2018-09-19 - KO CHANGE TT-STD-GE Traveller base on cy and FIZA TEST REQ
'4000660 2018-09-06 - Quah - Block Deviceno deletion if lots exists in wip/fg/ship.
'4000659 2018-09-04 - KO - REOPEN AMPHENOL RMSTT CY PROJECT PAPER CUT TT-QFN & TT-STD-FT ,REMOVE THE RMSFOUND ='Y'
'4000658 2018-09-04 - KO - DEBUG RELEASE CY PROJECT PAPER CUT TT-QFN & TT-STD-FT ,REMOVE THE RMSFOUND ='Y'
'4000657 2018-09-04 - KO - DEBUG RELEASE CY PROJECT PAPER CUT TT-QFN & TT-STD-FT ,REMOVE THE RMSFOUND ='Y'
'4000656 2018-09-04 - KO - CY PROJECT PAPAER CUT TT-QFN & TT-STD-FT
'3000655 2018-08-29 - Quah - CS Report (Cycletime Rpt) fix for untested lots.
'3000654 2018-08-20 - Quah - Aplustek APT show 'Clear All Die'.
'3000653 2018-08-10 - Quah - Add Golden Unit for QWAVE, same as MAXTEK-QWAVE.
'3000652 2018-08-01 - Quah - Add B3 for SPLIT form.
'3000651 2018-07-26 - Quah - Fix typo in TT.
'3000650 2018-07-23 - Quah - Add additional TT for NPA-1-C02142 (AIC), with calculation for 16 CYCLES, req by Fizah Test.
'3000649 2018-07-09 - Quah - Fix bug for Splitlot seq > 100.
'3000648 2018-07-06 - Quah - Amphenol APT format for Dual Port, Dual Dispensing WASSYPT-GE-2P2D.rpt
'3000647 2018-05-04 - Quah - APT 3010 add "CLEANING", req by KC.
'3000646 2018-04-24 - Quah - SKIP Miniapt for GE RWK.
'3000645 2018-04-16 - Quah - APT print double side.
'3000644 2018-04-04 - Quah - Add WaferNo. for oper 1300 in GE_PROCESS_MASTER (Wipstream.mdb) of Amphenol APT.
'3000643 2018-03-21 - Quah - Mini report change path to aicwksvr2016
'3000642 2018-03-13 - Quah - Fix bugs in SubAssy Recv (for saving Amphenol).
'3000641 2018-03-08 - Quah - Add Custname in LotSearch query result.
'3000640 2018-02-22 - Quah - Remove logic for Tactilis auto-deduct Kitting Inventory during APT Print.
'3000639 2018-02-13 - Quah - Lot enquiry (OTHER) for AD/MC takes from CLS_LOT_INFO, instead of BAIC_LOT_INFO.
'3000638 2018-02-08 - Quah - Tactilis print APT: Deduct qty from BAIC_KITTING_HEADER, BAIC_KITTING_DETAIL.
'3000637 2018-01-19 - Quah - Dont allow SPLIT at ASSY, to prevent abnormal YIELD REPORT.
'3000636 2018-01-04 - KO - RELEASE FOR TL reject EDC system
'3000635 2017-12-22 - KO - RELEASE FOR TL reject EDC system
'3000634 2017-12-12 - Quah - Block RVSE for Microchip, due to B2B trans.
'3000633 2017-12-06 - Quah - Microchip APT waferlotno show total die qty. Req by Nabila.
'3000633 2017-12-06 - Quah - Remove APT Remarks for Diodes ("USE BLACK REEL" for targetdevice "AP" and package 3X3) req by Hidthir.
'3000632 2017-11-21 - Quah - TT2-SCRF, SCRF test traveller add in TEST PROGRAM barcode (unicode), req by Akmal.
'4000631 2017-11-14 - Quah - Fix bugs (auto focus to CANCEL button after key in MVOU Reject Qty).
'4000630 2017-11-14 - Quah - Use Arial Narrow font on TargetDevice for WASSYPT-GE.rpt (some housing device is long).
'4000629 2017-10-23 - Quah - Amphenol Port/Disensing picture on APT: include back for ENGR lots.
'4000628 2017-10-11 - Quah - Separate TT (with Foundry Lot) for Microchip.
'4000627 2017-10-11 - Quah - Block APT print if Microchip load using Atmel waferlotno, req by Nabila.
'4000626 2017-10-03 - Quah - Check for missing MARK FILE on APT, skip if lot not go through EOL oper.
'4000625 2017-09-26 - Quah - Update columns for TT-GE-VISION.
'4000624 2017-09-19 - Quah - Set TT for Microchip, follow Atmel SICM/SCRF setting.
'4000623 2017-09-12 - Quah - Exclude Microchip lots in B2B recv, moveout. Will be process by auto-schedule B2B program.
'4000622 2017-09-07 - Quah - Add Microchip (MC) to Atmel (AD/AS) MOVEOUT logic.
'4000621 2017-09-07 - Quah - Fix bugs.
'4000620 2017-09-07 - Quah - Fix bugs.
'4000619 2017-09-07 - Quah - APT print setup for MICROCHIP (new Atmel).
'4000618 2017-08-23 - Quah - SRCV will be default to CLOSE for CAP ("SOIR14.NPA-CAP.GR00","SOIR14.NPX-CAP.GR00"), req by Azizah. Because no issuance of qty will be done after that.
'4000616 2017-08-21 - Quah - Release for APT and Mini-APT for Tactilis.
'4000615 2017-08-08 - Quah - Skip ENGR for (Check Marktemplate file against IT_BARCODE_MARK.)
'4000614 2017-08-03 - Quah - Skip NS for (Check Marktemplate file against IT_BARCODE_MARK.)
'4000613 2017-08-31 - Quah - Check Marktemplate file against IT_BARCODE_MARK.
'4000612 2017-07-31 - Quah - Skip for CAP for Device registration checking in APT.
'4000611 2017-07-18 - Quah - Amphenol use generic APT for NPX, HSE (without Port/Dispensing pictures).
'4000610 2017-06-21 - Quah - CS Report: Add 'ALL WIP ASSY CT'
'4000609 2017-06-21 - Quah - Fix bugs.
'4000608 2017-06-20 - Quah - Amphenol check for Port Type and Dispensing Type in comtbl, for APT format.
'4000607 2017-06-15 - Quah - Fix bug. Amphenol device NPA-700M-C02105 use 2nd Dispensing APT format.
'4000606 2017-06-14 - Quah - Amphenol device NPA-700M-C02105 use 2nd Dispensing APT format.
'4000605 2017-05-18 - Quah - Improve query for WIP Dimension Report.
'4000604 2017-05-11 - Quah - Tactilis APT - insert additional process columns for Oper 1300, 1400.
'4000603 2017-05-11 - Quah - CS Yield Report fix deduct merge lots only for mother lots. (eg lot NK6340401).
'4000602 2017-04-18 - Quah - Fix bugs.
'4000601 2017-04-17 - Quah - Fixup missed-out logic for Atmel SJ, to deduct IQA inventory during APT print.
'4000600 2017-04-17 - Quah - Remove bypass.
'4000599 2017-04-14 - Quah - temporary bypass IQA wafer checking, for urgent loading, req by Nabila.
'4000598 2017-04-07 - Quah - remove NPA port,base no. in APT, req by Syamil.
'4000597 2017-04-03 - Quah - fix logic bug, only SCM need to trunctate after dash for TEST TARGET DEVICE, refer LY.
'4000596 2017-03-21 - Quah - Block REVERSAL for lots already SRCV, to prevent mismatch with SUBASSY transactions.
'4000595 2017-03-21 - Quah - SQA for Atmel SOIC14 (ASJ) change from 100% to 1100, requested by Akhmal.
'4000594 2017-02-24 - Quah - Logic to handle P2 Wip Sheet2 EOF.
'4000593 2017-02-23 - Quah - GoodArk GA01 SQL enquiry.
'4000592 2017-01-26 - Quah - Fix APT printing bugs on IQA wafer qty check.
'4000591 2017-01-18 - Quah - ATMEL SJ (SOIC14) match by IQA cuslotno during APT Print.
'4000590 2017-01-12 - Quah - Include TST for waferlot DBANK inventory checking (req by Nabila, ATMEL SJ is TST).
'4000590 2017-01-12 - Quah - Release WASSYPT-GE.rpt, which include NPA PORT and NPA BASE partno.
'4000589 2017-01-12 - Quah - additional 2-9 running no. for Atmel SJ marking $, request by Khairul.
'4000588 2017-01-10 - Quah - Initialize form after merge successfully in Atmel SJ Merge.
'4000587 2017-01-05 - Quah - Add TargetDevice to Wip Hold report.
'4000586 2016-11-17 - Quah - Amphenol (RWK, TST) 3010 MVOU skip checking for Subassy Lot, req by Anita.
'4000585 2016-11-07 - Quah - Assy report skip certain reject code, to avoid run time error.
'4000584 2016-10-27 - Quah - fix bugs.
'4000583 2016-10-26 - Quah - TT modifications (TT-VISION.rpt, TT-STD-FT.rpt, TT-STD-GE.rpt)
'4000582 2016-10-18 - Quah - add SPLIT function for Atmel SJ (SOIC14).
'4000581 2016-10-12 - Quah - set 100% SQA for Atmel SJ (SOIC14).
'4000580 2016-10-06 - Quah - Phase verification for ATMEL SOIC14 Splitting.
'4000579 2016-10-05 - Quah - TT for Atmel SJ (SOIC14).
'4000578 2016-09-29 - Quah/Ko - Update for TT-STD-GE.rpt
'4000577 2016-09-28 - Quah - Link change lotno instead of lotno9 for TT-STD-GE.rpt
'4000576 2016-09-26 - Quah/Ko - Release for TT-GE-NPX-3PRG.rpt
'4000575 2016-09-15 - Quah/Ko - Add #TMARK4# for NPX (instead of #TMARK3#)
'4000574 2016-09-08 - Quah/Ko - Improve logic: Cuslotno #TMARK3# populate from TOPM3 (NPA), TOPM4 (NPX)
'4000573 2016-09-07 - KO - NV custlono replace with Topm3 req. by KP TAN & Anita
'4000572 2016-08-24 - Quah - Subassy Withdrawal: Add logic to verify Pcell-APT# match with ASIC-Part# in Aic_Wip_Header.
'4000571 2016-08-15 - Quah - Bourns SOIK MD03 (Oper 2100) rej trigger limit at 0.3%, req by Khairul/Cust.
'4000570 2016-08-10 - Quah - Modify MINIAPT_GE_HOUSING.rpt, req by KP Kok.
'4000569 2016-08-05 - Quah - minor fix for rec sequence in BAIC_GE_TEST_SUBLOT.
'4000568 2016-08-04 - Quah - Remove sorting for TT-STD-GE.rpt, to follow ori sequence in BAIC_GE_TEST_SUBLOT.
'4000567 2016-07-13 - Quah - Bugs.
'4000566 2016-07-13 - Quah - New format for TT-STD-GE.rpt
'4000565 2016-07-13 - Quah - DB Mvou List exclude On Hold lots.
'4000564 2016-06-27 - Quah - Update query for CS Yield report (include TRLT lots).
'4000563 2016-06-08 - Quah - Add filter out ACTIVE parts only in bom, for version 4000562.
'4000562 2016-05-31 - KO&Quah - GE PORT FORMAT AS PER KP TAN REQUESTED
'4000561 2016-05-23 - Quah - Fix bugs.
'4000560 2016-05-20 - Quah - Fix bugs.
'4000559 2016-05-20 - Quah - CV function open to Planners, because Production is block if CV more than 1%. Checking excl M-Circuits. Req by Mon/Nabila.
'4000557 2016-05-13 - Quah - Validation for certain RMA lots can only select RMA Test Reference.
'4000556 2016-05-11 - Quah - Modify and release for TT-GE-3PRG format.
'4000555 2016-04-29 - Quah - Fix bugs.
'4000554 2016-04-29 - Quah - Link to BLOCK-LIST in MVOU form, to prevent MVOU for problem lot.
'4000553 2016-04-26 - Quah - Add 'PRCV' to output report.
'4000552 2016-04-20 - Quah - Improve logic for Atmel diebank lot size for split.
'4000551 2016-03-31 - Quah - Count Variance for Prod lots only allow for 1%.
'4000550 2016-03-24 - Quah - Add "ALL" to Plant (P1/P2) Combobox in Output report.
'4000549 2016-03-09 - Quah - Atmel SL Number on APT pull from ABI (cls_lot_info).
'4000548 2016-03-08 - Quah - Fix bugs.
'4000547 2016-02-12 - Quah - Add 018.0004.0001 for "68 UF" watermark.
'4000545 2016-01-29 - Quah - Link Amphenol APT (WASSYPT-GE-port.rpt) if Bom contains Dual-Port (025.0006.0000)
'4000544 2016-01-26 - Quah - Add Hold Code to Release report.
'4000543 2016-01-14 - Quah - Modify Hold Release report to display CT as per individual hold tranx, req by Ahnaf.
'4000542 2015-12-14 - Quah - Reactive mini-apt for Ge/Amphenol, req by Anita,Azizah.
'4000541 2015-12-14 - Quah - Reactive mini-apt for Ge/Amphenol, req by Anita,Azizah.
'4000540 2015-12-01 - kO - Request by Vivian add SOIC AND NPA/NPX AT TT-VISION TRAVELLER
'4000539 2015-11-30 - Quah - Lot enquiry reject add in Desc for Defect Mode.
'4000538 2015-11-26 - Quah - Include MC in Cls System logic.
'4000537 2015-11-25 - Quah - Add IDESYN for Sample Split.
'4000536 2015-11-18 - Quah - Bypass CLASS SYSTEM checking for Microchip 6020, due to history mismatch.
'4000535 2015-11-05 - Quah - Include Microchip to Micrel (Class System) logic.
'4000533 2015-10-29 - Quah - IDESYN Golden Unit - 20 at Test Proc.
'4000532 2015-10-23 - Quah - CS Report (CT2) calculation error due to null test date.
'4000531 2015-10-22 - Quah - AIMS Query for RMS modified to let user select customer. Previously fix for Amphenol only.
'4000530 2015-09-14 - Quah - SICM Test Report - add date filtering (why not there???).
'4000529 2015-08-25 - Quah - Remove SelamatHariRaya.
'4000528 2015-07-31 - Quah - Add remarks to M-Circuits APT in Rpt.
'4000527 2015-07-13 - Diana - Selamat Hari Raya :)
'4000526 2015-07-09 - Quah - Bugs.
'4000525 2015-07-09 - Quah - Increase TopMark spacing, req by SE Lim, to cater to MQ manual underline.
'4000524 2015-07-06 - Quah - New CT report (all customers) in CS-Report screen, req by EK.
'4000523 2015-07-03 - Quah - Remove NIKO RMA block, req by LY.
'4000522 2015-07-01 - DIANA - TT-STD-GE add work order for retest lot and handler.
'4000521 2015-07-02 - Quah - Atmel RVSE not allowed due to B2B Reporting.
'4000520 2015-06-30 - QUAH - APT for TRAY lots to incude 2 dummy operations before Oper 3000 (no need transactions), req by SE LIM.
'4000519 2015-06-30 - QUAH - NEW 'CT REPORT2' in CS Enquiry, req by EK.
'4000518 2015-06-26 - QUAH - Temp block NIKO Po 3515-104060176 to clear old lots first, due to marking different. Refer Alice email.
'4000517 2015-06-23 - QUAH - Validate Print-TT 'Test-Reference' must be RMA for UPI retest lots. Req by LY.
'4000516 2015-06-17 - DIANA - MINI-CIRCUITS ADD REMARK AT APT "TO CLEAR ALL DIE"
'4000515 2015-06-05 - Quah - Fix bug in SAMPLE SPLIT. Get cusname in AIC_WIP_HEADER by mother lot.
'4000514 2015-05-27 - Diana - GE Test Traveller new format, refer Ang.
'4000513 2015-05-26 - Quah - Released FOLYAS-NV.
'4000513 2015-05-26 - Quah - Released input function for BOURNS Scribe# (starting from BE522...).
'4000512 2015-05-20 - Quah - Temporary Block NV 24 lots due to discolouration.
'4000512 2015-05-20 - Quah - TT-QFN add in Datecode info (useful for UPI combine rules).
'4000512 2015-05-12 - DIANA - GE Family change APT, wassypt-ge.rpt & folyas-nv.rpt, remove mini-apt link from GE housing
'4000510 2015-05-11 - Diana - APT auto to printer.
'4000509 2015-05-11 - Diana - Bourns APT changes, wafer# to scribe lot.
'4000508 2015-05-06 - Quah - Improvement: APT format added input box for Bourns DTM/Scribe record.
'4000507 2015-05-06 - Quah - TestMaster setup block for 3x3 RU devices, requested by LY, for Orientation Setting follow-up.
'4000506 2015-05-06 - Quah - APT format added input box for Bourns DTM/Scribe record.
'4000505 2015-05-05 - Quah - Syamil req to unblock.
'4000504 2015-04-30 - Quah - Syamil req all GE,AMPHENOL lots blocked at 3010.
'4000503 2015-04-20 - Quah - TT-QFN.rpt separate the Merge, Residue to avoid overlapping info.
'4000502 2015-04-14 - Quah - Enquiry PN01 Refreshh SAP, trim the GRN# to avoid length error.
'4000501 2015-04-14 - DIANA - GE- ASSY, TEST, REJS REPORT ERROR DUE TO FUNNY MACHIINE..
'4000500 2015-04-03 - Quah - Block AD and AS for SPLIT and MERGE (due to B2B manual handling for transactions).
'4000499 2015-04-02 - Quah - Bourns 3010 - remove block (for AG Peeling case), requested by Syamil.
'4000498 2015-03-16 - Quah - Atmel lots block at 3500 if ABI OWNER <> MFG (req by Vanitha).
'                   - Quah - LY req Golden Unit remark2 should be independent of SQA condition.
'        2015-02-26 - Diana - Add Maxtek for calling chkGoldenUnit.
'4000497 2015-02-09 - Diana - HAPPY CNY! GONG XI FA CAI!
'4000496 2015-01-22 - Quah - Only allow transactions if lot_status='ACT' (instead of just base on ltm_deleted='N')
'4000495 2015-01-17 - Diana - Recordset still open, trans not committed for below query.
'4000494 2015-01-16 - Quah - Special function (Prd_Hold_Status.frm) for Desmond to release Bourns lot (blocked at 3010 due to AG-Peeling issue).
'4000493 2015-01-15 - Diana - Atmel SJ marking, if-endif closure.
'                           - TI's APT need watermark to show capacitor output requirement.
'4000492 2015-01-14 - Quah - Set a bypass logic for selective Bourns lots currently block at 3010.
'4000491 2015-01-14 - Quah - Update Aic_Wip_Header.Orderno to Refno"X" if lot TERMINATE at 900, to prevent old-lot linking.
'4000490 2015-01-13 - Quah - Atmel marking conversion check for NULL.
'4000489 2015-01-09 - Quah - Bourns change block to 3010, req by Anita.
'4000488 2015-01-08 - Diana - BOURNS ON-HOLD
'4000487 2014-12-29 - Diana - HAPPY NEW YEAR!
'4000486 2014-12-23 - Diana - Atmel check MVOU and PROC against ABI.
'                           - Atmel marking conversion buggy.
'4000485 2014-11-28 - Diana - Atmel marking conversion buggy.
'        2014-12-12 - Diana - Zuraini Bourns + machine report.
'        2014-12-15 - Diana - Diodes APT print remark "USE BLACK REEL" for targetdevice "AP" and package 3X3.
'4000484 2014-11-27 - Diana - Micrel blocked lots can be released. ref: v.4000480
'4000483 2014-11-25 - Diana - QFN new traveller format + Atmel marking conversion new logic.
'4000482 2014-11-20 - Quah - Revise logic for Atmel PLCC marking conversion YYWW (check for top or bottom).
'4000481 2014-11-20 - Quah - Revise logic for Atmel PLCC marking conversion YYWW (check for top or bottom).
'4000480 2014-11-14 - Quah - QA req to block MU4420803, MU4420804 for Split, Merge.
'4000479 2014-11-14 - Quah - Marking Conversion for ATMEL SJ at 2100 MVOU (5+1 new ABI requirement).
'4000478 2014-11-13 - DIANA - PO 1MM12-141000001 can only select RMA (UPi) at test traveller
'4000477 2014-11-03 - Diana - GE SUB_ASSY add 'Export to Excel' and summary of total by device
'4000476 2014-10-30 - Diana - MQ follow AP golden unit logic & add Maxtek, refer LY Chin
'4000475 2014-10-14 - DIANA - GMT NEED CHECKING FOR VACUUM MBB BAG OR NOT, REFER QA
'4000474 2014-10-07 - ko - Modify for customer lbl for TOPSYSTEMS,AMPHENOL
'4000473 2014-09-22 - Quah - Print APT add 2 new col during INSERT AIC_INVENTORY_MASTER.
'4000472 2014-09-19 - Diana - Disable option to change cuslot due to no parallelity with label.
'4000471 2014-08-28 - Diana - re-compile project due to missing codes :(
'4000470 2014-08-27 - Diana - NPX take TOP4, NPA take TOP3 marking.
'4000469 2014-08-25 - Diana - Undo by pass of 5 lots (refer below(4000468)).
'4000468 2014-08-15 - KO/Diana - by pass checking for device SCM2.AT2961009.4T01EC,AT29610-094T-03E-,s/b AT29610-094T-01E
                    '- AT29610-094T-03E for 5 lots requested by Anita AD4100901,AD4101001,AD4101201,AD4101301,AD4101801
'4000467 2014-08-06 - Diana - add PN06 to PN03 devices no.
'4000466 2014-08-06 - Quah - Clear bugs.
'4000465 2014-07-25 - Diana Bourns FTVM report add BIN3 and BIN4 (refer to Vivien)
'4000464 2014-07-24 - Diana/Quah - BOURNS Yield Rpt fix filtering for completed lots.
'4000463 2014-07-22 - Quah - Get devices from PN01 to PN03. Diana - Fix 'Refresh SAP Data' checking on report.
'4000462 2014-07-16 - Quah - Add MRPK 103070005 for APOWER rework.
'4000461 2014-07-15 - Quah - Add DIODES to take Markfile from Lot_Info during APT Print.
'4000460 2014-07-08 - Diana - PN06-Monthly Loading report.
'4000459 2014-07-04 - Diana - Add remove LI at oper >=3500 (SQL Report).
'4000458 2014-07-03 - Quah - APOWER rework special handling in TT for 3 PO (MRPK 103060016,MRPK 103060017,MRPK 103060018)
'4000456 2014-07-02 - Quah - Hardcode remark on APT for Apower rework 3 PO, refer LY.
'4000455 2014-07-02 - Quah/Diana - Add EOH and CONV_FACTOR to PN03.
'4000454 2017-06-27 - Diana - PN03 add to auto-delete from AIC_LI_RESERVED for oper>3500.
'4000453 2014-06-27 - Quah/Diana - Update changes in PN03, remove by Device only and BAIC_LOTMAST table conn.
'4000452 2014-06-19 - Diana - Change cell format to date for PN03.
'4000451 2014-06-18 - Quah/Diana - SAP refresh add grouping for ItemCode and sum Qty (Refer to EK Tan).
'4000450 2014-06-17 - Diana - New report PN05 and edit PN03 (Refer to EK Tan).
'4000449 2014-06-04 - Diana - Print NA for GE and GE_newpix APT for 3 part numbers.
'4000448 2014-06-04 - Quah -  Set SSPL access for Sample Split form.
'4000447 2014-06-03 - Diana - Disable Fitipower from chkGoldenUnit-TestRoute (Refer LY Chin).
'4000446 2014-06-02 - Diana - Exclude Fitipower from Sample Lot Split.
'4000445 2014-05-29 - Quah - Include 'LOT LABEL' for TopSystem, Amphenol to BAIC_LOT_INFO during APT Print.
'4000444 2014-05-29 - Din - set A-power HF_INDICATION in baic_lot_info to 'NO' if REM4 in AIC_LI_LABEL_INFO is 'NO'.
'4000443 2014-05-23 - Quah - SQL Query for PLN03, GRN data refresh from SAP.
'4000442 2014-05-22 - Quah - Add TA and NV to TT-GE-VISION logic.
'4000441 2014-05-22 - Din/Diana  -  GoldenUnit Sample Lot function (Split,TT,Mvou,TransView) Note: MVOU to 7000 for Sample lot.
'4000441 2014-05-22 - KO  -  FIN01 Aims Query for Product List, refer Lalitha.
'4000440 2014-05-20 - KO  -  Add TT-GE-VISION Traveller at vision , remove the tt-ge-n2 traveller at final test
'4000438 2014-04-29 - Quah - Add TopSystem,Nova to subassy Recv,Withdraw forms.
'4000437 2014-04-24 - Quah - Minor change to format TT-STD-GE-N2.rpt, refer Fifi.
'4000436 2014-04-24 - Quah - SQL Query: Amphenol Yield Report (Anita).
'4000435 2014-04-23 - Quah - SQL Query for PLN, data refresh from SAP.
'4000435 2014-04-23 - Ko   - SOIC TT residue bugs.
'4000435 2014-04-23 - Ko   - NPA/NPX TT additional page.
'4000434 2014-04-21 - Quah - Improve Micrel ClassSystem checking on Residue Lot Mvou.
'4000433 2014-04-21 - Quah - clear bugs.
'4000431 2014-04-15 - Quah - PLN03 Sql Query: Material Listing by multiple Device (req by EK)
'4000430 2014-04-09 - Quah - Impinj update Estimate Ship Date (for Wip Rpt) from REM9 (Aic_Loading_Instruction_Remark) during APT Print.
'4000429 2014-04-07 - Quah - Mini AIMS Report (run by Query).
'4000428 2014-04-02 - Quah - APT setting for GE-TOPSYSTEM.
'4000427 2014-03-26 - Quah - CS Report add Shipqty.
'4000426 2014-03-25 - Quah - Block Micrel at 6020 for lot >= 11 chars, multiple split create unmatch data in CLASS System.
'4000425 2014-02-20 - KO -GE Test traveller for NPX ADVANTEST
'4000424 2014-02-14 - Quah - Unblock NPA at 5520.
'4000423 2014-02-13 - KO - Ge NPX, Indicate ETEC, ADVANTEC PROGRAM
'4000422 2014-01-29 - Quah - GS req unblock for MX,MQ.
'4000421 2014-01-28 = Quah - GS req block for MX,MQ due to label info concern.
'4000420 2014-01-23 - Quah - GS req block for GE-NPA at 5520 due to quality problem.
'4000418 2014-01-22 - Quah - Baic_TT_Result add columns for Finance Test Time.
'4000417 2014-01-08 - Quah - TT GE-NPX format add Dry Bake period (24 hrs).
'4000416 2013-12-27 - Ko - Test TIme, test program for finance fairchild billing invoice
'4000415 2013-12-23 - Quah - Add BIN2 to CS report.
'4000414 2013-12-19 - Quah - Prodmast setup, max length 30, segment 9+14+7.
'4000413 2013-12-18 - Quah - Anita req, for NN and AS (PLCC) same cuslot must MERGE first before allow MVOU Oper 3000 as single lot.
'4000412 2013-11-29 - Quah - DIODES update Cuslotno and LOT LABEL follow AICLOT (during APT Print). Refer Chua/Alice.
'4000411 2013-11-29 - Quah - Add in Bourns Yield Rpt (for Finance Reject Billing), refer Celine/Mary.
'4000410 2013-11-26 - Quah - DB Mvou unblock for OSRAM lots.
'4000409 2013-11-22 - Quah - Release for FOLYAS-WGLASS (Osram).
'4000408 2013-11-22 - Quah - Include OSRAM for Planner to Split/APT Print. (due to WINDOW-GLASS Project).
'4000407 2013-11-21 - Quah - Undo Bourns Test Yield changes, due to TestEngr using it.
'4000406 2013-11-20 - Quah - Bourns Test Yield change query to pull base on TestOut date, and display by AIC Lot (req by Mary, Celine).
'4000406 2013-11-20 - Quah/Ko - Positive count variance not allow for GE (req by Christine).
'4000406 2013-11-20 - Quah/Ko - GE-NPX TestTraveller add additional boxes at the bottom (req by Rosina/EK Tan).
'4000405 2013-11-08 - Quah/Ko GE-NPX Special Edc only for oper 4400.
'                      Note: To debug on MVOU EDC parameter, prd_mvou_edc point on the active form and load form,
'                      need to click on NextPage & PreviousPage button.
'4000404 2013-11-07 - Quah - Req by GSYU unblock GE lots at 5520 MVOU.
'4000403 2013-11-01 - Quah/Ko - GE NPX TT remove SQA remark.
'4000402 2013-10-29 - Quah/Ko 1620 default LowYield set at < 99%, for those Cust/Pkg not registered in Comtbl.
'4000401 2013-10-29 - Quah/Ko 1620 default LowYield set at 99%, for those Cust/Pkg not registered in Comtbl.
'4000400 2013-10-24 - Quah - Logic for GE APT,MiniApt differentiate by NPA,NPX.
'4000399 2013-10-23 - Quah - Amend formula for GE-NPX marking running# follow Workweek,Mother7 (exclude TgtDevice). Refer KP Tan.
'4000398 2013-10-23 - KO - GE Test Traveller printing for vision only 'NPX'
'4000397 2013-10-22 - Quah - GE APT marking running# YYWW$$ for NPX, by TgtDevice, WW, Mother7
'4000396 2013-10-18 - Quah - New TT format for GE-NPX.
'4000395 2013-10-17 - KO- GE Final TEST EDC, Same operation 4400,categ=groupcateg_lbl ="NPX"
'4000394 2013-10-11 - Quah - Increase FT Program input length in TestMaster Setup.
'4000393 2013-10-09 - Ko - 2013-Oct-09 add pono Column at assembly yield report and operation desc at po tracking LOT# BY Eric
'4000392 2013-10-01 - Quah - Alps Label Printing Enquiry.
'4000391 2013-09-24 - Quah - Block if running from DEPLOY.
'4000391 2013-09-24 - Quah - Update Prodmast, to include data for PDM_PACKAGE_DIMS.
'4000390 2013-09-23 - Quah - Include Oper7000 (Staging) in Internal Wip report.
'4000389 2013-09-19 - Quah - GE NPX auto running# by week, by device in APT.
'4000388 2013-09-18 - Quah - GE NPX auto running# by week, by device in APT.
'4000387 2013-09-04 - Quah - Machine inquiry by lot - include Mold machines.
'4000386 2013-08-29 - Quah - Increase spacing on FinalTest TT DutBoard field.
'4000385 2013-08-26 - Quah - FT Special Program auto-insert child lot, if mother lot is SPLIT.
'4000384 2013-08-16 - Quah - GE NPX Topm3 $$ runningno. (group by YYWW, DEVICE Topm2).
'4000383 2013-08-14 - Quah - If Golden Unit, also must have SQA.
'4000382 2013-08-12 - Quah - Special FT Test Program linking (TMH_DATA_2) for Traveller.
'4000381 2013-08-06 - Quah - Update logic: IDA CV block only if new qty > creation qty.
'4000381 2013-08-06 - Quah - Unblock APE8890 lots at 5520.
'4000380 2013-08-02 - Quah - GSYU req to block APE8890 lots at 5520.
'4000379 2013-07-26 - Quah - Taggle TT-Vision TT.
'4000378 2013-07-24 - Quah - Fairchild lot > 1 year, allow MVOU if customer give disposition (register in Alarm table).
'4000377 2013-07-23 - Quah - CS Yield Report - include new filtering by Loading Date.
'4000376 2013-07-23 - Quah - Bourns Test Yield Report - add DEVICENO (to diffentiate Copper/Gold), req by Celine.
'4000375 2013-07-23 - Quah - APT 3010 add "REMOVE REMNANT", req by SK Hong / Siva.
'4000374 2013-07-19 - Quah - Add Shipdate to CS Yield report.
'4000373 2013-07-18 - Quah - Bourns update PO Shipdate function.
'4000372 2013-07-18 - Quah - Fairchild merging, also check for mother lot for > 365 days.
'4000372 2013-07-18 - Quah - IDA wafer cannot allow positive BYLT, req by Lawrence.
'4000372 2013-07-18 - Quah - FITIPOWER TT golden unit for every cuslot (30 unit).
'4000370 2013-07-05 - Quah - GS Yu req temporary block for GE lots at 5520.
'4000369 2013-07-04 - Quah - Lot Enquiry (OTHERS) - link to Baic_Lot_Info instead of aic_li_labelinfo.
'4000368 2013-07-04 - Quah - CS CT report - include SFO time, req by SE LIM, for CT monitoring.
'4000367 2013-07-01 - Quah - Fairchild Merge, Mvou block if lot > 1 year.
'4000366 2013-06-18 - Quah - Add BINSTOCK to Wipreport.
'4000365 2013-06-18 - Quah/Ko - Transfer to 9100 BINSTOCK.
'4000364 2013-06-11 - Quah - Include BIN in wiplot report.
'4000364 2013-06-11 - Quah - Add Diepartno to Miniapt, req by Heng/Kang.
'4000364 2013-06-11 - Quah - Change Bourns SQA every 1,6 lots, req by EK Tan.
'4000364 2013-06-11 - Quah - Add Attach-Doc function for all lots, req by SHYeo (see comment for ver 4000361)
'4000363 2013-05-24 - Quah - CS Plant CT report include PO#.
'4000362 2013-05-23 - Quah - Include 950 (BG) in report oper range.
'4000361 2013-05-22 - Quah - Document attachement function for Test Engr during Lot Release (req by SH Yeo).
'4000360 2013-05-16 - Quah - Add SKIPWFR recordtype in [baic_customer_addsetup] to bypass Wafer qty check during APT print.
'4000359 2013-05-16 - Quah - No need RMS traverller during ReleaseHold.
'4000358 2013-05-13 - Quah - APT temp MDB, increase cuslotno to 30 (M-Circuit long cuslotno).
'4000357 2013-04-26 - Quah - Improve GoldenUnit logic (MAXTEK/APOWER) to skip if same cuslotno already done.
'4000355 2013-04-24 - Quah - Lot creation insert to PO_Master : add date eff,exp during Price linking.
'4000354 2013-04-23 - Quah - Improve logic to handle Micrel residue-residue combination.
'4000353 2013-04-17 - Quah - NIKO Rwk special case allow MRLT for 3 special PO.
'4000352 2013-04-16 - Quah - Improve logic for Micrel MVOU 6020 [for link to CLASS SYSTEM].
'4000351 2013-04-15 - Quah - Reset GoldenUnit logic for Apower, Maxtek-Qwave (refer Hidayah).
'4000350 2013-04-12 - Quah - Improve Golden Unit logic for APower, Qwave.
'4000349 2013-04-11 - Quah - Micrel Untested Mvou from 3010 to 7000.
'4000348 2013-04-08 - Quah - ENE marking replace YYMMDD (APT print date).
'4000347 2013-04-03 - Quah - Improve PCELL (GE-Housing) info in Enquiry Function.
'4000346 2013-04-02 - Quah - APT exclude Capillary 052.
'4000346 2013-04-02 - Quah - Add Mapfile Find function in Enquiry.
'4000345 2013-03-27 - Quah - Minor bugs.
'4000344 2013-03-21 - Quah - Block back for untested Micrel, need further testing at 3010 MVOU.
'4000343 2013-03-21 - Quah - Mvou 3010 for Micrel Untested, update to CLASS SYSTEM - ASSY to PACK.
'4000342 2013-03-13 - Quah - Add XIRKA to TT logic, follow Atmel COL
'4000341 2013-03-12 - Quah - Fix minor bugs.
'4000339 2013-03-04 - Quah - Micrel SOIC use MOLDOUT date for YYWW in Marking and Datecode.
'4000338 2013-02-08 - Quah - Release for CNY greeting.
'4000337 2013-01-31 - Quah - Engr Rej Hold Back TT - include in Reprint RMS form.
'4000336 2013-01-25 - Quah - Impinj APT Print:  ,e3, convert to ,, in marking.
'4000336 2013-01-24 - Quah - GE NPA/NPX identified by PDM_Targetdevice instead of Deviceno.
'4000335 2013-01-24 - Quah - Set GE NPA 28 per strip, NPX 14 per strip, req by Vernon.
'4000334 2013-01-22 - Quah - Micrel bugs.
'4000333 2013-01-22 - Quah - GE MVOU bugs.
'4000332 2013-01-18 - Quah - bugs.
'4000331 2013-01-17 - Quah - Micrel 4700, block if Residue lot MVOU before mother lot (Mother Lot must MVOU first, so that the Child record can be registerd in CLASS System Cls_Lotmast)
'4000330 2013-01-17 - Quah - GE Sub Withdraw for CAP (no need to link to PCELL Apt).
'4000329 2013-01-16 - Quah - GE NPA,NPX detect by CAP,HSE only for SP lot generation.
'4000329 2013-01-16 - Quah - GE 3100 main lot Mvou check for already perform Sub-Assy Withdraw.
'4000328 2013-01-11 - Quah - Add Catg, Detape to RLLT report.
'4000327 2013-01-09 - Quah - Set AVT=NIKO for all customisation (eg. APT, TT, SQA)
'4000326 2012-12-26 - Ko   - Hold Release category selection, req by EK.
'4000325 2012-12-21 - Quah - QFN Test Rpt set HW-1 = TESTOUT for Combine ChildMother lots.
'4000324 2012-12-20 - Quah - Changes in APT FOLYAS, as per latest Engr spec.
'4000323 2012-12-17 - Quah - Include TargetDevice in GE APT remarks process by NPA,NPX.
'4000322 2012-12-14 - Quah - Lot Creation 'SP' detect by -CAP, -HSE for GE. Don't fix NPA-CAP or NPX-CAP.
'4000321 2012-12-14 - Quah - Remove SQA msg updating bugs.
'                            GE customisation for NPX APT.
'4000320 2012-11-27 - Quah - Race merging checking for NULL bin indicator.
'4000319 2012-11-27 - Quah - GoldenUnit added for Maxtek-Qwave 20units.
'                             Apower, Maxtek SQA 500 units.
'                             Remove QFN/SOIC filter for Apower Golden Unit and apply for all pkg.
'4000318 2012-11-23 - Quah - Clear bugs.
'4000317 2012-11-23 - Quah - APT Print skip NS for targetdevice checking if loaded before (due to pending NOPB BD change).
'4000316 2012-11-12 - Quah - Niko Datecode: replace ASSY#L2# with APT last 2 chars.
'4000314 2012-11-06 - Quah - Combine P1+P2 WipSummary.
'4000313 2012-11-06 - Quah - 4700 High-Yield check only apply for mother lots only.
'4000312 2012-11-05 - Quah - Apower GoldenUnit include for QFN as well, previouly exclude QFN.
'4000312 2012-11-05 - Quah - Add 4700 QFN for 99.9 HighYld control.
'4000311 2012-11-01 - Quah - Race RC MRLT, must also check for same marking.
'4000310 2012-10-25 - Quah - Race RC, check by Label Format during Merge.
'4000309 2012-10-15 - Quah - Enquiry - Machine by Lot, req by EK.
'4000308 2012-10-10 - Quah - EK req to disable SBL trigger.
'4000307 2012-10-09 - Quah - Blanket Load Base setting, for Fairchild.
'4000306 2012-10-05 - Quah - Heedayu req temp block for STF8211(GREEN) at Oper 5520.
'4000305 2012-10-02 - Quah - Bugs.
'4000304 2012-10-01 - Quah - Add ENGR HOLD STATUS.
'4000303 2012-09-27 - Quah - Add ENGR HOLD STATUS.
'4000303 2012-09-27 - Quah - CLASS system link to BAIC_LOT_INFO, to search for residue lot combine with another waferlot.
'4000232 2012-09-25 - Quah - Block GE Sub Withdrawal, if PCELL lot not active.
'4000230 2012-09-21 - Quah - Class System check for multiple 8500 records.
'                          - Temp unblock GE Sub Withdraw backdated transactions.
'4000298 2012-09-11 - Quah - Remove block for FS Mold A02 lots.
'4000296 2012-09-10 - Quah - Block FS A02 for MultiMerge form also.
'4000295 2012-09-10 - Quah - Clear bugs.
'4000294 2012-09-10 - Quah - MU Class System include check for Merge Qty.
'4000293 2012-09-10 - Quah - GS req to block Fairchild Mold A02 lots at 5520.
'4000292 2012-09-07 - Quah - KP Kok req Merging Control, A02 cannot merge with A04.
'4000291 2012-09-06 - Quah - GS req temporary block FS,FP at 5520, due to Assy issue.
'4000290 2012-09-05 - Quah - MU Class check for mother lot cannot be combine with another mother.
'4000289 2012-08-29 - Quah - Product master setup segment2 length increase from 10 to 14.
'4000286 2012-08-17 - Quah - Clear bugs.
'4000285 2012-08-16 - Quah - Heedayu change SQA for Niko QFN, by Waferlotno every 5 lots (lot1 & 6).
'4000284 2012-08-15 - Quah - 6000 (Qfn) shld not TransInput (T) on Apt.
'4000283 2012-08-15 - Quah - Include E1 indicator for Engineering Chargeable.
'4000282 2012-08-14 - Quah - Bourns SOIK new APT format, with marking orientation image.
'4000281 2012-08-14 - Quah - Add Scrap Residue tick on Zerolize form.
'4000281 2012-08-14 - Quah - Undo 4000279, due to follow SOIC flow, QA ACCEPT in AISS.
'4000280 2012-08-13 - Quah - Block DB Mvou if APT not printed yet.
'4000280 2012-08-13 - Quah - Format HOLDDATE dd/mm/yyyy in Release Rpt.
'4000279 2012-08-10 - Quah - APT print add 6000,6020 for GE.
'4000278 2012-08-10 - Quah - Add in ViewTrans, for GE Housing (by main APT lot).
'4000277 2012-08-09 - Quah - GE Sub Withdraw strip calculation, and View History function.
'4000276 2012-08-08 - Ramzul - Apower Binsplit, insert Cuslotno+A to Baic_lot_info for B1 lot.
'4000276 2012-08-08 - Quah - APT change bigger font and layout, format# 06-170-07/01 REV: B
'4000275 2012-08-06 - Quah - Hold/Release report formatted date. DD/MM/YYYY
'4000274 2012-08-03 - Quah - GE withdrawal by Strips, FutureHold trigger also insert HLLT to Alarm table.
'4000273 2012-08-02 - Quah - GE Subassy withdrawal, filter lot with qty > 0.
'4000272 2012-08-02 - Quah - Hold Rlse report (EK Tan).
'4000271 2012-07-30 - Quah  -  GE sub withdraw, skip Sec_AuthOper, due to no ops code for withdrawal process.
'4000268 2012-07-30 - Ramzul - Hold Reason improvement
'4000267 2012-07-27 - Ramzul - Hold Reason improvement (only for FT, and cannot blank).
'4000266 2012-07-24 - Quah - TestMaster report - improve logic.
'4000265 2012-07-19 - Quah - Subassy withdrawal bugs.
'4000264 2012-07-19 - Quah - View Trans (EDC data), format time to HH:MM for easier readibility.
'4000263 2012-07-17 - Quah - FT TT Page# bigger fontsize.
'4000262 2012-07-17 - Quah - FT TT Test-Ref bigger fontsize.
'4000261 2012-07-17 - Quah - M2 shows **Two Insertion** on FT TT.
'4000260 2012-07-13 - Quah - FT-QA TT, especially for Residual lots, req by Hidayah (temporary visible false)
'4000259 2012-07-11 - Quah - Insert to BAIC_LOT_INFO for Release retest, if RT flag is true.
'4000258 2012-07-09 - Quah - Hold and Release reason limit at 100 chars long.
'4000257 2012-07-09 - Quah - Release Form - allow RT indication. Also show RT in Wiplot Report.
'4000256 2012-07-09 - ZUL - HOLD REASON PRE-RELEASED
'4000255 2012-07-06 - Quah - TT-STD allow multiple test prg in crystal.
'4000254 2012-07-05 - Quah - During APT print, use lin_data_alpha to save alpha and numeric data, to prevent zero truncate.
'4000253 2012-07-04 Test Master Report filtering by Test Program.
'4000253 2012-07-04 Hold History - include cols for Wafer, Tgt Device.
'4000252 2012-06-26 CS Report fix bug, dont add split lot qty unless mother lot completed FT.
'4000251 2012-06-26 GE marking YYMDD use loaddate, $$ base on running counter by date.
'4000250 2012-06-22 GE SP lots, APT running number continue from last loading, is same ww.
'4000249 2012-06-20 Change transactions code for ATMEL CLASS insert.
'4000247 2012-06-19 Tables changes for NEW INVOICE SYSTEM.
'4000246 2012-06-13 clear bugs AD-Class insertion.
'4000245 2012-06-13 clear bugs.
'4000244 2012-06-13 CS Yield Report - remove BIN 1 (col 14) column.
'4000243 2012-06-13 clear Micrel-Class bugs.
'4000242 2012-06-12 clear bugs.
'4000241 2012-06-12 Enforce 900 Mvou must use Dbank MVOU (due to logic for AD/MU populate to CLASS).
'4000240 2012-06-12 Logic improvement for AD auto populate to CLASS System during Dbank MVOU.
'4000239 2012-06-12 New Test Traveller format, req by EK.
'4000239 2012-06-12 Add in Lotsearch by Machine.
'4000238 2012-06-08 AD for Class System Transaction Code = STRT instead of MVOU.
'4000237 2012-06-07 Add CP_Custprefix to CLASS query.
'4000236 2012-06-07 ATMEL Insert to CLS tables, for Dbank MVOU, other oper no need.
'4000235 2012-06-05 Add Hold Ct to Test Report
'4000234 2012-06-01 CLS_MVOU addin for Atmel.
'4000233 2012-06-01 Update default B1 when MVOU FT, if null.
'4000232 2012-05-31 Remove APT SkipWaferChecking.
'4000231 2012-05-31 Revenue Category in Aims Product Setup.
'4000230 2012-05-31 ANPEC can indicate "QC Sample" during Lot Splitting, for label printing.
'4000229 2012-05-29 ATMEL AD lots populate to CLASS System (for new Atmel-LTS system)
'4000229 2012-05-29 CS report, +Variance qty to Assyin.
'4000228 2012-05-29 Lot Creation include a few more POMaster fields.
'4000227 2012-05-28 SRCV update ltm_qty.
'4000226 2012-05-24 Function to update PO shipdate.
'4000226 2012-05-24 Reminder to change Process Ref when changing Route.
'4000225 2012-05-23 Remove double comma in IT_BARCODE_MARK for IMPINJ, during APT print.
'4000224 2012-05-23 Block RVSE to Diebank.
'4000223 2012-05-23 FOI Sample QA, prompt for 315 unit, but dont block.
'4000222 2012-05-22 Fix bug for release 4000221.
'4000221 2012-05-22 Control SPLT at FT,5500  MRLT at 3500,5500 for Non-QFN, SPLT at 4700, MRLT at 3500,4700 for QFN.
'4000220 2012-05-21 CS Hardbin Report: for QFN, add back SPLT, minus contra-MRLT.
'4000220 2012-05-21 FOI Sample Size 315, req by Selvi.
'4000219 2012-05-19 CS Report: assyin fix bug.
'4000218 2012-05-18 Fix Bug for Mac# edc input.
'4000217 2012-05-17 Relese new EDC input for FOL/EOL (similar as WB) req by Heng.
'4000216 2012-05-16 Allow APT print for same deviceno diff targetdevice (if old lots no longer active). BLOCK already in place at ProdMast update.
'4000215 2012-05-15 Fix bug: CS Yield Report assy yield wrong column reference.
'4000214 2012-05-15 Fix bug: Impinj update Shipdate as 0 (not '') during APT print.
'4000213 2012-05-14 TRLT adj to DBank Inventory: Insert to alarm table trim wafer to 13 chars only.
'4000213 2012-05-14 CS Hardbin report - remove Remark and AWT Test cols.
'4000212 2012-05-14 Block Prodmast update /APT Print if Targetdevice changed.
'4000211 2012-05-11 Impinj shipdate - dont auto-populate during creation. LL will update separately later.
'4000210 2012-05-11 Restrict SubAssy Recv/Draw only for GE lots.
'4000209 2012-05-10 Clear bug: APT Creation skip for device not found in Baic_Invoice_Price_master.
'4000208 2012-05-10 LY Chin TEST REPORT (detail & summary) in TEST SETUP Screen.
'4000207 2012-05-07 CS Yield Report for Hard/Soft bin, also selection by TEST OUT Date.
'4000206 2012-05-02 Remove NA from APT Time column. Req by Heng, for implementation of new EDC MAC-input (FOL,EOL process).
'4000204 2012-04-27 GE-SUBPROCESS no need SQL CONNECTION
'4000203 2012-04-27 GE-SUBPROCESS no need to print MINI-APT.
'4000203 2012-04-27 GE-SUBPROCESS include NPA-HSE.
'4000202 2012-04-27 RE-COMPLIED GE SUB LOTS TEST PROC
'4000201 2012-04-26 GE SUB LOTS TEST PROC
'4000199 2012-04-24 Impinj - insert Est Ship Date to Baic_Lot_Info, for CUST WIP Report.
'4000197 2012-04-23 CS report: Yield by Shipment-Date (Eric format).
'4000196 2012-04-20 Block ('FSQ211','FSDL0165RN') due do wrong FS/FP loading.
'4000195 2012-04-20 Release block for AAT 5500 Mvou
'4000194 2012-04-18 Clear bugs.
'4000193 2012-04-18 Merging condition addin to include same DeviceType matching, req by EK.
'4000192 2012-04-18 Temp block AAT at 5500 due to change of Packing Method.
'4000192 2012-04-18 GE TT & Sublot
'4000191 2012-04-16 Merging criteria - add DeviceType matching (so that diff TestProgram cannot be merge)
'4000190 2012-04-16 KO ADD THE DEVICE_LBL <> 'NA' matching for different test program
'4000189 2012-04-13 Special tracking for 'MAC# (W/B)' input.
'4000187 2012-04-09 GE APT no need extra rows for 1400.
'4000187 2012-04-09 Test Master additional tick to indicate TargetDevice for Mold-Out tracking.
'4000187 2012-04-09 Validation: cannot cancel EDC form if first time 1400 input, also cannot accept 00 time.
'4000186 2012-04-06 GE-APT.
'4000185 2012-04-06 GE-APT.
'4000184 2012-04-06 Improve logic for Fairchild IC/Mosfet wafer matching in Diebank qty deduction (in APT Printing).
'4000183 2012-04-04 Add '_NO' to 1400 EDC MAC#.
'4000182 2012-04-03 APT Fairchild Diebank qty matchby Labelinfo waferlotno first.
'4000181 2012-04-03 Allow Release for person who set FHOLD (regardless of P/Q operation).
'4000180 2012-03-30 Set compulsory input for 1400 EDC mc.
'4000179 2012-03-28 Release for EDC multiple mc input for 1400 (req by Heng).
'4000178 2012-03-27 Niko datecode generated and inserted into Baic_Lot_Info, previously generated during Label printing.
'4000178 2012-03-26 Split, BinSplit duplicates the Fhold_master, if mother lots splits to child.
'4000177 2012-03-26 Test Master detailed Report by lot, req by LY CHIN.
'4000177 2012-03-26 APT Oper column heading change to START DATE, START TIME.
'4000176 2012-03-20 Fairchild RMA noneed SQA.
'4000176 2012-03-20 Test Master Report done for LY Chin (in Setup form).
'4000175 2012-03-16 Mini APT change condition: oper 1000 to 1210, due to some wafer are pre-sawn, do go thru SAW.
'4000174 2012-03-12 APT Diebankqty matching, match by exact qty first.
'4000173 2012-03-08 AAT MoldOut datecode no need to insert to Baic_Lot_Info.
'4000171 2012-02-29 New APT, request by Mr Heng, for 1400 multiple machines.
'4000171 2012-02-29 MU Class matching include LIREF# to calculate sum qty.
'4000170 2012-02-24 QFN Dimension report - group by dimen.
'4000168 2012-02-21 Add On-Hold report (ByDevice) for P1 - requested by Siva.
'4000168 2012-02-21 Cls_Micrel_Update check for MergeLot location.
'4000167 2012-02-02 APT (1400) revert back to ori crstal report.
'4000166 2012-02-02 APT (1400) revert back to ori, due to Mr Heng need more time to brief Production on new format.
'4000165 2012-02-01 Add BCD for TT logic (BCD previously Aura).
'4000164 2012-02-01 Expand rows for APT WireBond 1400. Also add in TIME column in APT.
'4000163 2012-01-31 For APT Print wafer qty deduct, AD match by left(waferlotno,6).
'4000162 2012-01-30 Improve for Micrel Class updating, for TRNS records, and RWK lot (skip ASSY process).
'4000161 2012-01-20 Testrun for Mr Heng Machine EDC.
'4000160 2012-01-20 For APT Print wafer qty deduct, only AD match by custlot, AC still match by waferlotno.
'4000158 2012-01-19 Fix bug for Micrel Class ASSY-START qty (Diebank MVOU form).
'4000157 2012-01-18 Skip Preload and proceed direct to print APT for ALL customers.
'4000156 2012-01-14 Fix bug for Micrel Class ASSY-START qty (Diebank MVOU form).
'4000153 2012-01-13 Fix for STD-TT, ltm_deviceno link to pdm_deviceno, instead of linking by TgtDevice.
'4000150 2012-01-10 BOURNS AND FAIRCHILD REMOVED PRE-LOAD CHECKING.
'4000149 2012-01-09 Enable QA Report to pull SOIC Packing QA qty.
'4000148 2012-01-06 CreatLot listview change to Yellow background, due to some PC cannot see the data.
'4000147 2012-01-05 Preload NIKO by left7.
'4000146 2012-01-05 Fix bug. MU CLASS update for DBANK OUT qty.
'4000144 2012-01-04 APT preload check for FS, link to AIC_LI_LABELINFO FOM_WAFER
'4000143 2012-01-03 Micrel Class system checking at MVOU, fix to cater for 4700 only.
'4000142 2011-12-30 APT wafer validation, Micrel match by first 7 char.
'4000141 2011-12-29 Micrel Class system updating, cater for RESIDUE lot 2nd-time split.
'4000140 2011-12-28 APT wafer validation, for FS (083) also match by first 6 char.
'4000139 2011-12-27 APT Crystal add DRY-PACK remark for RACE-TW TSSOP.
'4000138 2011-12-27 Telefunken marking YYWW during 2100 MVOU - follow AIC ww. (ww-1 for yr 2011 only)
'4000137 2011-12-23 Temporary skip on TESTREPORT bug.
'4000136 2011-12-22 Fixbug: DB Pre-Load compulsory for all cust, req by LL.
'4000135 2011-12-22 DB Pre-Load compulsory for all cust, req by LL.
'4000134 2011-12-21 Improve on 4700-Micrel MVOU, check for 8500 Transfer completed?
'4000133 2011-12-13 Improve on 6020-Micrel MVOU for CLASS-SYSTEM integration.
'4000132 2011-12-12 Improve on 6020-Micrel MVOU for CLASS-SYSTEM integration.
'4000131 2011-12-08 Micrel CLASS System improvement for 6020 integration.
'4000130 2011-12-08 Hold code include Assy Reject codes - req by Mr Heng.
'4000129 2011-12-07 Include MQ (QWAVE) for linking to BAIC_LOT_INFO for Markfile Underline in APT printing.
'4000127 2011-12-06 Block change cuslotno for Micrel, Fairchild due to Interface File.
'4000127 2011-12-06 SQA for Fairchild, fix bugs. 200unit every 3 lots FS & FP. (refer Chong)
'4000127 2011-12-06 TestReport fix bug for Remarks field.
'4000126 2011-12-02 Change Msgbox for Micrel wafer endlot alert (tMVOU_Edc).
'4000124 2011-12-01 Extra Hold info for Test Report.
'4000124 2011-12-01 Micrel last wafer sublot - need to transfer to 8500 first, before allow lot to proceed VMI.
'4000123 2011-12-01 PO Tracking Report - Wsdb lock Optimistic
'4000123 2011-12-01 MU Class System updating set in DB-Mvou.
'4000122 2011-11-30 APT Print, allow ITADM to tick Bypass-WaferCheck.
'4000122 2011-11-30 Micrel Mvou reminder message to TRNS last sublot to 8500 (for CLASS System tracking).
'4000120 2011-11-29 Update bugs for 10th-lot LOTNO runnning num (DB Split module).
'4000119 2011-11-24 CLASS system allow proceed if lot already MRLT status.
'4000118 2011-11-18 debug split lot in diebank
'4000117 2011-11-17 Allow Fairchild scrap for R lots.
'4000116 2011-11-16 DB-Split Lotno generation looping variable changed (due to fix for Atmel dash to cuslot).
'4000115 2011-11-16 SQA-Fairchild and Telefunken requirement changed
'4000114 2011-11-14 Fairchild Merging - up to 7080 only.
'4000113 2011-11-09 Add FAC to Fairchild FOM for EON lots.
'4000112 2011-11-09 Rpt_DBlocked adodc recordsource and DBLK/UBLK update sql change to trim space behind PDM_TARGETDEVICE.
'4000111 2011-11-08 Block for BYLT full qty.
'4000110 2011-11-02 CLASS-System trigger at 6020 for Micrel Lots, for old lots not registered in CLASS, then allow proceed.
'4000109 2011-11-02 Add TOPMARK in Wiplot report.
'4000109 2011-11-02 CLASS-System trigger at 6020 for Micrel Lots, allow proceed if process = WHSE or SHIP.
'4000108 2011-11-02 Device Block Function input method modified into multiline input.
'4000107 2011-11-01 CLASS-System trigger at 6020 for Micrel Lots.
'4000106 2011-10-31 CLASS-System trigger at 5500 for Micrel Lots.
'4000105 2011-10-28 Telefunken Merging rules - clear bugs.
'4000104 2011-10-25 Improve logic on Micrel insert to CLASS System (to avoid begin-commit-trans conflict)
'4000103 2011-10-25 Insert Targetdevice to CLASS System.
'4000102 2011-10-24 Micrel insert to CLASS system during DB Mvou, Prod_Mvou, Edc_Mvou, Trans 8500
'4000102 2011-10-24 MAXTEK Apt markfile take from LI label info, if available.
'4000101 2011-10-21 GMT Apt Remark 'BLACK REEL NEW MOULD' for 2x2, 3x3, 4x4
'4000100 2011-10-19 Fairchild put back control for merging same datecode.
'4000099 2011-10-19 Telefunken merging rules change back to single cuslotno (pending for further instructions).
'4000098 2011-10-18 Add APT preload setting for Impinj.
'4000097 2011-10-18 Add Telefunken merging rules.
'4000096 2011-10-17 Minor correction in FOLYAS sheet.
'4000095 2011-10-17 Release new APT-FOLYAS (for revised 2OI reject codes).
'4000094 2011-10-17 Add machine list as per process mapping
'4000093 2011-10-14 Add APT preload setting for Telefunken.
'4000092 11/10/2011 IRC TT datalog.txt dot in cuslotno replaced with _
'4000091 10/10/2011 Remove (U) in IT_BARCODE_MARK for Mini_Circuit.
'4000090 06/10/2011 Blocking Function Change from Deviceno Level to Target Device level.
'4000089 03/10/2011 Block Device and Enquiry for Device Blocked function Added.
'4000088 29/09/2011 Remove IRC EL due to new requirement from IRC SAP System.
'4000087 28/09/2011 Remove (-) AnalogPower from IT_BARCODE_MARK topmark.
'4000086 28/09/2011 Impinj {e3 Logo} retain the comma due to need blank data in the marking template row.
'4000085 28/09/2011 Add Oper matching in BAIC_SAMPLE_LOTS deletion during REVERSAL.
'4000084 26/09/2011 Telefunken YYWW in EDC-MVOU.
'4000083 21/09/2011 Telefunken YYWW conversion during MOLD-OUT Mvou.
'4000082 12/09/2011 Quah EDC WB machine length--> change to 3 to prevent NA
'4000081 09/09/2011 Disable EDC machine no. 4 char length checking. In future, Mr Heng will provide machine list for checking.
'4000080 09/09/2011 Report 'TTVISION' format changed. PO auto-close function updated in Zerorize and Terminate form.
'4000079 24/08/2011 Exclude 'CANCEL' lots in Baic_Fg_Header. DBMvou update YYYYMMDD to Aic_wip_header!AIC_HEAD_COMMENT1 for DB_Listing filtering.
'4000078 23/08/2011 Clear bug - FOM Scrap (checking for TRLT before).
'4000077 22/08/2011 Impinj remove {e3 Logo} from IT_BARCODE_MARK.
'4000076 19/08/2011 Allow zerorize at Test operation and allow charging reject under ENGEVA
'4000075 17/08/2011 Fairchild FOM scrap - check for other lot of the same cuslot must not scrap before.
'4000074 15/08/2011 Add FGPack qty to Wiplot report filtering.
'4000073 12/08/2011 NIKO lot creation, remove '(LOWER LINE)' in IT_BARCODE_MARK.
'4000072 12/08/2011 New function- Fairchild FOM Terminate Residue.
'4000071 11/08/2011 Add info for Combine-Reel in View Lot Trans.
'4000069 09/08/2011 FS & FP SQA on TT follow Apower logic : every 5 lots (1,6,11,16,21,....) - req by EK.
'4000068 09/08/2011 add machine # and operator , wb date as per heng req. --not compile yet
'4000067 05/08/2011 Add WAFERNO# to the results of Lot Search (by Wafer).
'4000066 29/07/2011 LI-Apower QFN PMPAK pkg.
'4000065 22/07/2011 Change BuildQty to DieQty field in MiniApt.
'4000064 21/07/2011 Add AURA to TT (multi Test Prog)
'4000063 21/07/2011 Remove if-eof for Split addnew.
'4000062 19/07/2011 RMS-TT bugs.
'4000061 19/07/2011 RMS-TT qty take from Baic_Loss_Qty old-new.
'4000060 18/07/2011 Enable RouteMaster screen for Planner for Viewing, but disable Add/Remove buttons.
'4000059 13/07/2011 Add Reprint RMS-TT .REJ QTY
'4000058 13/07/2011 Add Reprint RMS-TT .
'4000057 06/07/2011 Hard code device AT "SCRF" as per fatah req.
'4000056 04/07/2011 All WIP Report, put TopLaser after PMC.
'4000055 04/07/2011 AMICCOM marking datecode based on APT Date, instead of MOLD Out, requested by Glyn.
'4000054 01/07/2011 FOM_MRLT - ignore is Lotno9 of both lots are same.
'4000053 30/06/2011 Add TRLT Reason to Lot Search results.
'4000051 28/06/2011 FOM_MRLT - exclude other than Fairchild.
'4000050 28/06/2011 Delete before insert: FOM_MRLT
'4000049 27/06/2011 Without PO No for fairchild 'as per ek tan feedback
'4000047 23/06/2011 MRLT insert ChildLot Datecode, Qty into BAIC_LOT_INFO for Fairchild FOM lot Merging.
'4000046 20/06/2011 allow FS skip wafer check (temporary).
'4000045 20/06/2011 update po_load_ymd upon DB Creation.
'4000044 15/06/2011 clear bug - replace FOM $$ lotcode to BAIC_LOT_INFO during APT printing.
'4000043 14/06/2011 replace FOM $$ lotcode to BAIC_LOT_INFO during APT printing.
'4000041 09/06/2011 add Zilltek rms -tt
'4000040 09/06/2011 add Zilltek to TT logic for multiple test-program.
'4000039 08/06/2011 add FS FOM auto numbering for PID, Marking.
'4000038 03/06/2011 sort listing in LI Create.
'4000036 02/06/2011 check & update PDM_DATA_1 before TT printing.
'4000034 27/05/2011 add trim to sml_lotno in PROC.
'4000033 23/05/2011 add oper check to PRCV click.
'4000030 17/05/2011 update APT bugs.
'4000029 16/05/2011 update APT bugs.
'4000028 16/05/2011 Include APT Print, Reprint.
'4000024 22/04/2011 For Skip FINAL TEST AUTO Trigger
'4000022 21/04/2011 Missing hist for TRANSFER from 8000 --> add in BeginEnd Transaction to trap errors.
'4000021 14/04/2011 MVOU EDC remove checking for Bedc_Tranx existed.
'4000020 13/04/2011 Add to TT MAXTEK 20 Golden unit.
'4000017 04/04/2011 Correction to Atmel VMI CT, and WIPLOT PACK CT.
'3000090 11/03/2011 REVERSE MVOU enabled for SHLT old lots.
'3000089 10/03/2011 Modify FGPACK formula for WIP REPORTS.
'3000086 04/03/2011 put back VIEW SHIPMENT logic for VIEW TRANS enquiry.
'3000083 03/03/2011 update LOGN in comtbl with 'AIMS'
'3000082 03/03/2011 SCRF SQA change back from 600 to 200.
'3000081 02/03/2011 Adjust TT-QFN spacing due to long ProgramName, Remove WAREHOUSE from top menu.
'3000078 01/03/2011 Add Terminate Osram lot form. Fix Wip Reports to avoid Excel clash.
'3000076 25/02/2011 CS Form: set Micrel wip to follow standard by-wafer logic.
'3000074 23/02/2011 Enabled button for LOWYLDQA.
'3000073 23/02/2011 Improve on View Trans (Ship info).
'3000072 23/02/2011 Allow change Shipform for PRCV (FGP) lots.
'3000071 22/02/2011 DEVICENO FOR SCRF 'AT56916
'3000070 22/02/2011 SPEED UP DBANK SPLIT.
'3000069 22/02/2011 FOR SCRF MULTITEST INSERTION & LOWYLD QA
'3000064 17/02/2011 Wip rpt amendments due to AISS implementations.
'3000061 16/02/2011 update bugs..
'3000060 16/02/2011 Use atmeloricuslot varible in DBank Split for Baic_lotmast
'3000059 16/02/2011 Improve query for QFN Eol/Test CT Report.
'3000058 16/02/2011 Chg Total_CT formula in QFN Eol/Test CT Report.
'3000057 14/02/2011 For SICM/SCRF add Me. to new cuslot (with dash prefix) - DB Split module.
'3000056 14/02/2011 For sicm /scrf traveller format changed
'3000053 09/02/2011 For Atmel SCRF/SICM at 1620, only QA can release lot, requested by ELLEN as per customer 8D control.
'3000051 01/02/2011 Die Bank - Create Lot Form : Edit message.
'3000048 24/01/2011 FOR NIKO AS PER SHLIM REQ WAFERLOT PER 1ST SUB LOT
'3000047 19/01/2011 In DB Split, change Baic_Apt_Routing back to Baic_Routing
'3000046 13/01/2011 for AP, if not 1 or 6, then NO NEED SQA
'3000045 12/01/2011 New report for Kang: QFN-FOL-CycleTime
'3000044 11/01/2011 Re-compiled.
'3000043 11/01/2011 P2 Wip Summary separated by QFN & SCARD lots.
'3000042 11/01/2011 FutureHold block for 1210,1610 QA Gate.
'3000041 11/01/2011 Add summary to QFN Dimension Report.
'3000040 10/01/2011 Mvou High-Yield for Scard lots -> 99.8 change to 99.95
'        10/01/2011 Update FOL lowyld logic to take from 2100-in instead of 1610-out (1610 no longer in route).
'3000038 07/01/2011 Anpec same tgt dev, diff test program, base on PDM_DATA_1 match with TMH_Marking (APL5606, APL5607)
'3000037 06/01/2011 Addback 8000,8500 qty to Custwip.
'3000036 05/01/2011 IRC label datecode added to LOTINFO, base on 2100 MVOU converted base on WW Calendar.
'3000035 04/01/2011 AS PER ARU REQ TMTECH TESTED SAMPLE SIZE DID 100%
'3000031 31/12/2010 HAPPY NEW YEAR 2011
'3000030 31/12/2010 ASJ add X for split lot, bug--> AD should be AS.
'3000025 30/12/2010 DBSplit temporary use BAIC_APT_ROUTING for Aic_Wip_Operation
'3000024 28/12/2010 set 2 digit ww for lot number creation in DBCreate.
'3000023 23/12/2010 Remove EpoxyCure col from Wip Reports....
'3000022 22/12/2010 Fix for WCAL_PLNG_YEAR query.
'3000021 22/12/2010 Fix for Atmel SCM,SCRF suffix during DB Split, ASJ14 suffix X during Prod Split, WW53= year yr ww1.
'3000020 21/12/2010 fix for tt_result additional checking.
'3000019 20/12/2010 3OI/OG change to 3OI in reports.
'3000018 17/12/2010 Fix the SJ Marking data retrieval in Test Traveller.
'3000017 17/12/2010 Fix the calculation for number of lots in DB Split.
'3000016 17/12/2010 Fix Test Traveller query - custname was incorrectly init to blank during matching with WIP_HEADER.
'3000013 14/12/2010 Fix minor calculation bug in DB Split.
'3000011 13/12/2010 Include only ACTIVE bom-part in Wip Report.
'3000010 13/12/2010 Add Bom Partno. to Wip by Device.
'3000009 09/12/2010 Add Wafersaw to WipReport.
'3000008 09/12/2010 Change wip report heading from OG to OI.
'3000006 08/12/2010 Block 1210,1610 for Output Report.
'3000000 08/12/2010 Operation Reduction - FOL P1
'2000176 06/12/2010 Additional QFN wip reports for SH Yeo (by Dimension, by CustDevice)
'2000174 02/12/2010 'TT RESULT ON FT & SQA
'2000171 30/11/2010 Add in 8500,8600 for FT Transfer (RES, TPENDING).
'2000169 29/11/2010 BBC_WK_ORDER_TEMP -> AIC_WK_ORDER_ALL, add in VIEW HOLD HIST to VIEWTRANS screen.
'2000165 25/11/2010 Clear minor bugs.
'2000164 25/11/2010 Allow RVSE TERMINATE at 8000, since TERMINATE can be done at 8000.
'        25/11/2010 Add in View Other Info (Aic_li_labelinfo) in View Trans form.   eg. for info such as Samhop Probetype.
'2000159 16/11/2010 BinSplit B2 qty must also not be zero.
'2000155 12/11/2010 Add PkgGrp to Wip/Output Report. Improve on DBSplit logic.
'2000154 10/11/2010 SCRF TT Comments
'2000151 10/11/2010 change to biacrouting with oper
'2000149 10/11/2010 change to biacrouting with oper
'2000148 09/11/2010 LOWYLD TRIGGER FOR OVERALL 3010 'KO/CHUA
'2000144 02/11/2010 SHLT baic_fg_header missing .update command.
'2000142 02/11/2010 prd_bin 2 verify.2010
'2000139 28/10/2010 Changes to WIP-Report query file, to show PkgGroup.
'2000135 28/10/2010 Wipreport-change back to PackageSummary
'2000132 27/10/2010 TEST Proc - add ORDER BY RTG_SEQ to get the correct next oper.
'2000131 27/10/2010 Add BOTTOM2 ,BOTTOM3 TO MARKING
'2000129 27/10/2010 Add catg2 to Package Summary Wip Report.
'2000126 26/10/2010 ADD BD# to Wipreport.
'2000124 22/10/2010 TRIGGER 95% AAT 98.5 AS PER ROEL REQ.
'2000123 22/10/2010 sbl release lot and moveout
'2000122 21/10/2010 overall assy yld fixed to 95 meeting from ke chee Room ,only atmel follow 98 system value
'2000121 21/10/2010 overall assy yld fixed to 95 meeting from ke chee Room ,maintain the fol yield eol yield
'2000120 21/10/2010 OVERALL ASSY YLD (MVOU 900 /3000/3010 1620)
'2000119 21/10/2010 SBL TRIGGER LIMIT AIC_SBL_TRIGGER FOR SM LIM BASE ON SOFTWARE BIN AND TARGET DEVICE
'2000118 20/10/2010 Add Marking to Lotsearch by Targetdevice, Add marking to View Transaction screen.
'2000115 19/10/2010 RMS & REMOVE THE NEXT OPER AT BAIC_ROUTING
'2000112 18/10/2010 'SETTLE QUERY PROBLEM IN SICM WEIGHT CHECKING
'2000111 18/10/2010 ' remove the next_oper follow the current oper for testtraveller & test route module
'2000110  15/10/10 :REMOVE THE SHLT MAINT BY MAINT DATE
'2000109 15/10/10 : Remove the db Split for baic_customer_po,and change for shlt for po
'2000108 14/10/10 : Add LI Ref# to View Trans., change lbl to textbox for allow highlight-copy.
'2000108 13/10/10 : Added in Internal BD in View Transactions.
'2000107 12/10/10 : Modify TT to support PDM_category other then 5 the groups.
'2000107 12/10/10 : Atmel lots dont allow MVOU if > 3 years. Requested by Christine.
'2000106 : sinpower flow dev_ext
'2000106 : Remove PO Customer
'2000103 : ltm_bin_indicator follow mother lot, plant ct report.
'2000102 : custwip bug
'2000098 : Clear bugs in Plant CT rpt
'2000097 : Plant CT rpt
'2000096 :
'2000091 : Matching by TOP MARK (IT BARCODE MARK)
'2000090 : bin split and atmel col test travller same target diff packageled
'2000089 : 20101001 for last sub lot 'ko / shah lastsub lot ='N'
'2000087 : 20100930 FOR Test traveller lotno change & sample lots for SQA & FT
'2000084 : proc 3500 & chk update improvement
'2000083 : Improve on IntWip & Output query, add deviceno to LotSearch.
'2000078 : ADD FOR SCRF 4420 NEED TO PERFORM SQA 100% 20100927
'2000074 : EXTEND LOSS QTY TO 20 AND ADD AMICCOM MARKING -KO 20100924 (FRIDAY)
'2000074 : TEST REPORT TO HANDLE SOIC -ZZ
'2000073 : IRC MARKING & LOCK FOR REJ COLUMN 13 TO 18
'2000071 : SpltInv qty not appear in Internal Wip report.
'2000069 : AtmelCT report, ASJ merge without CustIntDev
'2000068 : fixed for no need perform SQA
'2000067 : Internal Wip Report - remove link to Aic_wip_header
'2000064 : MIRCEL CHANGE TO EVERY LOT
'2000064 : Base on Mold Out Instead of Mold IN as per engineer vernon and shivi 20100921
'2000062 : MVOU with 0 qty at 5500 + TRLT
'2000058 : MRLT condition, add ltm_bin_indicator =''
'2000057 : remove backup files..
'2000052 : MRLT - addin Marking matching.
'2000048 : A-POWER 3 DATECODE PER CUSLOTNO 20100915
'2000047 : IRC CHANGE SHIP FORM TNR TO TT
'2000045 : merge array increase to 30, Insert TRNS trans to history.
'2000044 : add VMIQA to WipRpt
'2000040 : add CARDLOGIX
'2000036 : bugs
'2000035 : WIPREPORT..EIAJ,MOSFET,etc
'2000034 : FIX INVALID Procedure loss qty
'2000033 : FIX Bug for Lotno atmel aic wip header 9 digit test travller
'2000032 : Travell atmel sj vision 20100909
'2000026 : Bugs in wip report.
'2000025 : lotsize change upon li_lbl for sicm/scrf atmel col & lot split 5520 not allow
'2000024 : ATMEL 3300 EOL MERGE NOT ALLOW AS PER ANITA 20100908
'2000020 : improve on wip report groupings -P1P2
'2000014 : Add DZ CARD UNDER ATMEL COL CONDITION PRINT TEST TRAVLLER 20100907
'2000013 : QA report sml_lotno(+)
'2000012 : MRLT for ASJ runtime error.
'2000011 : osram db mvou
'2000008 : IRC Marking
'2000006 : Holdflag isempty condition (for MVOU)
'2000005 : MergeAtmel TT, Atmel Cust-Int-Dev matching for Merging,
'2000000 : First attempt at overthrowing Workstream... <(-_-)<   >(^o^)>   v(-o-)v   ^(^_^)^
'----------------------------------------------------------------------------------------------------

'1000049 : DB Split: enable lotsize (for Engrlot) and startdate
'1000047 : merge if array > 15 skip.
'1000045 : IMPROVE ON LOT CREATION LOGIC
'1000042 : txt_lotno = trim(txt_lotno)
'1000041 : CS Report updated as Cindy requested.
'1000039 : Future hold -check on authorisation oper base on future hold oper, not current lot oper.
'1000038 : CS Yield repor.
'1000037 : Lowyield msgbox - shorter message text.
'1000029 : set Lotype default to N
'1000028 :  fix null histseq
'1000027 :  fixed runnum --> gtsdat/lotreg
'1000025 :  ENGRLOT STATUS
'1000024 : Fix for MultiMerge, Null and B1 inter-merge-able, otherwise not.
'1000023 : FIX FOR SAMPLE LOTS message .
'1000020 : Fix BaicLossQty LSQ_LOSS_CATGRY data input.
'1000019 : Keyboard > Mouse improvement on SHLT, MVOU, MVOU_EDC, some Reports
'          Fixes for conflicting LotNo with Workstream
'1000016 :  Keyboard > Mouse improvement on SHLT, MVOU, MVOU_EDC
'           Visibility of frame NONQFNTESTINFO
'1000015 : Keyboard > Mouse improvement by = sign
'1000014 : TT-QFN barcode
'           Reverse Terminate - RVTR
'1000013 :
'1000012 : Keyboard > Mouse
'1000011 - 20100804 -   Improve on custwip query.
'1000006 - 20100802 -   Allow release lot for SYSTEM AUTO HOLD.
'                       DB Splt startdate take from LI startdate.
'1000005 - 20100802 -   Reset lowyld trigger.

'1000001 - AIMS GO LIVE - August-2010.  ^^^^^^^^^^^^^
'----------------------------------------------------------------------------------------------------


'*** 2013-09-26 block if run from deploy folder.
If Left(Trim(App.Path), 20) = "\\aicwksvr2016\deploy" Then
    MsgBox "Error: System running from Deploy folder.", vbExclamation, "Incorrect Path Setup"
    Unload Me
    End
End If



'***Check for multiple instance
Dim m_hWnd As Long
m_hWnd = FindWindow(MesSysName & " (*")
If m_hWnd > 0 Then
    MsgBox "Please close all AIMS program before starting a new one.", vbExclamation, "AIMS already started"
    Unload Me
    End
    Exit Sub
End If
'***

Me.Caption = MesSysName & " " & MesSysVer

   User_Login.Show     '---> LOGIN SCREEN
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
    
 chkupdate 'ignore updating version inorder to solve hidthir 20201224
   
 '  MsgBox "ONLY USE FOR TEST MASTER UPDATE !!!", vbExclamation, "Message"
   
    
    Select Case ButtonMenu.Key
    
        '------------------------------------
        'MASTER SETUP
        '------------------------------------
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
            
        '------------------------------------
        'IQA/QA
        '------------------------------------
        Case "QAIN":
            Call DisableAllMenu
            Prd_QA_Screen.Show
        Case "BGRD":
            Call DisableAllMenu
            Prd_QA_BackGrind.Show
        Case "ARMS":                'Quah added 20150720
            Call DisableAllMenu
            Prd_Rms_Assy.Show
            
            
        '------------------------------------
        'DIEBANK
        '------------------------------------
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
            
         '------------------------------------
        'DIEBANK ST PRELOAD TEST
        '------------------------------------
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
        
        '------------------------------------
        'PRODUCTION
        '------------------------------------
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
'            Prd_Merge.Show
            Prd_MultiMerge.Show
        Case "HOLD":
            Call DisableAllMenu
            Prd_Hold.Show

        '------------------------------------
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
            
        Case "KTIN"                 'Kitting In.
'            Call DisableAllMenu
'            Kitting_Recv.Show
            
        Case "KTOU"                 'Kitting Out.
            Call DisableAllMenu
            Kitting_Wdraw.Show
            
        Case "MATCNT":              'Stock Count 25/10/2021
            Call DisableAllMenu
            Prd_StockCount.Show
            
        Case "MCPC"
            Call DisableAllMenu
            Microchip_FOL_Combine.Show
            
            
'KO ADD 20130313 For Production Bin2 Split lot
' ???? Still need or not????
'       Case "B2SP":
'            Call DisableAllMenu
'            Prd_B2Split.Show
            
        '------------------------------------
        Case "SRCV":
            Call DisableAllMenu
            Sub_Receive.Show
        Case "DRAW":
            Call DisableAllMenu
            Sub_Withdraw.Show
        
        
        '------------------------------------
        'WAREHOUSE
        '------------------------------------
        Case "SHLT":
            Call DisableAllMenu
            Prd_Shlt.Show
        Case "RSHT":
            Call DisableAllMenu
            Rvse_Trn.lbl_title.Caption = "REVERSE SHLT"
            Rvse_Trn.Show
        
        '------------------------------------
        'OTHER FUNCTIONS
        '------------------------------------
        Case "FHLD":
            Call DisableAllMenu
            Prd_Future_Hold.Show
        Case "RLSE":
            Call DisableAllMenu
            Prd_Release.Show
        Case "XDOC":
            Call DisableAllMenu
            Prd_Attach_Doc.Show
        '----------------------------------------
        Case "CGTP":                                                'testprocess
            Call DisableAllMenu
            Prd_Chg_LotInfo.lbl_title.Caption = "CHANGE TEST PROGRAM"
            Prd_Chg_LotInfo.txt_transtype.Text = "CGTP"
            Prd_Chg_LotInfo.new_testprog.Enabled = True
            Prd_Chg_LotInfo.new_testprog.BackColor = &HC0FFFF
            Prd_Chg_LotInfo.Show
        Case "CGRT":                                                'route
            Call DisableAllMenu
            Prd_Chg_LotInfo.lbl_title.Caption = "CHANGE PROCESS ROUTE"
            Prd_Chg_LotInfo.txt_transtype.Text = "CGRT"
            Prd_Chg_LotInfo.new_processroute.Enabled = True
            Prd_Chg_LotInfo.new_processroute.BackColor = &HC0FFFF
            Prd_Chg_LotInfo.new_routetype.Enabled = True
            Prd_Chg_LotInfo.new_routetype.BackColor = &HC0FFFF
            Prd_Chg_LotInfo.Show
        Case "CGTD":                                                'target device
            Call DisableAllMenu
            Prd_Chg_LotInfo.lbl_title.Caption = "CHANGE TARGET DEVICE"
            Prd_Chg_LotInfo.txt_transtype.Text = "CGTD"
            Prd_Chg_LotInfo.new_targetdevice.Enabled = True
            Prd_Chg_LotInfo.new_targetdevice.BackColor = &HC0FFFF
            Prd_Chg_LotInfo.Show
        
        '20140919 Diana -Disable changing cuslot due to no parallelity with label
'        Case "CGCL":                                                'cuslot
'            Call DisableAllMenu
'            Prd_Chg_LotInfo.lbl_title.Caption = "CHANGE CUSTOMER LOTNO."
'            Prd_Chg_LotInfo.txt_transtype.Text = "CGCL"
'            Prd_Chg_LotInfo.new_cuslotno.Enabled = True
'            Prd_Chg_LotInfo.new_cuslotno.BackColor = &HC0FFFF
'            Prd_Chg_LotInfo.Show
        Case "CGSF":                                                'shipform
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
        '----------------------------------------
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
    '            Call DisableAllMenu
    '            Rvse_Trn.lbl_title.Caption = "REVERSE MRLT"
    '            Rvse_Trn.Show
        '----------------------------------------
        Case "ASMG":    'Quah added 20161011
            Call DisableAllMenu
            Prd_Atmel_Merge.Show
        Case "STBI":    'Quah added 20220427
            Call DisableAllMenu
            STMICRO_BINCONVERT.Show
        Case "SCAC":
            Call DisableAllMenu
            Prd_Terminate.lbl_title = "TERMINATE FAIRCHILD RESIDUE - FOM LOTS"
            Prd_Terminate.Show
        Case "FSMR":
            Call DisableAllMenu
            Prd_Fairchild_Mrlt.Show
'        Case "MUTR":
'            MsgBox "Function Not Ready", vbCritical, "Message"
'            Call DisableAllMenu
'            Prd_Transfer_Micrel.Show
        '----------------------------------------
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
'            MsgBox "Atmel B2B Not Ready", vbCritical, "Message"
            Call DisableAllMenu
           B2B_Update.Show
        Case "PMSS":    'ASYRAF added 20231123 SCRAP
            Call DisableAllMenu
            Prd_Mvou_STSCRAP.Show
        Case "STBIS":    'ASYRAF added 20231123
            Call DisableAllMenu
            STMICRO_BINCONVERT_SCRAP.Show
            
        '------------------------------------
        'ENQUIRY / REPORTS
        '------------------------------------
'        Case "WEBQ":
'            Dim enq
'            enq = Shell("C:\Program Files\Internet Explorer\IEXPLORE.EXE http://191.1.2.200/shiftoutputwip/mainmenu.aspx", vbNormalNoFocus)
'            Call userPermission(login_id)
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
