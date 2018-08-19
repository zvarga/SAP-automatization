Sub SAP_Main()

'''''IMPORTANT! Before you start running the macro you have to be logged into SAP.

'''''''''''''''''''''''''''''''''       CONECTING INTO SAP        '''''''''''''''''''''''''''''''''''''
'Get the first active SAP session:
    Set SapGuiAuto = GetObject("SAPGUI")                    'Get the SAP GUI Scripting object
    Set SAPApp = SapGuiAuto.GetScriptingEngine              'Get the currently running SAP GUI
    Set SAPCon = SAPApp.Children(0)                         'Get the first system connected
    Set session = SAPCon.Children(0)                        'Get the first session (window of connection)
    session.findById("wnd[0]").maximize                     'Maximize SAP screen


''''''''''''''''''''''''''''''''''''       PREPARATION        ''''''''''''''''''''''''''''''''''''''''''
Set objExcel = GetObject(, "Excel.Application")
Set objSheet = objExcel.ActiveWorkbook.ActiveSheet

'Read Initial date into variable to be run FBL1N
    initialdate = Trim(CStr(objSheet.Cells(5, 2).Value & "." & objSheet.Cells(5, 3).Value & "." _
    & objSheet.Cells(5, 4).Value))

'Input Final date into variable to be run FBL1N
    finaldate = Trim(CStr(objSheet.Cells(8, 2).Value & "." & objSheet.Cells(8, 3).Value & "." _
    & objSheet.Cells(8, 4).Value))

'Input Company Code into variable to be run FBL1N
    CoCd = Trim(CStr(objSheet.Cells(10, 4).Value))

'Set path and file name in which output will be saved
    strPath = Application.ThisWorkbook.Path                 'File path
    flname = Trim(CStr("FBL1N_output.xls"))                 'Filename


''''''''''''''''''''''''''''''''       RUN SAP TRANSACTION        ''''''''''''''''''''''''''''''''''''''
'Initate FBL1N transaction in SAP:
With session
    .findById("wnd[0]/tbar[0]/okcd").Text = "/nFBL1n"       'Initiate FBL1N transaction
    .findById("wnd[0]").sendVKey 0                          'Press enter

'Input all Vendors with *
     .findById("wnd[0]/usr/ctxtKD_LIFNR-LOW").Text = "*"

'Input one company code *
    .findById("wnd[0]/usr/ctxtKD_BUKRS-LOW").Text = CoCd

'Put initial and final posring date
    .findById("wnd[0]/usr/radX_AISEL").Select
    .findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").Text = initialdate
    .findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").Text = finaldate

'Select all ticks in SAP
    On Error Resume Next
    .findById("wnd[0]/usr/chkX_SHBV").Selected = True
    .findById("wnd[0]/usr/chkX_MERK").Selected = True
    .findById("wnd[0]/usr/chkX_PARK").Selected = True
    .findById("wnd[0]/usr/chkX_APAR").Selected = True
    On Error GoTo 0

'Press execute
    .findById("wnd[0]/tbar[1]/btn[8]").press                'Press execute clock


''''''''''''''''''''''''''''      DOWNLOAD FILE AND FINALIZE        ''''''''''''''''''''''''''''''''''''
'Select export format
    .findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
    .findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    .findById("wnd[1]/tbar[0]/btn[0]").press

'Input filename and export
    .findById("wnd[1]/usr/ctxtDY_FILENAME").Text = flname   'Copy file name into SAP
    .findById("wnd[1]/usr/ctxtDY_PATH").Text = strPath      'Copy file path into SAP

'Press the button to save
    .findById("wnd[1]/tbar[0]/btn[0]").press                'Pressing ENTER

'Note in excel that it was done
    MsgBox ("Done. FBL1N_output.xls outcome was saved in the " & strPath & " folder")

End With

End Sub
