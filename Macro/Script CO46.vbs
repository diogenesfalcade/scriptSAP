If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nCO46"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSTRIP_T1/tabpTAB30").select
session.findById("wnd[0]/usr/tabsTABSTRIP_T1/tabpTAB30/ssub%_SUBSCREEN_T1:PP_ORDER_PROGRESS:0030/chkP_PSHIR").selected = true
session.findById("wnd[0]/usr/tabsTABSTRIP_T1/tabpTAB30/ssub%_SUBSCREEN_T1:PP_ORDER_PROGRESS:0030/ctxtP_POSID").text = "104564-05BR"
session.findById("wnd[0]/usr/tabsTABSTRIP_T1/tabpTAB30/ssub%_SUBSCREEN_T1:PP_ORDER_PROGRESS:0030/ctxtP_WRK30").text = "5260"
session.findById("wnd[0]/usr/tabsTABSTRIP_T1/tabpTAB30/ssub%_SUBSCREEN_T1:PP_ORDER_PROGRESS:0030/chkP_PSHIR").setFocus
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell[0]").pressContextButton "&LOAD"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell[0]").selectContextMenuItem "&LOAD"
session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").setCurrentCell 2,"TEXT"
session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").selectedRows = "2"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell[0]").pressContextButton "&PRINT_BACK"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell[0]").selectContextMenuItem "&PRINT_PREV_ALL"
session.findById("wnd[0]/mbar/menu[3]/menu[5]/menu[2]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\363649\\Desktop"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "TESTESCRIPT.txt"
'session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 11
session.findById("wnd[1]/tbar[0]/btn[11]").press
'session.findById("wnd[0]/tbar[0]/okcd").text = "/nCO46"
'session.findById("wnd[0]").sendVKey 0
'session.findById("wnd[0]/tbar[0]/btn[15]").press
'session.findById("wnd[0]/tbar[0]/btn[15]").press
'session.findById("wnd[0]/tbar[0]/btn[0]").press
