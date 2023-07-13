def script_CO46(session, WBS, Temp):
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nCO46"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/tabsTABSTRIP_T1/tabpTAB30").select()
    session.findById("wnd[0]/usr/tabsTABSTRIP_T1/tabpTAB30/ssub%_SUBSCREEN_T1:PP_ORDER_PROGRESS:0030/chkP_PSHIR").selected = True
    session.findById("wnd[0]/usr/tabsTABSTRIP_T1/tabpTAB30/ssub%_SUBSCREEN_T1:PP_ORDER_PROGRESS:0030/ctxtP_POSID").text = str(WBS)
    session.findById("wnd[0]/usr/tabsTABSTRIP_T1/tabpTAB30/ssub%_SUBSCREEN_T1:PP_ORDER_PROGRESS:0030/chkP_PSHIR").setFocus()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell[0]").pressContextButton("&LOAD")
    session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell[0]").selectContextMenuItem("&COL0")
    session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").currentCellRow = 0
    session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").SelectAll()
    session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0620/btnAPP_FL_SING").press()
    session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = -1
    session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectColumn("SELTEXT")
    session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").pressColumnHeader("SELTEXT")
    session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").SelectAll()
    session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0620/btnAPP_WL_SING").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell[0]").pressContextButton("&PRINT_BACK")
    session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell[0]").selectContextMenuItem("&PRINT_PREV_ALL")
    session.findById("wnd[0]/mbar/menu[3]/menu[5]/menu[2]/menu[2]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = str(Temp)
    txtname = "CO46_" + str(WBS) + ".txt"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = txtname
    session.findById("wnd[1]/tbar[0]/btn[11]").press()