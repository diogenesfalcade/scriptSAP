def script_MB25(session, path, name, plant="5260"):
    session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = plant
    session.findById("wnd[0]/usr/ctxtWERKS-LOW").setFocus
    session.findById("wnd[0]/usr/ctxtWERKS-LOW").caretPosition = 4
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = path
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = name
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 14
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    
    