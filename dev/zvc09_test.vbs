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
session.findById("wnd[0]/tbar[0]/okcd").text = "/nzvc09"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/radP_NONE").select
session.findById("wnd[0]/usr/ctxtA_MATNR").text = "custom-ahu"
session.findById("wnd[0]/usr/ctxtA_PLANT").text = "1170"
session.findById("wnd[0]/usr/txtA_VBELN").text = "9069544"
session.findById("wnd[0]/usr/txtA_POSNR").text = "20"
session.findById("wnd[0]/usr/radP_NONE").setFocus
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlG_CONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlG_CONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/radRB_2").setFocus
session.findById("wnd[1]/usr/radRB_2").select
session.findById("wnd[1]/usr/radRB_OTHERS").setFocus
session.findById("wnd[1]/usr/radRB_OTHERS").select
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\timmerman\Documents\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "scripting_zvc09.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 20
session.findById("wnd[1]/tbar[0]/btn[0]").press
