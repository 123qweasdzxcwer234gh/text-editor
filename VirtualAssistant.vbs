Set Sapi=Wscript.CreateObject("SAPI.spVoice")
Sapi.speak "Hello i,m Your Computer, wat is your name"
name=InputBox("","","")
Sapi.speak "nice to meet you" & " " & name
Sapi.speak "What to you want me to do next"
Set shell=Wscript.CreateObject("Wscript.Shell")
X=InputBox("","","")
If X = "open command prompt" Then
   shell.Run "cmd.exe"
ElseIf X = "open cmd.exe" Then
   shell.Run "cmd.exe"
ElseIf X = "open calc." Then
   shell.run "calc.exe"
Elseif X = "go fly a kite" Then
    Sapi.Speak "What??"
Elseif X = "shut down my computer" Then
   Sapi.speak "i can't exactly do that, but i can open slidetoshutdown.exe"
   shell.run "SlideToShutDown.exe"
Elseif X = "open a program" Then
   Sapi.speak "what program do you want me to open"
   Y=InputBox("","","")
   shell.run Y
Elseif X = "open notepad" Then
   shell.run "notepad.exe" 
ElseIf X = "restart" Then
   shell.run "VirtualAssistant.vbs"
ElseIf X = "say something" Then
   Sapi.speak "wht do you want me to say"
   Y=InputBox("","","")
   Sapi.speak Y
ElseIf X = "turn on my screensaver" Then
   shell.run "Mystify.scr"
Elseif X = "tell me a poem" Then
   Sapi.speak "Marry Had A Little Lam, It's Fleece Was White As Snow, And Where Ever Mary Went, The Lam Was Shure To Go, It Followed Her To School One Day, Witch Was Against The Rule, It Made The Children Laugh And Play, To See A Lam At School."
ElseIf X = "open wordpad" Then
   shell.run "write.exe"
ElseIf X = "close a program" Then
   Sapi.speak "what program do you want me to close"
   Y=InputBox("","","")
   Dim strComputer, strProcessToKill, objWMIService, colProcess, objProcess

   strComputer = "."
   strProcessToKill = Y
   Set objWMIService = GetObject("winmgmts:" _ 
      & "{impersonationLevel=impersonate}!\\" _ 
      & strComputer _ 
      & "\root\cimv2") 
   Set colProcess = objWMIService.ExecQuery _
      ("Select * from Win32_Process Where Name = '" & strProcessToKill & "'")
   For Each objProcess in colProcess
      objProcess.Terminate()
   Next
ElseIf X = "open cb." Then
   shell.run "CB.bat"
ElseIf X = "open file explorer" Then
   shell.run "explorer.exe"
Else
   Sapi.speak "try something difrent next time"
End If
