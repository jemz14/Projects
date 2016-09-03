Dim Message, Speak
Message = InputBox("Hello Jemz Quitlong","Speak")
Set Speak = CreateObject("sapi.spvoice")
Speak.Speak Message