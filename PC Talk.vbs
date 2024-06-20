Dim Message, speak
            Message=InputBox("Enter text", "PC Talk")
            Set Speak=CreateObject("sapi.spvoice")
            Speak.Speak Message