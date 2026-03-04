Set wmi = GetObject("winmgmts:\\.\root\cimv2")

Do
    Set processos = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name='nbblo.exe'")
    
    For Each p In processos
        p.Terminate
    Next
    
    WScript.Sleep 1000 'espera 1 segundo
Loop