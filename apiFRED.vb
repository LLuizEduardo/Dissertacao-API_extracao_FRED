Sub apiFred()
    Dim URL As String
    i = 1
    
    Do Until Range("G2").Value = Range("H2").Value
        i = i + 1
        Range("G2") = i
        j = 1
        
        Do Until Range("G3").Value = Range("H3").Value
            j = j + 1
            Range("G3") = j
            
            'Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
            URL = Range("A9").Value
            ActiveWorkbook.FollowHyperlink URL
                Application.Wait Now + TimeSerial(0, 0, 5)
            
            k = 2
            
            AppActivate "Excel"
            Range("B4").Copy
                Application.Wait Now + TimeSerial(0, 0, 3)
            AppActivate "Google Chrome"
                Application.Wait Now + TimeSerial(0, 0, k)
            Application.SendKeys "^s"
                Application.Wait Now + TimeSerial(0, 0, k)
            Call SendKeys("^v", True)
                Application.Wait Now + TimeSerial(0, 0, k)
            Call SendKeys("{ENTER}")
                Application.Wait Now + TimeSerial(0, 0, k)
            AppActivate "Google Chrome"
            Application.SendKeys "^w"
                Application.Wait Now + TimeSerial(0, 0, k)
        Loop
    Loop
End Sub
