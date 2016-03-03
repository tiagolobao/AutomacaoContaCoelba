Attribute VB_Name = "Módulo2"
Function DATAR(entrada As Variant) As String

Dim txt As String

txt = txt & Mid(CStr(entrada), 1, 1)
txt = txt & Mid(CStr(entrada), 2, 1) 'DIA
txt = txt & "/"

txt = txt & Mid(CStr(entrada), 3, 1)
txt = txt & Mid(CStr(entrada), 4, 1) 'MES
txt = txt & "/"

txt = txt & "2"
txt = txt & "0"
txt = txt & Mid(CStr(entrada), 5, 1)
txt = txt & Mid(CStr(entrada), 6, 1) 'ANO

DATAR = txt

End Function
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    

