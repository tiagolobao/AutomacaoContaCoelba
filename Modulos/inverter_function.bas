Attribute VB_Name = "Inverter"
Function INVERTER(inverte As Variant) As String
Dim txt As String
For i = 0 To Len(inverte) - 1
txt = txt & Mid(CStr(inverte), Len(CStr(inverte)) - i, 1)
Next
INVERTER = txt
End Function

                    
                    
                    
                    
                    
                    
                    
                    
                    

