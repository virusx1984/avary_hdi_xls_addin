Attribute VB_Name = "hdi_xls_funs"

Public Function LNF_MPN(ByVal pn As String) As String
    
    If hdi_xls_funs.LNF_IsSapPN(pn) Then
        LNF_MPN = Mid(pn, 3, 6)
    Else
        If pn Like "*TBD*" Then
            LNF_MPN = Mid(pn, 5, 8)
        Else
            LNF_MPN = Mid(pn, 6, 6)
        End If
    End If
    
End Function
    

Public Function LNF_SPN3(ByVal pn As String) As String
    
    If hdi_xls_funs.LNF_IsSapPN(pn) Then
        LNF_SPN3 = Left(pn, 11)
    Else
        If pn Like "*TBD*" Then
            LNF_SPN3 = Left(pn, 13)
        Else
            LNF_SPN3 = Left(pn, 13)
        End If
    End If
    
    

End Function

Public Function LNF_SPN2(ByVal pn As String) As String
    
    If hdi_xls_funs.LNF_IsSapPN(pn) Then
        LNF_SPN2 = Left(pn, 9)
    Else
        If pn Like "*TBD*" Then
            LNF_SPN2 = Left(pn, 12)
        Else
            LNF_SPN2 = Left(pn, 12)
        End If
    End If
    
    

End Function

Public Function LNF_SPN(ByVal pn As String) As String
    
    If hdi_xls_funs.LNF_IsSapPN(pn) Then
        LNF_SPN = Left(pn, 8)
    Else
        If pn Like "*TBD*" Then
            LNF_SPN = Left(pn, 12)
        Else
            LNF_SPN = Left(pn, 11)
        End If
    End If
    
End Function


Public Function LNF_IsSapPN(ByVal pn As String) As Boolean

    Dim regx As RegExp
    
    Set regx = New RegExp
    
    regx.Pattern = "^[Y|Q|Z|F|H|R|M][0-9][0,7,A-Z]{2,2}[0-9,A-Z]{4,4}"
    regx.Global = True
    regx.IgnoreCase = True
    
    Set matches = regx.Execute(pn)
    
    LNF_IsSapPN = (matches.Count > 0)
    
    If LNF_IsSapPN = False Then
        
        regx.Pattern = "^[Y|Q|Z|F|H|R|M][0-9]{7,7}"
        regx.Global = True
        regx.IgnoreCase = True
        
        Set matches = regx.Execute(pn)
        
        LNF_IsSapPN = (matches.Count > 0)
        
        If LNF_IsSapPN = False Then
            regx.Pattern = "^[Y|Q|Z|F|H|R|M]CQU[0-9,A-Z]{4,4}"
            regx.Global = True
            regx.IgnoreCase = True
            
            Set matches = regx.Execute(pn)
            
            LNF_IsSapPN = (matches.Count > 0)
            
        End If
        
    End If

End Function
