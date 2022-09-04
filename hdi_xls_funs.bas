Attribute VB_Name = "hdi_xls_funs"
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
