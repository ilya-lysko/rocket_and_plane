Sub main()
    
    Dim boom1, past1 As Boolean
    
    Call NewGraph
    
    Call SborInfy(vsam, k1, k2, xrock2, yrock2, xsam2, ysam2, pe, wsam, hsam, maxa)
    
    Call Process(vsam, k1, k2, xrock2, yrock2, xsam2, ysam2, pe, wsam, hsam, maxa)
        
End Sub
