Private Sub WaitOneSecond()

    ' Процедура для задержки на одну секунду, для лучшего визуального восприятия
    
    Dim Start As Single
        Start = Timer
    Do While Start + 1 > Timer
       DoEvents
    Loop
    
End Sub
Private Sub NewGraph()

    ' Процедура создания поля для графа и подготовки его к работе (размеры и положение)

    Charts.Add
    With ActiveChart
        .ChartType = xlXYScatterSmooth
        .Location Where:=xlLocationAsObject, Name:="Лист1"
    End With
    
    With Worksheets(1).ChartObjects(1)
        .Width = 850
        .Height = 350
    End With

    Лист1.ChartObjects(1).Left = 10
    Лист1.ChartObjects(1).Top = 10
    
End Sub
Private Sub Plane()

    ' Процедура управления самолетиком

    For i = 2 To 20
        ActiveChart.Shapes.AddConnector(msoConnectorStraight, 25, 25, 25 * i, 25) _
            .Select
            Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadOpen
        Call WaitOneSecond
    Next
    
End Sub
Private Sub Rocket()
    
    ' Процедура управления ракеткой
    
    For i = 1 To 12
        ActiveChart.Shapes.AddConnector(msoConnectorStraight, 25, 325, 25, 325 - 25 * i) _
            .Select
            Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadOpen
        Call WaitOneSecond
    Next
    
End Sub
Private Sub SborInfy(vsam, k1, k2, xrock2, yrock2, xsam2, ysam2, pe, wsam, hsam, maxa)

    vsam = Cells(1, 2).Value ' скорость самолета
    k1 = Cells(2, 2).Value ' во сколько раз скорость ракеты больше скорости самолета
    k2 = Cells(3, 2).Value ' коэффицент усиления
    
    xrock2 = Cells(4, 2).Value ' начальные координаты ракетки
    yrock2 = Cells(5, 2).Value
    xsam2 = Cells(6, 2).Value
    ysam2 = Cells(7, 2).Value
    pe = Cells(8, 2).Value
    wsam = Cells(9, 2).Value
    hsam = Cells(10, 2).Value
    maxa = Cells(11, 2).Value
    
End Sub
Private Sub Process(vsam, k1, k2, xrock2, yrock2, xsam2, ysam2, pe, wsam, hsam, maxa)
    
    i = 1
    ad1 = 0
    
    While (boom1 = False) And (past1 = False) ' тут будет до того случая, как ракетка не попадет в самолетик (until)
        
        ' угол отклонения, известный ракетке, его рассчитываю непосредственно в координатах, так как тригонометрические функции в vba - неадекватны
        
        ' Координаты самолетика
        
        xsam1 = xsam2
        ysam1 = ysam2
        xsam2 = xsam1 + vsam * pe
        ysam2 = ysam1
        
         ' Координаты ракетки
        
        ' radians = degrees / ( 180 / Pi )
        ' degrees = radians * ( 180 / Pi )
        'Application.WorksheetFunction.Asin
        
        xrock1 = xrock2
        yrock1 = yrock2
        
        ad0 = ad1
        ad1 = Application.WorksheetFunction.Asin(Abs(xsam1 - xrock1) / Sqr((xsam1 - xrock1) ^ 2 + (ysam1 - yrock1) ^ 2)) * (180 / 3.14)
        maxad = maxa / (180 / 3.14)
        a0 = ad0 / (180 / 3.14)
        a1 = ad1 / (180 / 3.14)
        
        If (Abs(ad1 - ad0) > maxa) Then
            j1 = Abs(Sin(a0 + maxad))
            j2 = Abs(Cos(a0 + maxad))
            ad1 = ad0 + maxa
        Else
            j1 = (Abs(xsam1 - xrock1) / Sqr((xsam1 - xrock1) ^ 2 + (ysam1 - yrock1) ^ 2))
            j2 = (Abs(ysam1 - yrock1) / Sqr((xsam1 - xrock1) ^ 2 + (ysam1 - yrock1) ^ 2))
        End If
        
        xrock2 = xrock1 + Sqr(k2) * k1 * vsam * pe * j1
        yrock2 = yrock1 - (1 / Sqr(k2)) * k1 * vsam * pe * j2
        
               
        If (yrock2 < 0) Then
            past1 = True ' Промах
            MsgBox ("Промах!")
        End If
        
        If ((xsam2 - xrock2) < hsam * 0.5) And ((ysam2 - yrock2) < wsam * 0.5) Then
            boom1 = True ' Попадание
            MsgBox ("Попадание!")
        End If
               
        If (boom1 = False) And (past1 = False) Then
            ' Система управления самолетиком
            
            ActiveChart.Shapes.AddConnector(msoConnectorStraight, xsam1, ysam1, xsam2, ysam2) _
                .Select
                'Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadOpen ' "Стрелочки"
                
            ' Система управления ракеткой
            ActiveChart.Shapes.AddConnector(msoConnectorStraight, xrock1, yrock1, xrock2, yrock2) _
                .Select ' возможно стоит запараметризировать отношение скоростей ракетки и самолетика
                'Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadOpen ' "Стрелочки"
            Selection.ShapeRange.ShapeStyle = msoLineStylePreset3
        Else
            ActiveChart.Shapes.AddShape(msoShapeRectangle, xsam1 - vsam * pe, ysam1 - 0.5 * hsam, 20, 10). _
        Select
        End If
         
        Call WaitOneSecond ' Задержка на одну секунду
        i = i + 1
        
    Wend
    
End Sub
Sub Main()
    
    Dim boom1, past1 As Boolean
    
    Call NewGraph
    
    Call SborInfy(vsam, k1, k2, xrock2, yrock2, xsam2, ysam2, pe, wsam, hsam, maxa)
    
    Call Process(vsam, k1, k2, xrock2, yrock2, xsam2, ysam2, pe, wsam, hsam, maxa)
        
End Sub
