Sub shcpgrapher()

    'Aqui se pueden alterar los paramteros basicos de estilo
    Const fontName As String = "Century Gothic"
    Const fontSize As Integer = 12
    fontColor = RGB(0, 0, 0)
    Const height As Double = 311.811023622
    Const width As Double = 283.4645669291

    'Si el objeto seleccionado no es una grafica el codigo barfea
    If Not TypeName(Selection) = "ChartArea" Then
        MsgBox "Esto no es una gráfica.", vbExclamation
        Exit Sub
    End If

    Dim selectedChart As ChartObject
    Set selectedChart = ActiveChart.Parent

    On Error GoTo ErrorHandler
    With selectedChart.Chart
        ' Se usan los parametros de arriba para alterar el estilo
         ' Modificaciones en el area
        .ChartArea.Font.Name = fontName
        .ChartArea.Font.Size = fontSize
        .ChartArea.Font.Color = fontColor
         ' Modificaciones en la leyenda
        .Legend.Font.Name = fontName
        .Legend.Font.Size = fontSize
        .Legend.Font.Color = fontColor
         ' Modificaciones en los ejes
        .Axes(xlCategory, xlPrimary).TickLabels.Font.Name = fontName
        .Axes(xlCategory, xlPrimary).TickLabels.Font.Size = fontSize
        .Axes(xlCategory, xlPrimary).TickLabels.Font.Color = fontColor
        .Axes(xlValue, xlPrimary).TickLabels.Font.Name = fontName
        .Axes(xlValue, xlPrimary).TickLabels.Font.Size = fontSize
        .Axes(xlValue, xlPrimary).TickLabels.Font.Color = fontColor

        ' Color de los ejes
        .Axes(xlCategory).Border.Color = fontColor
        .Axes(xlValue).Border.Color = fontColor

        ' ticks de los ejes
        .Axes(xlCategory).MajorTickMark = xlCross
        .Axes(xlValue).MajorTickMark = xlInside
        
        ' Borrar el tìtulo si existe
        If .HasTitle Then
             .ChartTitle.Delete
        End If

        ' Borrar la malla si exite
        If .Axes(xlCategory).HasMajorGridlines Then
            .Axes(xlCategory).MajorGridlines.Delete
        End If
        If .Axes(xlValue).HasMajorGridlines Then
            .Axes(xlValue).MajorGridlines.Delete
        End If
        
    End With

    With selectedChart
        'Cambiar el tamaño de la gràfica a 10x11
        .height = height
        .width = width
        
            'Hacer la gràfica transparente
        .Chart.ChartArea.Fill.Visible = msoFalse
        .Chart.PlotArea.Fill.Visible = msoFalse
        
        .Border.LineStyle = msoLineStyleNone
    End With

    Exit Sub

ErrorHandler:
    MsgBox "Hubo un error: " & Err.Description
End Sub
