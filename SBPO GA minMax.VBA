Sub ConstruirModeloSCM()
    Dim ws As Worksheet
    Dim chtObj As ChartObject
    
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Otimizacao_SCM_Fix_"
    
    ' ==========================================
    ' CABEÇALHOS
    ' ==========================================
    Dim headers As Variant
    headers = Array("k", "d", "u1", "u2", "u3", "EA", "ED", "EV", "PEA", "PED", "PEV", _
                    "PRP", "PRA", "PRD", "PPA", "PPD", "PPV", "VP", "VA", "VD", "VV", _
                    "IA", "ID", "IV", "mEA", "mED", "mEV", "e1", "e2", "e3", _
                    "J_Track", "J_StockMax", "Pen_Neg", "Pen_VV", "J_Smooth")
    
    ws.Range("A1:AI1").Value = headers
    ws.Range("A1:AI1").Font.Bold = True
    ws.Range("A1:AI1").Interior.Color = RGB(200, 220, 240)
    
    ' ==========================================
    ' ESTADO INICIAL (k = 0) LINHA 2
    ' ==========================================
    ws.Range("A2").Value = 0
    ws.Range("B2").Value = 20 ' d inicial
    ws.Range("C2:E2").Value = 100 ' u warm start
    ws.Range("F2:K2").Value = 100 ' EA..PEV = 100
    ws.Range("L2:X2").Value = 0   ' PRP..IV = 0
    ' As colunas intermediárias (Y até AI) ficam vazias no k=0
    
    ' ==========================================
    ' EQUAÇÕES DINÂMICAS (k >= 1) LINHA >= 3
    ' ==========================================
    ws.Range("A3").Formula = "=A2+1"
    ws.Range("B3").Formula = "=IF(A3<=30, 20, 10)"
    ws.Range("C3:E3").Value = 100 ' Células Variáveis (Decisão)
    
    ' Variáveis intermediárias mE (Calculadas com estados de k-1)
    ws.Range("Y3").Formula = "=MIN(F2,O2)"
    ws.Range("Z3").Formula = "=MIN(G2,P2)"
    ws.Range("AA3").Formula = "=MIN(H2,Q2)"
    
    ' Erros (Calculados com entradas k, estados k-1 e mE atual)
    ws.Range("AB3").Formula = "=C3-I2-L2+Y3"
    ws.Range("AC3").Formula = "=D3-J2-Y3+Z3"
    ws.Range("AD3").Formula = "=E3-K2-Z3+AA3"
    
    ' Estados Principais
    ws.Range("F3").Formula = "=F2+R2-Y3"
    ws.Range("G3").Formula = "=G2+S2-Z3"
    ws.Range("H3").Formula = "=H2+T2-AA3"
    ws.Range("I3").Formula = "=I2+L2-Y3"
    ws.Range("J3").Formula = "=J2+Y3-Z3"
    ws.Range("K3").Formula = "=K2+Z3-AA3"
    
    ' Controladores PID Saturados
    ws.Range("L3").Formula = "=MAX(0, MIN(2000, V2 + 1*AB3))"
    ws.Range("M3").Formula = "=MAX(0, MIN(2000, W2 + 1*AC3))"
    ws.Range("N3").Formula = "=MAX(0, MIN(2000, X2 + 1*AD3))"
    
    ' Posição de Pedidos Pendentes
    ws.Range("O3").Formula = "=MAX(0, MIN(2000, O2-Y3+M3))"
    ws.Range("P3").Formula = "=MAX(0, MIN(2000, P2-Z3+N3))"
    ws.Range("Q3").Formula = "=MAX(0, MIN(2000, Q2-AA3+B3))" ' <--- B3 (Demanda k)
    
    ' Atrasos de Transporte e Venda Realizada
    ws.Range("R3").Formula = "=L3"
    ws.Range("S3").Formula = "=Y3"
    ws.Range("T3").Formula = "=Z3"
    ws.Range("U3").Formula = "=AA3"
    
    ' Integradores
    ws.Range("V3").Formula = "=MAX(-2000, MIN(2000, V2 + 0*AB3))"
    ws.Range("W3").Formula = "=MAX(-2000, MIN(2000, W2 + 0*AC3))"
    ws.Range("X3").Formula = "=MAX(-2000, MIN(2000, X2 + 0*AD3))"
    
    ' Cálculos de Custo
    ws.Range("AE3").Formula = "=(U3-B3)^2"
    ws.Range("AF3").Formula = "=MAX(F3, G3, H3)"
    ws.Range("AG3").Formula = "=MAX(0, -F3)^2 + MAX(0, -G3)^2 + MAX(0, -H3)^2"
    ws.Range("AH3").Formula = "=MAX(0, U3-B3)^2"
    
    ' Smoothness 
    ws.Range("AI3").Formula = "=IF(A3=1, 0, ABS(C3-C2) + ABS(D3-D2) + ABS(E3-E2))"
    
    ' ==========================================
    ' 60 MESES
    ' ==========================================
    ws.Range("A3:AI3").AutoFill Destination:=ws.Range("A3:AI62"), Type:=xlFillDefault
    
    ' ==========================================
    ' OBJETIVO
    ' ==========================================
    ws.Range("AK1").Value = "Custo Total (J)"
    ws.Range("AK1").Font.Bold = True
    ws.Range("AK2").Formula = "=500*SUM(AE3:AE62) + 250*SUM(AF3:AF62) + 10000*SUM(AG3:AG62) + 10000*SUM(AH3:AH62) + 0.1*SUM(AI3:AI62)"
    ws.Range("AK2").Interior.Color = RGB(255, 255, 153)
    ws.Range("AK2").Font.Bold = True
    
    ' ==========================================
    ' GRÁFICOS
    ' ==========================================
    ws.Range("AM1:AP1").Value = Array("Tempo", "u1", "u2", "u3")
    ws.Range("AM1:AP1").Font.Bold = True
    ws.Range("AM1:AP1").Interior.Color = RGB(242, 242, 242)
    
    ws.Range("AM2").Formula = "=INT((ROW()-1)/2)"
    ws.Range("AN2").Formula = "=INDEX($C$2:$C$62, INT((ROW()-2)/2) + 1)"
    ws.Range("AO2").Formula = "=INDEX($D$2:$D$62, INT((ROW()-2)/2) + 1)"
    ws.Range("AP2").Formula = "=INDEX($E$2:$E$62, INT((ROW()-2)/2) + 1)"
    ws.Range("AM2:AP2").AutoFill Destination:=ws.Range("AM2:AP122"), Type:=xlFillDefault
    
    Set chtObj = ws.ChartObjects.Add(Left:=ws.Range("AR2").Left, Top:=ws.Range("AR2").Top, Width:=450, Height:=280)
    With chtObj.Chart
        .ChartType = xlXYScatterLinesNoMarkers
        .SetSourceData Source:=ws.Range("AM1:AP122")
        .HasTitle = True
        .ChartTitle.Text = "Controles"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Amostras (k)"
        .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(0, 112, 192)
        .SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(192, 0, 0)
        .SeriesCollection(3).Format.Line.ForeColor.RGB = RGB(0, 176, 80)
    End With
    
    ws.Columns("A:AP").AutoFit
    ws.Range("C3:E62").Interior.Color = RGB(226, 239, 218)
    
    MsgBox "Modelo pronto! Inserir chutes iniciais em C3:E62", vbInformation
End Sub

