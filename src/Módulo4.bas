Attribute VB_Name = "Módulo4"
Sub Consulta_EQUIPAMENTO()
    Dim targetWb As Workbook
    Dim sourceWb As Workbook
    Dim targetWs As Worksheet
    Dim sourceWs As Worksheet
    Dim instruments As String
    Dim DataRange As Range
    Dim lookupValue As Variant
    Dim Lab As Variant
    Dim results As Variant
    Dim i As Long
    
    Set targetWb = ThisWorkbook
    Set targetWs = targetWb.Sheets("Informacoes")
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False
    
    On Error GoTo ErrorHandler
    
    With targetWs
        .Range("C24").Value = ""
        .Range("D8:D19").Value = ""
        .Range("G34:G35").Value = ""
        .Range("D21:D22").Value = ""
        .Range("C29:C32").Value = ""
    End With
    
    instruments = "https://sesirs.sharepoint.com/sites/gdms-ISISistemasdeSensoriamento/Documentos%20Compartilhados/ISI%20SIM%20-%20Metrologia/%23P%C3%BAblico/Desenvolvimento/Automa%C3%A7%C3%A3o%20planilhas/planilhas/Planilhas%20automatizadas/databasek/2-instruments.xlsx"
    
    Set sourceWb = Workbooks.Open(instruments, ReadOnly:=True, UpdateLinks:=0)
    Set sourceWs = sourceWb.Sheets("_2_new_calibrations")
    
    lookupValue = targetWs.Range("G33").Value
    If IsEmpty(lookupValue) Or lookupValue = "" Then
        GoTo CleanUp
    End If
    
    Set DataRange = sourceWs.Range("B1:X100000")
    
    On Error Resume Next
    
    ReDim results(1 To 16)
    results(1) = Application.VLookup(lookupValue, DataRange, 10, False)  ' D8
    results(2) = Application.VLookup(lookupValue, DataRange, 2, False)   ' code
    results(3) = Application.VLookup(lookupValue, DataRange, 5, False)   ' D10
    results(4) = Application.VLookup(lookupValue, DataRange, 6, False)   ' D11
    results(5) = Application.VLookup(lookupValue, DataRange, 4, False)   ' D12
    results(6) = Application.VLookup(lookupValue, DataRange, 3, False)   ' D13
    results(7) = Application.VLookup(lookupValue, DataRange, 7, False)   ' D14
    results(8) = Application.VLookup(lookupValue, DataRange, 9, False)   ' D15
    results(9) = Application.VLookup(lookupValue, DataRange, 20, False)  ' D16
    results(10) = Application.VLookup(lookupValue, DataRange, 21, False) ' D17
    results(11) = Application.VLookup(lookupValue, DataRange, 11, False) ' D18
    results(12) = Application.VLookup(lookupValue, DataRange, 12, False) ' D19
    results(13) = Application.VLookup(lookupValue, DataRange, 16, False) ' G35
    results(14) = Application.VLookup(lookupValue, DataRange, 17, False) ' C24
    results(15) = Application.VLookup(lookupValue, DataRange, 19, False) ' Lab
    
    On Error GoTo 0
    
    With targetWs
        .Range("D8").Value = results(1)
        .Range("D9").Value = results(2)
        .Range("D10").Value = results(3)
        .Range("D11").Value = results(4)
        .Range("D12").Value = results(5)
        .Range("D13").Value = results(6)
        .Range("D14").Value = results(7)
        .Range("D15").Value = results(8)
        .Range("D16").Value = results(9)
        .Range("D17").Value = results(10)
        .Range("D18").Value = results(11)
        .Range("D19").Value = results(12)
        .Range("G35").Value = results(13)
        .Range("C24").Value = results(14)
        .Range("G34").Value = Date
    End With
    
    Lab = results(15)
    
    
    If Not IsError(Lab) Then
        On Error Resume Next
        targetWs.Range("D21").Value = Application.VLookup(Lab, targetWs.Range("M7:P16"), 2, False)
        targetWs.Range("D22").Value = Application.VLookup(Lab, targetWs.Range("M7:P16"), 3, False) & " e " & _
                                     Application.VLookup(Lab, targetWs.Range("M7:P16"), 4, False)
        On Error GoTo 0
    Else
        With targetWs
            .Range("C24").Value = ""
            .Range("D7:D19").Value = ""
            .Range("G34:G35").Value = ""
            .Range("D21:D22").Value = ""
        End With
        MsgBox "Número de certificado não encontrado.", vbCritical
        
        sourceWb.Close SaveChanges:=False
        Exit Sub
        
    End If
    
    Call Consulta_ESTRUTURA

CleanUp:
    On Error Resume Next
    If Not sourceWb Is Nothing Then
        sourceWb.Close SaveChanges:=False
    End If
    
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    
    Exit Sub

ErrorHandler:
    MsgBox "Erro no caminho da fonte de dados.", vbCritical
    Resume CleanUp
End Sub

Sub Consulta_TAG()

    Dim targetWb As Workbook
    Dim sourceWb As Workbook
    Dim targetWs As Worksheet
    Dim sourceWs As Worksheet
    Dim standards As String
    Dim incomplete As String
    Dim i As Long
    Dim cell As Range
    Dim DataRange As Range
    
    
    Set targetWb = ThisWorkbook                                             ' Declara arquivo de saída da query
    Set targetWs = targetWb.Sheets("Informacoes")                           ' Declara aba de saída da query
    
    targetWs.Range("G8:J28").ClearContents
    
    Application.ScreenUpdating = False                                      ' Booleano para ocultar atualizações gráficas de operações em segundo plano do vba
        
    standards = "https://sesirs.sharepoint.com/sites/gdms-ISISistemasdeSensoriamento/Documentos%20Compartilhados/ISI%20SIM%20-%20Metrologia/%23P%C3%BAblico/Desenvolvimento/Automa%C3%A7%C3%A3o%20planilhas/planilhas/Planilhas%20automatizadas/databasek/2-standards.xlsx"
    incomplete = "[VAZIO]"
    
    On Error GoTo ErrorHandler                                              ' Declara método para não encontrar caminho da fonte de dados
    
    Set sourceWb = Workbooks.Open(standards)                                ' Declara função para abrir fonte de dados
    Set sourceWs = sourceWb.Sheets("_2_standards")                          ' Declara aba de busca da fonte de dados
    
    On Error Resume Next

    Set DataRange = sourceWs.Range("B1:U100000")
    
    
    For Each cell In targetWs.Range("J8:J28")
     
        cell.Interior.Color = (RGB(217, 217, 217))
        cell.Font.Color = (RGB(0, 0, 0))
        
    Next cell
    
    For Each cell In targetWs.Range("F8:F28")
        If cell.Value <> "" Then
            targetWs.Cells(cell.Row, "G").Value = Application.WorksheetFunction.VLookup(cell.Value, DataRange, 8, False)        ' itname
            targetWs.Cells(cell.Row, "I").Value = Application.WorksheetFunction.VLookup(cell.Value, DataRange, 2, False)        ' certificate_code
            targetWs.Cells(cell.Row, "H").Value = Application.WorksheetFunction.VLookup(cell.Value, DataRange, 7, False)        ' supplier
            targetWs.Cells(cell.Row, "J").Value = Application.WorksheetFunction.VLookup(cell.Value, DataRange, 5, False)        ' supplier
        End If
    Next cell
    
    For Each cell In targetWs.Range("J8:J28")
    If cell.Value <> "" Then
        'Check if the value is a date and if it's greater than today
        If IsDate(cell.Value) Then
            If CDate(cell.Value) < Date Then
                cell.Interior.Color = RGB(255, 0, 0)
                cell.Font.Color = RGB(255, 255, 255)
            End If
        End If
    End If
    Next cell
    
    On Error GoTo 0

    sourceWb.Close SaveChanges:=False                                       ' Ao encerrar consulta, fecha standards sem salvar alterações
    Application.ScreenUpdating = True

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True                                       ' Restore screen updating on error
    MsgBox "Erro no caminho da fonte de dados.", vbCritical
End Sub

Sub Consulta_ESTRUTURA()

    Dim targetWb As Workbook
    Dim sourceWb As Workbook
    Dim targetWs As Worksheet
    Dim sourceWs As Worksheet
    Dim standards As String
    Dim incomplete As String
    Dim i As Long
    Dim cell As Range
    Dim DataRange As Range
    
    Set targetWb = ThisWorkbook                                             ' Declara arquivo de saída da query
    Set targetWs = targetWb.Sheets("Informacoes")                           ' Declara aba de saída da query
    
    targetWs.Range("O7:P16").ClearContents
    
    Application.ScreenUpdating = False                                      ' Booleano para ocultar atualizações gráficas de operações em segundo plano do vba

    sectors = "https://sesirs.sharepoint.com/sites/gdms-ISISistemasdeSensoriamento/Documentos%20Compartilhados/ISI%20SIM%20-%20Metrologia/%23P%C3%BAblico/Desenvolvimento/Automa%C3%A7%C3%A3o%20planilhas/planilhas/Planilhas%20automatizadas/databasek/3-sectors.xlsx"
    incomplete = "[VAZIO]"
    
    On Error GoTo ErrorHandler                                              ' Declara método para não encontrar caminho da fonte de dados
    
    Set sourceWb = Workbooks.Open(sectors)                                  ' Declara função para abrir fonte de dados
    Set sourceWs = sourceWb.Sheets("_3_sectors")                            ' Declara aba de busca da fonte de dados
    
    On Error Resume Next

    Set DataRange = sourceWs.Range("C1:U20")
    
    For Each cell In targetWs.Range("M7:M16")
        If cell.Value <> "" Then
            targetWs.Cells(cell.Row, "O").Value = Application.WorksheetFunction.VLookup(cell.Value, DataRange, 2, False)        ' temperature
            targetWs.Cells(cell.Row, "P").Value = Application.WorksheetFunction.VLookup(cell.Value, DataRange, 3, False)        ' humidity
        End If
    Next cell
    
    On Error GoTo 0

    sourceWb.Close SaveChanges:=False                                       ' Ao encerrar consulta, fecha sectors sem salvar alterações
    Application.ScreenUpdating = True

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True                                       ' Restore screen updating on error
    MsgBox "Erro no caminho da fonte de dados.", vbCritical
End Sub
