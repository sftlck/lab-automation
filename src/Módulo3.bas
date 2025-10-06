Attribute VB_Name = "Módulo3"
' ==============================
' PROJETO: GERAÇÃO DE CERTIFICADO DE CALIBRAÇÃO
' AUTOMAÇÃO VBA EXCEL E UPLOAD PARA SHAREPOINT
' ==============================

Option Explicit

' Variáveis Globais
Public CertNum As String
Public Certificado As Integer
Public NomeArq As String
Public Cert As String
Public dataEmissao As String
Public Ano As String
Public pontos As Integer
Public pdfNomeArq As String
Public onedriveDiretorio As String
Public pdfsuffix_NomeArq As String
Public OS As String
Public Diretorio As String
Public Username As String
Public pdffinal_NomeArq As String
Public OSfolder As String
Public Std_rmdNomeArq As String
Public pdffinal_rmdNomeArq  As String
Public Std_wordNomeArqOneDrive As String
Public Std_excelNomeArq As String

msgbox "ESTE SUPLEMENTO FOI IMPORTADO DO GITHUB"

' -------------------------------
' FUNÇÃO CENTRAL: Obtem instância do Word
' -------------------------------
Function GetWordApp() As Word.Application
    On Error Resume Next
    Set GetWordApp = GetObject(, "Word.Application")
    If GetWordApp Is Nothing Then
        Set GetWordApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    GetWordApp.Visible = True
End Function

Function ExtraiCertNum(inputText As String) As String
Dim result As String
    If InStr(inputText, "/") > 0 Then
        ExtraiCertNum = Split(inputText, "/")(0)
    Else
        ExtraiCertNum = inputText
    End If
End Function
' -------------------------------
' ABERTURA DO DOCUMENTO-MODELO
' -------------------------------
Sub AbreDocumentoModelo()
    Dim oWord As Word.Application
    Set oWord = GetWordApp()
    Dim oDoc As Word.Document ' adicionado

    Dim caminhoModelo As String
    Dim Std_NomeArq As String
    Dim doc_NomeArq As String
    Dim Tipo As String
    Dim sourceWb As Workbook
    Dim result As String
    Dim AnoCert As String
    Dim OSNum As String
    Dim Diretorio2 As String
    Dim suffix_NomeArq As String
    Dim doc_suffixNomeArq As String
    Dim Std_wordNomeArq As String
             
    get_Username
    
    dataEmissao = Sheets("Informacoes").Cells(34, 7).Text
    
    Diretorio = "https://sesirs.sharepoint.com/sites/gdms-ISISistemasdeSensoriamento/Documentos%20Compartilhados/ISI%20SIM%20-%20Metrologia/%23P%C3%BAblico/Desenvolvimento/Automa%C3%A7%C3%A3o%20planilhas/Salvar/"
    
    With Sheets("Informacoes")
        Tipo = .Range("D7").Text
        CertNum = .Cells(33, 7).Text
        Cert = ExtraiCertNum(CertNum)
        OSNum = .Cells(8, 4).Text
        OS = ExtraiCertNum(OSNum)
        
        If Tipo = "" Or IsError(Tipo) Then
            MsgBox "Célula 'Tipo' não preenchido.", vbCritical
            Exit Sub
        End If
        
        If Tipo = .Range("N23").Text Then
            caminhoModelo = Diretorio & "Modelo Padrão RBC - 2025-28.dotx"
        End If
        If Tipo = .Range("N24").Text Then
            caminhoModelo = Diretorio & "Modelo Padrão - 2025-28.dotx"
        End If
    End With
    
    Set oDoc = oWord.Documents.Open(caminhoModelo)
    oWord.Activate
    
    Ano = Right(dataEmissao, 4)
    AnoCert = Right(Ano, 2)
    Diretorio = "https://sesirs.sharepoint.com/sites/gdms-ISISistemasdeSensoriamento/Documentos%20Compartilhados/ISI%20SIM%20-%20Metrologia/%23P%C3%BAblico/Desenvolvimento/Automa%C3%A7%C3%A3o%20planilhas/Salvar/"
    OS = CharOS_Normalize(OS)
    Std_NomeArq = Diretorio & "Ca-" & Ano & "\" & OS & "\Cliente\" & Cert & "-" & AnoCert
    Std_wordNomeArq = Diretorio & "Ca-" & Ano & "\" & OS & "\" & Cert & "-" & AnoCert & ".doc"
    Std_wordNomeArqOneDrive = "C:\Users\" & Username & "\Sistema Fiergs\GDMS - ISI SIM - Metrologia - Salvar\" & "Ca-" & Ano & "\" & OS & "\" & Cert & "-" & AnoCert & ".doc"
    Std_excelNomeArq = Diretorio & "Ca-" & Ano & "\" & OS & "\" & Cert & "-" & AnoCert & ".xlsm"
    
    Std_rmdNomeArq = Diretorio & "Ca-" & Ano & "\" & OS & "\rmd " & Cert & "-" & AnoCert
    suffix_NomeArq = "Ca-" & Ano & "\" & OS & "\" & Cert & "-" & AnoCert
    doc_NomeArq = Std_NomeArq & ".doc"
    
    pdffinal_rmdNomeArq = Diretorio & "Ca-" & Ano & "\" & OS & "\rmd " & Cert & "-" & AnoCert
    
    pdfNomeArq = Std_NomeArq & ".pdf"
    pdfsuffix_NomeArq = suffix_NomeArq & ".pdf"
    pdffinal_NomeArq = OS & "\" & Cert & "-" & AnoCert & ".pdf"
    pdffinal_rmdNomeArq = Std_rmdNomeArq & ".pdf"
    
    doc_suffixNomeArq = suffix_NomeArq & ".doc"
    
    onedriveDiretorio = "C:\Users\" & Username & "\Sistema Fiergs\GDMS - ISI SIM - Metrologia - Salvar\" & pdfsuffix_NomeArq
    OSfolder = "C:\Users\" & Username & "\Sistema Fiergs\GDMS - ISI SIM - Metrologia - Salvar\" & "Ca-" & Ano & "\" & OS
    
    oDoc.SaveAs Std_wordNomeArq
End Sub
Function get_Username()
    
    Username = Environ("Username")
    
End Function
Sub Check_SaveAs2(onedriveDiretorio As String, pdffinal_NomeArq As String)
    OSfolder = OSfolder & "\"
    
    If Dir(OSfolder) = "" Then
        MkDir (OSfolder)
        MkDir (OSfolder & "Cliente\")
        Application.Wait (Now + TimeValue("0:00:02"))                                   'CASTRO: DAR TEMPO PARA ONEDRIVE SINCRONIZAR CRIAÇÃO DE PASTA VIA CAMINHO ABSOLUTO COM O SHAREPOINT PARA SALVAR ARQUIVOS
        Export_pdf
        GetWordApp().Application.Quit SaveChanges:=wdDoNotSaveChanges
        MsgBox "Certificado gerado com sucesso!"
    Else
        If Dir(onedriveDiretorio) = "" Then
            Export_pdf
            GetWordApp().Application.Quit SaveChanges:=False
            Application.Wait (Now + TimeValue("0:00:02"))                           ' CASTRO: É NECESSÁRIO USAR TEMPO PARA O SHELL ENCERRAR A APLICAÇÃO ANTES DE EXCLUIR O REGISTRO
            Kill Std_wordNomeArqOneDrive
            MsgBox "Certificado gerado com sucesso!"
            Else:
            Dim userResponse As VbMsgBoxResult
            userResponse = MsgBox("Gostaria de sobrescrever o arquivo da OS " & pdffinal_NomeArq & "?", vbYesNo + vbQuestion, "Sobrescrever arquivo")
            If userResponse = vbYes Then
                Export_pdf
                MsgBox "Certificado gerado com sucesso!"
            ElseIf userResponse = vbNo Then
                MsgBox "Salvamento de arquivo .pdf cancelado"
                Exit Sub
            End If
        End If
    End If
End Sub
Sub Export_pdf()
    GetWordApp().ActiveDocument.SaveAs2 pdfNomeArq, wdFormatPDF
    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, _
                                      filename:=pdffinal_rmdNomeArq, _
                                      Quality:=xlQualityStandard, _
                                      IncludeDocProperties:=True, _
                                      IgnorePrintAreas:=False
    ActiveWorkbook.SaveAs Std_excelNomeArq, xlOpenXMLWorkbookMacroEnabled, "5"
    
End Sub

Function CharOS_Normalize(OS As String)
    Dim LenOS As String
    LenOS = Len(OS)
    If LenOS = 1 Then
        CharOS_Normalize = "000" & OS
    End If
    If LenOS = 2 Then
        CharOS_Normalize = "00" & OS
    End If
    If LenOS = 3 Then
        CharOS_Normalize = "0" & OS
    End If
    If LenOS = 4 Then
        CharOS_Normalize = OS
    End If
    
End Function
' -------------------------------
' COLAGEM DE TEXTO FORMATADO
' -------------------------------
Sub ColaTextoFormatado(texto As String, Optional negrito As Boolean = False, Optional italico As Boolean = False, Optional corFonte As Long = -16777216, Optional centralizar As Boolean = False)
    ' corFonte padrão: preto (cor RGB = -16777216)
    Dim oWord As Word.Application
    Set oWord = GetWordApp()

    With oWord.Selection
        .Font.Bold = negrito
        .Font.Italic = italico
        .Font.Color = corFonte
        .ParagraphFormat.Alignment = IIf(centralizar, wdAlignParagraphCenter, wdAlignParagraphLeft)
        .TypeText Text:=texto
    End With
End Sub
' -------------------------------
' INSERÇÕES NO DOCUMENTO
' -------------------------------
Sub InsereParagrafoWord()
    GetWordApp().Selection.TypeParagraph
End Sub

Sub InserePaginaNova()
    With GetWordApp().Selection
        .InsertBreak Type:=wdPageBreak
        .InsertParagraphAfter
    End With
End Sub

Sub SelecionaIndicador(indicador As String)
    GetWordApp().ActiveDocument.Bookmarks(indicador).Select
End Sub

' -------------------------------
' CLIENTE E REQUERENTE
' -------------------------------
Sub InsereCliente()
    Dim Cliente As String, EndCliente As String
    Dim Requerente As String, EndRequerente As String
    Dim CONTAR As Integer: CONTAR = 0

    With Sheets("Informacoes")
        Cliente = .Range("D14").Text
        EndCliente = .Range("D15").Text
        Requerente = .Range("D16").Text
        EndRequerente = .Range("D17").Text
    End With

    SelecionaIndicador "Cliente"
    
    If EndRequerente = EndCliente Then
        ColaTextoFormatado Cliente
        InsereParagrafoWord
        ColaTextoFormatado EndCliente
        With Sheets("Informacoes")
            SelecionaIndicador "Responsavel": ColaTextoFormatado .Range("D18").Text
            SelecionaIndicador "Contato": ColaTextoFormatado .Range("D19").Text
        End With
    Else
        ColaTextoFormatado Cliente
        InsereParagrafoWord
        ColaTextoFormatado EndCliente
        With Sheets("Informacoes")
            SelecionaIndicador "Responsavel": ColaTextoFormatado .Range("D18").Text
            SelecionaIndicador "Contato": ColaTextoFormatado .Range("D19").Text
        End With
        InsereParagrafoWord
        InsereParagrafoWord
        ColaTextoFormatado "Solicitante: ", True
        InsereParagrafoWord
        ColaTextoFormatado Requerente
        InsereParagrafoWord
        ColaTextoFormatado EndRequerente
    End If

    'Do While CONTAR < 4
    '    GetWordApp().Selection.Delete Unit:=wdCharacter, Count:=1
    '    CONTAR = CONTAR + 1
    'Loop
End Sub

' -------------------------------
' PROCEDIMENTOS
' -------------------------------
Sub InsereProcedimento()
    Dim i As Integer, texto As String
    With Sheets("Informacoes")
        SelecionaIndicador "PC"
        For i = 24 To 26
            texto = .Cells(i, 3).Text
            'texto = .Cells(i, 3).Text & " - Revisão: " & .Cells(i, 4).Text             ' CASTRO: Ocultado pois <procedure_instrument> não tem entidade <revision>
            If .Cells(i, 3).Text = "" Then
                Exit For
            ElseIf .Cells(i, 3).Text <> "" And i = 24 Then
                ColaTextoFormatado texto
            ElseIf .Cells(i, 3).Text <> "" Then
                InsereParagrafoWord
                ColaTextoFormatado texto
            End If
        Next i
    End With
    'GetWordApp().Selection.TypeBackspace
End Sub

' -------------------------------
' MÉTODOS
' -------------------------------
Sub InsereMetodo()
    Dim i As Integer, texto As String
    With Sheets("Informacoes")
        SelecionaIndicador "Métodos"
        For i = 29 To 32
            texto = .Cells(i, 3).Text
            If texto = "" Then
                Exit For
            ElseIf texto <> "" And i = 29 Then
                ColaTextoFormatado texto
            ElseIf texto <> "" Then
                InsereParagrafoWord
                ColaTextoFormatado texto
            End If
        Next i
    End With
    'GetWordApp().Selection.TypeBackspace
End Sub

' -------------------------------
' OBSERVAÇÕES
' -------------------------------
Sub InsereObservacoes()
    With Sheets("Informacoes")
            SelecionaIndicador "Observacoes": ColaTextoFormatado .Range("D34").Text
    End With
End Sub
' -------------------------------
' LOCALIZAÇÃO
' -------------------------------
Sub InsereLocalizacao()

Dim LocalizacaoInfo As String
Dim LocalizacaoRange As Range
Dim Localizacao As String

With Sheets("Informacoes")
        Localizacao = .Range("D35").Text
        SelecionaIndicador "Localizacao": ColaTextoFormatado Localizacao
End With

End Sub
Sub InsereSignatario()
    With Sheets("Informacoes")
            SelecionaIndicador "Signatario": ColaTextoFormatado .Range("G35").Text, centralizar:=True
    End With
End Sub

' -------------------------------
' IMAGEM
' -------------------------------
Sub InsereImagem()

Dim wImagens As Worksheet
Set wImagens = ThisWorkbook.Sheets("Imagens")
Dim a As String

If Check_Imagem(wImagens, wImagens.Range("Imagem0")) = True And Check_Imagem(wImagens, wImagens.Range("Imagem1")) = True Then
    SelecionaIndicador "Figura": ColaTextoFormatado "IMAGENS 1 E 2", True
End If

If Check_Imagem(wImagens, wImagens.Range("Imagem0")) = True And Check_Imagem(wImagens, wImagens.Range("Imagem1")) = False Then
    SelecionaIndicador "Figura": ColaTextoFormatado "IMAGEM 1", True
End If

With wImagens
If Check_Imagem(wImagens, wImagens.Range("Imagem0")) = True Then
    SelecionaIndicador "Imagem0": ColaImagem wImagens, .Range("Imagem0"), , , True
    'InsereParagrafoWord
    'GetWordApp().Selection.TypeParagraph

If Check_Imagem(wImagens, wImagens.Range("Imagem1")) = True Then
    SelecionaIndicador "Imagem1": ColaImagem wImagens, .Range("Imagem1"), , , True
End If
End If
End With

End Sub
Function Check_Imagem(sheet As Worksheet, cell As Range)
Dim a As String
Dim b As String
If cell.MergeCells Then
        Set cell = cell.MergeArea(1, 1)
    End If
    
    Dim img As Object
    Set img = Nothing
    
    On Error Resume Next
    For Each img In sheet.Shapes
        If Not Intersect(img.TopLeftCell, cell) Is Nothing Then
            Exit For
        End If
    Next img
    On Error GoTo 0
    
    If Not img Is Nothing Then
        Check_Imagem = True
    Else
        Check_Imagem = False
    End If

End Function
' -------------------------------
' OBTER IMAGEM
' -------------------------------
Sub ColaImagem(sheet As Worksheet, cell As Range, Optional largura As Double = -1, Optional altura As Double = -1, Optional centralizar As Boolean = False)
    
    If cell.MergeCells Then
        Set cell = cell.MergeArea(1, 1)
    End If
    
    Dim img As Object
    Set img = Nothing
    
    On Error Resume Next
    For Each img In sheet.Shapes
        If Not Intersect(img.TopLeftCell, cell) Is Nothing Then
            Exit For
        End If
    Next img
    On Error GoTo 0
    
    If Not img Is Nothing Then
        img.Copy
        
        Dim oWord As Word.Application
        Set oWord = GetWordApp()
        
        With oWord.Selection
            .PasteAndFormat (wdPasteDefault)
            
            If centralizar Then
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
            End If
            
            If .InlineShapes.Count > 0 Then
                With .InlineShapes(1)
                    .Width = 10
                    .Height = 10
                    'If largura > 0 Then .Width = largura
                    'If altura > 0 Then .Height = altura
                End With
            End If
            
            .TypeParagraph
        End With
    Else
        MsgBox "Nenhuma imagem encontrada na célula " & cell.Address, vbCritical
        Exit Sub
    End If
End Sub
' -------------------------------
' PADRÕES
' -------------------------------
Sub InserePadroes()
    Dim CONTAR As Integer: CONTAR = 1
    Dim texto As String
    With Sheets("Informacoes")
        SelecionaIndicador "TAGP1"
        Do While .Cells(7 + CONTAR, 6).Text <> ""
            texto = .Cells(7 + CONTAR, 7).Text & ", identificação " & .Cells(7 + CONTAR, 6).Text & ", certificado número " & .Cells(7 + CONTAR, 8).Text & " emitido por " & .Cells(7 + CONTAR, 9).Text & ", com validade até " & .Cells(7 + CONTAR, 10).Value
            If .Cells(7 + CONTAR, 10).Value < .Cells(34, 7).Value Then
                ColaTextoFormatado texto, True, False, vbRed
            Else
                ColaTextoFormatado texto
            End If
            InsereParagrafoWord
            CONTAR = CONTAR + 1
        Loop
    End With
    GetWordApp().Selection.TypeBackspace
End Sub

' -------------------------------
' SUBSTITUIÇÃO DE SÍMBOLOS
' -------------------------------
Sub SubstituiTexto(findText As String, replaceText As String)
    With GetWordApp().Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replaceText
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
    End With
    GetWordApp().Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub SubstituiOhm()
    SubstituiTexto "Ohm", ChrW(937)
End Sub

Sub SubstituiInfinito()
    SubstituiTexto "Infinito", ChrW(8734)
End Sub

' -----------------------------------
' PERÍODO DE CALIBRAÇÃO
' -----------------------------------
Sub PeriodoCalibracao()
    Dim dtIni As Variant, dtFim As Variant
    Dim texto As String
    With Sheets("Informacoes")
        dtIni = .Range("G31").Value
        dtFim = .Range("G32").Value
        If dtIni = dtFim Then
            '.Range("BH1") = "Data da calibração: "
            '.Range("BK1") = "Data da calibração: " & dtIni
            texto = "Data da calibração: " & dtIni
        Else
            '.Range("BH1") = "Período de calibração: "
            '.Range("BK1") = "Período de calibração: " & dtIni & " a " & dtFim
            texto = "Período de calibração: " & dtIni & " a " & dtFim
        End If
    End With
    SelecionaIndicador "PeriodoCalibracao"
    ColaTextoFormatado texto
End Sub

' -------------------------------
' INSERÇÃO DE PLANILHAS DOS RESULTADOS
' -------------------------------
Sub InsereTabelaPlanilha()
    On Error GoTo TrataErro

    Dim intervaloTexto As String
    Dim wsResultados As Worksheet
    Dim rngTabela As Range
    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim a As Integer
    Dim b As Integer
    
    ' === Etapa 1: Obter a referência da planilha "Resultados" ===
    Set wsResultados = ThisWorkbook.Sheets("Resultados")


    For a = 1 To 4
        For b = 2 To 6
            ' === Etapa 2: Obter o intervalo de células descrito na célula B2 ===
            intervaloTexto = Trim(wsResultados.Cells(b, 2 * a).Value)
            If intervaloTexto = "" Then
                MsgBox "A célula com dados da aba 'Resultados' está vazia. Nenhum intervalo foi especificado.", vbExclamation
                Exit Sub
            End If
            ' === Etapa 3: Tentar criar o intervalo a partir do texto ===
            On Error Resume Next
            Set rngTabela = wsResultados.Range(intervaloTexto)
            On Error GoTo TrataErro

            If rngTabela Is Nothing Then
                MsgBox "O intervalo '" & intervaloTexto & "' informado em dados é inválido.", vbCritical
                Exit Sub
            End If

            ' === Etapa 4: Copiar intervalo ===
            rngTabela.Copy

            ' === Etapa 5: Obter o Word ativo e o documento aberto ===
            Set oWord = GetObject(, "Word.Application")
            Set oDoc = oWord.ActiveDocument

            ' === Etapa 6: Colar a tabela no marcador "Planilhas" ===
            If oDoc.Bookmarks.Exists("Planilhas") Then
                With oDoc.Bookmarks("Planilhas").Range
                    .PasteExcelTable LinkedToExcel:=False, WordFormatting:=False, RTF:=False
                End With
            Else
                MsgBox "O marcador 'Planilhas' não foi encontrado no documento Word.", vbCritical
            End If
        Next b
    Next a
            Application.CutCopyMode = False
            Exit Sub

TrataErro:
    MsgBox "Erro ao inserir a tabela no Word: " & Err.Description, vbCritical
End Sub

' -------------------------------
' MÓDULO DE EXECUÇÃO FINAL
' -------------------------------
Sub GerarCertificadoCompleto1()

    Dim Cert As String
    
    Application.ScreenUpdating = False

    ' Exemplo: valores devem vir da interface (UserForm ou célula)
    CertNum = Sheets("Informacoes").Cells(33, 7).Text
    dataEmissao = Format(Sheets("Informacoes").Cells(34, 7).Value, "dd/mm/yyyy")
    Ano = Right(dataEmissao, 4)
    Cert = Left(CertNum, 5)
    Certificado = 1
    pontos = 0

    ' Abre Word com documento modelo
    AbreDocumentoModelo
    If Certificado = 0 Then Exit Sub

    ' Cabeçalho do certificado
    SelecionaIndicador "Número": ColaTextoFormatado texto:=CertNum, centralizar:=True, negrito:=True
'    SelecionaIndicador "Data": ColaTextoFormatado dataEmissao

InsereCliente
    ' Dados principais
    With Sheets("Informacoes")
        SelecionaIndicador "Equipamento": ColaTextoFormatado .Range("D10").Text
        SelecionaIndicador "Fabricante": ColaTextoFormatado .Range("D11").Text
        SelecionaIndicador "CondAmb": ColaTextoFormatado texto:=.Range("D22").Text
        SelecionaIndicador "Modelo": ColaTextoFormatado .Range("D12").Text
        SelecionaIndicador "Série": ColaTextoFormatado .Range("D13").Text 'não retornou
        If .Range("D9").Text <> "" Then
            SelecionaIndicador "TAG": ColaTextoFormatado .Range("D9").Text
        End If
        PeriodoCalibracao
        SelecionaIndicador "Data": ColaTextoFormatado dataEmissao
        SelecionaIndicador "Protocolo": ColaTextoFormatado .Range("D8").Text 'não retornou
    End With

    ' Procedimentos e métodos
    InserePadroes
    InsereProcedimento
    InsereMetodo
    InsereObservacoes
    InsereLocalizacao
    InsereSignatario
    
    InsereImagem
        
'    SubstituiOhm
'    SubstituiInfinito

    InsereTabelaPlanilha
    'GetWordApp().ActiveDocument.Save
    GetWordApp().ActiveWindow.View.Type = wdPrintView
    
    Check_SaveAs2 onedriveDiretorio, pdffinal_NomeArq
    
    Application.ScreenUpdating = True
End Sub



