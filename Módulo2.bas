Attribute VB_Name = "Módulo2"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveWorkbook.SaveAs filename:= _
        "https://sesirs.sharepoint.com/sites/gdms-ISISistemasdeSensoriamento/Documentos%20Compartilhados/ISI%20SIM%20-%20Metrologia/%23Público/%23Laboratórios/Proposta%20planilha1.xlsm" _
        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
End Sub

Sub InsereTabelaPlanilha1(oDoc As Word.Document)
    On Error GoTo TrataErro  ' Ativa tratamento de erros

    Dim intervaloTexto As String         ' Armazena o texto da célula B2 (ex: "A1:F20")
    Dim wsResultados As Worksheet        ' Referência à planilha "Resultados"
    Dim rngTabela As Range               ' Intervalo a ser copiado
    Dim oWord As Word.Application        ' Instância do Word

    ' === ETAPA 1: Obter referência da planilha "Resultados" ===
    Set wsResultados = ThisWorkbook.Sheets("Resultados")

    ' === ETAPA 2: Ler da célula B2 o intervalo a ser copiado ===
    intervaloTexto = Trim(wsResultados.Range("B2").Value)

    ' === ETAPA 3: Validar o intervalo ===
    If intervaloTexto = "" Then
        MsgBox "A célula B2 da aba 'Resultados' está vazia.", vbExclamation
        Exit Sub
    End If

    ' === ETAPA 4: Definir o intervalo a partir do texto lido ===
    ' Ex: se B2 = "A1:F20", então pega esse intervalo da planilha "Resultados"
    Set rngTabela = wsResultados.Range(intervaloTexto)

    ' === ETAPA 5: Copiar intervalo ===
    rngTabela.Copy

    ' === ETAPA 6: Inserir no Word no marcador "Planilhas" ===
    Set oWord = oDoc.Application  ' Garante que temos o Word ativo

    If oDoc.Bookmarks.Exists("Planilhas") Then
        With oDoc.Bookmarks("Planilhas").Range
            .PasteExcelTable LinkedToExcel:=False, WordFormatting:=False, RTF:=False
        End With
    Else
        MsgBox "O marcador 'Planilhas' não foi encontrado no documento Word.", vbCritical
    End If

    ' === ETAPA 7: Limpar modo de cópia ===
    Application.CutCopyMode = False

    Exit Sub

TrataErro:
    MsgBox "Erro ao inserir tabela no Word: " & Err.Description, vbCritical
End Sub
Sub GerarCertificadoCompleto()
    On Error GoTo TrataErro

    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim caminhoModelo As String
    Dim NomeArq As String
    Dim Diretorio As String
    Dim CertNum As String
    Dim Ano As String
    Dim pastaDestino As String

    ' Caminho do modelo
    caminhoModelo = ThisWorkbook.path & "\ModeloCertificado.dotx"
    
    ' Obtem dados da aba "Informacoes"
    CertNum = Sheets("Informacoes").Range("G33").Value
    dataEmissao = Sheets("Informacoes").Range("G34").Value
    Ano = Right(dataEmissao, 4)

    ' Diretório base (ajustar se necessário)
    Diretorio = ThisWorkbook.path & "\"
    pastaDestino = Diretorio & "Ca-" & Ano & "\"

    ' Cria pasta se não existir
    If Dir(pastaDestino, vbDirectory) = "" Then MkDir pastaDestino
    
    ' Caminho completo do arquivo final
    NomeArq = pastaDestino & CertNum & "-" & Ano & ".docx"

    ' Instancia o Word e abre o modelo
    Set oWord = GetWordApp()
    Set oDoc = oWord.Documents.Open(caminhoModelo)
    oWord.Visible = True

    ' Preenche o certificado com os módulos existentes
    PeriodoCalibracao
    InsereProcedimento
    InsereMetodo
    InserePadroes
    InsereCliente
    Call InsereTabelaPlanilha(oDoc)
    
    'Call InsereCliente(oDoc)
    'Call InsereMetodo(oDoc)
    'Call InsereInstrumento(oDoc)
    'Call InsereCertificado(oDoc)
    'Call InsereNormas(oDoc)
    'Call InsereTabelaPlanilha(oDoc)
    'Call InsereAssinatura(oDoc)

    ' Salva como .docx
    oDoc.SaveAs2 NomeArq
    Debug.Print "Documento salvo em: " & NomeArq

    ' Gera PDF opcional
    Dim NomePDF As String
    NomePDF = Replace(NomeArq, ".docx", ".pdf")
    oDoc.ExportAsFixedFormat OutputFileName:=NomePDF, ExportFormat:=17 ' 17 = wdExportFormatPDF
    Debug.Print "PDF salvo em: " & NomePDF

    ' Fecha o Word (opcional: comentar se quiser manter aberto para revisão)
    oDoc.Close SaveChanges:=False
    oWord.Quit

    MsgBox "Certificado gerado com sucesso!", vbInformation

    Exit Sub

TrataErro:
    MsgBox "Erro ao gerar certificado: " & Err.Description, vbCritical
    If Not oDoc Is Nothing Then oDoc.Close SaveChanges:=False
    If Not oWord Is Nothing Then oWord.Quit
End Sub


