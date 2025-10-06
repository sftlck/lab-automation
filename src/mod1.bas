Attribute VB_Name = "Módulo1"
Sub OpenWordDoc()
'
Dim wdApp As Word.Application, wdDoc As Word.Document
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If Err.Number <> 0 Then 'Word isn't already running
    Set wdDoc = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    Set wdDoc = Word.Application.Documents.Open("C:\Users\carlos.nadaletti\Desktop\IST PGE\Novos Serviços\Eletricidade\Planilhas\Elet\Modelo Padrão-2024.dotx")
    Word.Application.Visible = True
    wdDoc.Activate
End Sub

Sub AbreWord()
'
Dim NumAno, CertInf, Diretorio As String
Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim oRange As Word.Range
    Certificado = 1
    Set oWord = CreateObject("Word.Application")
        With oWord
            .Visible = True
            .Activate
            .WindowState = wdWindowStateNormal
        End With
    Set oDoc = oWord.Documents.Open("C:\Users\carlos.nadaletti\Desktop\IST PGE\Novos Serviços\Eletricidade\Planilhas\Elet\Modelo Padrão-2024.dotx")
    oDoc.Activate
    
    Diretorio = "C:\Users\carlos.nadaletti\Desktop\IST PGE\Novos Serviços\Eletricidade\Planilhas"
    NumAno = Right(CertNum, 4)
    CertInf = Left(CertNum, 5)
    
    ActiveWorkbook.Saved = True
    ActiveWorkbook.Save
        
    NomeArq = Diretorio & "\Ca-" & NumAno & "\" & CertInf & "-" & NumAno & ".doc"
    If Dir(NomeArq) = CertInf & "-" & NumAno & ".doc" Then

        If MsgBox("Já existe o Arquivo ' " & NomeArq & "'. Deseja Substituí-lo ?", vbYesNo + vbQuestion + vbDefaultButton2, "Planilha On-Line - Confirmação de Nº de Certificado") = vbNo Then
          Certificado = 0
          
          Exit Sub
        End If
     End If
     MsgBox (NomeArq)
     oDoc.SaveAs NomeArq
End Sub

Sub CloseWordDoc()
'
Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim oRange As Word.Range
    Set oWord = GetObject(, "Word.Application")
    oWord.Selection.TypeParagraph
    'oWord.Selection.Paste
    oWord.Activate
    oWord.Visible = True
    If oWord.ActiveWindow.View.SplitSpecial = wdPaneNone Then
        oWord.ActiveWindow.ActivePane.View.Type = wdPrintView
    Else
        oWord.ActiveWindow.View.Type = wdPrintView
    End If
    'oWord.Application.Run MacroName:="Vencidos"
    oWord.Documents.Save
    'oWord.Documents.Close
End Sub

Sub CloseDoc()
'
    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range
    Set oWord = GetObject(, "Word.Application")
    oWord.Activate
    oWord.Visible = True
    oWord.Documents.Close
End Sub

Sub PasteToWord()
'
Dim AppWord As Word.Application
Set AppWord = GetObject(, "Word.Application")
    AppWord.Visible = True
    AppWord.Selection.PasteAndFormat wdFormatOriginalFormatting
    AppWord.Selection.Tables(1).Rows.Alignment = wdAlignRowCenter
    AppWord.Selection.Tables(1).Rows.HeightRule = wdRowHeightAtLeast
    AppWord.Selection.Tables(1).PreferredWidthType = wdPreferredWidthPercent
    AppWord.Selection.Tables(1).PreferredWidth = 100
    AppWord.Selection.Tables(1).Rows.Borders.OutsideLineStyle = wdLineStyleSingle
    Application.CutCopyMode = False
End Sub

Sub ColaTextoWord(texto As String)
'
On Error GoTo MergeButton_Err
    Dim objWord As Word.Application
    Set objWord = GetObject(, "Word.Application")
    With objWord
        'torna a aplicaçao visivel
        .Visible = True
        .Selection.TypeText Text:=texto
    End With
        Set objWord = Nothing
        Exit Sub
MergeButton_Err:
    'campo em branco
        If Err.Number = 94 Then
            objWord.Selection.Text = ""
            Resume Next
        Else
            MsgBox Err.Number & vbCr & Err.Description
        End If
    Exit Sub
End Sub

Sub deleta()
'
Dim objWord As Word.Application
Dim oDoc As Word.Document
    Set objWord = GetObject(, "Word.Application")
    objWord.Selection.TypeBackspace
End Sub

Sub deletar()
'
Dim objWord As Word.Application
Dim oDoc As Word.Document
    Set objWord = GetObject(, "Word.Application")
    objWord.Selection.Delete Unit:=wdCharacter, Count:=1
End Sub

Sub ColaTitulo(texto As String)
'Dim texto As String
Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim oRange As Word.Range
Dim pontos As Integer
pontos = 0
Set oWord = GetObject(, "Word.Application")
    ActiveSheet.Select
    'texto = Range("J95").Text
    oWord.Selection.Font.Bold = True
    oWord.Selection.Font.Italic = False
    oWord.Selection.Font.Color = wdColorBlack
    oWord.Selection.Font.Size = 11
    oWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    ColaTextoWord (texto)
    oWord.Visible = True
End Sub

Sub ColaNormal(texto As String)
'
Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim oRange As Word.Range
Set oWord = GetObject(, "Word.Application")
    ActiveSheet.Select
    oWord.Selection.Font.Bold = False
    oWord.Selection.Font.Italic = False
    ColaTextoWord (texto)
    oWord.Visible = True
End Sub

Sub ColaNegritoVermelho(texto As String)
'
Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim oRange As Word.Range
Set oWord = GetObject(, "Word.Application")
    ActiveSheet.Select
    oWord.Selection.Font.Bold = False
    oWord.Selection.Font.Italic = False
    oWord.Selection.Font.Color = wdColorRed
    ColaTextoWord (texto)
    oWord.Visible = True
End Sub

Sub ColaNegrito(texto As String)
'
Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim oRange As Word.Range
Set oWord = GetObject(, "Word.Application")
    ActiveSheet.Select
    oWord.Selection.Font.Bold = True
    ColaTextoWord (texto)
    oWord.Visible = True
End Sub

Sub ColaNegritoItalico(texto As String)
'
Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim oRange As Word.Range
Set oWord = GetObject(, "Word.Application")
    ActiveSheet.Select
    oWord.Selection.Font.Bold = True
    oWord.Selection.Font.Italic = True
    ColaTextoWord (texto)
    oWord.Visible = True
End Sub

Sub InserePaginaNova()
'
Dim AppWord As Word.Application
Set AppWord = GetObject(, "Word.Application")
    AppWord.Visible = True
    AppWord.Selection.InsertBreak Type:=wdPageBreak
    AppWord.Selection.InsertParagraphAfter
    Application.CutCopyMode = False
End Sub

Sub InsereParagrafoWord()
      
On Error GoTo MergeButton_Err
    Dim objWord As Word.Application
    Set objWord = GetObject(, "Word.Application")
    With objWord
        'torna a aplicaçao visivel
        .Visible = True
        .Selection.TypeParagraph
        '.Selection.text = Chr(13) & Chr(10)
    End With
        Set objWord = Nothing
        Exit Sub
MergeButton_Err:
    ' campo em branco
        If Err.Number = 94 Then
            objWord.Selection.Text = ""
        Resume Next
        Else
            MsgBox Err.Number & vbCr & Err.Description
        End If
    Exit Sub
End Sub

Sub SelecionaIndicador(indicador As String)

On Error GoTo MergeButton_Err
    Dim objWord As Word.Application
    ' inicia o word
    Set objWord = GetObject(, "Word.Application")
    With objWord
        .Visible = True
        .ActiveDocument.Bookmarks(indicador).Select
    End With
        'imprime o documento
        Set objWord = Nothing
        Exit Sub
MergeButton_Err:
    ' campo em branco
        If Err.Number = 94 Then
            objWord.Selection.Text = ""
        Resume Next
        Else
            MsgBox Err.Number & vbCr & Err.Description
        End If
    Exit Sub
End Sub

Sub InsereCliente()
'
Dim Cliente As String
Dim EndCliente As String
Dim Requerente As String
Dim EndRequerente As String
Dim CONTAR As Integer
Sheets("Dados").Select
SelecionaIndicador ("Cliente")
Cliente = Range("I6").Text
EndCliente = Range("J6").Text & " - " & Range("K6").Text & " - " & Range("L6").Text & " - " & Range("M6").Text
Requerente = Range("O6").Text
EndRequerente = Range("P6").Text & " - " & Range("Q6").Text & " - " & Range("R6").Text & " - " & Range("S6").Text
'MsgBox (Cliente & Requerente)
If Cliente = Requerente Then
    ColaTextoWord (Cliente)
    InsereParagrafoWord
    'AlinhaEndereco
    ColaTextoWord (EndCliente)
    'AlinhaEndereco
Else
    ColaTextoWord (Cliente)
    InsereParagrafoWord
    'AlinhaEndereco
    ColaTextoWord (EndCliente)
    InsereParagrafoWord
    InsereParagrafoWord
    ColaNegrito ("Solicitante: ")
    InsereParagrafoWord
    ColaNormal (Requerente)
    InsereParagrafoWord
    'AlinhaEndereco
    ColaTextoWord (EndRequerente)
End If

Do While (CONTAR < 4)
    deletar
    CONTAR = CONTAR + 1
    'MsgBox ("Delete número=" & contar)
    Loop
End Sub

Sub AlinhaEndereco()
Dim objWord As Word.Application
Dim oDoc As Word.Document
    Set objWord = GetObject(, "Word.Application")
    objWord.Selection.TypeText Text:=vbTab & vbTab & vbTab & vbTab
End Sub

Sub InserePC()
'
Dim texto As String
Sheets("Cadastro_Dados").Select
SelecionaIndicador ("PC")
If Range("C21") <> "" Then
    texto = "Procedimento de calibração " & Range("C21") & " - Revisão " & Range("D21") & " "
    ColaTextoWord (texto)
End If
If Range("C22") <> "" Then
    texto = "Procedimento de calibração " & Range("C22") & " - Revisão " & Range("D22") & " "
    InsereParagrafoWord
    ColaTextoWord (texto)
End If
If Range("C23") <> "" Then
    texto = "Procedimento de calibração " & Range("C23") & " - Revisão " & Range("D23") & " "
    InsereParagrafoWord
    ColaTextoWord (texto)
End If
If Range("C24") <> "" Then
    texto = "Procedimento de calibração " & Range("C24") & " - Revisão " & Range("D24") & " "
    InsereParagrafoWord
    ColaTextoWord (texto)
End If
If Range("C25") <> "" Then
    texto = "Procedimento de calibração " & Range("C25") & " - Revisão " & Range("D25") & " "
    InsereParagrafoWord
    ColaTextoWord (texto)
End If
If Range("C26") <> "" Then
    texto = "Procedimento de calibração " & Range("C26") & " - Revisão " & Range("D26") & " "
    InsereParagrafoWord
    ColaTextoWord (texto)
End If
deleta
End Sub

Sub InsereMetodo()
'
Dim texto As String
Sheets("Cadastro_Dados").Select
SelecionaIndicador ("Métodos")
If Range("C39") <> "" Then
    texto = Range("C39")
    ColaTextoWord (texto) & "  "
End If
If Range("C40") <> "" Then
    InsereParagrafoWord
    texto = Range("C40")
    ColaTextoWord (texto) & "  "
End If
If Range("C41") <> "" Then
    InsereParagrafoWord
    texto = Range("C41")
    ColaTextoWord (texto) & "  "
End If
If Range("C42") <> "" Then
    InsereParagrafoWord
    texto = Range("C42")
    ColaTextoWord (texto) & "  "
End If
If Range("C43") <> "" Then
    InsereParagrafoWord
    texto = Range("C43")
    ColaTextoWord (texto) & "  "
End If
If Range("C44") <> "" Then
    InsereParagrafoWord
    texto = Range("C44")
    ColaTextoWord (texto) & "  "
End If
If Range("C45") <> "" Then
    InsereParagrafoWord
    texto = Range("C45")
    ColaTextoWord (texto) & "  "
End If
If Range("C46") <> "" Then
    InsereParagrafoWord
    texto = Range("C46")
    ColaTextoWord (texto) & "  "
End If
deleta
End Sub

Sub InserePadroes()
'
Dim texto As String
Dim CONTAR As Integer
'MsgBox ("Inserir padroes")
CONTAR = 1
Sheets("Dados").Select
SelecionaIndicador ("TAGP1")
texto = Range("Q42").Text

Do While (texto <> "")
    'MsgBox ("Texto de teste: " & texto)
    'MsgBox ("texto da planilha: " & Worksheets("Dados").Cells(41 + contar, 16).text)
    If texto = "" Then
    Else
        If Right(Worksheets("Dados").Cells(41 + CONTAR, 17).Text, 7) = "VENCIDO" Then
        'MsgBox ("Planilha vencido=" & texto)
        texto = Worksheets("Dados").Cells(41 + CONTAR, 17).Text
        vencido (texto)
        InsereParagrafoWord
        Else
        texto = Worksheets("Dados").Cells(41 + CONTAR, 17).Text
        'MsgBox ("Planilha Não vencido=" & texto)
        'MsgBox (texto)
        NaoVencido (texto)
        InsereParagrafoWord
    End If
    CONTAR = CONTAR + 1
    texto = Worksheets("Dados").Cells(41 + CONTAR, 17).Text
    End If
    Loop
deleta
End Sub

Sub PeriodoCal()
'
Dim data_inicial As Variant
Dim data_final As Variant
    Sheets("Cadastro_Dados").Select
    Range("D14:D14").Select
    'Selection.NumberFormat = "dd/mm/yyyy"
    Range("D15:D15").Select
    'Selection.NumberFormat = "dd/mm/yyyy"
    data_inicial = Range("Cadastro_Dados!D14:D14").Value
    data_final = Range("Cadastro_Dados!D15:D15").Value
    
    If data_inicial = data_final Then
        Sheets("Dados").Select
        Range("Dados!BH1:BH1") = "Data da calibração: "
        Range("Dados!BK1:BK1").Value = Range("Dados!BH1:BH1") & data_inicial
    End If
    If data_inicial <> data_final Then
        Sheets("Cadastro_Dados").Select
        Range("Dados!BH1:BH1") = "Período de calibração: "
        Range("Dados!BK1:BK1").Value = Range("Dados!BH1:BH1") & data_inicial & " a " & data_final
    End If
    
    Sheets("Dados").Select
    Range("BH1").Select
    Selection.Font.Bold = False
    Selection.Font.Bold = True
    Selection.Font.Italic = False
    Sheets("Dados").Select
    Range("BM1").Select
    Selection.Font.Bold = False
    Selection.Font.Bold = True
    Selection.Font.Italic = False
    Sheets("Cadastro_Dados").Select
End Sub

Sub vencido(texto As String)
'
Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim oRange As Word.Range
Set oWord = GetObject(, "Word.Application")
    oWord.Selection.Font.Color = wdColorRed
    oWord.Selection.Font.Bold = True
    ActiveSheet.Select
    ColaTextoWord (texto)
    oWord.Visible = True
End Sub

Sub NaoVencido(texto As String)
'
Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim oRange As Word.Range
Set oWord = GetObject(, "Word.Application")
    'oWord.Selection.TypeParagraph
    oWord.Selection.Font.Bold = False
    oWord.Selection.Font.Color = wdColorAutomatic
    'MsgBox ("Trocando de cor")
    ActiveSheet.Select
    ColaTextoWord (texto)
    oWord.Selection.Font.Bold = False
    oWord.Selection.Font.Color = wdColorBlack
    oWord.Visible = True
End Sub

Sub Ultimapagina()
'
Dim CONTAR As Integer
'MsgBox ("Inserir padroes")
CONTAR = 1
Sheets("Dados").Select
SelecionaIndicador ("Convencao")
Do While (CONTAR < 46)
    deleta
    CONTAR = CONTAR + 1
    'MsgBox ("Delete número=" & contar)
    Loop
    
    SelecionaIndicador ("Convencao")
    corrigeultimapagina
    SelecionaIndicador ("Paginafinal")
    InserePaginaNova
End Sub

Sub corrigeultimapagina()
'
Dim objWord As Word.Application
Dim oDoc As Word.Document
    Set objWord = GetObject(, "Word.Application")
    objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    objWord.Selection.TypeText Text:=vbTab & vbTab
End Sub

Sub EliminaPontosVazios(pontos As Integer)
'
' EliminaPontosVazios Macro

Dim i As Integer
Dim objWord As Word.Application
Dim oDoc As Word.Document
    Set objWord = GetObject(, "Word.Application")

Do While (i < pontos)
    With objWord.Selection.Find
        .Text = ""
        '.Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    objWord.Selection.Find.Execute
    If objWord.Selection.Find.Found = False Then
    Else
        objWord.Selection.Rows.Delete
    End If
        i = i + 1
    Loop
    
End Sub

Sub SubstituiOhm()
'
Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim oRange As Word.Range
Set oWord = GetObject(, "Word.Application")
    oWord.Selection.Find.ClearFormatting
    oWord.Selection.Find.Replacement.ClearFormatting
    With oWord.Selection.Find
        .Text = "Ohm"
        .Replacement.Text = ChrW(937)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    oWord.Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub SubstituiInfinito()
'
Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim oRange As Word.Range
Set oWord = GetObject(, "Word.Application")
    oWord.Selection.Find.ClearFormatting
    oWord.Selection.Find.Replacement.ClearFormatting
    With oWord.Selection.Find
        .Text = "Infinito"
        .Replacement.Text = ChrW(8734)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    oWord.Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub Montacertificado()
    Cert.TextBox1.Value = Range("D16").Text
    Cert.TextBox2.Value = Range("D17").Text
    Cert.Show
End Sub

Sub Montacert()
'
Dim texto As String
Dim texto2 As String
Dim metodo1 As String
Dim metodo2 As String

Dim contador As Integer
Dim NomePlan As String
Dim servico As String
Dim data_inicial As Variant
Dim data_final As Variant
Dim data As Variant
Dim aux As Integer
Dim padroesvencidos As String
Dim padroes As String
Dim CONTAR As Integer
Dim vencido As Integer
CONTAR = 1
vencido = 0
Sheets("Dados").Select
padroes = Range("Q42").Text
Do While (padroes <> "")
    If padroes = "" Then
    Else
        If Right(Worksheets("Dados").Cells(41 + CONTAR, 17).Text, 7) = "VENCIDO" Then
            padroesvencidos = Left(Worksheets("Dados").Cells(41 + CONTAR, 17).Text, 6) & " ; " & padroesvencidos
            vencido = vencido + 1
            padroes = Worksheets("Dados").Cells(41 + CONTAR, 17).Text
        Else
            padroes = Worksheets("Dados").Cells(41 + CONTAR, 17).Text
        End If
    End If
    CONTAR = CONTAR + 1
Loop

If vencido = 0 Then

Else
    MsgBox ("Existem " & vencido & " padrões vencidos")
    MsgBox ("Padrões vencidos: " & padroesvencidos)
    Sheets("Cadastro_Dados").Select
    Exit Sub
End If


Application.ScreenUpdating = False
    AbreWord
    Sheets("Cadastro_Dados").Select
    If Certificado = 1 Then
    Else
        CloseDoc
        Exit Sub
    End If

       
    Range("E14").Select
    If Range("E14") = "" Then
    texto2 = Range("E10").Text & " - " & Range("E11").Text & " - " & Range("E12").Text & " - " & Range("E13").Text
    End If
    If Range("E14") <> "" Then
    texto2 = Range("E10").Text & " - " & Range("E11").Text & " - " & Range("E12").Text & " - " & Range("E13").Text & " - " & Range("E14").Text
    End If
    'MsgBox (texto2)
    
    ActiveSheet.Next.Select
    'Range("H2").Select
    ActiveSheet.Next.Select
    texto = CertNum
    'MsgBox ("Número do certificado" & CertNum)
       
    SelecionaIndicador ("Número")
    ColaTextoWord (texto)
    SelecionaIndicador ("Data")
    'MsgBox (dataEmissao)
    texto = dataEmissao
    ColaTextoWord (texto)
    
    Sheets("Dados").Select
        
    
    SelecionaIndicador ("Protocolo")
    texto = Range("B6:B6").Text
    ColaTextoWord (texto)
    
    SelecionaIndicador ("Equipamento")
    texto = Range("D6:D6").Text
    ColaTextoWord (texto)
    
    SelecionaIndicador ("Fabricante")
    texto = Range("E6:E6").Text
    ColaTextoWord (texto)
    
    SelecionaIndicador ("Modelo")
    texto = Range("F6:F6").Text
    ColaTextoWord (texto)
    
    SelecionaIndicador ("Série")
    texto = Range("G6:G6").Text
    ColaTextoWord (texto)
    
    If Range("V6:V6").Text <> "" Then
    SelecionaIndicador ("TAG")
    texto = Range("V6:V6").Text
    ColaTextoWord (texto)
    End If
    
    '**************************************
    Sheets("Cadastro_Dados").Select
    NomePlan = ActiveSheet.Name
    SelecionaIndicador ("Planilhas")
    Do While (NomePlan <> "Dados")
       ActiveSheet.Next.Select
       NomePlan = ActiveSheet.Name
       servico = Range("H2").Text
    If NomePlan = "Dados" Then
       
    Else
        If Range("F2") = "" Then
            MsgBox ("Preencha a data da calibração da planilha " & NomePlan)
            Exit Sub
        End If
        If Range("F1") = "" Then
            MsgBox ("Preencha o Executor da planilha " & NomePlan)
            Exit Sub
        End If
        
        InsereParagrafoWord
        InsereParagrafoWord
        ColaTitulo (Range("M38").Text)
        InsereParagrafoWord
        InsereParagrafoWord
        InsereParagrafoWord
        '*****Colagem dos dados**********
        Range("C34:H47").Select
        Application.CutCopyMode = False
        Selection.Copy
        PasteToWord
        pontos = pontos + 12
                
       ' InserePaginaNova
    End If
    
    Loop
    '***************************************
'CxMsg ("data inicial" & data_inicial & "e data final" & data_final)

PeriodoCal
Sheets("Dados").Select
texto = Range("BK1").Text
SelecionaIndicador ("PeriodoCalibracao")
ColaTextoWord (texto)
InsereMetodo
deleta
InserePadroes
InserePC
'Ultimapagina

InsereCliente
EliminaPontosVazios (pontos)
SubstituiOhm
SubstituiInfinito
CloseWordDoc
Application.ScreenUpdating = True

Sheets("Cadastro_Dados").Select
Range("A1").Select
End Sub



