' -------------------------------
' CASTRO1 - IMPORTAR TODOS OS MÓDULOS DE UMA PASTA
' -------------------------------
Sub AAIMPORTMODULE()
    On Error GoTo ErrorHandler
    
    Dim vbProj As Object
    Dim pasta As String, arquivo As String
    Dim fso As Object
    Dim dialogo As FileDialog
    Dim resultadoDialogo As Integer
    Dim contador As Integer
    
    ' Criar objetos
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set vbProj = ThisWorkbook.VBProject
    contador = 0
    
    Set dialogo = Application.FileDialog(msoFileDialogFolderPicker)
    
    With dialogo
        .Title = "Selecionar Pasta com Módulos VBA"
        .InitialFileName = "C:\"  
        .AllowMultiSelect = False 
        
        resultadoDialogo = .Show
        
        If resultadoDialogo <> -1 Then 
            MsgBox "Nenhuma pasta foi selecionada.", vbInformation
            Exit Sub
        End If
        
        pasta = .SelectedItems(1) & "\"
    End With
    
    If Not fso.FolderExists(pasta) Then
        MsgBox "Pasta não encontrada: " & pasta, vbExclamation
        Exit Sub
    End If
    
    arquivo = Dir(pasta & "*.bas")
    If arquivo = "" Then
        MsgBox "Nenhum arquivo .bas encontrado na pasta selecionada.", vbExclamation
        Exit Sub
    End If
    
    Do While arquivo <> ""
        On Error Resume Next
        vbProj.VBComponents.Import pasta & arquivo
        If Err.Number = 0 Then
            Debug.Print "? Importado: " & arquivo
            contador = contador + 1
        Else
            Debug.Print "? Erro ao importar " & arquivo & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo ErrorHandler
        arquivo = Dir
    Loop
    
    arquivo = Dir(pasta & "*.cls")
    Do While arquivo <> ""
        On Error Resume Next
        vbProj.VBComponents.Import pasta & arquivo
        If Err.Number = 0 Then
            Debug.Print "? Importado: " & arquivo
            contador = contador + 1
        Else
            Debug.Print "? Erro ao importar " & arquivo & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo ErrorHandler
        arquivo = Dir
    Loop
    
    If contador > 0 Then
        MsgBox contador & " módulos importados com sucesso!" & vbCrLf & _
               "Pasta: " & pasta, vbInformation
    Else
        MsgBox "Nenhum módulo foi importado.", vbExclamation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erro ao importar módulos: " & Err.Description & vbCrLf & _
           "Linha: " & Erl, vbCritical
End Sub

' -------------------------------
' CASTRO2 - ADICIONAR REFERÊNCIAS AUTOMATICAMENTE
' -------------------------------
Sub AAIMPORTLIBRARIES()
    On Error GoTo ErrorHandler
    
    Dim vbProj As Object
    Dim refExtensibilidade As Object
    Dim refWord As Object
    Dim referencias As Object
    Dim referencia As Object
    Dim refAdicionadas As Integer
    
    Set vbProj = ThisWorkbook.VBProject
    Set referencias = vbProj.References
    refAdicionadas = 0
    
    If Not ReferenciaExiste("{0002E157-0000-0000-C000-000000000046}") Then
        Set refExtensibilidade = referencias.AddFromGuid("{0002E157-0000-0000-C000-000000000046}", 5, 3)
        refAdicionadas = refAdicionadas + 1
        Debug.Print "Referência Extensibilidade 5.3 adicionada"
    Else
        Debug.Print "Referência Extensibilidade 5.3 já existe"
    End If
    
    If Not ReferenciaExiste("{00020905-0000-0000-C000-000000000046}") Then
        Set refWord = referencias.AddFromFile(EncontrarBibliotecaWord())
        refAdicionadas = refAdicionadas + 1
        Debug.Print "Referência Word adicionada"
    Else
        Debug.Print "Referência Word já existe"
    End If
    
    If refAdicionadas > 0 Then
        MsgBox refAdicionadas & " referências foram adicionadas com sucesso!", vbInformation
    Else
        MsgBox "Todas as referências já estão adicionadas.", vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erro ao adicionar referências: " & Err.Description & vbCrLf & _
           "Certifique-se de que o acesso ao modelo VBA está habilitado.", vbCritical
End Sub

' -------------------------------
' VERIFICAR SE REFERÊNCIA JÁ EXISTE
' -------------------------------
Function ReferenciaExiste(guid As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ref As Object
    Dim vbProj As Object
    
    Set vbProj = ThisWorkbook.VBProject
    
    For Each ref In vbProj.References
        If ref.guid = guid Then
            ReferenciaExiste = True
            Exit Function
        End If
    Next ref
    
    ReferenciaExiste = False
    Exit Function
    
ErrorHandler:
    ReferenciaExiste = False
End Function

' -------------------------------
' ENCONTRAR CAMINHO DA BIBLIOTECA WORD
' -------------------------------
Function EncontrarBibliotecaWord() As String
    Dim caminhos() As Variant
    Dim i As Integer
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    caminhos = Array("C:\Program Files\Microsoft Office\root\Office16\MSWORD.OLB", "C:\Program Files (x86)\Microsoft Office\root\Office16\MSWORD.OLB", "C:\Program Files\Microsoft Office\Office15\MSWORD.OLB", "C:\Program Files (x86)\Microsoft Office\Office15\MSWORD.OLB", "C:\Program Files\Microsoft Office\Office14\MSWORD.OLB", "C:\Program Files (x86)\Microsoft Office\Office14\MSWORD.OLB")
    
    For i = LBound(caminhos) To UBound(caminhos)
        If fso.FileExists(caminhos(i)) Then
            EncontrarBibliotecaWord = caminhos(i)
            Exit Function
        End If
    Next i
    
    EncontrarBibliotecaWord = ObterCaminhoWordDoRegistro()
End Function

' -------------------------------
' OBTER CAMINHO DO WORD VIA REGISTRO
' -------------------------------
Function ObterCaminhoWordDoRegistro() As String
    On Error GoTo ErrorHandler
    
    Dim shell As Object
    Dim regPath As String
    Dim wordPath As String
    
    Set shell = CreateObject("WScript.Shell")
    
    Dim versoes() As Variant
    versoes = Array("16.0", "15.0", "14.0", "12.0")
    
    For Each ver In versoes
        On Error Resume Next
        regPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\" & ver & "\Word\InstallRoot\"
        wordPath = shell.RegRead(regPath & "Path") & "MSWORD.OLB"
        
        If Err.Number = 0 Then
            If CreateObject("Scripting.FileSystemObject").FileExists(wordPath) Then
                ObterCaminhoWordDoRegistro = wordPath
                Exit Function
            End If
        End If
        On Error GoTo 0
    Next
    
ErrorHandler:
    ObterCaminhoWordDoRegistro = "C:\Program Files\Microsoft Office\root\Office16\MSWORD.OLB"
End Function