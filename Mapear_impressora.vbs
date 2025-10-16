' =====================================================================
' NOME: InstalaImpressoras.vbs
' DESCRIÇÃO: Script para instalar impressoras de rede de forma robusta,
'            verificando se já existem e tratando erros.
' AUTOR: Gemini
' VERSÃO: 2.0
' =====================================================================

Option Explicit ' Força a declaração de todas as variáveis.

' --- CONFIGURAÇÃO ---
' Adicione ou remova os caminhos de rede (UNC) das impressoras na lista abaixo.
Dim arrPrinters
arrPrinters = Array("\\printserver\Tecnico", "\\printserver\controladoria", "\\printserver\Marketing")

' Defina se deseja que a última impressora da lista se torne a padrão.
' Mude para True se desejar ativar esta funcionalidade.
Const SET_LAST_AS_DEFAULT = False
' --------------------


' Declaração de variáveis
Dim objNetwork, strPrinter, oPrinters, i, bFound, intInstalledCount
Set objNetwork = CreateObject("WScript.Network")
intInstalledCount = 0

' Inicia o processo com uma mensagem
MsgBox "O processo de instalação de impressoras será iniciado.", vbInformation, "Instalador de Impressoras"

' Loop principal para percorrer a lista de impressoras
For Each strPrinter In arrPrinters
    On Error Resume Next ' Ativa o tratamento de erros para a verificação

    bFound = False ' Reseta a variável de verificação para cada impressora
    
    ' Enumera as impressoras já instaladas para evitar duplicidade
    Set oPrinters = objNetwork.EnumPrinterConnections
    For i = 0 To oPrinters.Count - 1 Step 2
        If LCase(oPrinters.Item(i+1)) = LCase(strPrinter) Then
            bFound = True
            Exit For
        End If
    Next

    ' Limpa qualquer erro que possa ter ocorrido na enumeração
    If Err.Number <> 0 Then Err.Clear
    
    ' Se a impressora não foi encontrada, tenta instalá-la
    If Not bFound Then
        objNetwork.AddWindowsPrinterConnection strPrinter
        
        ' Verifica se a instalação foi bem-sucedida
        If Err.Number = 0 Then
            MsgBox "Impressora instalada com sucesso:" & vbCrLf & strPrinter, vbInformation, "Sucesso"
            intInstalledCount = intInstalledCount + 1
        Else
            ' Se houve um erro, informa o usuário com detalhes
            MsgBox "Falha ao instalar a impressora: " & strPrinter & vbCrLf & vbCrLf & _
                   "Motivo: " & Err.Description & " (Código: " & Err.Number & ")", _
                   vbCritical, "Erro de Instalação"
            Err.Clear ' Limpa o erro para não afetar o próximo loop
        End If
    Else
        ' Informa que a impressora já estava instalada
        MsgBox "A impressora já está instalada:" & vbCrLf & strPrinter, vbInformation, "Informação"
    End If
Next

' --- (OPCIONAL) DEFINIR IMPRESSORA PADRÃO ---
' Se a constante SET_LAST_AS_DEFAULT for True e pelo menos uma impressora foi instalada,
' define a última impressora da lista como padrão.
If SET_LAST_AS_DEFAULT AND UBound(arrPrinters) >= 0 Then
    On Error Resume Next ' Ativa o tratamento de erros para esta operação
    Dim lastPrinter
    lastPrinter = arrPrinters(UBound(arrPrinters))
    
    objNetwork.SetDefaultPrinter lastPrinter
    If Err.Number <> 0 Then
        MsgBox "Não foi possível definir " & lastPrinter & " como padrão. Motivo: " & Err.Description, vbExclamation, "Aviso"
        Err.Clear
    End If
End If
' --------------------------------------------

' Exibe um relatório final
MsgBox "Processo finalizado." & vbCrLf & vbCrLf & _
       intInstalledCount & " nova(s) impressora(s) foram instalada(s).", _
       vbInformation, "Fim da Execução"

' Limpa o objeto da memória
Set objNetwork = Nothing
