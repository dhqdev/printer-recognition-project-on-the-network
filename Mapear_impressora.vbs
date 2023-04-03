On Error Resume Next

Set objNetwork = CreateObject("WScript.Network")

'Adicione aqui os nomes das impressoras que você deseja instalar
arrPrinters = Array("\\printserver\Tecnico", "\\printserver\controladoria")

For Each strPrinter in arrPrinters
    'Instala a impressora
    On Error Resume Next
    objNetwork.AddWindowsPrinterConnection strPrinter
    If Err.Number = 0 Then
        MsgBox "A impressora " & strPrinter & " foi instalada com sucesso."
    End If
Next

Set objNetwork = Nothing