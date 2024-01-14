Set WshNetwork = CreateObject("WScript.Network")

On Error Resume Next

PrinterPath = "\\forum-print\HPcolor550"

WshNetwork.AddWindowsPrinterConnection PrinterPath
