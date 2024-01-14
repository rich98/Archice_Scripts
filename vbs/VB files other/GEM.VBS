Set WshNetwork = CreateObject("WScript.Network")

PrinterPath = "\\forum-print\GENICOM COLOUR cl160"

WshNetwork.AddWindowsPrinterConnection PrinterPath
