Set WshNetwork = CreateObject("WScript.Network")

PrinterPath = "\\forum-print\GENICOM COLOUR cl160"

WshNetwork.AddWindowsPrinterConnection PrinterPath

PrinterPath = "\\forum-print\HP 8500 Colour - PS"

WshNetwork.AddWindowsPrinterConnection PrinterPath