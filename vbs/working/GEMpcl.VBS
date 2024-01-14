Set WshNetwork = CreateObject("WScript.Network")

PrinterPath = "\\forum-print\genicom2"

WshNetwork.AddWindowsPrinterConnection PrinterPath
