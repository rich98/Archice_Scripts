Set WshNetwork = CreateObject("WScript.Network")

PrinterPath = "\\forum-print\HPcolor5500"

WshNetwork.AddWindowsPrinterConnection PrinterPath
