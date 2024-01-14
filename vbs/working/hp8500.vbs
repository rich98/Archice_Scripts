Set WshNetwork = CreateObject("WScript.Network")


PrinterPath = "\\forum-print\HP 8500 colour - PS"


WshNetwork.AddWindowsPrinterConnection PrinterPath

