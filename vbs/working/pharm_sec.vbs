Set WshNetwork = CreateObject("WScript.Network")


PrinterPath = "\\forum-print\pharm_sec"


WshNetwork.AddWindowsPrinterConnection PrinterPath


WshNetwork.SetDefaultPrinter "\\forum-print\pharm_sec"