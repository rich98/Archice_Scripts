Set WshNetwork = CreateObject("WScript.Network")


PrinterPath = "\\forum-print\pharm_reg"


WshNetwork.AddWindowsPrinterConnection PrinterPath


WshNetwork.SetDefaultPrinter "\\forum-print\pharm_reg"