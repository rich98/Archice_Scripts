Set WshNetwork = CreateObject("WScript.Network")


PrinterPath = "\\forum-print\finance_hp4100"


WshNetwork.AddWindowsPrinterConnection PrinterPath


WshNetwork.SetDefaultPrinter "\\forum-print\finance_hp4100"