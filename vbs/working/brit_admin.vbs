Set WshNetwork = CreateObject("WScript.Network")


PrinterPath = "\\forum-print\brit_admin"


WshNetwork.AddWindowsPrinterConnection PrinterPath


WshNetwork.SetDefaultPrinter "\\forum-print\brit_admin"