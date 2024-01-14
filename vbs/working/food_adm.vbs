Set WshNetwork = CreateObject("WScript.Network")


PrinterPath = "\\forum-print\food_admin"


WshNetwork.AddWindowsPrinterConnection PrinterPath


WshNetwork.SetDefaultPrinter "\\forum-print\food_admin"