Set WshNetwork = CreateObject("WScript.Network")


PrinterPath = "\\forum-print\CreditCtrl2"


WshNetwork.AddWindowsPrinterConnection PrinterPath


WshNetwork.SetDefaultPrinter "\\forum-print\CreditCtrl2"