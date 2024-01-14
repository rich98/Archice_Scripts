Set WshNetwork = CreateObject("WScript.Network")


PrinterPath = "\\forum-print\food_div1"


WshNetwork.AddWindowsPrinterConnection PrinterPath


WshNetwork.SetDefaultPrinter "\\forum-print\food_div1"