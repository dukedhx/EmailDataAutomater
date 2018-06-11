

Set xl = CreateObject("Excel.Application")
Set xlBk = xl.Workbooks.Open (Wscript.Arguments(0))

Wscript.Echo xlbk.Worksheets(2).Shapes(Wscript.Arguments(1)).ControlFormat.Value 



xl.DisplayAlerts = False

xlBk.Close(True)
xl.Quit

