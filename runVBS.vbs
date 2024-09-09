'--- ADD
set WshShell = WScript.CreateObject("WScript.Shell")

Dim excelApp,macro

file = WScript.Arguments(0)
macro = WScript.Arguments(1)

Set excelApp = CreateObject("Excel.Application")

excelApp.Visible = False        'Excelを非表示にする
excelApp.DisplayAlerts = False  'ポップアップメッセージを非表示にする
excelApp.AutomationSecurity = 1 'マクロを有効にする

'Excelファイルを読み取り専用で開く
excelApp.Workbooks.Open file,3,False

'WScript.Echo "---マクロを実行します---"

'マクロを実行する
excelApp.Run macro

'WScript.Echo "---マクロの実行が完了しました---"

'--- ADD Start
excelApp.Visible = True
excelApp.DisplayAlerts = True

'WScript.Sleep( 3000 )

'WshShell.SendKeys "%{TAB}"
'WshShell.SendKeys "{ENTER}"
'--- ADD End

'Excelを終了する
excelApp.Quit

Set excelApp = Nothing