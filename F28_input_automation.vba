' ==============================================================================
' Project:      SAP Financial Partial Payment Posting Automation
' Author:       Krisztián Princzinger
' Description:  Automates the allocation of partial payments in SAP (F-28).
'               Handles pagination/scrolling on SAP table controls automatically.
' ==============================================================================

Option Explicit

Sub InputInvoicesToSelectionScreen()

	'Set up SAP connection
    Dim SapGuiAuto As Object
    Dim App As Object
    Dim Connection As Object
    Dim Session As Object

	On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    If SapGuiAuto Is Nothing Then
        MsgBox "SAP GUI is not open!", vbCritical
        Exit Sub
    End If
    
    Set App = SapGuiAuto.GetScriptingEngine
    Set Connection = App.Children(0)
    Set Session = Connection.Children(0)
    On Error GoTo 0

	'Variables for the logic
	Dim itemCount As Integer 'The record limit for a document is around 990 in SAP, so Integer is enough.
	Const MAX_SAP_ROWS As Integer = 990 'Customize this as needed if the limit is different
	Dim inputLastRow As Long 'The max row count in Excel is 1 048 576, so it's advised to use Long here to avoid overflow errors.
	Dim i As Integer
	Dim uiRow As Integer
	Dim scroll As Integer
	Dim ws As Worksheet
	Dim invoiceArray() As Variant

	Set ws = ThisWorkbook.Sheets(1) 'Change this if the source is on another sheet
	
	'Count the amount of invoices from D4 cell downwards
	inputLastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
	If inputLastRow < 4 Then
        MsgBox "No data found in Column D!", vbExclamation
        Exit Sub
    End If
	
	' Calculate item count (Row 4 is the first data row)
    itemCount = inputLastRow - 3
	
	'990 item limit check
    If itemCount > MAX_SAP_ROWS Then
        MsgBox "You have " & itemCount & " items!" & vbCrLf & _
               "The limit in SAP is " & MAX_SAP_ROWS & " per posting in this transaction.", vbCritical
        Exit Sub
    End If

	ReDim invoiceArray(0 To itemCount - 1)

	'From D4 downwards put the invoice numbers/references/other parameters into the array
	For i = 0 To itemCount - 1
		invoiceArray(i) = ws.Cells(i + 4, 4).Value 
	Next i
	
	'Start filling the UI rows from the array
	'SAP Limit: The selection screen typically fits 27 lines before pressing submit and getting new, empty input fields (Enter in this case). Customize this as needed.
	uiRow = 0
	For i = LBound(invoiceArray) To UBound(invoiceArray)
		If (i Mod 27 = 0) And i > 0 Then
			Session.FindById("wnd[0]").SendVKey 0
			uiRow = 0
			If Session.FindById("wnd[0]/sbar").MessageType = "W" Then
				MsgBox "There is an invalid invoice in the list!"
				Exit Sub
			End If
		End If
		Session.FindById("wnd[0]/usr/sub:SAPMF05A:0731/txtRF05A-SEL01[" & uiRow & ",0]").Text = invoiceArray(i)
		uiRow = uiRow + 1
	Next i

	Session.FindById("wnd[0]").SendVKey 0

	Session.FindById("wnd[0]/tbar[1]/btn[16]").Press 'Payment processing
	Session.FindById("wnd[0]/usr/tabsTS/tabpPART").Select 'Partial payment

	'At this point, the current content of invoiceArray is not needed, so it can be deleted/overwritten.
	'In order to optimize memory usage, I decided to overwrite the memory data.
	'From E4 downwards put the payment amounts into the array
	For i = 0 To itemCount - 1
		invoiceArray(i) = ws.Cells(i + 4, 5).Value
	Next i

	'Start filling the UI rows from the array
	'SAP Limit: The payment screen typically fits 21 lines before needing to page down. Customize this as needed.
	'Important note: Page down is not possible here and you have to use the VerticalScrollbar.Position property to move down on the list, as used in this code. 
	uiRow = 0
	scroll = 0
	For i = LBound(invoiceArray) To UBound(invoiceArray)
		If (i Mod 21 = 0) And i > 0 Then
			Session.FindById("wnd[0]/usr/tabsTS/tabpPART/ssubPAGE:SAPDF05X:6104/tblSAPDF05XTC_6104").VerticalScrollbar.Position = scroll * 21 + 21
			scroll = scroll + 1
			uiRow = 0
		End If
		Session.FindById("wnd[0]/usr/tabsTS/tabpPART/ssubPAGE:SAPDF05X:6104/tblSAPDF05XTC_6104/txtDF05B-PSZAH[7," & uiRow & "]").Text = invoiceArray(i)
		uiRow = uiRow + 1
	Next i
	 
End Sub

Sub ClearInputCells()
	Dim lastRowD As Long
	Dim lastRowE As Long
	Dim lastRow As Long
	Dim i As Integer
	Dim ws As Worksheet
	Set ws = ThisWorkbook.Sheets(1) 'Change this if the source is on another sheet
	
	lastRowD = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
	lastRowE = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
	lastRow = Application.WorksheetFunction.Max(lastRowD, lastRowE) 'Using Excel Max function to compare values
	
	For i = 4 To lastRow
		ws.Range("D" & i & ":E" & i).ClearContents
	Next i
End Sub
