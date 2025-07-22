REM  *****  BASIC  *****
Option VBASupport 1
Option Explicit
Sub Main

'	On Error GoTo EH

    Dim Calculadora as Worksheet
    Set Calculadora = thisWorkbook.Sheets("Calculadora")
    
    Dim totalDate as Date, totalPayment as Double
    totalDate = Calculadora.Cells(3,4)
    totalPayment = Calculadora.Cells(3,5)
    
    If totalDate = 0 Or totalPayment = 0 Then
		msgbox "Error: Valores a pagar incompletos."
		Stop
	End If
    
    'Counts the amount of rows from row 3
    Dim chequesLastRow as Long
	If Calculadora.Cells(Rows.Count, 1).End(xlUp).Row <= Calculadora.Cells(Rows.Count, 2).End(xlUp).Row Then
		chequesLastRow = Calculadora.Cells(Rows.Count, 2).End(xlUp).Row
	Else
		chequesLastRow = Calculadora.Cells(Rows.Count, 1).End(xlUp).Row
	End If


    'Loops through the cheques table
    Dim xlrow as Long, xlcol as Long
    Dim proportion as Double, daysDifference as Long, daysOffset as Double


    'Checks the total length of the table against the first row location
    If chequesLastRow < 3 Then
    	msgbox "Error: Tabla de cheques incompleta."
		Stop
    End If
    
    For xlrow = 3 to chequesLastRow

		'Checks for empty cells on the cheques table    
		If Calculadora.Cells(xlrow, 1) = "" Or Calculadora.Cells(xlrow, 2) = "" Then
			msgbox "Error: Tabla de cheques incompleta."
			Stop
		End If
    
		proportion = Calculadora.Cells(xlrow, 2) / totalPayment
		daysDifference = Calculadora.Cells(xlrow, 1) - totalDate
		daysOffset = (proportion * daysDifference) + daysOffset
	Next xlrow


	'Adds all the cheque amounts entered
	Dim partialPayment as Double
	For xlrow = 3 to chequesLastRow
		partialPayment = Calculadora.Cells(xlrow, 2) + partialpayment
	Next xlrow


	'Logic that checks whether we enter an incomplete payment or a complete one to check.
	If partialPayment < totalPayment Then
		Dim toPay as Double
		toPay = totalPayment - partialPayment
		proportion = toPay / totalPayment
				
		
	ElseIf partialPayment > totalpayment Then
		MsgBox "El monto total ingresado en cheques es mayor al importe a pagar."
		End
		
	Else
		If daysOffset < 0 Then
			msgbox "Estás pagando la factura " & (daysOffset * (-1)) & " día(s) adelantada."
			End

		ElseIf daysOffset > 0 Then
			msgbox "Estás pagando la factura " & daysOffset & " día(s) atrasada."
			End
			
		Else
			msgbox "Estás pagando la factura exactamente al día"
			End
			
		End If
		
	End If
	
	
	'After logic, if a date for a new cheque needs to be calculated.
	Dim daysToPay as Long, dateNewCheque as Date
	daysToPay = daysOffset / proportion
	dateNewCheque = totalDate - daysToPay
	
	MsgBox "El monto restante para cubrir el pago es $" & Format(toPay,"Standard") & " a fecha " & dateNewCheque

Done:
	Exit Sub
		
EH:
	msgbox "Error: " & Err.Description			

End Sub