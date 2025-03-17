'********************************************************************
'*
'* Copyright Kecskeméti Lajos
'*
'* Module Name:    sql2xls_new.vbs 
'*
'* SQL eredmény XLS-be töltése
'*
'*
'********************************************************************

' alap beállítások
Option Explicit

ON ERROR RESUME NEXT
Err.Clear
'On Error GoTo ErrorHandler

' valtozók
Dim objRootDSE, strDNSDomain, adoConnection
Dim strBase, strFilter, strAttributes, strQuery, adoRecordset
Dim strName, strDN, objManagerList, strManagerDN
Dim objExcel, objExcel_read, objWorkbook, objWorkbook_read,sorn, oszlopn, sork, oszlopk, eField, objRange, objRange2
Dim strExcelPath, strExcelPath_read, konyvtar
Dim strCon, strsql, i, ii, eltolas, eltolas_tomb, eltsz, sornn, olvass_el
Dim eredmeny_tomb(2) 
Dim ki_xls_neve(50),	munkalapnev(50),	sql(50),	kezd_sor(50), kezd_oszlop(50), fejlec(50),	szamolo(50),	r_nev(50), eltolasok(50), kapcsolatok(50)
Dim kezdx, kezdy, munkalapsz, eredmeny

Const xlAscending = 1
Const xlDescending = 2
Const xlYes = 1

sornn = 1

'--****************************************** Vált **********************************
konyvtar = Left(WScript.ScriptFullName,(Len(WScript.ScriptFullName) - (Len(WScript.ScriptName) + 1)))
' eredmeny xls megadása csak nevvel (alapértelmezett könyvtárba dolgozik  "d:\Kecskemet1L314\My Documents\sql_xls.xlsx"  )
strExcelPath      = konyvtar & "\master.xlsx"
strExcelPath_read = konyvtar & "\xls_main.xlsx"
'--********************************************************************************

WSCript.Echo "---------- xls nevek  -----------"
WSCript.Echo konyvtar
WSCript.Echo strExcelPath
WSCript.Echo strExcelPath_read

'-------------------------------- Excel olvasás kezdet------------------------------------------------
Set objExcel_read = CreateObject("Excel.Application")
	' open Excel 2003
'Set objExcel = CreateObject("Excel.Application.11")
	' open Excel 2007
'Set objExcel = CreateObject("Excel.Application.12")
objExcel_read.Visible = FALSE
objExcel_read.ScreenUpdating = FALSE
objExcel_read.DisplayAlerts = FALSE 
Set objWorkbook_read = objExcel_read.Workbooks.Open(strExcelPath_read)

munkalapsz = "SQL"
objWorkbook_read.Worksheets(munkalapsz).Activate

For i = 1 To 16   '' excel olvasó ciklus kezdet

  ki_xls_neve(i) = objExcel_read.Cells( i+1, 2).Value
  munkalapnev(i) = objExcel_read.Cells( i+1, 3).Value
  sql(i) = objExcel_read.Cells( i+1, 4).Value
  kezd_sor(i) = objExcel_read.Cells( i+1, 5).Value
  kezd_oszlop(i) = objExcel_read.Cells( i+1, 6).Value
  fejlec(i) = objExcel_read.Cells( i+1, 7).Value
  szamolo(i) = objExcel_read.Cells( i+1, 8).Value
  r_nev(i) = objExcel_read.Cells( i+1, 1).Value
  eltolasok(i) = objExcel_read.Cells( i+1, 9).Value
  kapcsolatok(i) = objExcel_read.Cells( i+1, 10).Value
  
Next				'' excel olvasó ciklus vég

olvass_el = objExcel_read.Cells(2,11).Value

munkalapsz = "kapcsolat"
objWorkbook_read.Worksheets(munkalapsz).Activate
strCon = objExcel_read.Cells( 2, 2).Value

' Erõforrás felszabadítása
objExcel_read.ActiveWorkbook.Save
objExcel_read.ActiveWorkbook.Close
objWorkbook_read = Nothing
objExcel_read.Application.Quit
objExcel_read = Nothing 

'WSCript.Echo "----------ora XLS-bõl ------------"
'WSCript.Echo strCon
'wscript.quit

'-------------------------------- Excel olvasás  vége------------------------------------------------

'--------------------------------  EXCEL írás  kezedet----------------------------------------------
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = FALSE
objExcel.ScreenUpdating = FALSE
objExcel.DisplayAlerts = FALSE 
Set objWorkbook = objExcel.Workbooks.Add

'--------- Oracle konnekció
Dim oCon: Set oCon = WScript.CreateObject("ADODB.Connection")
Dim oRs: Set oRs = WScript.CreateObject("ADODB.Recordset")
oCon.Open strCon
WSCript.Echo "---------- ora kapcs -----------"

For i = 1 To 16    	''  riport tömbön lépkedés kezdet

 munkalapsz = munkalapnev(i)
 if  munkalapsz > "" then 				''' csak munkalap név kitöltésnél kezdet

WSCript.Echo "---------- xls munkalap  -----------"
WSCript.Echo munkalapnev(i)

WSCript.Echo "---------- xls riport nev -----------"
WSCript.Echo r_nev(i)

WSCript.Echo "---------- idõ -----------"
wscript.echo now 
 
sork = kezd_sor(i) 'Y
oszlopk = kezd_oszlop(i)  'x

objWorkbook.Sheets.Add
'Konténerhez új objektumot hozzáfûzni az Add metódussal lehet
'ActiveWorkbook.Sheets.Add {[Before|, After]}[,[Count][,Type]] 
'Az elsõ elé két munkalapot, az utolsóként beszúrt lesz aktív
'Worksheets.Add Worksheets.Item(1), , 2

objWorkbook.ActiveSheet.Name = r_nev(i)
' objWorkbook.Worksheets(munkalapsz).Activate
oszlopn = 0
sorn = 0

Set oRs = oCon.Execute(sql(i))

eltolas_tomb = Split(eltolasok(i),",")

 'If oRs.RecordCount <> 0 Then  
 If Not (oRs.BOF And oRs.EOF) Then   '' van eredmény kezdet
   WSCript.Echo "---------- Van eredmény -----------"
   WSCript.Echo oRs.RecordCount

 do until oRs.EOF 					'' eredmény tömb sorolvasás kezdet

  for each eredmeny in oRs.Fields   '' egy eredmeny sor kezdet
    
   ' eltolás használat	  
        eltolas = 0
	      if Len(eltolasok(i)) > 1 Then 		  
		    eltolas = eltolas_tomb(oszlopn)*1	  
		  End If
   
	  If sorn = 0  Then  '' elsor sor kezdet
	  
' fejléc

       objExcel.Cells( (sork + sorn),(oszlopk + oszlopn + eltolas)).Value = eredmeny.name
	   
	   'Fejléc formázása	   
        'With objExcel.Selection
		With objExcel.Cells( (sork + sorn),(oszlopk + oszlopn + eltolas))
		   With .Font
	   	   	.Bold = TRUE
		    .Size = 12
		    .ColorIndex = 3 
			End With
		  .Interior.ColorIndex =  6
		  '.Columns.Autofit
		  '.WrapText = True
		  '.VerticalAlignment = -4108
          '.Borders(xlTop).LineStyle = xlNone 
          .Borders(xlBottom).LineStyle = xlContinuous
          '.Borders(xlLeft).LineStyle = xlNone
          '.Borders(xlRight).LineStyle = xlNone
          '.ColumnWidth = 8.43 
          '.RowHeight = 14 
		  '.Orientation = xlHorizontal
		  '.Orientation = xlVertical''' vagy ez TJ
          '.Orientation = xlUpward   ''' ez TJ
          '.Orientation = xlDownward
          '.Orientation = 45
          .Orientation = 90

		End With	
		
       if  eredmeny.value > "" then   '' van eredmeny 1 kezdet
' 1. érték adás	   	   
	   objExcel.Cells( (sork + sorn + 1),(oszlopk + oszlopn + eltolas)).Value = eredmeny.value

	   'Elsõ adatsor formázása	   
        'With objExcel.Selection
		With objExcel.Cells( (sork + sorn + 1),(oszlopk + oszlopn + eltolas))
		   With .Font
	   	   	.Bold = FALSE
		    .Size = 10
		    .ColorIndex = 1 
			'.Color = vbRed
			End With
		  '.Interior.ColorIndex =  8
		End With	
		
	   End if					'' van eredmeny 1 vég
	   
	  else						'' elsor sor kezdet különben
 
      if  eredmeny.value > "" then   ''  van eredmény 2. kezdete
	  
'további érték adás
	   objExcel.Cells( (sork + sorn + 1), (oszlopk + oszlopn + eltolas)).Value = eredmeny.value
	   
	   
'további adatsor formázása	   
        'With objExcel.Selection
		With objExcel.Cells( (sork + sorn + 1),(oszlopk + oszlopn + eltolas))
		   With .Font
	   	   	.Bold = FALSE
		    .Size = 10
		    .ColorIndex = 1 
			'.Color = Red
			End With
		  '.Interior.ColorIndex =  8
		  
		End With	
		
	  End If						''  van eredmény 2. vég	  
	  
	 End If							'' elsor sor vég
	 
    oszlopn = oszlopn +1
	
  Next								' egy eredmeny sor veg
  
     oszlopn = 0    
     sorn = sorn + 1  	 
	 sornn = sornn + 1  
	 
	 if sornn > 200 Then
	    sornn = 1
        WSCript.Echo "----------sorszam-----------"
        WSCript.Echo sorn
	 End If
	 
   oRs.MoveNext
   
  loop 								'' eredmény tömb sorolvasás vég
  
 else  					'			''van eredmény vizsgálat különben ág
  
  ' Nincs eredmény
  WSCript.Echo "---------- Nincs eredmény -----------"
    
  End If	 						'' van eredmény vége
 
 End If           					''' csak munkalap név kitöltésnél vég
 
Next								''  riport tömbön lépkedés vége

WSCript.Echo "---------- sql2xls riport készítés vége -----------"
wscript.echo now 

 if olvass_el = "1" then 
   olvass_el_m()
   WSCript.Echo "---------- Olvass el munkalap elkészült -----------"
 else
   WSCript.Echo "---------- Olvass el munkalap nem kell  -----------"
 End If

 ' Erõforrás felszabadítása
'objExcel.ActiveWorkbook.Save
strExcelPath = konyvtar & "\" & ki_xls_neve(1)
objExcel.ActiveWorkbook.SaveAs strExcelPath

objExcel.ActiveWorkbook.Close
objWorkbook = Nothing
objExcel.Application.Quit
objExcel = Nothing 
 
''--------------------------------  EXCEL írás  vég----------------------------------------------

' db kapcsolat felszabadítás
adoRecordset.Close
adoConnection.Close 

oCon.Close
Set oRs = Nothing
Set oCon = Nothing
' ----------------------------------------- SUB-ok
Sub olvass_el_m()

' ----   olvass el
  objWorkbook.Sheets.Add
  objWorkbook.ActiveSheet.Name = "olvass_el"
  objExcel.Cells( 1,1).Value = " Készült az SQL2XLS programmal"
  objExcel.Cells( 2,1).Value = "Készítési dátum : " & now    '& date & " \ " & time 
  'objExcel.Cells( 2,1).Columns.Autofit
  
  '  objWorkbook.ActiveSheet.Range("D5").Select
  '  Selection.BorderAround LineStyle:=xlDouble, Color:=vbBlue
  '  ActiveSheet.Range("D1").Select
  
' ----    olvass el formázása	   
       With objExcel.Range("A1","A5")
	      .Font.Size = 12
		  .Font.Color = vbRed
   	      .Interior.ColorIndex = 8
		  .Columns.Autofit
		  '.BorderAround LineStyle:=xlDouble, Color:=vbBlue
  	   End With	
    
End Sub



' vége