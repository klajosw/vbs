'********************************************************************
'*
'* Copyright Kecskem�ti Lajos
'*
'* Module Name:    sql2xls_upg.vbs 
'*
'* SQL eredm�ny XLS-be t�lt�se
'*
'*
'********************************************************************

' alap be�ll�t�sok
Option Explicit

ON ERROR RESUME NEXT
Err.Clear
'On Error GoTo ErrorHandler

' valtoz�k
Dim objRootDSE, strDNSDomain, adoConnection
Dim strBase, strFilter, strAttributes, strQuery, adoRecordset
Dim strName, strDN, objManagerList, strManagerDN
Dim objExcel, objExcel_read, objWorkbook, objWorkbook_read,sorn, oszlopn, sork, oszlopk, eField, objRange, objRange2
Dim strExcelPath, strExcelPath_read, konyvtar
Dim strCon, strsql, i, ii, eltolas, eltolas_tomb, eltsz, sornn, olvass_el
Dim eredmeny_tomb(2) 
Dim ki_xls_neve(150),	munkalapnev(150),	sql(150),	kezd_sor(150), kezd_oszlop(150), fejlec(150),	szamolo(150),	r_nev(150), eltolasok(150), kapcsolatok(150)
Dim kezdx, kezdy, munkalapsz, eredmeny

Const xlAscending = 1
Const xlDescending = 2
Const xlYes = 1

'--****************************************** V�lt **********************************
konyvtar = Left(WScript.ScriptFullName,(Len(WScript.ScriptFullName) - (Len(WScript.ScriptName) + 1)))
' eredmeny xls megad�sa csak nevvel (alap�rtelmezett k�nyvt�rba dolgozik  "d:\Kecskemet1L314\My Documents\sql_xls.xlsx"  )
strExcelPath      = konyvtar & "\master.xlsx"
strExcelPath_read = konyvtar & "\xls_main.xlsx"
'--********************************************************************************

WSCript.Echo "---------- xls nevek  -----------"
WSCript.Echo konyvtar
WSCript.Echo strExcelPath
WSCript.Echo strExcelPath_read

'-------------------------------- Excel olvas�s kezdet------------------------------------------------
Set objExcel_read = CreateObject("Excel.Application")
objExcel_read.Visible = FALSE
objExcel_read.ScreenUpdating = FALSE
objExcel_read.DisplayAlerts = FALSE 
Set objWorkbook_read = objExcel_read.Workbooks.Open(strExcelPath_read)

munkalapsz = "SQL"
objWorkbook_read.Worksheets(munkalapsz).Activate

For i = 1 To 100   '' excel olvas� ciklus kezdet

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
  
Next				'' excel olvas� ciklus v�g
olvass_el = objExcel_read.Cells(2,11).Value

munkalapsz = "kapcsolat"
objWorkbook_read.Worksheets(munkalapsz).Activate
strCon = objExcel_read.Cells( 2, 2).Value

' Er�forr�s felszabad�t�sa
objExcel_read.ActiveWorkbook.Save
objExcel_read.ActiveWorkbook.Close
objWorkbook_read = Nothing
objExcel_read.Application.Quit
objExcel_read = Nothing 

'WSCript.Echo "----------ora XLS-b�l ------------"
'WSCript.Echo strCon
'wscript.quit

'-------------------------------- Excel olvas�s  v�ge------------------------------------------------

'--------------------------------  EXCEL �r�s  kezedet----------------------------------------------
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = FALSE
objExcel.ScreenUpdating = FALSE
objExcel.DisplayAlerts = FALSE 
Set objWorkbook = objExcel.Workbooks.Open(strExcelPath)

'--------- Oracle konnekci�
Dim oCon: Set oCon = WScript.CreateObject("ADODB.Connection")
Dim oRs: Set oRs = WScript.CreateObject("ADODB.Recordset")
oCon.Open strCon
 WSCript.Echo "---------- ora kapcs -----------"
 

 For i = 1 To 100			''  riport t�mb�n l�pked�s kezdet

  munkalapsz = munkalapnev(i)
 if  munkalapsz > "" then    ''' csak munkalap n�v kit�lt�sn�l kezdet
 
WSCript.Echo "---------- xls munkalap  -----------"
WSCript.Echo munkalapnev(i)

WSCript.Echo "---------- xls riport nev -----------"
WSCript.Echo r_nev(i)

sork = kezd_sor(i) 'Y
oszlopk = kezd_oszlop(i)  'x

objWorkbook.Worksheets(munkalapsz).Activate

oszlopn = 0
sorn = 0

Set oRs = oCon.Execute(sql(i))

eltolas_tomb = Split(eltolasok(i),",")

 'If oRs.RecordCount <> 0 Then  
  If Not (oRs.BOF And oRs.EOF) Then		'' van eredm�ny kezdet
   WSCript.Echo "---------- Van eredm�ny -----------"
   WSCript.Echo oRs.RecordCount

 do until oRs.EOF						' eredm�ny t�mb sorolvas�s kezdet

  for each eredmeny in oRs.Fields		'' egy eredmeny sor kezdet
  
     ' eltol�s haszn�lat	  
        eltolas = 0
	      if Len(eltolasok(i)) > 1 Then 		  
		    eltolas = eltolas_tomb(oszlopn)*1	  
		  End If
   
	  If sorn = 0  Then  				'' elsor sor kezdet
' fejl�c
'�rt�k ad�s
		if  eredmeny.value > "" or eredmeny.value = 0 then   '' van eredmeny 1 kezdet
	      objExcel.Cells( (sork + sorn),(oszlopk + oszlopn + eltolas)).Value = eredmeny
		End if							'' van eredmeny 1 v�g
  
	  else								'' elsor sor kezdet k�l�nben
	  
		 if  eredmeny.value > "" or eredmeny.value = 0 then   ''  van eredm�ny 2. kezdete
'�rt�k ad�s
	   objExcel.Cells( (sork + sorn), (oszlopk + oszlopn + eltolas)).Value = eredmeny
	   
	   End If							''  van eredm�ny 2. v�g	  	  
	 
	End If								'' elsor sor v�g	
	
    oszlopn = oszlopn +1
	
  Next
     oszlopn = 0    
     sorn = sorn + 1  
	 sornn = sornn + 1  
	 
	 if sornn > 200 Then
	    sornn = 1
        WSCript.Echo "----------sorszam-----------"
        WSCript.Echo sorn
	 End If

   oRs.MoveNext
   
  loop 								'' eredm�ny t�mb sorolvas�s v�g
  
 else  					'			''van eredm�ny vizsg�lat k�l�nben �g
  ' Nincs eredm�ny
  WSCript.Echo "---------- Nincs eredm�ny -----------"
 
  End If	 						'' van eredm�ny v�ge
 
 End If           					''' csak munkalap n�v kit�lt�sn�l v�g
 
Next								''  riport t�mb�n l�pked�s v�ge


 if olvass_el = "1" then 
   olvass_el_m()
   WSCript.Echo "---------- Olvass el munkalap elk�sz�lt -----------"
 else
   WSCript.Echo "---------- Olvass el munkalap nem kell  -----------"
 End If
 
 
 ' Er�forr�s felszabad�t�sa
'objExcel.ActiveWorkbook.Save
strExcelPath = konyvtar & "\" & ki_xls_neve(1)
objExcel.ActiveWorkbook.SaveAs strExcelPath

objExcel.ActiveWorkbook.Close
objWorkbook = Nothing
objExcel.Application.Quit
objExcel = Nothing 
 
''--------------------------------  EXCEL �r�s  v�g----------------------------------------------

' db kapcsolat felszabad�t�s
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
  objExcel.Cells( 1,1).Value = " K�sz�lt az SQL2XLS programmal"
  objExcel.Cells( 2,1).Value = "K�sz�t�si d�tum : " & now    '& date & " \ " & time 
  'objExcel.Cells( 2,1).Columns.Autofit
  
  '  objWorkbook.ActiveSheet.Range("D5").Select
  '  Selection.BorderAround LineStyle:=xlDouble, Color:=vbBlue
  '  ActiveSheet.Range("D1").Select
  
' ----    olvass el form�z�sa	   
       With objExcel.Range("A1","A5")
	      .Font.Size = 12
		  .Font.Color = vbRed
   	      .Interior.ColorIndex = 8
		  .Columns.Autofit
		  '.BorderAround LineStyle:=xlDouble, Color:=vbBlue
  	   End With	
    
End Sub

' v�ge