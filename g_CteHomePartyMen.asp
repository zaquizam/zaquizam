<%@language=vbscript%>
<!--#include file="conexion.asp"-->

<%
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	'response.write "<br>84 LLEGO"
	'response.end
	'08Mar2021 - 2
	response.buffer = true
	Server.ScriptTimeout=1000
	Dim idCliente
	Dim gMeses
	dim iMostrar 
	iMostrar = 0
	dim sTam 
	dim sTamG
	dim iTam
	dim sCat
	dim sAre
	dim sFab
	dim sMar
	dim sSeg
	dim sRan
	dim sInd
	dim iAre
	dim iFab
	dim iMar
	dim iSeg
	dim iRan
	dim iInd
	dim TotalFab 
	dim TotalMar
	dim TotalSeg
	dim TotalRan
	dim idSemana
	dim TotalDias
	'26Ene2021-2
	dim TotalFabricante 
	dim TotalArea 
	dim gProductosTotal
	dim gProductosTotalNacional
	dim Contador
	Contador = 0
	idCliente = Session("idCliente")
	if idCliente = "" then request.cookie("cliente")

	sCat=Request.QueryString("cat")
	if sCat = "" Then sCat = 1
	
	sAre=Request.QueryString("are")
	sFab=Request.QueryString("fab")
	sMar=Request.QueryString("mar")
	sSeg=Request.QueryString("seg")
	sRan=Request.QueryString("ran")
	sInd=Request.QueryString("ind")
	'08Mar2021 - 1
	sTam=Request.QueryString("tam")
	sTamG=Request.QueryString("tam")
	
	'09Feb2021-8
	TotalArea = "NO"
	if sAre <> "" then
		if Mid(sAre,1,1) = "0" then
			TotalArea = "SI"
			sAre = mid(sAre,2)
			if Mid(sAre,1,1) = "," then
				sAre = mid(sAre,2)
			end if
		end if
	end if
	'response.write "<br>57 " & sAre
	'	if  sAre = "" Then sAre = "1,2,3,4,5,6,7"
	'26Ene2021-8
	'TotalFabricante = "NO"
	if sFab <> "" then
		TotalArea = "NO"
		if Mid(sFab,1,1) = "0" then
			TotalFabricante = "SI"
			sFab = mid(sFab,2)
			if Mid(sFab,1,1) = "," then
				sFab = mid(sFab,2)
			end if
		end if
	end if
	
	'if sSeg <> "" and sFab = "" and sMar = "" then 
	'	sFab = "0"
	'	sMar = "0"
	'end if
	
	
	
	'response.write "<br>84 LLEGO" & sFab
	'response.end
	'if sFab = "" then sFab = 0
	'if sMar = "" then sMar = 0
	'if sSeg = "" then sSeg = 0
	'if sRan = "" then sRan = 0
	
	dim gProductos
	dim gIndicadores
	dim Indicador
	dim Valor
	
	dim gDatos1
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 1 'adLockOptimistic 

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_PeriodoDesde, "
	sql = sql & " Id_PeriodoPub "
	sql = sql & " FROM "
	sql = sql & " ss_ClienteCategoria "
	sql = sql & " WHERE "
	sql = sql & " Id_Cliente = " & idCliente
	sql = sql & " AND Id_Categoria = " & sCat
	sql = sql & " AND Ind_Mensual = 1 "
	sql = sql & " AND Ind_Activo = 1 "
	'response.write "<br>108 sql:=" & sql
	'response.end
	'rsx1.Open sql ,conexion
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gDatos1 = rsx1.GetRows
		rsx1.close
		iMesDes = gDatos1(0,0)
		iMesHas = gDatos1(1,0)
	end if

	'response.write "<br>310 Semana iMesDes:= " &  iMesDes
	'response.write "<br>310 Semana iMesHas:= " &  iMesHas

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " IdPeriodo, "
	sql = sql & " PeriodoCorto, "
	sql = sql & " Semanas "
	sql = sql & " FROM "
	sql = sql & " ss_Periodo "
	sql = sql & " WHERE "
	sql = sql & " IdPeriodo >= " & iMesDes
	sql = sql & " And IdPeriodo <= " & iMesHas
	'response.write "<br>108 sql:=" & sql
	'response.end
	'rsx1.Open sql ,conexion
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gMeses = rsx1.GetRows
		rsx1.close
	end if
	'if idCliente = 10 then
	'Manpa
	
	'Categorias Trimestrales Feb-Mar-Abr 2021
	if ((sCat = 12 or sCat = 93 and idCliente = 21) or (sCat = 106) or (sCat = 72) or (sCat = 27) or (sCat = 29) or (sCat = 30) or (sCat = 31) or (sCat = 73) or (sCat = 35) or (sCat = 19) or (sCat = 38) or (sCat = 41) or sCat = 97 or sCat = 10 or sCat = 146 or sCat = 44 or sCat = 45 or sCat = 54) then 
		if idCliente <> 8 then 
			'response.write "<br>pasoooooooooooooooo"
			'erase gMeses
			'redim gMeses(2,0)
			'gMeses(1,0) = "Trim. Feb-Mar-Abr/2021"
			'gMeses(2,0) = "20,21,22,23,24,25,26,27,28,29,30,31,32"
			'Manpa
			if (sCat = 54) and idCliente = 22 then 
				'response.write "<br>pasoooooooooooooooo"
				erase gMeses
				redim gMeses(2,4)
				gMeses(1,0) = "Abril 2021"
				gMeses(2,0) = "29,30,31,32"
				gMeses(1,1) = "Mayo 2021"
				gMeses(2,1) = "33,34,35,36"
				gMeses(1,2) = "Junio 2021"
				gMeses(2,2) = "37,38,39,40"
				gMeses(1,3) = "Diciembre 2021"
				gMeses(2,3) = "63,64,65,66,67"
				gMeses(1,4) = "Enero 2022"
				gMeses(2,4) = "68,69,70,71"
			end if
			'Central el Palmar
			if (sCat = 97) and idCliente = 20 then 
				'response.write "<br>pasoooooooooooooooo"
				erase gMeses
				redim gMeses(2,2)
				gMeses(1,0) = "Trim. Mar-Abr-May/2021"
				gMeses(2,0) = "24,25,26,27,28,29,30,31,32,33,34,35,36"
				gMeses(1,1) = "Trim. Jun-Jul-Ago/2021"
				gMeses(2,1) = "37,38,39,40,41,42,43,44,45,46,47,48,49"
				gMeses(1,2) = "Trim. Sep-Oct-Nov/2021"
				gMeses(2,2) = "50,51,52,53,54,55,56,57,58,59,60,61,62"
				'gMeses(1,3) = "Trim. Oct-Nov-Dic/2021"
				'gMeses(2,3) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
			end if
			if (sCat = 106 or sCat = 72) and idCliente = 20 then 
				'response.write "<br>pasoooooooooooooooo"
				erase gMeses
				redim gMeses(2,0)
				gMeses(1,0) = "Semestre Mar-Ago/2021"
				gMeses(2,0) = "24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49"
			end if
			'Capri
			if (sCat = 10) and idCliente = 29 then 
				'response.write "<br>pasoooooooooooooooo"
				erase gMeses
				redim gMeses(2,4)
				gMeses(1,0) = "Trim. Mar-Abr-May/2021"
				gMeses(2,0) = "24,25,26,27,28,29,30,31,32,33,34,35,36"
				gMeses(1,1) = "Trim. May-Jun-Jul/2021"
				gMeses(2,1) = "33,34,35,36,37,38,39,40,41,42,43,44,45"
				gMeses(1,2) = "Trim. Jul-Ago-Sep/2021"
				gMeses(2,2) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
				gMeses(1,3) = "Trim. Sep-Oct-Nov/2021"
				gMeses(2,3) = "50,51,52,53,54,55,56,57,58,59,60,61,62"
				gMeses(1,4) = "Trim. Nov-Dic/2021 Ene/2022"
				gMeses(2,4) = "59,60,61,62,63,64,65,66,67,68,69,70,71"
			end if
			if (sCat = 44 or sCat = 45)and idCliente = 29 then 
				erase gMeses
				redim gMeses(2,4)
				gMeses(1,0) = "Mayo 2021"
				gMeses(2,0) = "33,34,35,36"
				gMeses(1,1) = "Julio 2021"
				gMeses(2,1) = "41,42,43,44,45"
				gMeses(1,2) = "Septiembre 2021"
				gMeses(2,2) = "50,51,52,53,54"
				gMeses(1,3) = "Noviembre 2021"
				gMeses(2,3) = "59,60,61,62"
				gMeses(1,4) = "Enero 2022"
				gMeses(2,4) = "68,69,70,71"
				'response.write "paso"
				'response.end
			end if
			'El Tunal
			if (sCat = 12 and idCliente = 21) then 
				'response.write "<br>pasoooooooooooooooo"
				erase gMeses
				redim gMeses(2,7)
				'gMeses(1,0) = "Trim. Feb-Mar-Abr/2021"
				'gMeses(2,0) = "20,21,22,23,24,25,26,27,28,29,30,31,32"
				'gMeses(1,0) = "Trim. Mar-Abr-May/2021"
				'gMeses(2,0) = "24,25,26,27,28,29,30,31,32,33,34,35,36"
				gMeses(1,0) = "Trim. Abr-May-Jun/2021"
				gMeses(2,0) = "29,30,31,32,33,34,35,36,37,38,39,40"
				gMeses(1,1) = "Trim. May-Jun-Jul/2021"
				gMeses(2,1) = "33,34,35,36,37,38,39,40,41,42,43,44,45"
				gMeses(1,2) = "Trim. Jun-Jul-Ago/2021"
				gMeses(2,2) = "37,38,39,40,41,42,43,44,45,46,47,48,49"
				gMeses(1,3) = "Trim. Jul-Ago-Sep/2021"
				gMeses(2,3) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
				gMeses(1,4) = "Trim. Ago-Sep-Oct/2021"
				gMeses(2,4) = "46,47,48,49,50,51,52,53,54,55,56,57,58"
				gMeses(1,5) = "Trim. Sep-Oct-Nov/2021"
				gMeses(2,5) = "50,51,52,53,54,55,56,57,58,59,60,61,62"
				gMeses(1,6) = "Trim. Oct-Nov-Dic/2021"
				gMeses(2,6) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
				gMeses(1,7) = "Trim. Nov-Dic/2021 Ene/2022"
				gMeses(2,7) = "59,60,61,62,63,64,65,66,67,68,69,70,71"
			end if
			'El Tunal
			if (sCat = 93 and idCliente = 21) then 
				'response.write "<br>pasoooooooooooooooo"
				erase gMeses
				redim gMeses(2,2)
				gMeses(1,0) = "Trim. Abr-May-Jun/2021"
				gMeses(2,0) = "29,30,31,32,33,34,35,36,37,38,39,40"
				gMeses(1,1) = "Trim. Jul-Ago-Sep/2021"
				gMeses(2,1) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
				gMeses(1,2) = "Trim. Oct-Nov-Dic/2021"
				gMeses(2,2) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
			end if
			'Baron
			if ((sCat = 19 or sCat = 38) and idCliente = 19)  then 
				'response.write "<br>pasoooooooooooooooo"
				erase gMeses
				redim gMeses(2,4)
				gMeses(1,0) = "Trim. Feb-Mar-Abr/2021"
				gMeses(2,0) = "20,21,22,23,24,25,26,27,28,29,30,31,32"
				gMeses(1,1) = "Trim. Mar-Abr-May/2021"
				gMeses(2,1) = "24,25,26,27,28,29,30,31,32,33,34,35,36"
				gMeses(1,2) = "Trim. Abr-May-Jun/2021"
				gMeses(2,2) = "29,30,31,32,33,34,35,36,37,38,39,40"
				gMeses(1,3) = "Trim. May-Jun-Jul/2021"
				gMeses(2,3) = "33,34,35,36,37,38,39,40,41,42,43,44,45"
				gMeses(1,4) = "Trim. Jun-Jul-Ago/2021"
				gMeses(2,4) = "37,38,39,40,41,42,43,44,45,46,47,48,49"
			end if
			if ((sCat = 41 or sCat = 146) and idCliente = 19)  then 
				'response.write "<br>pasoooooooooooooooo"
				erase gMeses
				redim gMeses(2,4)
				gMeses(1,0) = "Trim. Feb-Mar-Abr/2021"
				gMeses(2,0) = "20,21,22,23,24,25,26,27,28,29,30,31,32"
				gMeses(1,1) = "Trim. Mar-Abr-May/2021"
				gMeses(2,1) = "24,25,26,27,28,29,30,31,32,33,34,35,36"
				gMeses(1,2) = "Trim. Abr-May-Jun/2021"
				gMeses(2,2) = "29,30,31,32,33,34,35,36,37,38,39,40"
				gMeses(1,3) = "Trim. May-Jun-Jul/2021"
				gMeses(2,3) = "33,34,35,36,37,38,39,40,41,42,43,44,45"
				gMeses(1,4) = "Trim. Jun-Jul-Ago/2021"
				gMeses(2,4) = "37,38,39,40,41,42,43,44,45,46,47,48,49"
			end if
			'SC Johnson
			if (sCat = 27) or (sCat = 30) or (sCat = 29) or (sCat = 31) or (sCat = 73) then 
				erase gMeses
				redim gMeses(2,3)
				gMeses(1,0) = "Trim. Ene-Feb-Mar/2021"
				gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28"
				gMeses(1,1) = "Trim. Feb-Mar-Abr/2021"
				gMeses(2,1) = "20,21,22,23,24,25,26,27,28,29,30,31,32"
				gMeses(1,2) = "Trim. Mar-Abr-May/2021"
				gMeses(2,2) = "24,25,26,27,28,29,30,31,32,33,34,35,36"
				gMeses(1,3) = "Trim. Abr-May-Jun/2021"
				gMeses(2,3) = "29,30,31,32,33,34,35,36,37,38,39,40"
				'response.write "paso"
				'response.end
			end if
			' Pepsico Alimento
			if (sCat = 35) and idCliente = 11 then 
				erase gMeses 
				redim gMeses(2,7)
				'gMeses(1,0) = "Trim. Ene-Feb-Mar/2021"
				'gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28"
				'gMeses(1,0) = "Trim. Feb-Mar-Abr/2021"
				'gMeses(2,0) = "20,21,22,23,24,25,26,27,28,29,30,31,32"
				'gMeses(1,0) = "Trim. Mar-Abr-May/2021"
				'gMeses(2,0) = "24,25,26,27,28,29,30,31,32,33,34,35,36"
				gMeses(1,0) = "Trim. Abr-May-Jun/2021"
				gMeses(2,0) = "29,30,31,32,33,34,35,36,37,38,39,40"
				gMeses(1,1) = "Trim. May-Jun-Jul/2021"
				gMeses(2,1) = "33,34,35,36,37,38,39,40,41,42,43,44,45"
				gMeses(1,2) = "Trim. Jun-Jul-Ago/2021"
				gMeses(2,2) = "37,38,39,40,41,42,43,44,45,46,47,48,49"
				gMeses(1,3) = "Trim. Jul-Ago-Sep/2021"
				gMeses(2,3) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
				gMeses(1,4) = "Trim. Ago-Sep-Oct/2021"
				gMeses(2,4) = "46,47,48,49,50,51,52,53,54,55,56,57,58"
				gMeses(1,5) = "Trim. Sep-Oct-Nov/2021"
				gMeses(2,5) = "50,51,52,53,54,55,56,57,58,59,60,61,62"
				gMeses(1,6) = "Trim. Oct-Nov-Dic/2021"
				gMeses(2,6) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
				gMeses(1,7) = "Trim. Nov-Dic/2021 Ene/2022"
				gMeses(2,7) = "59,60,61,62,63,64,65,66,67,68,69,70,71"
				'response.write "paso"
				'response.end
			end if
		end if
	end if
	'Nestle
	if (sCat = 9) then 
		'response.write "<br>pasoooooooooooooooo"
		erase gMeses
		redim gMeses(2,0)
		gMeses(1,0) = "Trim. Abr-May-Jun/2021"
		gMeses(2,0) = "29,30,31,32,33,34,35,36,37,38,39,40"
	end if
	if (sCat = 14) then 
		'response.write "<br>pasoooooooooooooooo"
		erase gMeses
		redim gMeses(2,3)
		gMeses(1,0) = "Trim. Feb-Mar-Abr/2021"
		gMeses(2,0) = "20,21,22,23,24,25,26,27,28,29,30,31,32"
		gMeses(1,1) = "Trim. Mar-Abr-May/2021"
		gMeses(1,1) = "Trim. Mar-Abr-May/2021"
		gMeses(2,1) = "24,25,26,27,28,29,30,31,32,33,34,35,36"
		gMeses(1,2) = "Trim. Abr-May-Jun/2021"
		gMeses(2,2) = "29,30,31,32,33,34,35,36,37,38,39,40"
		gMeses(1,3) = "Trim. May-Jun-Jul/2021"
		gMeses(2,3) = "33,34,35,36,37,38,39,40,41,42,43,44,45"
	end if
	if (sCat = 8)then 
		'response.write "<br>pasoooooooooooooooo"
		erase gMeses 
		redim gMeses(2,7) 
		'gMeses(1,0) = "Trim. Feb-Mar-Abr/2021"
		'gMeses(2,0) = "20,21,22,23,24,25,26,27,28,29,30,31,32"
		'gMeses(1,0) = "Trim. Mar-Abr-May/2021"
		'gMeses(2,0) = "24,25,26,27,28,29,30,31,32,33,34,35,36"
		gMeses(1,0) = "Trim. Abr-May-Jun/2021"
		gMeses(2,0) = "29,30,31,32,33,34,35,36,37,38,39,40"
		gMeses(1,1) = "Trim. May-Jun-Jul/2021"
		gMeses(2,1) = "33,34,35,36,37,38,39,40,41,42,43,44,45"
		gMeses(1,2) = "Trim. Jun-Jul-Ago/2021"
		gMeses(2,2) = "37,38,39,40,41,42,43,44,45,46,47,48,49"
		gMeses(1,3) = "Trim. Jul-Ago-Sep/2021"
		gMeses(2,3) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
		gMeses(1,4) = "Trim. Ago-Sep-Oct/2021"
		gMeses(2,4) = "46,47,48,49,50,51,52,53,54,55,56,57,58"
		gMeses(1,5) = "Trim. Sep-Oct-Nov/2021"
		gMeses(2,5) = "50,51,52,53,54,55,56,57,58,59,60,61,62"
		gMeses(1,6) = "Trim. Oct-Nov-Dic/2021"
		gMeses(2,6) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
		gMeses(1,7) = "Trim. Nov-Dic/2021 Ene/2022"
		gMeses(2,7) = "59,60,61,62,63,64,65,66,67,68,69,70,71"
	end if
	'Unilever
	if (sCat = 37 or sCat = 36) and idCliente = 36 then 
		erase gMeses
		redim gMeses(2,2)
		gMeses(1,0) = "Trim. Abr-May-Jun/2021"
		gMeses(2,0) = "29,30,31,32,33,34,35,36,37,38,39,40"
		gMeses(1,1) = "Trim. Jul-Ago-Sep/2021"
		gMeses(2,1) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
		gMeses(1,2) = "Trim. Oct-Nov-Dic/2021"
		gMeses(2,2) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
		'response.write "paso"
		'response.end
	end if
	if (sCat = 40 or sCat = 42) and idCliente = 36 then 
		erase gMeses
		redim gMeses(2,2)
		gMeses(1,0) = "Trim. Abr-May-Jun/2021"
		gMeses(2,0) = "29,30,31,32,33,34,35,36,37,38,39,40"
		gMeses(1,1) = "Trim. Jul-Ago-Sep/2021"
		gMeses(2,1) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
		gMeses(1,2) = "Trim. Oct-Nov-Dic/2021"
		gMeses(2,2) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
		'response.write "paso"
		'response.end
	end if
	if (sCat = 35) and idCliente = 36 then 
		erase gMeses
		redim gMeses(2,1)
		gMeses(1,0) = "Semestre Ene-Jun/2021"
		gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
		gMeses(1,1) = "Semes. Jul - Dic/2021"
		gMeses(2,1) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
		'response.write "paso"
		'response.end
	end if

	'Categorias Semestrales Ene-Feb-Mar-Abr/2021
	if (sCat = 6) or (sCat = 5) then 
		'response.write "<br>pasoooooooooooooooo"
		erase gMeses
		redim gMeses(2,1)
		gMeses(1,0) = "Semes. Ene-Jun/2021"
		gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
		gMeses(1,1) = "Semes. Feb - Jul/2021"
		gMeses(2,1) = "20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45"
	end if
	if (sCat = 11) or (sCat = 57) then 
		'response.write "<br>pasoooooooooooooooo"
		erase gMeses
		redim gMeses(2,0)
		gMeses(1,0) = "Semes. Ene-Feb-Mar-Abr/2021"
		gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32"
	end if
	'Categorias Semestrales Ene-Feb-Mar-Abr/2021
	if (sCat = 19 or sCat = 18 or sCat = 41 or sCat = 87 or sCat = 36) and idCliente = 27 then 
		'response.write "<br>pasoooooooooooooooo"
		erase gMeses
		redim gMeses(2,0)
		gMeses(1,0) = "Semes. Ene-Feb-Mar-Abr-May/2021"
		gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36"
	end if

	'Categorias Semestrales Ene-Feb-Mar-Abr-May-JUn/2021
	'Del Monte
	if (sCat = 105 or sCat = 70 or sCat = 10 or sCat = 22 ) and idCliente = 33 then 
		'response.write "<br>pasoooooooooooooooo"
		erase gMeses
		redim gMeses(2,1)
		gMeses(1,0) = "Semes. Ene-Jun/2021"
		gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
		gMeses(1,1) = "Semes. Jul - Dic/2021"
		gMeses(2,1) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
	end if
	if (sCat = 55 ) and idCliente = 35 then 
		'response.write "<br>pasoooooooooooooooo"
		erase gMeses
		redim gMeses(2,0)
		gMeses(1,0) = "Semes. Ene-Feb-Mar-Abr-May/2021"
		gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36"
	end if
	'Dimassi
	if idCliente = 8 then 
		if (sCat = 41 or sCat = 146) then 
			'response.write "<br>pasoooooooooooooooo"
			erase gMeses
			redim gMeses(2,3)
			gMeses(1,0) = "Trim. Abr-May-Jun/2021"
			gMeses(2,0) = "28,29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,1) = "Trim. May-Jun-Jul/2021"
			gMeses(2,1) = "33,34,35,36,37,38,39,40,41,42,43,44,45"
			gMeses(1,2) = "Trim. Jun-Jul-Ago/2021" 
			gMeses(2,2) = "37,38,39,40,41,42,43,44,45,46,47,48,49"
			gMeses(1,3) = "Trim. Jul-Ago-Sep/2021"
			gMeses(2,3) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
		end if
	end if
	
	if idCliente = 45 then 
		'3 Agua Mineral
		'22 Jugos Corta - Larga Duracion
		'1 Refresco
		if (sCat = 3 or sCat = 22 or sCat = 1) then
			'response.write "<br>pasoooooooooooooooo2"
			erase gMeses
			redim gMeses(2,3) 
			gMeses(1,0) = "Q1<br>Ene-Feb-Mar<br>2021"
			gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28"
			gMeses(1,1) = "Q2<br>Abr-May-Jun<br>2021"
			gMeses(2,1) = "29,30,31,32,33,34,35,36,33,34,35,36"
			gMeses(1,2) = "Q3<br>Jul-Ago-Sep<br>2021"
			gMeses(2,2) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
			gMeses(1,3) = "Q4<br>Oct-Nov-Dic<br>2021"
			gMeses(2,3) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
			'gMeses(1,4) = "Acumulado<br>2021"
			'gMeses(2,4) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
		end if
	end if
	if idCliente = 13 then 
		'89 Arroz
		if (sCat = 89) then
			'response.write "<br>pasoooooooooooooooo2"
			'response.end
			erase gMeses
			redim gMeses(2,7) 
			gMeses(1,0) = "Abril 2021"
			gMeses(2,0) = "29,30,31,32"
			gMeses(1,1) = "Mayo 2021"
			gMeses(2,1) = "33,34,35,36"
			gMeses(1,2) = "Junio 2021"
			gMeses(2,2) = "37,38,39,40"
			gMeses(1,3) = "Julio 2021"
			gMeses(2,3) = "41,42,43,44,45"
			gMeses(1,4) = "Agosto 2021"
			gMeses(2,4) = "46,47,48,49"
			gMeses(1,5) = "Septiembre 2021"
			gMeses(2,5) = "50,51,52,53,54"
			gMeses(1,6) = "Octubre 2021"
			gMeses(2,6) = "55,56,57,58"
			gMeses(1,7) = "Enero 2022"
			gMeses(2,7) = "68,69,70,71"
		end if
	end if

	'response.write "<br>idCliente:= " & idCliente
	'response.write "<br>sCat:= " & sCat
	if idCliente = 3 then 
		'Pepsico Bebidas
		' 3 = Agua Mineral
		' 22 = Jugos Corta/Larga Duracion
		if (sCat = 3) or (sCat = 22) then 
			erase gMeses
			redim gMeses(2,10)
			gMeses(1,0) = "Ene-Feb-Mar/21"
			gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28"
			gMeses(1,1) = "Feb-Mar-Abr/21"
			gMeses(2,1) = "20,21,22,23,24,25,26,27,28,29,30,31,32"
			gMeses(1,2) = "Mar-Abr-May/21"
			gMeses(2,2) = "24,25,26,27,28,29,30,31,32,33,34,35,36"
			gMeses(1,3) = "Abr-May-Jun/21"
			gMeses(2,3) = "29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,4) = "May-Jun-Jul/21"
			gMeses(2,4) = "33,34,35,36,37,38,39,40,41,42,43,44,45"
			gMeses(1,5) = "Jun-Jul-Ago/21"
			gMeses(2,5) = "37,38,39,40,41,42,43,44,45,46,47,48,49"
			gMeses(1,6) = "Jul-Ago-Sep/21"
			gMeses(2,6) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
			gMeses(1,7) = "Ago-Sep-Oct/21"
			gMeses(2,7) = "46,47,48,49,50,51,52,53,54,55,56,57,58"
			gMeses(1,8) = "Sep-Oct-Nov/21"
			gMeses(2,8) = "50,51,52,53,54,55,56,57,58,59,60,61,62"
			gMeses(1,9) = "Oct-Nov-Dic/21"
			gMeses(2,9) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
			gMeses(1,10) = "Nov-Dic/21 Ene/22"
			gMeses(2,10) = "59,60,61,62,63,64,65,66,67,68,69,70,71"
			'response.write "paso"
			'response.end
		end if
	end if
	
	'Atenas Agrupado
	if idCliente = 35 then 
		'Malta
		if (sCat = 4) then 
			erase gMeses
			redim gMeses(2,4)
			gMeses(1,0) = "Trim Sep-Oct-Nov/2021"
			gMeses(2,0) = "50,51,52,53,54,55,56,57,58,59,60,61,62"
			gMeses(1,1) = "Trim Oct-Nov-Dic/2021"
			gMeses(2,1) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
			gMeses(1,2) = "Trim Nov-Dic/2021 Ene/2022"
			gMeses(2,2) = "59,60,61,62,63,64,65,66,67,68,69,70,71"
			gMeses(1,3) = "Sem. Ene-Jun/2021"
			gMeses(2,3) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,4) = "Sem. Jul-Dic/2021"
			gMeses(2,4) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
			
			'response.write "paso"
			'response.end
		end if
		'Bebidas Energizantes
		if (sCat = 91) then 
			erase gMeses
			redim gMeses(2,2)
			gMeses(1,0) = "Sem. Ene-Jun/2021"
			gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,1) = "Sem. Jul-Dic/2021"
			gMeses(2,1) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
			gMeses(1,2) = "Ene-Dic/2021"
			gMeses(2,2) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
			'response.write "paso"
			'response.end
		end if
		'Sodas/Mezzcladores
		if (sCat = 109) then 
			erase gMeses
			redim gMeses(2,2)
			gMeses(1,0) = "Sem. Ene-Jun/2021"
			gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,1) = "Sem. Jul-Dic/2021"
			gMeses(2,1) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
			gMeses(1,2) = "Ene-Dic/2021"
			gMeses(2,2) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
			'response.write "paso"
			'response.end
		end if
		'Agua Mineral
		if (sCat = 3) then 
			erase gMeses
			redim gMeses(2,4)
			gMeses(1,0) = "Sem. Ene-Jun/2021"
			gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,1) = "Sem. Jul-Dic/2021"
			gMeses(2,1) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
			gMeses(1,2) = "Ene-Dic/2021"
			gMeses(2,2) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
			gMeses(1,3) = "Trim Jul-Sep/2021"
			gMeses(2,3) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
			gMeses(1,4) = "Trim Oct-Dic/2021"
			gMeses(2,4) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
			'response.write "paso"
			'response.end
		end if
		'Prot Femenina
		if (sCat = 18) then 
			erase gMeses
			redim gMeses(2,2)
			gMeses(1,0) = "Sem. Ene-Jun/2021"
			gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,1) = "Sem. Jul-Dic/2021"
			gMeses(2,1) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
			gMeses(1,2) = "Ene-Dic/2021"
			gMeses(2,2) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
			'response.write "paso"
			'response.end
		end if
		'Tintes
		if (sCat = 68) then 
			erase gMeses
			redim gMeses(2,2)
			gMeses(1,0) = "Sem. Ene-Jun/2021"
			gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,1) = "Sem. Jul-Dic/2021"
			gMeses(2,1) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
			gMeses(1,2) = "Ene-Dic/2021"
			gMeses(2,2) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
			'response.write "paso"
			'response.end
		end if
		'Cepillos Dentales
		if (sCat = 56) then 
			erase gMeses
			redim gMeses(2,2)
			gMeses(1,0) = "Sem. Ene-Jun/2021"
			gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,1) = "Sem. Jul-Dic/2021"
			gMeses(2,1) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
			gMeses(1,2) = "Ene-Dic/2021"
			gMeses(2,2) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
			'response.write "paso"
			'response.end
		end if
		'Enjuague Bucal
		if (sCat = 64) then 
			erase gMeses
			redim gMeses(2,2)
			gMeses(1,0) = "Sem. Ene-Jun/2021"
			gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,1) = "Sem. Jul-Dic/2021"
			gMeses(2,1) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
			gMeses(1,2) = "Ene-Dic/2021"
			gMeses(2,2) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
			'response.write "paso"
			'response.end
		end if
		'Pan Industrial
		if (sCat = 48) then 
			erase gMeses
			redim gMeses(2,2)
			gMeses(1,0) = "Trim. Jun-Jul-Ago/2021"
			gMeses(2,0) = "37,38,39,40,41,42,43,44,45,46,47,48,49"
			gMeses(1,1) = "Trim. Sep-Oct-Nov/2021"
			gMeses(2,1) = "50,51,52,53,54,55,56,57,58,59,60,61,62"
			gMeses(1,2) = "Trim. Oct-Nov-Dic/2021"
			gMeses(2,2) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
			'response.write "paso"
			'response.end
		end if
	
		if (sCat = 95) then 
			erase gMeses
			redim gMeses(2,0)
			gMeses(1,0) = "Desde Ene Hasta Nov/2021"
			gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62"
			'response.write "paso"
			'response.end
		end if
		if (sCat = 25) then 
			erase gMeses
			redim gMeses(2,1)
			gMeses(1,0) = "Trim. Jun-Jul-Ago/2021"
			gMeses(2,0) = "37,38,39,40,41,42,43,44,45,46,47,48,49"
			gMeses(1,1) = "Trim. Sep-Oct-Nov/2021"
			gMeses(2,1) = "50,51,52,53,54,55,56,57,58,59,60,61,62"
			'response.write "paso"
			'response.end
		end if
		if (sCat = 47) then 
			'Mayonesa
			erase gMeses
			redim gMeses(2,6)
			gMeses(1,0) = "Trim. Ene-Feb-Mar/2021"
			gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28"
			gMeses(1,1) = "Trim. Abr-May-Jun/2021"
			gMeses(2,1) = "29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,2) = "Trim. Jun-Jul-Ago/2021"
			gMeses(2,2) = "37,38,39,40,41,42,43,44,45,46,47,48,49"
			gMeses(1,3) = "Trim. Jul-Ago-Sep/2021"
			gMeses(2,3) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
			gMeses(1,4) = "Trim. Ago-Sep-Oct/2021"
			gMeses(2,4) = "46,47,48,49,50,51,52,53,54,55,56,57,58"
			gMeses(1,5) = "Trim. Sep-Oct-Nov/2021"
			gMeses(2,5) = "50,51,52,53,54,55,56,57,58,59,60,61,62"
			gMeses(1,6) = "Trim. Oct-Nov-Dic/2021"
			gMeses(2,6) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
			'response.write "paso"
			'response.end
		end if
		'Cereales_Complem Alimenticios 
		if (sCat = 13) then  
			'response.write "<br>pasoooooooooooooooo"
			erase gMeses
			redim gMeses(2,2)
			gMeses(1,0) = "Enero - Noviembre 2021"
			gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62"
			gMeses(1,1) = "Trim Jul-Ago-Sep/2021"
			gMeses(2,1) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
			gMeses(1,2) = "Trim Oct-Nov-Dic/2021"
			gMeses(2,2) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
		end if
		if (sCat = 26 ) then 
			'response.write "<br>pasoooooooooooooooo"
			erase gMeses
			redim gMeses(2,1)
			gMeses(1,0) = "Semes. Mar-Abr-May-Jun/2021"
			gMeses(2,0) = "24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,1) = "Enero-Noviembre/2021"
			gMeses(2,1) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62"
		end if
		if (sCat = 20 ) then 
			'response.write "<br>pasoooooooooooooooo"
			erase gMeses
			redim gMeses(2,0) 
			gMeses(1,0) = "Mayo-Octubre/2021" 
			gMeses(2,0) = "33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58"
		end if
		if (sCat = 15 ) then 
			'response.write "<br>pasoooooooooooooooo"
			erase gMeses
			redim gMeses(2,0) 
			gMeses(1,0) = "Mayo-Octubre/2021" 
			gMeses(2,0) = "33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58"
		end if
		'Refresco
		if (sCat = 1 ) then 
			'response.write "<br>pasoooooooooooooooo"
			erase gMeses
			redim gMeses(2,6) 
			gMeses(1,0) = "Trimestre May-Jun-Jul/2021" 
			gMeses(2,0) = "33,34,35,36,37,38,39,40,41,42,43,44,45"
			gMeses(1,1) = "Trimestre Ago-Sep-Oct/2021" 
			gMeses(2,1) = "46,47,48,49,50,51,52,53,54,55,56,57,58"
			gMeses(1,2) = "Trimestre Jun-Jul-Ago/2021" 
			gMeses(2,2) = "37,38,39,40,41,42,43,44,45,46,47,48,49"
			gMeses(1,3) = "Trimestre Sep-Oct-Nov/2021" 
			gMeses(2,3) = "50,51,52,53,54,55,56,57,58,59,60,61,62"
			gMeses(1,4) = "Sem. Ene-Jun/2021"
			gMeses(2,4) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,5) = "Sem. Jul-Dic/2021"
			gMeses(2,5) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
			gMeses(1,6) = "Ene-Dic/2021"
			gMeses(2,6) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
		end if
		if (sCat = 54 ) then 
			'response.write "<br>pasoooooooooooooooo"
			erase gMeses
			redim gMeses(2,0) 
			gMeses(1,0) = "Semestre Abr-Sep/2021" 
			gMeses(2,0) = "29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54"
		end if
		'Bebidas Instantaneas
		if (sCat = 21) then 
			'response.write "<br>pasoooooooooooooooo"
			'Bebidas Instantaneas
			erase gMeses
			redim gMeses(2,6)
			gMeses(1,0) = "Trim. Jun-Jul-Ago/2021"
			gMeses(2,0) = "37,38,39,40,41,42,43,44,45,46,47,48,49"
			gMeses(1,1) = "Trim. Jul-Ago-Sep/2021"
			gMeses(2,1) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
			gMeses(1,2) = "Trim. Sep-Oct-Nov/2021"
			gMeses(2,2) = "50,51,52,53,54,55,56,57,58,59,60,61,62"
			gMeses(1,3) = "Trim. Oct-Nov-Dic/2021"
			gMeses(2,3) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
			gMeses(1,4) = "Sem. Ene-Jun/2021"
			gMeses(2,4) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,5) = "Sem. Jul-Dic/2021"
			gMeses(2,5) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
			gMeses(1,6) = "Ene-Dic/2021"
			gMeses(2,6) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
		end if
		if (sCat = 66) then 
			'response.write "<br>pasoooooooooooooooo"
			erase gMeses
			redim gMeses(2,1)
			gMeses(1,0) = "Trim. May-Jun-Jul-Ago/2021"
			gMeses(2,0) = "33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49"
			gMeses(1,1) = "Trim. Jul-Ago-Sep/2021"
			gMeses(2,1) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
		end if

		if (sCat = 34 ) then 
			'response.write "<br>pasoooooooooooooooo"
			'Chocolate
			erase gMeses
			redim gMeses(2,1)
			gMeses(1,0) = "Trim. Jun-Jul-Ago/2021" 
			gMeses(2,0) = "37,38,39,40,41,42,43,44,45,46,47,48,49"
			gMeses(1,1) = "Trim. Oct-Nov-Dic/2021"
			gMeses(2,1) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
		end if
		if (sCat = 106) then 
			'response.write "<br>pasoooooooooooooooo"
			erase gMeses
			redim gMeses(2,1)
			gMeses(1,0) = "Semes. Mar - Ago/2021"
			gMeses(2,0) = "24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49"
			gMeses(1,1) = "Enero - Octubre 2021"
			gMeses(2,1) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58"
		end if
		if (sCat = 23) then  
			'response.write "<br>pasoooooooooooooooo"
			'Lacteos Larga Duracion
			erase gMeses
			redim gMeses(2,2)
			gMeses(1,0) = "Enero - Octubre 2021"
			gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58"
			gMeses(1,1) = "Trim. Jul-Ago-Sep/2021"
			gMeses(2,1) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
			gMeses(1,2) = "Trim. Oct-Nov-Dic/2021"
			gMeses(2,2) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
		end if
		if (sCat = 72) then 
			'response.write "<br>pasoooooooooooooooo"
			erase gMeses
			redim gMeses(2,0)
			gMeses(1,0) = "Semes. Mar - Ago/2021"
			gMeses(2,0) = "24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49"
		end if
		if (sCat = 97) then 
			'Mezcla para Postres
			'response.write "<br>pasoooooooooooooooo"
			erase gMeses
			redim gMeses(2,3)
			gMeses(1,0) = "Trim. Mar-Abr-May/2021"
			gMeses(2,0) = "24,25,26,27,28,29,30,31,32,33,34,35,36"
			gMeses(1,1) = "Trim. Jun-Jul-Ago/2021"
			gMeses(2,1) = "37,38,39,40,41,42,43,44,45,46,47,48,49"
			gMeses(1,2) = "Trim. Sep-Oct-Nov/2021"
			gMeses(2,2) = "50,51,52,53,54,55,56,57,58,59,60,61,62"
			gMeses(1,3) = "Trim. Oct-Nov-Dic/2021"
			gMeses(2,3) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
		end if
		if (sCat = 51) then
			'response.write "<br>pasoooooooooooooooo"
			erase gMeses
			redim gMeses(2,0)
			gMeses(1,0) = "Trim. Feb-Jul/2021"
			gMeses(2,0) = "20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45"
		end if
		if (sCat = 122) then
			'response.write "<br>pasoooooooooooooooo"
			erase gMeses
			redim gMeses(2,0)
			gMeses(1,0) = "Trim. May-Jun-Jul/2021"
			gMeses(2,0) = "33,34,35,36,37,38,39,40,41,42,43,44,45"
		end if
		if (sCat = 41) then
			'response.write "<br>pasoooooooooooooooo2"
			erase gMeses
			redim gMeses(2,8)
			gMeses(1,0) = "Trim. Feb-Mar-Abr/2021"
			gMeses(2,0) = "20,21,22,23,24,25,26,27,28,29,30,31,32"
			gMeses(1,1) = "Trim. Mar-Abr-May/2021"
			gMeses(2,1) = "24,25,26,27,28,29,30,31,32,33,34,35,36"
			gMeses(1,2) = "Trim. Abr-May-Jun/2021"
			gMeses(2,2) = "29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,3) = "Trim. May-Jun-Jul/2021"
			gMeses(2,3) = "33,34,35,36,37,38,39,40,41,42,43,44,45"
			gMeses(1,4) = "Trim. Jun-Jul-Ago/2021"
			gMeses(2,4) = "37,38,39,40,41,42,43,44,45,46,47,48,49"
			gMeses(1,5) = "Trim. Jul-Ago-Sep/2021"
			gMeses(2,5) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
			gMeses(1,6) = "Sem. Ene-Jun/2021"
			gMeses(2,6) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,7) = "Sem. Jul-Dic/2021"
			gMeses(2,7) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
			gMeses(1,8) = "Ene-Dic/2021"
			gMeses(2,8) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
		end if
		if (sCat = 146) then
			'response.write "<br>pasoooooooooooooooo2"
			erase gMeses
			redim gMeses(2,5)
			gMeses(1,0) = "Trim. Feb-Mar-Abr/2021"
			gMeses(2,0) = "20,21,22,23,24,25,26,27,28,29,30,31,32"
			gMeses(1,1) = "Trim. Mar-Abr-May/2021"
			gMeses(2,1) = "24,25,26,27,28,29,30,31,32,33,34,35,36"
			gMeses(1,2) = "Trim. Abr-May-Jun/2021"
			gMeses(2,2) = "29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,3) = "Trim. May-Jun-Jul/2021"
			gMeses(2,3) = "33,34,35,36,37,38,39,40,41,42,43,44,45"
			gMeses(1,4) = "Trim. Jun-Jul-Ago/2021"
			gMeses(2,4) = "37,38,39,40,41,42,43,44,45,46,47,48,49"
			gMeses(1,5) = "Trim. Jul-Ago-Sep/2021"
			gMeses(2,5) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
		end if
		'Caf??
		if (sCat = 17) then
			'response.write "<br>pasoooooooooooooooo2"
			erase gMeses
			redim gMeses(2,4) 
			gMeses(1,0) = "Q1<br>Ene-Feb-Mar<br>2021"
			gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28"
			gMeses(1,1) = "Q2<br>Abr-May-Jun<br>2021"
			gMeses(2,1) = "29,30,31,32,33,34,35,36,33,34,35,36"
			gMeses(1,2) = "Q3<br>Jul-Ago-Sep<br>2021"
			gMeses(2,2) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
			gMeses(1,3) = "Q4<br>Oct-Nov-Dic<br>2021"
			gMeses(2,3) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
			gMeses(1,4) = "Acumulado<br>2021"
			gMeses(2,4) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
		end if
		'Bebidas Isotonicas
		if (sCat = 92) then 
			'response.write "<br>pasoooooooooooooooo"
			erase gMeses
			redim gMeses(2,3)
			gMeses(1,0) = "Trim. Mar-Abr-May/2021"
			gMeses(2,0) = "24,25,26,27,28,29,30,31,32,33,34,35,36"
			gMeses(1,1) = "Sem. Ene-Jun/2021"
			gMeses(2,1) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,2) = "Sem. Jul-Dic/2021"
			gMeses(2,2) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
			gMeses(1,3) = "Ene-Dic/2021"
			gMeses(2,3) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
		end if
		if (sCat = 27) or (sCat = 29) or (sCat = 31) or (sCat = 73)   then 
			erase gMeses
			redim gMeses(2,4)
			gMeses(1,0) = "Trim. Ene-Feb-Mar/2021"
			gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28"
			gMeses(1,1) = "Trim. Feb-Mar-Abr/2021"
			gMeses(2,1) = "20,21,22,23,24,25,26,27,28,29,30,31,32"
			gMeses(1,2) = "Trim. Mar-Abr-May/2021"
			gMeses(2,2) = "24,25,26,27,28,29,30,31,32,33,34,35,36"
			gMeses(1,3) = "Trim. Abr-May-Jun/2021"
			gMeses(2,3) = "29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,4) = "Trim. May-Jun-Jul/2021"
			gMeses(2,4) = "33,34,35,36,37,38,39,40,41,42,43,44,45"
			'response.write "paso"
			'response.end
		end if
		'Limpiadores
		if (sCat = 30) then 
			erase gMeses
			redim gMeses(2,7)
			gMeses(1,0) = "Trim. Ene-Feb-Mar/2021"
			gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28"
			gMeses(1,1) = "Trim. Feb-Mar-Abr/2021"
			gMeses(2,1) = "20,21,22,23,24,25,26,27,28,29,30,31,32"
			gMeses(1,2) = "Trim. Mar-Abr-May/2021"
			gMeses(2,2) = "24,25,26,27,28,29,30,31,32,33,34,35,36"
			gMeses(1,3) = "Trim. Abr-May-Jun/2021"
			gMeses(2,3) = "29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,4) = "Trim. May-Jun-Jul/2021"
			gMeses(2,4) = "33,34,35,36,37,38,39,40,41,42,43,44,45"
			gMeses(1,5) = "Trim. Sep-Oct-Nov/2021"
			gMeses(2,5) = "50,51,52,53,54,55,56,57,58,59,60,61,62"
			gMeses(1,6) = "Sem. Ene-Jun/2021"
			gMeses(2,6) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,7) = "Sem. Jul-Dic/2021"
			gMeses(2,7) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
			'response.write "paso"
			'response.end
		end if
		if (sCat = 14) or (sCat = 19) or (sCat = 38) then  
			erase gMeses
			redim gMeses(2,5)
			gMeses(1,0) = "Trim. Feb-Mar-Abr/2021"
			gMeses(2,0) = "20,21,22,23,24,25,26,27,28,29,30,31,32"
			gMeses(1,1) = "Trim. Mar-Abr-May/2021"
			gMeses(2,1) = "24,25,26,27,28,29,30,31,32,33,34,35,36"
			gMeses(1,2) = "Trim. Abr-May-Jun/2021"
			gMeses(2,2) = "29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,3) = "Trim. May-Jun-Jul/2021"
			gMeses(2,3) = "33,34,35,36,37,38,39,40,41,42,43,44,45"
			gMeses(1,4) = "Trim. Jun-Jul-Ago/2021"
			gMeses(2,4) = "37,38,39,40,41,42,43,44,45,46,47,48,49"
			gMeses(1,5) = "Trim. Jul-Ago-Sep/2021"
			gMeses(2,5) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
			'response.write "paso"
			'response.end
		end if
		if (sCat = 8) then  
			'Sopas Deshidratadas
			erase gMeses
			redim gMeses(2,7)
			'gMeses(1,0) = "Trim. Mar-Abr-May/2021"
			'gMeses(2,0) = "24,25,26,27,28,29,30,31,32,33,34,35,36"
			gMeses(1,0) = "Trim. Abr-May-Jun/2021"
			gMeses(2,0) = "29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,1) = "Trim. May-Jun-Jul/2021"
			gMeses(2,1) = "33,34,35,36,37,38,39,40,41,42,43,44,45"
			gMeses(1,2) = "Trim. Jun-Jul-Ago/2021"
			gMeses(2,2) = "37,38,39,40,41,42,43,44,45,46,47,48,49"
			gMeses(1,3) = "Trim. Jul-Ago-Sep/2021"
			gMeses(2,3) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
			gMeses(1,4) = "Trim. Ago-Sep-Oct/2021"
			gMeses(2,4) = "46,47,48,49,50,51,52,53,54,55,56,57,58"
			gMeses(1,5) = "Trim. Sep-Oct-Nov/2021"
			gMeses(2,5) = "50,51,52,53,54,55,56,57,58,59,60,61,62"
			gMeses(1,6) = "Trim. Oct-Nov-Dic/2021"
			gMeses(2,6) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
			gMeses(1,7) = "Trim. Nov-Dic/2021 Ene/2022"
			gMeses(2,7) = "59,60,61,62,63,64,65,66,67,68,69,70,71"
			'response.write "paso"
			'response.end
		end if
		if (sCat = 12) then  
			'Te Listo_Beb Saborizadas RTD
			erase gMeses
			redim gMeses(2,7)
			'gMeses(1,0) = "Trim. Mar-Abr-May/2021"
			'gMeses(2,0) = "24,25,26,27,28,29,30,31,32,33,34,35,36"
			'gMeses(1,0) = "Trim. Abr-May-Jun/2021"
			'gMeses(2,0) = "29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,0) = "Trim. May-Jun-Jul/2021"
			gMeses(2,0) = "33,34,35,36,37,38,39,40,41,42,43,44,45"
			gMeses(1,1) = "Trim. Jun-Jul-Ago/2021"
			gMeses(2,1) = "37,38,39,40,41,42,43,44,45,46,47,48,49"
			gMeses(1,2) = "Trim. Jul-Ago-Sep/2021"
			gMeses(2,2) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
			gMeses(1,3) = "Trim. Ago-Sep-Oct/2021"
			gMeses(2,3) = "46,47,48,49,50,51,52,53,54,55,56,57,58"
			gMeses(1,4) = "Mayo-Octubre/2021" 
			gMeses(2,4) = "33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58"
			gMeses(1,5) = "Trim. Sep-Oct-Nov/2021"
			gMeses(2,5) = "50,51,52,53,54,55,56,57,58,59,60,61,62"
			gMeses(1,6) = "Trim. Oct-Nov-Dic/2021"
			gMeses(2,6) = "55,56,57,58,59,60,61,62,63,64,65,66,67"			
			gMeses(1,7) = "Trim. Nov-Dic/2021 Ene/2022"
			gMeses(2,7) = "59,60,61,62,63,64,65,66,67,68,69,70,71"
			'response.write "paso"
			'response.end
		end if
		if (sCat = 35) then 
			'Helados
			erase gMeses
			redim gMeses(2,7)
			'gMeses(1,0) = "Trim. Abr-May-Jun/2021"
			'gMeses(2,0) = "29,30,31,32,33,34,35,36,37,38,39,40"
			'gMeses(1,0) = "Trim. May-Jun-Jul/2021"
			'gMeses(2,0) = "33,34,35,36,37,38,39,40,41,42,43,44,45"
			gMeses(1,0) = "Semestre Ene-Jun/2021"
			gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,1) = "Trim. Jun-Jul-Ago/2021"
			gMeses(2,1) = "37,38,39,40,41,42,43,44,45,46,47,48,49"
			gMeses(1,2) = "Trim. Jul-Ago-Sep/2021"
			gMeses(2,2) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
			gMeses(1,3) = "Trim. Ago-Sep-Oct/2021"
			gMeses(2,3) = "46,47,48,49,50,51,52,53,54,55,56,57,58"
			gMeses(1,4) = "Trim. Sep-Oct-Nov/2021"
			gMeses(2,4) = "50,51,52,53,54,55,56,57,58,59,60,61,62"
			gMeses(1,5) = "Trim. Oct-Nov-Dic/2021"
			gMeses(2,5) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
			gMeses(1,6) = "Semes. Jul - Dic/2021"
			gMeses(2,6) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
			gMeses(1,7) = "Trim. Nov-Dic/2021 Ene/2022"
			gMeses(2,7) = "59,60,61,62,63,64,65,66,67,68,69,70,71"
			'response.write "paso"
			'response.end
		end if
		if (sCat = 6) or (sCat = 5) then 
			'response.write "<br>pasoooooooooooooooo"
			erase gMeses
			redim gMeses(2,2)
			gMeses(1,0) = "Semes. Ene - Jun/2021"
			gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,1) = "Semes. Feb - Jul/2021"
			gMeses(2,1) = "20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45"
			gMeses(1,2) = "Semes. Mar - Ago/2021"
			gMeses(2,2) = "24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49"
		end if
		if (sCat = 105) then 
			'Granos Empacados
			'response.write "<br>pasoooooooooooooooo"
			erase gMeses
			redim gMeses(2,1)
			gMeses(1,0) = "Semes. Ene - Jun/2021"
			gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,1) = "Semes. Jul - Dic/2021"
			gMeses(2,1) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
		end if
		if (sCat = 70) then 
			'response.write "<br>pasoooooooooooooooo"
			'Granos y Vegetales
			erase gMeses
			redim gMeses(2,1)
			gMeses(1,0) = "Semes. Ene - Jun/2021"
			gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,1) = "Semes. Jul - Dic/2021"
			gMeses(2,1) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
		end if
		if (sCat = 10) then 
			'response.write "<br>pasoooooooooooooooo"
			'Salsa para Pastas
			erase gMeses
			redim gMeses(2,1)
			gMeses(1,0) = "Semes. Ene - Jun/2021"
			gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,1) = "Semes. Jul - Dic/2021"
			gMeses(2,1) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
		end if
		' Jugos Corta - Larga Duracion
		if (sCat = 22)then 
			'response.write "<br>pasoooooooooooooooo"
			'Jugos Corta - Larga Duracion
			erase gMeses
			redim gMeses(2,2)
			gMeses(1,0) = "Sem. Ene-Jun/2021"
			gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,1) = "Sem. Jul-Dic/2021"
			gMeses(2,1) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
			gMeses(1,2) = "Ene-Dic/2021"
			gMeses(2,2) = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
		end if
		'Embutidos
		if (sCat = 93) then  
			'response.write "<br>pasoooooooooooooooo" 
			erase gMeses
			redim gMeses(2,2)
			gMeses(1,0) = "Trim. May-Jun-Jul/2021"
			gMeses(2,0) = "33,34,35,36,37,38,39,40,41,42,43,44,45"
			gMeses(1,1) = "Trim. Ago-Sep-Oct/2021"
			gMeses(2,1) = "46,47,48,49,50,51,52,53,54,55,56,57,58"
			gMeses(1,2) = "Trim. Nov-Dic/2021 Ene/2022"
			gMeses(2,2) = "59,60,61,62,63,64,65,66,67,68,69,70,71"
		end if
		if (sCat = 37) then 
			'response.write "<br>pasoooooooooooooooo"
			'Acondicionadores
			erase gMeses
			redim gMeses(2,1)
			gMeses(1,0) = "Trim. Abr-May-Jun/2021"
			gMeses(2,0) = "29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,1) = "Trim. Oct-Nov-Dic/2021"
			gMeses(2,1) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
		end if
		if (sCat = 36) then 
			'response.write "<br>pasoooooooooooooooo"
			'Champu
			erase gMeses
			redim gMeses(2,1)
			gMeses(1,0) = "Trim. Abr-May-Jun/2021"
			gMeses(2,0) = "29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,1) = "Trim. Oct-Nov-Dic/2021"
			gMeses(2,1) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
		end if
		if (sCat = 42) then 
			'response.write "<br>pasoooooooooooooooo"
			'Jabon Tocador
			erase gMeses
			redim gMeses(2,1)
			gMeses(1,0) = "Trim. Abr-May-Jun/2021"
			gMeses(2,0) = "29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,1) = "Trim. Oct-Nov-Dic/2021"
			gMeses(2,1) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
		end if
		if (sCat = 40) then 
			'response.write "<br>pasoooooooooooooooo"
			'Desodorantes
			erase gMeses
			redim gMeses(2,3)
			gMeses(1,0) = "Trim. Abr-May-Jun/2021"
			gMeses(2,0) = "29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,1) = "Trim. Ago-Sep-Oct/2021"
			gMeses(2,1) = "46,47,48,49,50,51,52,53,54,55,56,57,58"
			gMeses(1,2) = "Trim. Sep-Oct-Nov/2021" 
			gMeses(2,2) = "50,51,52,53,54,55,56,57,58,59,60,61,62"
			gMeses(1,3) = "Trim. Oct-Nov-Dic/2021"
			gMeses(2,3) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
		end if		
		if (sCat = 2) then 
			'response.write "<br>pasoooooooooooooooo"
			'Galletas
			erase gMeses
			redim gMeses(2,6)
			gMeses(1,0) = "Trim. Ene-Feb-Mar/2021"
			gMeses(2,0) = "16,17,18,19,20,21,22,23,24,25,26,27,28"
			gMeses(1,1) = "Trim. Abr-May-Jun/2021"
			gMeses(2,1) = "29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,2) = "Trim. Jun-Jul-Ago/2021"
			gMeses(2,2) = "37,38,39,40,41,42,43,44,45,46,47,48,49"
			gMeses(1,3) = "Trim. Jul-Ago-Sep/2021"
			gMeses(2,3) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
			gMeses(1,4) = "Trim. Ago-Sep-Oct/2021"
			gMeses(2,4) = "46,47,48,49,50,51,52,53,54,55,56,57,58"
			gMeses(1,5) = "Trim. Sep-Oct-Nov/2021"
			gMeses(2,5) = "50,51,52,53,54,55,56,57,58,59,60,61,62"
			gMeses(1,6) = "Trim. Oct-Nov-Dic/2021"
			gMeses(2,6) = "55,56,57,58,59,60,61,62,63,64,65,66,67"
		end if		
		if (sCat = 44) then 
			'Pastas Alimenticias
			erase gMeses
			redim gMeses(2,5)
			gMeses(1,0) = "Mayo 2021"
			gMeses(2,0) = "33,34,35,36"
			gMeses(1,1) = "Julio 2021"
			gMeses(2,1) = "41,42,43,44,45"
			gMeses(1,2) = "Trim. Ene-Feb-Mar/2021"
			gMeses(2,2) = "16,17,18,19,20,21,22,23,24,25,26,27,28"
			gMeses(1,3) = "Trim. Abr-May-Jun/2021"
			gMeses(2,3) = "29,30,31,32,33,34,35,36,37,38,39,40"
			gMeses(1,4) = "Trim. Jul-Ago-Sep/2021"
			gMeses(2,4) = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
			gMeses(1,5) = "Trim. Oct-Nov-Dic/2021"
			gMeses(2,5) = "55,56,57,58,59,60,61,62,63,64,65,66,67"			
			'response.write "paso"
			'response.end
		end if
		if (sCat = 45) then 
			erase gMeses
			redim gMeses(2,1)
			gMeses(1,0) = "Mayo 2021"
			gMeses(2,0) = "33,34,35,36"
			gMeses(1,1) = "Julio 2021"
			gMeses(2,1) = "41,42,43,44,45"
			'response.write "paso"
			'response.end
		end if
	end if

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Indicador, "
	sql = sql & " Abreviatura, "
	sql = sql & " UnidadMedida "
	sql = sql & " FROM "
	sql = sql & " PH_Indicadores "
	sql = sql & " WHERE "
	sql = sql & " Ind_Men = 1 " 
	if sInd <> "" then
		sql = sql & " And Id_Indicador in (" & sInd & ")"
	end if
	if cint(sCat) = 56  or cint(sCat) = 29 or cint(sCat) = 28 or cint(sCat) = 19 or cint(sCat) = 41 or cint(sCat) = 146   then
		'response.write "<br>pasoooooooooooooooo3"
		'15Oct2021
		sql = sql & " and Id_Indicador in(6,10,11,12,15,16,25,26,29,30,32,35,36,37,39) "
		'response.end
	end if
	if (cint(sCat) = 40 or cint(sCat) = 42  or cint(sCat) = 41 or cint(sCat) = 18 or cint(sCat) = 146) and idCliente = 8 then
		'response.write "<br>pasoooooooooooooooo4"
		sql = sql & " and Id_Indicador in(6,10,11,12,15,16,25,26,29,30,32,35,36,37,39) "
	end if
	if (cint(sCat) = 40 or cint(sCat) = 42 ) and idCliente = 36 then
		sql = sql & " and Id_Indicador in(6,10,11,12,15,16,25,26,29,30,32) "
	end if
	if (cint(sCat) = 37 or cint(sCat) = 35 or cint(sCat) = 36) and idCliente = 36 then
		sql = sql & " and Id_Indicador in(6,9,10,11,12,14,15,16,24,25,26,28,29,30,32) "
	end if
	if (cint(sCat) = 41 or cint(sCat) = 146) and idCliente = 19 then 
		'response.write "<br>pasoooooooooooooooo"
		sql = sql & " and Id_Indicador in(6,10,11,12,15,16,25,26,29,30,32) "
	end if
	if (cint(sCat) = 40 or cint(sCat) = 42) and idCliente = 16 then
		sql = sql & " and Id_Indicador in(6,10,11,12,15,16,25,26,29,30,32,35,36,37,39) "
	end if
	'Pharsana
	if (cint(sCat) = 19 or cint(sCat) = 41 or cint(sCat) = 18 or cint(sCat) = 146) and idCliente = 27 then
		'response.write "<br>pasoooooooooooooooo2"
		sql = sql & " and Id_Indicador in(6,10,11,12,15,16,25,26,29,30,32,35,36,37,39) "
	end if
	'Central
	if (cint(sCat) = 72 or cint(sCat) = 106 or cint(sCat) = 97) and idCliente = 20 then
		'response.write "<br>pasoooooooooooooooo"
		sql = sql & " and Id_Indicador in(6,9,10,11,12,14,15,16,24,25,26,28,29,30,32) "
	end if
	'Tunal
	if (cint(sCat) = 12) and idCliente = 21 then
		'response.write "<br>pasoooooooooooooooo"
		sql = sql & " and Id_Indicador in(6,9,10,11,12,14,15,16,24,25,26,28,29,30,32) "
	end if
	'Pepsico Alimento
	if (cint(sCat) = 35) and idCliente = 11 then
		'response.write "<br>pasoooooooooooooooo"
		sql = sql & " and Id_Indicador in(6,9,10,11,12,14,15,16,24,25,26,28,29,30,32) "
	end if
	'Dimassi
	if (cint(sCat) = 41 or cint(sCat) = 146) and idCliente = 8 then
		'response.write "<br>pasoooooooooooooooo"
		sql = sql & " and Id_Indicador in(6,9,10,11,12,14,15,16,24,25,26,28,29,30,32) "
	end if
	'Nestle
	if (cint(sCat) = 9 or cint(sCat) = 8 or cint(sCat) = 5 or cint(sCat) = 6 or cint(sCat) = 14) and idCliente = 17 then
		'response.write "<br>pasoooooooooooooooo"
		sql = sql & " and Id_Indicador in(6,9,10,11,12,14,15,16,24,25,26,28,29,30,32) "
	end if
	'Tunal
	if (cint(sCat) = 93)  and idCliente = 21 then
		'response.write "<br>pasoooooooooooooooo"
		sql = sql & " and Id_Indicador in(6,9,10,12,14,15,24,26,28,29,32) "
	end if
	'Alimex
	if (cint(sCat) = 93)  and idCliente = 28 then
		'response.write "<br>pasoooooooooooooooo"
		sql = sql & " and Id_Indicador in(6,9,10,12,14,15,24,26,28,29,32,35,36,37,39) "
	end if
	'MANPA
	if (cint(sCat) = 54)  and idCliente = 22 then
		'response.write "<br>pasoooooooooooooooo"
		sql = sql & " and Id_Indicador in(6,10,11,12,15,16,25,26,29,30,32) "
	end if
	

	sql = sql & " ORDER BY "
	sql = sql & " Id_Indicador "
	'response.write "<br>372 Combo1:=" & sql
	'response.end 
	rsx1.Open sql ,conexion
	'response.write "<br>Paso 117<br>"
	if rsx1.eof then
		rsx1.close
	else
		gIndicadores = rsx1.GetRows
		rsx1.close
	end if
	'response.write sInd
	'response.end
	
	'09Feb2021-Todo Query
	
	sql = ""
    sql = sql & " SELECT "
	sql = sql & " Id_Area, "
	sql = sql & " Area, "
	sql = sql & " Id_Fabricante, "
	sql = sql & " Fabricante, "
	sql = sql & " Id_Marca, "
	sql = sql & " Marca, "
	sql = sql & " Id_Segmento, "
	sql = sql & " Segmento, "
	sql = sql & " Id_RangoTamano, "
	sql = sql & " RangoTamano "
	sql = sql & " FROM "
	sql = sql & " PH_DataCrudaMensual "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria = " & sCat
	sql = sql & " GROUP BY "
	sql = sql & " Id_Area, "
	sql = sql & " Area, "
	sql = sql & " Id_Fabricante, "
	sql = sql & " Fabricante, "
	sql = sql & " Id_Marca, "
	sql = sql & " Marca, "
	sql = sql & " Id_Segmento, "
	sql = sql & " Segmento, "
	sql = sql & " Id_RangoTamano, "
	sql = sql & " RangoTamano "
	sql = sql & " HAVING "
	sql = sql & " Id_Area = 0 "
	sql = sql & " AND Id_Fabricante = 0 "
	sql = sql & " AND Id_Marca = 0 "
	sql = sql & " AND Id_Segmento = 0 "
	sql = sql & " AND Id_RangoTamano = 0 "
	sql = sql & " ORDER BY "
	sql = sql & " Id_Fabricante "
	'response.write "<br>157 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	'response.write "<br>Paso 164<br>"
	iExiste = 0
	'response.write "<br>84 LLEGO"
	'response.end
	if rsx1.eof then
		rsx1.close
	else
		gProductosTotalNacional = rsx1.GetRows
		rsx1.close
	end if
	
	'response.write "<br>172 sFab:=" & sFab
	sql = ""
    sql = sql & " SELECT "
	sql = sql & " Id_Area, "
	sql = sql & " Area "
	if sFab <> "" then
		sql = sql & " ,Id_Fabricante "	
		sql = sql & " ,Fabricante "		
	end if
	if sMar <> "" then
		sql = sql & " ,Id_Marca "
		sql = sql & " ,Marca "
	end if
	if sSeg <> "" then
		sql = sql & " ,Id_Segmento "
		sql = sql & " ,Segmento "
	end if
	if sRan <> "" then
		sql = sql & " ,Id_RangoTamano "
		sql = sql & " ,RangoTamano "
	end if
	'08Mar2021 - 4
	if sTam <> "" then
		sql = sql & " ,Id_Tamano "
		sql = sql & " ,Tamano "
	end if

	sql = sql & " FROM PH_DataCrudaMensual "

	sql = sql & " WHERE "
	sql = sql & " Id_Categoria = " & sCat

	sql = sql & " GROUP BY "
	sql = sql & " Id_Area, "
	sql = sql & " Area "
	if sFab <> "" then
		sql = sql & " ,Id_Fabricante "	
		sql = sql & " ,Fabricante "		
	end if
	if sMar <> "" then
		sql = sql & " ,Id_Marca "
		sql = sql & " ,Marca "
	end if
	if sSeg <> "" then
		sql = sql & " ,Id_Segmento "
		sql = sql & " ,Segmento "
	end if
	if sRan <> "" then
		sql = sql & " ,Id_RangoTamano "
		sql = sql & " ,RangoTamano "
	end if
	'08Mar2021 - 4
	if sTam <> "" then
		sql = sql & " ,Id_Tamano "
		sql = sql & " ,Tamano "
	end if
	'response.write "<br>335 paso" & sAre
	'response.end
	isw = 0
	if sAre <> "" then
		if isw = 0 then
			sql = sql & " HAVING "
			isw = 1
		else
			sql = sql & " AND "
		end if
		'sql = sql & " Id_Area in (" & sAre & ")"
		if iAre <> 0 and idCliente  = 16 then
			sql = sql & " Id_Area in(2,3,5)"
		else
			if iAre <> 0 and idCliente  = 19 then
				sql = sql & " Id_Area in(2,3)"
			else
				sql = sql & " Id_Area in(" & sAre & ")"
			end if
		end if
		'response.write "<br>310 paso"
		'response.write "<br>313 paso"
	else
		if isw = 0 then
			sql = sql & " HAVING "
			isw = 1
		else
			sql = sql & " AND "
		end if
		
		if TotalArea = "NO" and sFab <> "" and sMar = "" and sRan = "" then 
			sql = sql & " Id_Area = 0 "
		else
			if TotalArea = "NO" and sFab <> "" and sMar <> "" and sRan = "" then 
				sql = sql & " Id_Area = 0 "
			else
				sql = sql & " Id_Area <>0 "
				if idCliente  = 16 then
					sql = sql & " and Id_Area in(2,3,5)"
				end if
				if idCliente  = 19 then
					sql = sql & " and Id_Area in(2,3)"
				end if
			end if
			'sql = sql & " Id_Area <>0 "
		end if
		'response.write "<br>335 paso" & sAre
		'response.write "<br>330 paso"
	end if
	if sFab <> "" then
		if isw = 0 then
			sql = sql & " HAVING "
			isw = 1
		else
			sql = sql & " AND "
		end if
		sql = sql & " Id_Fabricante in (" & sFab & ")"
	else
		'if isw = 0 then
		'	sql = sql & " HAVING "
		'	isw = 1
		'else
		'	sql = sql & " AND "
		'end if
		'sql = sql & " Id_Fabricante <>0 "
		'response.write "<br>222 Paso"
	end if
	
	if sMar <> "" then
		if isw = 0 then
			isw = 1
			sql = sql & " HAVING "
		else
			sql = sql & " AND "
		end if
		sql = sql & " Id_Marca in (" & sMar & ")"
	end if
	if sSeg <> "" then
		if isw = 0 then
			isw = 1
			sql = sql & " HAVING "
		else
			sql = sql & " AND "
		end if
		sql = sql & " Id_Segmento in (" & sSeg & ")"
	end if
	if sRan <> "" then
		if isw = 0 then
			sql = sql & " HAVING "
			isw = 1
		else
			sql = sql & " AND "
		end if
		sql = sql & " Id_RangoTamano in (" & sRan & ")"
	end if
	'08Mar2021 - 9
	if sTam <> "" then
		if isw = 0 then
			sql = sql & " HAVING "
			isw = 1
		else
			sql = sql & " AND "
		end if
		sql = sql & " Id_Tamano in (" & sTam & ")"
	end if

	sql = sql & " ORDER BY "
	sql = sql & " Id_Area "
	if sFab <> "" then
		sql = sql & " ,Id_Fabricante "
	end if
	if sMar <> "" then
		sql = sql & " ,Id_Marca "
	end if
	if sSeg <> "" then
		sql = sql & " ,Id_Segmento "
	end if
	if sRan <> "" then
		sql = sql & " ,Id_RangoTamano "
	end if
	'08Mar2021 - 3
	if sTam <> "" then
		sql = sql & " ,Id_Tamano "
	end if

	'response.write "<br>313 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	'response.write "<br>84 LLEGO" 
	'response.end
	'response.write "<br>Paso 305<br>"
	iExiste = 0
	'response.write "<br>84 LLEGO"
	'response.end
	if rsx1.eof then
		rsx1.close
		%>
		<center>
		<h2>No hay Data para Mostrar</h2>
		</center>
		<div class="limiter">
			<div class="container-table100">
				<div class="wrap-table100">
					<div class="table100 ver1 m-b-110">
							<div class="table100-head">
								<table>
									<thead>
										<tr class="row100 head">
											<th class="cell100 column1 text-left">??rea</th>
											<th class="cell100 column2 text-left">Fabricante</th>
											<th class="cell100 column3 text-center">Marca</th>
											<th class="cell100 column4 text-center">Segmento</th>
											<th class="cell100 column5 text-center">Rango Tama??o</th>
											<th class="cell100 column6 text-center">Tama??o</th>
											<th class="cell100 column7 text-center">Indicador</th>
											<th class="cell100 column8 text-center">UniMed</th>
											<%	
												for iMes = 0 to  ubound(gMeses,2) 
													sx = gMeses(1,iMes) 
													%>
													<th class="cell100 column9 text-center"><%=sx%></th>
													<% 
												next 
												if ubound(gMeses,2) = 0 then
													%>
													<th class="cell100 column10 text-center"></th>
													<% 
												end if
											%>
										</tr>
									</thead>
								</table>
							</div>
					</div>
				</div>
			</div>
		</div>
			
		<%
	else
		'response.write "<br>84 LLEGO"
		'response.end
		gProductos = rsx1.GetRows
		rsx1.close
		%>
		<div class="limiter">
			
			<div class="container-table100">
			
				<div class="wrap-table100">
								
					<div class="table100 ver1 m-b-110">
						
							<div class="table100-head">
							
								<table border=0>
									<thead>
										<tr class="row100 head">
											<th class="cell100 column1 text-left">??rea</th>
											<th class="cell100 column2 text-left">Fabricante</th>
											<th class="cell100 column3 text-center">Marca</th>
											<th class="cell100 column4 text-center">Segmento</th>
											<th class="cell100 column5 text-center">Rango Tama??o</th>
											<th class="cell100 column6 text-center">Tama??o</th>
											<th class="cell100 column7 text-center">Indicador</th>
											<th class="cell100 column8 text-center">UniMed</th>
											<%	
												for iMes = 0 to  ubound(gMeses,2) 
													sx = gMeses(1,iMes) 
													%>
													<th class="cell100 column9 text-center"><%=sx%></th>
													<% 
												next 
												if ubound(gMeses,2) = 0 then
													%>
													<th class="cell100 column10 text-center"></th>
													<% 
												end if
											%>
										</tr>
									</thead>
								</table>
								
							</div>
												
							<div class="table100-body js-pscroll">
								<table border=0>
									<tbody>					
										<% 
										'response.write "<br>499 TotalArea:= " & TotalArea
										if iMostrar = 1 then response.write "<br>sAre := " & sAre
										if iMostrar = 1 then response.write "<br>sFab := " & sFab
										if iMostrar = 1 then response.write "<br>sMar := " & sMar
										if iMostrar = 1 then response.write "<br>sSeg := " & sSeg
										if iMostrar = 1 then response.write "<br>sRan := " & sRan
										if iMostrar = 1 then response.write "<br>sTam := " & sTam
										if TotalArea = "SI" and sFab <> "" and sMar = "" and sRan = "" and sTam = "" then 
											if iMostrar = 1 then response.write "<br>429 PasoLR1"
											'response.write "<br>386 Total Area"
											sAre = "0"
											'iAre = 1
											for iPro = 0 to  ubound(gProductosTotalNacional,2)
												
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductosTotalNacional(1,iPro)
													response.write "</td>"
													
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
														response.write gProductosTotalNacional(3,iPro)
													response.write "</td>"

													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
													response.write "</td>"

													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
													response.write "</td>"

													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
													response.write "</td>"

													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"

													response.write "</td>"

													response.write "<td width=10% class='cell100 column7 text-center'>"
													response.write "</td>"

													response.write "<td width=10% class='cell100 column8 text-center'>"
													response.write "</td>"

													response.write "<td width=10% class='cell100 column9 text-center'>"
													response.write "</td>"
													
													response.write "<td width=10% class='cell100 column10 text-center'>"
													
													response.write "</td>"

												response.write "</tr>"
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td colspan=6 >"
														response.write "</td>"
														response.write "<td width=0% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd) 
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														iAre = gProductosTotalNacional(0,iPro)
														iFab = gProductosTotalNacional(2,iPro)
														'iMar = gProductosTotalNacional(4,iPro)
														'iSeg = gProductosTotalNacional(6,iPro)
														'iRan = gProductosTotalNacional(8,iPro)
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next					
											response.end
										end if

										if TotalArea = "SI" and sFab = "" and sMar <> "" and sSeg = "" and sRan = "" and sTam = "" then 
											if iMostrar = 1 then response.write "<br>521 PasoLR22"
											'response.write "<br>386 Total Area"
											sAre = "0"
											sql = ""
											sql = sql & " SELECT "
											sql = sql & " Id_Area, "
											sql = sql & " Area "
											sql = sql & " ,Id_Marca "
											sql = sql & " ,Marca "
											sql = sql & " FROM PH_DataCrudaMensual "
											sql = sql & " WHERE "
											sql = sql & " Id_Categoria = " & sCat
											'sql = sql & " AND "
											'sql = sql & " Id_Fabricante = 0 "
											sql = sql & " GROUP BY "
											sql = sql & " Id_Area, "
											sql = sql & " Area "
											sql = sql & " ,Id_Marca "
											sql = sql & " ,Marca "
											sql = sql & " HAVING "
											sql = sql & " Id_Area in (" & sAre & ")"
											sql = sql & " AND "
											sql = sql & " Id_Marca in (" & sMar & ")"
											sql = sql & " ORDER BY "
											sql = sql & " Id_Area "
											sql = sql & " ,Id_Marca "
											'response.write "<br>686 sql:=" & sql
											'response.end
											rsx1.Open sql ,conexion
											gProductos = rsx1.GetRows
											rsx1.close

											for iPro = 0 to  ubound(gProductos,2)
												
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductos(1,iPro)
													response.write "</td>"
													
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
													response.write "</td>"

													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
														response.write gProductos(3,iPro)
													response.write "</td>"

													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
														
													response.write "</td>"

													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
													response.write "</td>"

													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"
													response.write "</td>"

													response.write "<td width=10% class='cell100 column7 text-center'>"
													response.write "</td>"

													response.write "<td width=10% class='cell100 column8 text-center'>"
													response.write "</td>"

													response.write "<td width=10% class='cell100 column9 text-center'>"
													response.write "</td>"
													
													response.write "<td width=10% class='cell100 column10 text-center'>"
													
													response.write "</td>"

												response.write "</tr>"
												'sTam1 = sTam
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td colspan=6 >"
														response.write "</td>"
														response.write "<td width=0% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd) 
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														iAre = 0
														iFab = 0
														sFab = ""
														iMar = gProductos(2,iPro)
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next			
											response.end 
										end if

										if TotalArea = "SI" and sFab = "" and sMar = "" and sSeg <> "" and sRan = "" and sTam = "" then 
											if iMostrar = 1 then response.write "<br>521 PasoLR21"
											'response.write "<br>386 Total Area"
											sAre = "0"
											sql = ""
											sql = sql & " SELECT "
											sql = sql & " Id_Area, "
											sql = sql & " Area "
											sql = sql & " ,Id_Segmento "
											sql = sql & " ,Segmento "
											sql = sql & " FROM PH_DataCrudaMensual "
											sql = sql & " WHERE "
											sql = sql & " Id_Categoria = " & sCat
											sql = sql & " GROUP BY "
											sql = sql & " Id_Area, "
											sql = sql & " Area "
											sql = sql & " ,Id_Segmento "
											sql = sql & " ,Segmento "
											sql = sql & " HAVING "
											sql = sql & " Id_Area in (" & sAre & ")"
											sql = sql & " AND "
											sql = sql & " Id_Segmento in (" & sSeg & ")"
											sql = sql & " ORDER BY "
											sql = sql & " Id_Area "
											sql = sql & " ,Id_Segmento "
											'response.write "<br>313 sql:=" & sql
											'response.end
											rsx1.Open sql ,conexion
											gProductos = rsx1.GetRows
											rsx1.close

											for iPro = 0 to  ubound(gProductos,2)
												
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductos(1,iPro)
													response.write "</td>"
													
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
													response.write "</td>"

													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
													response.write "</td>"

													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
														response.write gProductos(3,iPro)
													response.write "</td>"

													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
													response.write "</td>"

													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"
													response.write "</td>"

													response.write "<td width=10% class='cell100 column7 text-center'>"
													response.write "</td>"

													response.write "<td width=10% class='cell100 column8 text-center'>"
													response.write "</td>"

													response.write "<td width=10% class='cell100 column9 text-center'>"
													response.write "</td>"
													
													response.write "<td width=10% class='cell100 column10 text-center'>"
													
													response.write "</td>"

												response.write "</tr>"
												'sTam1 = sTam
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td colspan=6 >"
														response.write "</td>"
														response.write "<td width=0% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd) 
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														iAre = 0
														iSeg = gProductos(2,iPro)
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next			
											response.end 
										end if
										'09Feb2021
										if TotalArea = "SI" and sFab = "" and sMar = "" and sSeg = "" and sRan = "" and sTam = "" then
											if iMostrar = 1 then response.write "<br>521 PasoLR2"
											'response.write "<br>386 Total Area"
											for iPro = 0 to  ubound(gProductosTotalNacional,2)
												
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductosTotalNacional(1,iPro)
													response.write "</td>"
													
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
													response.write "</td>"

													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
													response.write "</td>"

													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
													response.write "</td>"

													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
													response.write "</td>"

													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"
													response.write "</td>"

													response.write "<td width=10% class='cell100 column7 text-center'>"
													response.write "</td>"

													response.write "<td width=10% class='cell100 column8 text-center'>"
													response.write "</td>"

													response.write "<td width=10% class='cell100 column9 text-center'>"
													response.write "</td>"
													
													response.write "<td width=10% class='cell100 column10 text-center'>"
													
													response.write "</td>"

												response.write "</tr>"
												'sTam1 = sTam
												'response.write "<br>749 Paso<br>"
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td colspan=6 >"
														response.write "</td>"
														response.write "<td width=0% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd) 
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														iAre = gProductosTotalNacional(0,iPro)
														iFab = gProductosTotalNacional(2,iPro)
														iMar = gProductosTotalNacional(4,iPro)
														iSeg = gProductosTotalNacional(6,iPro)
														iRan = gProductosTotalNacional(8,iPro)
														
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															'response.write "<br> idSemana:=" & idSemana
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next					
											
										end if
										if TotalArea = "SI" and sFab = "" and sMar = "" and sSeg = "" and sRan = "" and sTam <> "" then
											if iMostrar = 1 then response.write "<br>607 PasoLR42"
											sql = ""
											sql = sql & " SELECT "
											sql = sql & " PH_DataCrudaMensual.Id_Area, "
											sql = sql & " PH_DataCrudaMensual.Area, "
											sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
											sql = sql & " PH_DataCrudaMensual.Fabricante, "
											sql = sql & " PH_DataCrudaMensual.Id_Marca, "
											sql = sql & " PH_DataCrudaMensual.Marca, "
											sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
											sql = sql & " PH_DataCrudaMensual.Segmento, "
											sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
											sql = sql & " PH_DataCrudaMensual.RangoTamano, "
											sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
											sql = sql & " PH_DataCrudaMensual.Tamano "
											sql = sql & " FROM PH_DataCrudaMensual "
											sql = sql & " WHERE "
											sql = sql & " PH_DataCrudaMensual.Id_Categoria = " & sCat
											sql = sql & " GROUP BY PH_DataCrudaMensual.Id_Area, "
											sql = sql & " PH_DataCrudaMensual.Area, "
											sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
											sql = sql & " PH_DataCrudaMensual.Fabricante, "
											sql = sql & " PH_DataCrudaMensual.Id_Marca, "
											sql = sql & " PH_DataCrudaMensual.Marca, "
											sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
											sql = sql & " PH_DataCrudaMensual.Segmento, "
											sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
											sql = sql & " PH_DataCrudaMensual.RangoTamano, "
											sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
											sql = sql & " PH_DataCrudaMensual.Tamano "
											sql = sql & " HAVING "
											sql = sql & " PH_DataCrudaMensual.Id_Area = 0 "
											sql = sql & " AND PH_DataCrudaMensual.Id_Fabricante = 0 "
											sql = sql & " AND PH_DataCrudaMensual.Id_Marca = 0 "
											sql = sql & " AND PH_DataCrudaMensual.Id_Segmento = 0 "
											sql = sql & " AND PH_DataCrudaMensual.Id_RangoTamano = 0 "
											sql = sql & " and PH_DataCrudaMensual.Id_Tamano in (" & sTam & ")"
											'response.write "<br>157 sql:=" & sql
											'response.end
											rsx1.Open sql ,conexion
											'response.write "<br>Paso 164<br>"
											iExiste = 0
											'response.write "<br>84 LLEGO"
											'response.end
											if rsx1.eof then
												rsx1.close
											else
												gProductosTotalNacional = rsx1.GetRows
												rsx1.close
											end if
											
											'response.end
											'response.write "<br>386 valor"
											for iPro = 0 to  ubound(gProductosTotalNacional,2)
												'response.write "<br>428 LLEGO"
												'response.end
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductosTotalNacional(1,iPro)
													response.write "</td>"
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
														
													response.write "</td>"
													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
													response.write "</td>"
													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
													response.write "</td>"
													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
													response.write "</td>"
													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"
														response.write gProductosTotalNacional(11,iPro)
													response.write "</td>"
													response.write "<td colspan=4  class='cell100'>"
													response.write "</td>"
												response.write "</tr>"
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td width=60% colspan=6 >"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd)
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														iAre = 0
														iFab = 0
														iMar = 0
														iSeg = 0
														iRan = 0
														iTam = gProductosTotalNacional(10,iPro)
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next					
											response.end
										else
										
										end if
										if sAre = "" and sFab <> "" and sMar = "" and sRan <> "" and sSeg = "" and sTam = "" then
											if iMostrar = 1 then response.write "<br>6077 PasoLR3350"
											'response.end
											'response.write "<br>386 Todos Blanco"
												sql = ""
												sql = sql & " SELECT "
												sql = sql & " PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												'sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												'sql = sql & " PH_DataCrudaMensual.Marca, "
												'sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												'sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano "
												'sql = sql & " PH_DataCrudaMensual.Id_Tamano "
												'sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " FROM PH_DataCrudaMensual "
												sql = sql & " WHERE "
												sql = sql & " PH_DataCrudaMensual.Id_Categoria = " & sCat
												sql = sql & " GROUP BY PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												'sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												'sql = sql & " PH_DataCrudaMensual.Marca, "
												'sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												'sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano "
												'sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												'sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " HAVING "
												sql = sql & " PH_DataCrudaMensual.Id_Area = 0 "
												sql = sql & " AND PH_DataCrudaMensual.Id_Fabricante in (" & sFab & ")"
												'sql = sql & " AND PH_DataCrudaMensual.Id_Marca = 0 "
												'sql = sql & " AND PH_DataCrudaMensual.Id_Segmento = 0 "
												sql = sql & " AND PH_DataCrudaMensual.Id_RangoTamano in (" & sRan & ")"
												'sql = sql & " and PH_DataCrudaMensual.Id_Tamano in (" & sTam & ")"
												'response.write "<br>157 sql:=" & sql
												'response.end
												rsx1.Open sql ,conexion
												'response.write "<br>Paso 164<br>"
												iExiste = 0
												'response.write "<br>84 LLEGO"
												'response.end
												if rsx1.eof then
													rsx1.close
												else
													gProductos = rsx1.GetRows
													rsx1.close
												end if
											for iPro = 0 to  ubound(gProductos,2)
												'response.write "<br>428 LLEGO"
												'response.end
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductos(1,iPro)
													response.write "</td>"
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
														response.write gProductos(3,iPro)
													response.write "</td>"
													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
														'response.write gProductos(5,iPro)
													response.write "</td>"
													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
														'response.write gProductos(7,iPro)
													response.write "</td>"
													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
														response.write gProductos(5,iPro)
													response.write "</td>"
													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"
														'response.write gProductos(11,iPro)
													response.write "</td>"
													response.write "<td colspan=4  class='cell100'>"
													response.write "</td>"
												response.write "</tr>"
												
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td width=60% colspan=6 >"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd)
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														sAre = ""
														iAre = 0
														iFab = gProductos(2,iPro)
														iMar = 0
														iSeg = 0
														iRan = gProductos(4,iPro)
														'iTam = gProductos(10,iPro)
														'response.end
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next					
											response.end
										else
										
										end if

										if sAre = "" and sFab <> "" and sMar = "" and sRan <> "" and sSeg <> "" and sTam = "" then
											if iMostrar = 1 then response.write "<br>6077 PasoLR3351"
											'response.end
											'response.write "<br>386 Todos Blanco"
												sql = ""
												sql = sql & " SELECT "
												sql = sql & " PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												'sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												'sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano "
												'sql = sql & " PH_DataCrudaMensual.Id_Tamano "
												'sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " FROM PH_DataCrudaMensual "
												sql = sql & " WHERE "
												sql = sql & " PH_DataCrudaMensual.Id_Categoria = " & sCat
												sql = sql & " GROUP BY PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												'sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												'sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano "
												'sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												'sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " HAVING "
												sql = sql & " PH_DataCrudaMensual.Id_Area = 0 "
												sql = sql & " AND PH_DataCrudaMensual.Id_Fabricante in (" & sFab & ")"
												'sql = sql & " AND PH_DataCrudaMensual.Id_Marca = 0 "
												sql = sql & " AND PH_DataCrudaMensual.Id_Segmento in (" & sSeg & ")"
												sql = sql & " AND PH_DataCrudaMensual.Id_RangoTamano in (" & sRan & ")"
												'sql = sql & " and PH_DataCrudaMensual.Id_Tamano in (" & sTam & ")"
												'response.write "<br>157 sql:=" & sql
												'response.end
												rsx1.Open sql ,conexion
												'response.write "<br>Paso 164<br>"
												iExiste = 0
												'response.write "<br>84 LLEGO"
												'response.end
												if rsx1.eof then
													rsx1.close
												else
													gProductos = rsx1.GetRows
													rsx1.close
												end if
											for iPro = 0 to  ubound(gProductos,2)
												'response.write "<br>428 LLEGO"
												'response.end
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductos(1,iPro)
													response.write "</td>"
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
														response.write gProductos(3,iPro)
													response.write "</td>"
													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
														'response.write gProductos(5,iPro)
													response.write "</td>"
													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
														response.write gProductos(5,iPro)
													response.write "</td>"
													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
														response.write gProductos(7,iPro)
													response.write "</td>"
													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"
														'response.write gProductos(11,iPro)
													response.write "</td>"
													response.write "<td colspan=4  class='cell100'>"
													response.write "</td>"
												response.write "</tr>"
												
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td width=60% colspan=6 >"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd)
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														sAre = ""
														iAre = 0
														iFab = gProductos(2,iPro)
														iMar = 0
														iSeg = gProductos(4,iPro)
														iRan = gProductos(6,iPro)
														'iTam = gProductos(10,iPro)
														'response.end
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next					
											response.end
										else
										
										end if

										if sAre = "" and sFab = "" and sMar <> "" and sRan <> "" and sSeg = "" and sTam = "" then
											if iMostrar = 1 then response.write "<br>6077 PasoLR3352"
											'response.end
											'response.write "<br>386 Todos Blanco"
												sql = ""
												sql = sql & " SELECT "
												sql = sql & " PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												'sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												'sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano "
												'sql = sql & " PH_DataCrudaMensual.Id_Tamano "
												'sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " FROM PH_DataCrudaMensual "
												sql = sql & " WHERE "
												sql = sql & " PH_DataCrudaMensual.Id_Categoria = " & sCat
												sql = sql & " GROUP BY PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												'sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												'sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano "
												'sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												'sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " HAVING "
												sql = sql & " PH_DataCrudaMensual.Id_Area = 0 "
												sql = sql & " AND PH_DataCrudaMensual.Id_Fabricante = 0"
												sql = sql & " AND PH_DataCrudaMensual.Id_Marca in (" & sMar & ")"
												'sql = sql & " AND PH_DataCrudaMensual.Id_Segmento in (" & sSeg & ")"
												sql = sql & " AND PH_DataCrudaMensual.Id_RangoTamano in (" & sRan & ")"
												'sql = sql & " and PH_DataCrudaMensual.Id_Tamano in (" & sTam & ")"
												'response.write "<br>157 sql:=" & sql
												'response.end
												rsx1.Open sql ,conexion
												'response.write "<br>Paso 164<br>"
												iExiste = 0
												'response.write "<br>84 LLEGO"
												'response.end
												if rsx1.eof then
													rsx1.close
												else
													gProductos = rsx1.GetRows
													rsx1.close
												end if
											for iPro = 0 to  ubound(gProductos,2)
												'response.write "<br>428 LLEGO"
												'response.end
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductos(1,iPro)
													response.write "</td>"
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
														response.write gProductos(3,iPro)
													response.write "</td>"
													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
														response.write gProductos(5,iPro)
													response.write "</td>"
													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
														'response.write gProductos(5,iPro)
													response.write "</td>"
													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
														response.write gProductos(7,iPro)
													response.write "</td>"
													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"
														'response.write gProductos(11,iPro)
													response.write "</td>"
													response.write "<td colspan=4  class='cell100'>"
													response.write "</td>"
												response.write "</tr>"
												
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td width=60% colspan=6 >"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd)
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														sAre = ""
														iAre = 0
														iFab = 0
														iMar = gProductos(4,iPro)
														'iSeg = 0
														iRan = gProductos(6,iPro)
														'iTam = gProductos(10,iPro)
														'response.end
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next					
											response.end
										else
										
										end if

										if sAre = "" and sFab <> "" and sMar <> "" and sRan <> "" and sSeg = "" and sTam = "" then
											if iMostrar = 1 then response.write "<br>6077 PasoLR3353"
											'response.end
											'response.write "<br>386 Todos Blanco"
												sql = ""
												sql = sql & " SELECT "
												sql = sql & " PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												'sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												'sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano "
												'sql = sql & " PH_DataCrudaMensual.Id_Tamano "
												'sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " FROM PH_DataCrudaMensual "
												sql = sql & " WHERE "
												sql = sql & " PH_DataCrudaMensual.Id_Categoria = " & sCat
												sql = sql & " GROUP BY PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												'sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												'sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano "
												'sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												'sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " HAVING "
												sql = sql & " PH_DataCrudaMensual.Id_Area = 0 "
												sql = sql & " AND PH_DataCrudaMensual.Id_Fabricante in (" & sFab & ")"
												sql = sql & " AND PH_DataCrudaMensual.Id_Marca in (" & sMar & ")"
												'sql = sql & " AND PH_DataCrudaMensual.Id_Segmento in (" & sSeg & ")"
												sql = sql & " AND PH_DataCrudaMensual.Id_RangoTamano in (" & sRan & ")"
												'sql = sql & " and PH_DataCrudaMensual.Id_Tamano in (" & sTam & ")"
												'response.write "<br>157 sql:=" & sql
												'response.end
												rsx1.Open sql ,conexion
												'response.write "<br>Paso 164<br>"
												iExiste = 0
												'response.write "<br>84 LLEGO"
												'response.end
												if rsx1.eof then
													rsx1.close
												else
													gProductos = rsx1.GetRows
													rsx1.close
												end if
											for iPro = 0 to  ubound(gProductos,2)
												'response.write "<br>428 LLEGO"
												'response.end
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductos(1,iPro)
													response.write "</td>"
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
														response.write gProductos(3,iPro)
													response.write "</td>"
													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
														response.write gProductos(5,iPro)
													response.write "</td>"
													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
														'response.write gProductos(5,iPro)
													response.write "</td>"
													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
														response.write gProductos(7,iPro)
													response.write "</td>"
													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"
														'response.write gProductos(11,iPro)
													response.write "</td>"
													response.write "<td colspan=4  class='cell100'>"
													response.write "</td>"
												response.write "</tr>"
												
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td width=60% colspan=6 >"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd)
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														sAre = ""
														iAre = 0
														iFab = gProductos(2,iPro)
														iMar = gProductos(4,iPro)
														'iSeg = 0
														iRan = gProductos(6,iPro)
														'iTam = gProductos(10,iPro)
														'response.end
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next					
											response.end
										else
										
										end if

										if sAre = "" and sFab = "" and sMar = "" and sRan <> "" and sSeg <> "" and sTam = "" then
											if iMostrar = 1 then response.write "<br>6077 PasoLR3331"
											'response.end
											'response.write "<br>386 Todos Blanco"
												sql = ""
												sql = sql & " SELECT "
												sql = sql & " PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano "
												'sql = sql & " PH_DataCrudaMensual.Id_Tamano "
												'sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " FROM PH_DataCrudaMensual "
												sql = sql & " WHERE "
												sql = sql & " PH_DataCrudaMensual.Id_Categoria = " & sCat
												sql = sql & " GROUP BY PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano "
												'sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												'sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " HAVING "
												sql = sql & " PH_DataCrudaMensual.Id_Area = 0 "
												sql = sql & " AND PH_DataCrudaMensual.Id_Fabricante = 0 "
												sql = sql & " AND PH_DataCrudaMensual.Id_Marca = 0 "
												sql = sql & " AND PH_DataCrudaMensual.Id_Segmento in (" & sSeg & ")"
												sql = sql & " AND PH_DataCrudaMensual.Id_RangoTamano in (" & sRan & ")"
												'sql = sql & " and PH_DataCrudaMensual.Id_Tamano in (" & sTam & ")"
												'response.write "<br>157 sql:=" & sql
												'response.end
												rsx1.Open sql ,conexion
												'response.write "<br>Paso 164<br>"
												iExiste = 0
												'response.write "<br>84 LLEGO"
												'response.end
												if rsx1.eof then
													rsx1.close
												else
													gProductos = rsx1.GetRows
													rsx1.close
												end if
											for iPro = 0 to  ubound(gProductos,2)
												'response.write "<br>428 LLEGO"
												'response.end
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductos(1,iPro)
													response.write "</td>"
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
														response.write gProductos(3,iPro)
													response.write "</td>"
													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
														response.write gProductos(5,iPro)
													response.write "</td>"
													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
														response.write gProductos(7,iPro)
													response.write "</td>"
													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
														response.write gProductos(9,iPro)
													response.write "</td>"
													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"
														'response.write gProductos(11,iPro)
													response.write "</td>"
													response.write "<td colspan=4  class='cell100'>"
													response.write "</td>"
												response.write "</tr>"
												
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td width=60% colspan=6 >"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd)
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														sAre = ""
														iAre = 0
														iFab = 0
														iMar = 0
														iSeg = gProductos(6,iPro)
														iRan = gProductos(8,iPro)
														'iTam = gProductos(10,iPro)
														'response.end
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next					
											response.end
										else
										
										end if

										if sAre = "" and sFab = "" and sMar <> "" and sRan <> "" and sSeg <> "" and sTam = "" then
											if iMostrar = 1 then response.write "<br>6077 PasoLR3332"
											'response.end
											'response.write "<br>386 Todos Blanco"
												sql = ""
												sql = sql & " SELECT "
												sql = sql & " PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano "
												'sql = sql & " PH_DataCrudaMensual.Id_Tamano "
												'sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " FROM PH_DataCrudaMensual "
												sql = sql & " WHERE "
												sql = sql & " PH_DataCrudaMensual.Id_Categoria = " & sCat
												sql = sql & " GROUP BY PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano "
												'sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												'sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " HAVING "
												sql = sql & " PH_DataCrudaMensual.Id_Area = 0 "
												sql = sql & " AND PH_DataCrudaMensual.Id_Fabricante = 0 "
												sql = sql & " AND PH_DataCrudaMensual.Id_Marca in (" & sMar & ")"
												sql = sql & " AND PH_DataCrudaMensual.Id_Segmento in (" & sSeg & ")"
												sql = sql & " AND PH_DataCrudaMensual.Id_RangoTamano in (" & sRan & ")"
												'sql = sql & " and PH_DataCrudaMensual.Id_Tamano in (" & sTam & ")"
												'response.write "<br>157 sql:=" & sql
												'response.end
												rsx1.Open sql ,conexion
												'response.write "<br>Paso 164<br>"
												iExiste = 0
												'response.write "<br>84 LLEGO"
												'response.end
												if rsx1.eof then
													rsx1.close
												else
													gProductos = rsx1.GetRows
													rsx1.close
												end if
											for iPro = 0 to  ubound(gProductos,2)
												'response.write "<br>428 LLEGO"
												'response.end
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductos(1,iPro)
													response.write "</td>"
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
														response.write gProductos(3,iPro)
													response.write "</td>"
													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
														response.write gProductos(5,iPro)
													response.write "</td>"
													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
														response.write gProductos(7,iPro)
													response.write "</td>"
													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
														response.write gProductos(9,iPro)
													response.write "</td>"
													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"
														'response.write gProductos(11,iPro)
													response.write "</td>"
													response.write "<td colspan=4  class='cell100'>"
													response.write "</td>"
												response.write "</tr>"
												
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td width=60% colspan=6 >"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd)
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														sAre = ""
														iAre = 0
														iFab = 0
														iMar = gProductos(4,iPro)
														iSeg = gProductos(6,iPro)
														iRan = gProductos(8,iPro)
														'iTam = gProductos(10,iPro)
														'response.end
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next					
											response.end
										else
										
										end if
										
										if sAre = "" and sFab = "" and sMar <> "" and sRan = "" and sSeg <> "" and sTam = "" then
											if iMostrar = 1 then response.write "<br>6077 PasoLR3333"
											'response.end
											'response.write "<br>386 Todos Blanco"
												sql = ""
												sql = sql & " SELECT "
												sql = sql & " PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento "
												'sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												'sql = sql & " PH_DataCrudaMensual.RangoTamano "
												'sql = sql & " PH_DataCrudaMensual.Id_Tamano "
												'sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " FROM PH_DataCrudaMensual "
												sql = sql & " WHERE "
												sql = sql & " PH_DataCrudaMensual.Id_Categoria = " & sCat
												sql = sql & " GROUP BY PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento "
												'sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												'sql = sql & " PH_DataCrudaMensual.RangoTamano "
												'sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												'sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " HAVING "
												sql = sql & " PH_DataCrudaMensual.Id_Area = 0 "
												sql = sql & " AND PH_DataCrudaMensual.Id_Fabricante = 0 "
												sql = sql & " AND PH_DataCrudaMensual.Id_Marca in (" & sMar & ")"
												sql = sql & " AND PH_DataCrudaMensual.Id_Segmento in (" & sSeg & ")"
												'sql = sql & " AND PH_DataCrudaMensual.Id_RangoTamano in (" & sRan & ")"
												'sql = sql & " and PH_DataCrudaMensual.Id_Tamano in (" & sTam & ")"
												'response.write "<br>157 sql:=" & sql
												'response.end
												rsx1.Open sql ,conexion
												'response.write "<br>Paso 164<br>"
												iExiste = 0
												'response.write "<br>84 LLEGO"
												'response.end
												if rsx1.eof then
													rsx1.close
												else
													gProductos = rsx1.GetRows
													rsx1.close
												end if
											for iPro = 0 to  ubound(gProductos,2)
												'response.write "<br>428 LLEGO"
												'response.end
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductos(1,iPro)
													response.write "</td>"
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
														response.write gProductos(3,iPro)
													response.write "</td>"
													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
														response.write gProductos(5,iPro)
													response.write "</td>"
													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
														response.write gProductos(7,iPro)
													response.write "</td>"
													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
														'response.write gProductos(9,iPro)
													response.write "</td>"
													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"
														'response.write gProductos(11,iPro)
													response.write "</td>"
													response.write "<td colspan=4  class='cell100'>"
													response.write "</td>"
												response.write "</tr>"
												
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td width=60% colspan=6 >"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd)
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														sAre = ""
														iAre = 0
														iFab = 0
														iMar = gProductos(4,iPro)
														iSeg = gProductos(6,iPro)
														'iRan = gProductos(8,iPro)
														'iTam = gProductos(10,iPro)
														'response.end
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next					
											response.end
										else
										
										end if
										
										if sAre = "" and sFab = "" and sMar = "" and sRan <> "" and sSeg = "" and sTam = "" then
											if iMostrar = 1 then response.write "<br>6077 PasoLR3334"
											'response.end
											'response.write "<br>386 Todos Blanco"
												sql = ""
												sql = sql & " SELECT "
												sql = sql & " PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano "
												'sql = sql & " PH_DataCrudaMensual.Id_Tamano "
												'sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " FROM PH_DataCrudaMensual "
												sql = sql & " WHERE "
												sql = sql & " PH_DataCrudaMensual.Id_Categoria = " & sCat
												sql = sql & " GROUP BY PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano "
												'sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												'sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " HAVING "
												sql = sql & " PH_DataCrudaMensual.Id_Area = 0 "
												sql = sql & " AND PH_DataCrudaMensual.Id_Fabricante = 0 "
												sql = sql & " AND PH_DataCrudaMensual.Id_Marca = 0 "
												sql = sql & " AND PH_DataCrudaMensual.Id_Segmento = 0 "
												sql = sql & " AND PH_DataCrudaMensual.Id_RangoTamano in (" & sRan & ")"
												'sql = sql & " and PH_DataCrudaMensual.Id_Tamano in (" & sTam & ")"
												'response.write "<br>157 sql:=" & sql
												'response.end
												rsx1.Open sql ,conexion
												'response.write "<br>Paso 164<br>"
												iExiste = 0
												'response.write "<br>84 LLEGO"
												'response.end
												if rsx1.eof then
													rsx1.close
												else
													gProductos = rsx1.GetRows
													rsx1.close
												end if
											for iPro = 0 to  ubound(gProductos,2)
												'response.write "<br>428 LLEGO"
												'response.end
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductos(1,iPro)
													response.write "</td>"
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
														response.write gProductos(3,iPro)
													response.write "</td>"
													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
														response.write gProductos(5,iPro)
													response.write "</td>"
													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
														response.write gProductos(7,iPro)
													response.write "</td>"
													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
														response.write gProductos(9,iPro)
													response.write "</td>"
													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"
														'response.write gProductos(11,iPro)
													response.write "</td>"
													response.write "<td colspan=4  class='cell100'>"
													response.write "</td>"
												response.write "</tr>"
												
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td width=60% colspan=6 >"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd)
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														sAre = ""
														iAre = 0
														iFab = 0
														iMar = 0
														iSeg = 0
														iRan = gProductos(8,iPro)
														'iTam = gProductos(10,iPro)
														'response.end
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next					
											response.end
										else
										
										end if
										
										if sAre = "" and sFab <> "" and sMar <> "" and sRan <> "" and sSeg <> "" and sTam <> "" then
											if iMostrar = 1 then response.write "<br>6077 PasoLR333"
											'response.end
											'response.write "<br>386 Todos Blanco"
												sql = ""
												sql = sql & " SELECT "
												sql = sql & " PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " FROM PH_DataCrudaMensual "
												sql = sql & " WHERE "
												sql = sql & " PH_DataCrudaMensual.Id_Categoria = " & sCat
												sql = sql & " GROUP BY PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " HAVING "
												sql = sql & " PH_DataCrudaMensual.Id_Area = 0 "
												sql = sql & " AND PH_DataCrudaMensual.Id_Fabricante in (" & sFab & ")"
												sql = sql & " AND PH_DataCrudaMensual.Id_Marca in (" & sMar & ")"
												sql = sql & " AND PH_DataCrudaMensual.Id_Segmento in (" & sSeg & ")"
												sql = sql & " AND PH_DataCrudaMensual.Id_RangoTamano in (" & sRan & ")"
												sql = sql & " and PH_DataCrudaMensual.Id_Tamano in (" & sTam & ")"
												'response.write "<br>157 sql:=" & sql
												'response.end
												rsx1.Open sql ,conexion
												'response.write "<br>Paso 164<br>"
												iExiste = 0
												'response.write "<br>84 LLEGO"
												'response.end
												if rsx1.eof then
													rsx1.close
												else
													gProductos = rsx1.GetRows
													rsx1.close
												end if
											for iPro = 0 to  ubound(gProductos,2)
												'response.write "<br>428 LLEGO"
												'response.end
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductos(1,iPro)
													response.write "</td>"
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
														response.write gProductos(3,iPro)
													response.write "</td>"
													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
														response.write gProductos(5,iPro)
													response.write "</td>"
													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
														response.write gProductos(7,iPro)
													response.write "</td>"
													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
														response.write gProductos(9,iPro)
													response.write "</td>"
													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"
														response.write gProductos(11,iPro)
													response.write "</td>"
													response.write "<td colspan=4  class='cell100'>"
													response.write "</td>"
												response.write "</tr>"
												
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td width=60% colspan=6 >"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd)
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														sAre = ""
														iAre = 0
														iFab = gProductos(2,iPro)
														iMar = gProductos(4,iPro)
														iSeg = gProductos(6,iPro)
														iRan = gProductos(8,iPro)
														iTam = gProductos(10,iPro)
														'response.end
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next					
											response.end
										else
										if sAre = "" and sFab <> "" and sMar <> "" and sRan <> "" and sSeg <> "" and sTam = "" then
											if iMostrar = 1 then response.write "<br>6077 PasoLR3338"
											'response.end
											'response.write "<br>386 Todos Blanco"
												sql = ""
												sql = sql & " SELECT "
												sql = sql & " PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano "
												'sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												'sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " FROM PH_DataCrudaMensual "
												sql = sql & " WHERE "
												sql = sql & " PH_DataCrudaMensual.Id_Categoria = " & sCat
												sql = sql & " GROUP BY PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano "
												'sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												'sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " HAVING "
												sql = sql & " PH_DataCrudaMensual.Id_Area = 0 "
												sql = sql & " AND PH_DataCrudaMensual.Id_Fabricante in (" & sFab & ")"
												sql = sql & " AND PH_DataCrudaMensual.Id_Marca in (" & sMar & ")"
												sql = sql & " AND PH_DataCrudaMensual.Id_Segmento in (" & sSeg & ")"
												sql = sql & " AND PH_DataCrudaMensual.Id_RangoTamano in (" & sRan & ")"
												'sql = sql & " and PH_DataCrudaMensual.Id_Tamano in (" & sTam & ")"
												'response.write "<br>157 sql:=" & sql
												'response.end
												rsx1.Open sql ,conexion
												'response.write "<br>Paso 164<br>"
												iExiste = 0
												'response.write "<br>84 LLEGO"
												'response.end
												if rsx1.eof then
													rsx1.close
												else
													gProductos = rsx1.GetRows
													rsx1.close
												end if
											for iPro = 0 to  ubound(gProductos,2)
												'response.write "<br>428 LLEGO"
												'response.end
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductos(1,iPro)
													response.write "</td>"
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
														response.write gProductos(3,iPro)
													response.write "</td>"
													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
														response.write gProductos(5,iPro)
													response.write "</td>"
													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
														response.write gProductos(7,iPro)
													response.write "</td>"
													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
														response.write gProductos(9,iPro)
													response.write "</td>"
													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"
														'response.write gProductos(11,iPro)
													response.write "</td>"
													response.write "<td colspan=4  class='cell100'>"
													response.write "</td>"
												response.write "</tr>"
												
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td width=60% colspan=6 >"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd)
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														sAre = ""
														iAre = 0
														iFab = gProductos(2,iPro)
														iMar = gProductos(4,iPro)
														iSeg = gProductos(6,iPro)
														iRan = gProductos(8,iPro)
														iTam = 0 'gProductos(10,iPro)
														'response.end
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next					
											response.end
										else
										
										end if
										
										
										end if
										if sAre = "" and sFab <> "" and sMar = "" and sRan <> "" and sSeg <> "" and sTam <> "" then
											if iMostrar = 1 then response.write "<br>6077 PasoLR334"
											'response.end
											'response.write "<br>386 Todos Blanco"
												sql = ""
												sql = sql & " SELECT "
												sql = sql & " PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " FROM PH_DataCrudaMensual "
												sql = sql & " WHERE "
												sql = sql & " PH_DataCrudaMensual.Id_Categoria = " & sCat
												sql = sql & " GROUP BY PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " HAVING "
												sql = sql & " PH_DataCrudaMensual.Id_Area = 0 "
												sql = sql & " AND PH_DataCrudaMensual.Id_Fabricante in (" & sFab & ")"
												'sql = sql & " AND PH_DataCrudaMensual.Id_Marca = 0 "
												sql = sql & " AND PH_DataCrudaMensual.Id_Segmento in (" & sSeg & ")"
												sql = sql & " AND PH_DataCrudaMensual.Id_RangoTamano in (" & sRan & ")"
												sql = sql & " and PH_DataCrudaMensual.Id_Tamano in (" & sTam & ")"
												'response.write "<br>157 sql:=" & sql
												'response.end
												rsx1.Open sql ,conexion
												'response.write "<br>Paso 164<br>"
												iExiste = 0
												'response.write "<br>84 LLEGO"
												'response.end
												if rsx1.eof then
													rsx1.close
												else
													gProductos = rsx1.GetRows
													rsx1.close
												end if
											for iPro = 0 to  ubound(gProductos,2)
												'response.write "<br>428 LLEGO"
												'response.end
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductos(1,iPro)
													response.write "</td>"
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
														response.write gProductos(3,iPro)
													response.write "</td>"
													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
														response.write gProductos(5,iPro)
													response.write "</td>"
													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
														response.write gProductos(7,iPro)
													response.write "</td>"
													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
														response.write gProductos(9,iPro)
													response.write "</td>"
													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"
														response.write gProductos(11,iPro)
													response.write "</td>"
													response.write "<td colspan=4  class='cell100'>"
													response.write "</td>"
												response.write "</tr>"
												
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td width=60% colspan=6 >"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd)
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														sAre = ""
														iAre = 0
														iFab = gProductos(2,iPro)
														iMar = gProductos(4,iPro)
														iSeg = gProductos(6,iPro)
														iRan = gProductos(8,iPro)
														iTam = gProductos(10,iPro)
														'response.end
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next					
											response.end
										else
										
										end if

										if sAre = "" and sFab <> "" and sMar = "" and sRan = "" and sSeg <> "" and sTam <> "" then
											if iMostrar = 1 then response.write "<br>6077 PasoLR335"
											'response.end
											'response.write "<br>386 Todos Blanco"
												sql = ""
												sql = sql & " SELECT "
												sql = sql & " PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " FROM PH_DataCrudaMensual "
												sql = sql & " WHERE "
												sql = sql & " PH_DataCrudaMensual.Id_Categoria = " & sCat
												sql = sql & " GROUP BY PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " HAVING "
												sql = sql & " PH_DataCrudaMensual.Id_Area = 0 "
												sql = sql & " AND PH_DataCrudaMensual.Id_Fabricante in (" & sFab & ")"
												sql = sql & " AND PH_DataCrudaMensual.Id_Segmento in (" & sSeg & ")"
												sql = sql & " and PH_DataCrudaMensual.Id_Tamano in (" & sTam & ")"
												'response.write "<br>157 sql:=" & sql
												'response.end
												rsx1.Open sql ,conexion
												'response.write "<br>Paso 164<br>"
												iExiste = 0
												'response.write "<br>84 LLEGO"
												'response.end
												if rsx1.eof then
													rsx1.close
												else
													gProductos = rsx1.GetRows
													rsx1.close
												end if
											for iPro = 0 to  ubound(gProductos,2)
												'response.write "<br>428 LLEGO"
												'response.end
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductos(1,iPro)
													response.write "</td>"
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
														response.write gProductos(3,iPro)
													response.write "</td>"
													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
														response.write gProductos(5,iPro)
													response.write "</td>"
													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
														response.write gProductos(7,iPro)
													response.write "</td>"
													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
														response.write gProductos(9,iPro)
													response.write "</td>"
													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"
														response.write gProductos(11,iPro)
													response.write "</td>"
													response.write "<td colspan=4  class='cell100'>"
													response.write "</td>"
												response.write "</tr>"
												
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td width=60% colspan=6 >"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd)
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														sAre = ""
														iAre = 0
														iFab = gProductos(2,iPro)
														iMar = gProductos(4,iPro)
														iSeg = gProductos(6,iPro)
														iRan = gProductos(8,iPro)
														iTam = gProductos(10,iPro)
														'response.end
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next					
											response.end
										else
										
										end if
										if sAre = "" and sFab <> "" and sMar <> "" and sRan = "" and sSeg = "" and sTam <> "" then
											if iMostrar = 1 then response.write "<br>6077 PasoLR336"
											'response.end
											'response.write "<br>386 Todos Blanco"
												sql = ""
												sql = sql & " SELECT "
												sql = sql & " PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												'sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												'sql = sql & " PH_DataCrudaMensual.Segmento, "
												'sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												'sql = sql & " PH_DataCrudaMensual.RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " FROM PH_DataCrudaMensual "
												sql = sql & " WHERE "
												sql = sql & " PH_DataCrudaMensual.Id_Categoria = " & sCat
												sql = sql & " GROUP BY PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " HAVING "
												sql = sql & " PH_DataCrudaMensual.Id_Area = 0 "
												sql = sql & " AND PH_DataCrudaMensual.Id_Fabricante in (" & sFab & ")"
												sql = sql & " AND PH_DataCrudaMensual.Id_Marca in (" & sMar & ")"
												sql = sql & " and PH_DataCrudaMensual.Id_Tamano in (" & sTam & ")"
												'response.write "<br>1477 sql:=" & sql
												'response.end
												rsx1.Open sql ,conexion
												'response.write "<br>Paso 164<br>"
												iExiste = 0
												'response.write "<br>84 LLEGO"
												'response.end
												if rsx1.eof then
													rsx1.close
												else
													gProductos = rsx1.GetRows
													rsx1.close
												end if
											for iPro = 0 to  ubound(gProductos,2)
												'response.write "<br>428 LLEGO"
												'response.end
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductos(1,iPro)
													response.write "</td>"
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
														response.write gProductos(3,iPro)
													response.write "</td>"
													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
														response.write gProductos(5,iPro)
													response.write "</td>"
													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
														'response.write gProductos(7,iPro)
													response.write "</td>"
													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
														'response.write gProductos(9,iPro)
													response.write "</td>"
													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"
														response.write gProductos(7,iPro)
													response.write "</td>"
													response.write "<td colspan=4  class='cell100'>"
													response.write "</td>"
												response.write "</tr>"
												
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td width=60% colspan=6 >"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd)
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														sAre = ""
														iAre = 0
														iFab = gProductos(2,iPro)
														iMar = gProductos(4,iPro)
														iSeg = 0 'gProductos(6,iPro)
														iRan = 0 'gProductos(8,iPro)
														iTam = gProductos(6,iPro)
														'response.end
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next					
											response.end
										else
										
										end if
										if sAre <> "" and sFab <> "" and sMar <> "" and sRan = "" and sSeg <> "" and sTam <> "" then
											if iMostrar = 1 then response.write "<br>6077 PasoLR337"
											'response.end
											'response.write "<br>386 Todos Blanco"
												sql = ""
												sql = sql & " SELECT "
												sql = sql & " PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " FROM PH_DataCrudaMensual "
												sql = sql & " WHERE "
												sql = sql & " PH_DataCrudaMensual.Id_Categoria = " & sCat
												sql = sql & " GROUP BY PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " HAVING "
												sql = sql & " PH_DataCrudaMensual.Id_Area  in (" & sAre & ")"
												sql = sql & " AND PH_DataCrudaMensual.Id_Fabricante in (" & sFab & ")"
												sql = sql & " AND PH_DataCrudaMensual.Id_Marca in (" & sMar & ")"
												sql = sql & " AND PH_DataCrudaMensual.Id_Segmento in (" & sSeg & ")"
												sql = sql & " and PH_DataCrudaMensual.Id_Tamano in (" & sTam & ")"
												'response.write "<br>1477 sql:=" & sql
												'response.end
												rsx1.Open sql ,conexion
												'response.write "<br>Paso 164<br>"
												iExiste = 0
												'response.write "<br>84 LLEGO"
												'response.end
												if rsx1.eof then
													rsx1.close
												else
													gProductos = rsx1.GetRows
													rsx1.close
												end if
											for iPro = 0 to  ubound(gProductos,2)
												'response.write "<br>428 LLEGO"
												'response.end
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductos(1,iPro)
													response.write "</td>"
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
														response.write gProductos(3,iPro)
													response.write "</td>"
													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
														response.write gProductos(5,iPro)
													response.write "</td>"
													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
														response.write gProductos(7,iPro)
													response.write "</td>"
													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
														response.write gProductos(9,iPro)
													response.write "</td>"
													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"
														response.write gProductos(11,iPro)
													response.write "</td>"
													response.write "<td colspan=4  class='cell100'>"
													response.write "</td>"
												response.write "</tr>"
												
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td width=60% colspan=6 >"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd)
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														sAre = ""
														iAre = gProductos(0,iPro)
														iFab = gProductos(2,iPro)
														iMar = gProductos(4,iPro)
														iSeg = gProductos(6,iPro)
														iRan = gProductos(8,iPro)
														iTam = gProductos(10,iPro)
														'response.end
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next					
											response.end
										else
										
										end if
										
										if sAre = "" and sFab <> "" and sMar <> "" and sRan = "" and sSeg <> "" and sTam <> "" then
											if iMostrar = 1 then response.write "<br>6077 PasoLR338"
											'response.end
											'response.write "<br>386 Todos Blanco"
												sql = ""
												sql = sql & " SELECT "
												sql = sql & " PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " FROM PH_DataCrudaMensual "
												sql = sql & " WHERE "
												sql = sql & " PH_DataCrudaMensual.Id_Categoria = " & sCat
												sql = sql & " GROUP BY PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " HAVING "
												sql = sql & " PH_DataCrudaMensual.Id_Area = 0 "
												sql = sql & " AND PH_DataCrudaMensual.Id_Fabricante in (" & sFab & ")"
												sql = sql & " AND PH_DataCrudaMensual.Id_Marca in (" & sMar & ")"
												sql = sql & " AND PH_DataCrudaMensual.Id_Segmento in (" & sSeg & ")"
												sql = sql & " and PH_DataCrudaMensual.Id_Tamano in (" & sTam & ")"
												'response.write "<br>1477 sql:=" & sql
												'response.end
												rsx1.Open sql ,conexion
												'response.write "<br>Paso 164<br>"
												iExiste = 0
												'response.write "<br>84 LLEGO"
												'response.end
												if rsx1.eof then
													rsx1.close
												else
													gProductos = rsx1.GetRows
													rsx1.close
												end if
											for iPro = 0 to  ubound(gProductos,2)
												'response.write "<br>428 LLEGO"
												'response.end
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductos(1,iPro)
													response.write "</td>"
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
														response.write gProductos(3,iPro)
													response.write "</td>"
													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
														response.write gProductos(5,iPro)
													response.write "</td>"
													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
														response.write gProductos(7,iPro)
													response.write "</td>"
													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
														response.write gProductos(9,iPro)
													response.write "</td>"
													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"
														response.write gProductos(11,iPro)
													response.write "</td>"
													response.write "<td colspan=4  class='cell100'>"
													response.write "</td>"
												response.write "</tr>"
												
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td width=60% colspan=6 >"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd)
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														sAre = ""
														iAre = 0
														iFab = gProductos(2,iPro)
														iMar = gProductos(4,iPro)
														iSeg = gProductos(6,iPro)
														iRan = gProductos(8,iPro)
														iTam = gProductos(10,iPro)
														'response.end
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next					
											response.end
										else
										
										end if
										'response.end

										if sAre = "" and sFab = "" and sMar <> "" and sRan = "" and sSeg <> "" and sTam <> "" then
											if iMostrar = 1 then response.write "<br>6077 PasoLR339"
											'response.end
											'response.write "<br>386 Todos Blanco"
												sql = ""
												sql = sql & " SELECT "
												sql = sql & " PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " FROM PH_DataCrudaMensual "
												sql = sql & " WHERE "
												sql = sql & " PH_DataCrudaMensual.Id_Categoria = " & sCat
												sql = sql & " GROUP BY PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " HAVING "
												sql = sql & " PH_DataCrudaMensual.Id_Area = 0 "
												'sql = sql & " AND PH_DataCrudaMensual.Id_Fabricante in (" & sFab & ")"
												sql = sql & " AND PH_DataCrudaMensual.Id_Marca in (" & sMar & ")"
												sql = sql & " AND PH_DataCrudaMensual.Id_Segmento in (" & sSeg & ")"
												sql = sql & " and PH_DataCrudaMensual.Id_Tamano in (" & sTam & ")"
												'response.write "<br>1477 sql:=" & sql
												'response.end
												rsx1.Open sql ,conexion
												'response.write "<br>Paso 164<br>"
												iExiste = 0
												'response.write "<br>84 LLEGO"
												'response.end
												if rsx1.eof then
													rsx1.close
												else
													gProductos = rsx1.GetRows
													rsx1.close
												end if
											for iPro = 0 to  ubound(gProductos,2)
												'response.write "<br>428 LLEGO"
												'response.end
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductos(1,iPro)
													response.write "</td>"
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
														response.write gProductos(3,iPro)
													response.write "</td>"
													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
														response.write gProductos(5,iPro)
													response.write "</td>"
													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
														response.write gProductos(7,iPro)
													response.write "</td>"
													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
														response.write gProductos(9,iPro)
													response.write "</td>"
													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"
														response.write gProductos(11,iPro)
													response.write "</td>"
													response.write "<td colspan=4  class='cell100'>"
													response.write "</td>"
												response.write "</tr>"
												
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td width=60% colspan=6 >"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd)
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														sAre = ""
														iAre = 0
														iFab = gProductos(2,iPro)
														iMar = gProductos(4,iPro)
														iSeg = gProductos(6,iPro)
														iRan = gProductos(8,iPro)
														iTam = gProductos(10,iPro)
														'response.end
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next					
											response.end
										else
										
										end if
										if sAre = "" and sFab = "" and sMar = "" and sRan = "" and sSeg <> "" and sTam <> "" then
											if iMostrar = 1 then response.write "<br>6077 PasoLR340"
											'response.end
											'response.write "<br>386 Todos Blanco"
												sql = ""
												sql = sql & " SELECT "
												sql = sql & " PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " FROM PH_DataCrudaMensual "
												sql = sql & " WHERE "
												sql = sql & " PH_DataCrudaMensual.Id_Categoria = " & sCat
												sql = sql & " GROUP BY PH_DataCrudaMensual.Id_Area, "
												sql = sql & " PH_DataCrudaMensual.Area, "
												sql = sql & " PH_DataCrudaMensual.Id_Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Fabricante, "
												sql = sql & " PH_DataCrudaMensual.Id_Marca, "
												sql = sql & " PH_DataCrudaMensual.Marca, "
												sql = sql & " PH_DataCrudaMensual.Id_Segmento, "
												sql = sql & " PH_DataCrudaMensual.Segmento, "
												sql = sql & " PH_DataCrudaMensual.Id_RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.RangoTamano, "
												sql = sql & " PH_DataCrudaMensual.Id_Tamano, "
												sql = sql & " PH_DataCrudaMensual.Tamano "
												sql = sql & " HAVING "
												sql = sql & " PH_DataCrudaMensual.Id_Area = 0 "
												'sql = sql & " AND PH_DataCrudaMensual.Id_Fabricante in (" & sFab & ")"
												'sql = sql & " AND PH_DataCrudaMensual.Id_Marca in (" & sMar & ")"
												sql = sql & " AND PH_DataCrudaMensual.Id_Segmento in (" & sSeg & ")"
												sql = sql & " and PH_DataCrudaMensual.Id_Tamano in (" & sTam & ")"
												'response.write "<br>1477 sql:=" & sql
												'response.end
												rsx1.Open sql ,conexion
												'response.write "<br>Paso 164<br>"
												iExiste = 0
												'response.write "<br>84 LLEGO"
												'response.end
												if rsx1.eof then
													rsx1.close
												else
													gProductos = rsx1.GetRows
													rsx1.close
												end if
											for iPro = 0 to  ubound(gProductos,2)
												'response.write "<br>428 LLEGO"
												'response.end
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductos(1,iPro)
													response.write "</td>"
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
														response.write gProductos(3,iPro)
													response.write "</td>"
													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
														response.write gProductos(5,iPro)
													response.write "</td>"
													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
														response.write gProductos(7,iPro)
													response.write "</td>"
													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
														response.write gProductos(9,iPro)
													response.write "</td>"
													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"
														response.write gProductos(11,iPro)
													response.write "</td>"
													response.write "<td colspan=4  class='cell100'>"
													response.write "</td>"
												response.write "</tr>"
												
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td width=60% colspan=6 >"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd)
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														sAre = ""
														iAre = 0
														iFab = gProductos(2,iPro)
														iMar = gProductos(4,iPro)
														iSeg = gProductos(6,iPro)
														iRan = gProductos(8,iPro)
														iTam = gProductos(10,iPro)
														'response.end
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next					
											response.end
										else
										
										end if
										
										if sAre <> "" and sFab = "" and sMar = "" and sRan = "" and sSeg = "" and sTam = "" then
											if iMostrar = 1 then response.write "<br>6077 PasoLR3"
											'response.end
											'response.write "<br>386 Todos Blanco"
											for iPro = 0 to  ubound(gProductos,2)
												'response.write "<br>428 LLEGO"
												'response.end
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductos(1,iPro)
													response.write "</td>"
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
														
													response.write "</td>"
													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
													response.write "</td>"
													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
													response.write "</td>"
													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
													response.write "</td>"
													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"

													response.write "</td>"
													response.write "<td colspan=4  class='cell100'>"
													response.write "</td>"
												response.write "</tr>"
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td width=60% colspan=6 >"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd)
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														iAre = gProductos(0,iPro)
														iFab = 0
														iMar = 0
														iSeg = 0
														iRan = 0
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next					
											
										else
										
										end if
										if sAre <> "" And sFab = "" and sMar = "" and sSeg <> "" and sRan = "" then
											if iMostrar = 1 then response.write "<br>6070 PasoLR31"
											'response.end
											'response.write "<br>770 sAre:= " & sAre
											for iPro = 0 to  ubound(gProductos,2)
												'response.write "<br>428 LLEGO"
												'response.end
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductos(1,iPro)
													response.write "</td>"
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
														
													response.write "</td>"
													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
													response.write "</td>"
													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
														response.write gProductos(3,iPro)
													response.write "</td>"
													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
													response.write "</td>"
													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"

													response.write "</td>"
													response.write "<td colspan=4  class='cell100'>"
													response.write "</td>"
												response.write "</tr>"
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td width=60% colspan=6 >"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd)
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														iAre = gProductos(0,iPro)
														iFab = 0
														iMar = 0
														iSeg = gProductos(2,iPro)
														iRan = 0
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next					
											
										else
										
										end if

										if sAre <> "" And sFab = "" and sMar = "" and sSeg <> "" and sRan <> "" then
											if iMostrar = 1 then response.write "<br>6070 PasoLR311"
											'response.end
											'response.write "<br>770 sAre:= " & sAre
											for iPro = 0 to  ubound(gProductos,2)
												'response.write "<br>428 LLEGO"
												'response.end
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductos(1,iPro)
													response.write "</td>"
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
														
													response.write "</td>"
													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
													response.write "</td>"
													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
														response.write gProductos(3,iPro)
													response.write "</td>"
													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
														response.write gProductos(5,iPro)
													response.write "</td>"
													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"

													response.write "</td>"
													response.write "<td colspan=4  class='cell100'>"
													response.write "</td>"
												response.write "</tr>"
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td width=60% colspan=6 >"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd)
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														iAre = gProductos(0,iPro)
														iFab = 0
														iMar = 0
														iSeg = gProductos(2,iPro)
														iRan = gProductos(4,iPro)
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next					
											response.end
										else
										
										end if
										
										if sFab = "" and sMar = "" and sSeg = "" and sRan = "" and sTam <> "" then
											if iMostrar = 1 then response.write "<br>607 PasoLR41"
											'response.end
											'response.write "<br>386 Todos Blanco"
											for iPro = 0 to  ubound(gProductos,2)
												'response.write "<br>428 LLEGO"
												'response.end
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductos(1,iPro)
													response.write "</td>"
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
														
													response.write "</td>"
													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"
													response.write "</td>"
													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"
													response.write "</td>"
													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"
													response.write "</td>"
													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"
														response.write gProductos(3,iPro)
													response.write "</td>"
													response.write "<td colspan=4  class='cell100'>"
													response.write "</td>"
												response.write "</tr>"
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td width=60% colspan=6 >"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd)
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														iAre = gProductos(0,iPro)
														iFab = 0
														iMar = 0
														iSeg = 0
														iRan = 0
														iTam = gProductos(2,iPro)
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next					
											response.end
										else
										
										end if

										'26Ene2021-Todo el IF
										if TotalFabricante = "SI" then 
											if iMostrar = 1 then response.write "<br>682 PasoLR4"
											for iPro = 0 to  ubound(gProductosTotal,2)
												
												response.write "<tr class='row100 body'>"
													'Area
													response.write "<td width=10% class='cell100 column1'>"
														response.write gProductosTotal(1,iPro)
													response.write "</td>"
													'Fabricante
													response.write "<td width=10% class='cell100 column2'>"
														response.write gProductosTotal(3,iPro)
													response.write "</td>"
													'Marca
													response.write "<td width=10% class='cell100 column3 text-center'>"

													response.write "</td>"
													'Segmento
													response.write "<td width=10% class='cell100 column4 text-center'>"

													response.write "</td>"
													'Rango
													response.write "<td width=10% class='cell100 column5 text-center'>"

													response.write "</td>"
													'Tama??o
													response.write "<td width=10% class='cell100 column6 text-center'>"

													response.write "</td>"
													response.write "<td colspan=4  class='cell100'>"
													response.write "</td>"
												response.write "</tr>"
												for iInd = 0 to  ubound(gIndicadores,2)
													response.write "<tr class='row100 body'>"
														response.write "<td width=60% colspan=6 >"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column7 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd)
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=10% class='cell100 column8 text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														iAre = gProductosTotal(0,iPro)
														iFab = gProductosTotal(2,iPro)
														iMar = gProductosTotal(4,iPro)
														iSeg = gProductosTotal(6,iPro)
														'response.write "<br>Ind = " & Indicador
														for iMes = 0 to  ubound(gMeses,2) 
															'idSemana = "16,17,18,19"
															idSemana = gMeses(2,iMes)
															TotalDias = 28
															CalcularIndicador
															response.write "<td width=10% class='cell100 column9 text-right'>"
																response.write Valor
															response.write "</td>"
														next 
														if ubound(gMeses,2) = 0 then
															response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write "</td>"
														end if
													response.write "</tr>"
												next
											next					
										end if
										
										if sFab = "" and sMar = "" and sRan = "" and sTam = "" then
										else
										if iMostrar = 1 then response.write "<br>1 PasoLR5"
										icontador = 0
										for iPro = 0 to  ubound(gProductos,2)
											
											icontador = icontador + 1
											if icontador > 100 then
												Response.flush 
												icontador = 0
											end if
											
											'response.write "<br>579 Paso"
											response.write "<tr class='row100 body'>"
												'Area
												response.write "<td width=10% class='cell100 column1'>"
													response.write gProductos(1,iPro)
												response.write "</td>"
												ix = 1
												'Fabricante
												response.write "<td width=10% class='cell100 column2'>"
													if sFab <> "" then
														ix = ix + 2
														response.write gProductos(ix,iPro)
													end if
												response.write "</td>"
												'Marca
												response.write "<td width=10% class='cell100 column3'>"
													if sMar <> "" then
														ix = ix + 2
														response.write gProductos(ix,iPro)
													end if
												response.write "</td>"
												'Segmento
												response.write "<td width=10% class='cell100 column4 text-center'>"
													if sSeg <> "" then
														ix = ix + 2
														response.write gProductos(ix,iPro)
													end if
												response.write "</td>"
												'Rango
												response.write "<td width=10% class='cell100 column5 text-center'>"
													if sRan <> "" then
														ix = ix + 2
														response.write gProductos(ix,iPro)
													end if
												response.write "</td>"
												'Tama??o
												response.write "<td width=10% class='cell100 column6 text-center'>"
													if sTam <> "" then
														ix = ix + 2
														response.write gProductos(ix,iPro)
													end if
												response.write "</td>"
											response.write "</tr>"
											response.write "<td colspan=4  class='cell100'>"
											response.write "</td>"
											
											for iInd = 0 to  ubound(gIndicadores,2)
												'Contador = Contador + 1
												'response.write "<br>965:= " & Contador

												response.write "<tr class='row100 body'>"
													response.write "<td width=60% colspan=6 >"
													response.write "</td>"
													response.write "<td width=10% class='cell100 column7 text-center'>"
														response.write "<b>"
														'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
														response.write gIndicadores(1,iInd) 
														response.write "</b>"
													response.write "</td>"
													response.write "<td width=10%  class='text-center'>"
														response.write "<b>"
														response.write gIndicadores(2,iInd)
														response.write "</b>"
													response.write "</td>"
													Indicador = gIndicadores(0,iInd)
													iAre = gProductos(0,iPro)
													ix = 0
													if sFab <> "" then 
														ix = ix + 2
														iFab = gProductos(ix,iPro)
													end if
													if sMar <> "" then
														ix = ix + 2
														iMar = gProductos(ix,iPro)
													end if
													if sSeg <> "" then
														ix = ix + 2
														iSeg = gProductos(ix,iPro)
													end if
													if sRan <> "" then
														ix = ix + 2
														iRan = gProductos(ix,iPro)
													end if
													if sTam <> "" then
														ix = ix + 2
														iTam = gProductos(ix,iPro)
													end if
													'response.write "<br>Ind = " & Indicador
													for iMes = 0 to  ubound(gMeses,2) 
														'idSemana = "16,17,18,19"
														idSemana = gMeses(2,iMes)
														TotalDias = 28
														CalcularIndicador
														response.write "<td width=10% class='cell100 column9 text-right'>"
															response.write Valor
														response.write "</td>"
													next 
													'response.write "<br>1395 iAre:= " & iAre
													'response.write "<br>1395 iFab:= " & iFab
													if ubound(gMeses,2) = 0 then
														response.write "<td width=10% class='cell100 column9 text-right'>"
														response.write "</td>"
													end if
												response.write "</tr>"
											next
										next					
										end if
										%>
									</tbody>
								</table>
							</div>
					</div>
				
				</div>
				
			</div>
			
		</div>
		<%
	end if
	
Sub CalcularIndicador
	
	Select Case Indicador
		Case 1 'CompVol 
			if iFab <> 0 then
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Tamano, "
				sql = sql & " Cantidad "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria = " & sCat
				sql = sql & " And Id_Fabricante = " & iFab 
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
				end if
				sql = sql & " And id_Semana in( " & idSemana & ")"
			else
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Tamano, "
				sql = sql & " Cantidad "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria = " & sCat
				sql = sql & " And Id_Fabricante = 0 "
				sql = sql & " And Id_Marca = 0"
				sql = sql & " And Id_Segmento = 0"
				sql = sql & " And Id_RangoTamano = 0"
				sql = sql & " And id_Semana in( " & idSemana & ")"
			end if
			'response.write "<br>36 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			'response.write "<br>257 LLEGO" 
			'response.end
			if rsx1.eof then
				rsx1.close
				Valor = 0
				Valor = FormatNumber(Valor,2)
			else
				'response.write "<br>84 LLEGO"
				'response.end
				gDatos1 = rsx1.GetRows
				rsx1.close
				Valor = 0
				for iDat = 0 to ubound(gDatos1,2)
					Valor = Valor + (cdbl(gDatos1(0,iDat)) *cdbl(gDatos1(1,iDat)))
				next
				Valor = FormatNumber((Valor)/1000,2)
			end if
			'response.write "<br> Calculo Indicador 1:= " & Valor
		
		Case 2 'CompVal
			if iFab <> 0 then
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Cantidad, "
				sql = sql & " Precio_Producto, "
				sql = sql & " Dolar, "
				sql = sql & " Precio_producto/Dolar AS PrecioDolar, "
				sql = sql & " (Precio_producto/Dolar)*Cantidad AS ComprasValor "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria =  " & sCat
				sql = sql & " And Id_Fabricante = " & iFab 
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
				end if
				sql = sql & " And id_Semana in( " & idSemana & ")"
			else
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Cantidad, "
				sql = sql & " Precio_Producto, "
				sql = sql & " Dolar, "
				sql = sql & " Precio_producto/Dolar AS PrecioDolar, "
				sql = sql & " (Precio_producto/Dolar)*Cantidad AS ComprasValor "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria = " & sCat
				sql = sql & " And Id_Fabricante = 0 "
				sql = sql & " And Id_Marca = 0 "
				sql = sql & " And Id_Segmento = 0 "
				sql = sql & " And Id_RangoTamano = 0 "
				sql = sql & " And id_Semana in( " & idSemana & ")"
			end if
			'response.write "<br>36 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			'response.write "<br>257 LLEGO" 
			'response.end
			if rsx1.eof then
				rsx1.close
				Valor = 0
				Valor = FormatNumber(Valor,2)
			else
				'response.write "<br>84 LLEGO"
				'response.end
				gDatos1 = rsx1.GetRows
				rsx1.close
				Valor = 0
				for iDat = 0 to ubound(gDatos1,2)
					Valor = Valor + cdbl(gDatos1(4,iDat))
				next
				Valor = FormatNumber(Valor,2) 
			end if
			
		Case 3 'CompUni
			if iFab <> 0 then
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Cantidad "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria =  " & sCat
				sql = sql & " And Id_Fabricante = " & iFab 
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
				end if
				sql = sql & " And id_Semana in( " & idSemana & ")"
			else
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Cantidad "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria =  " & sCat
				sql = sql & " And Id_Fabricante = 0 "
				sql = sql & " And Id_Marca = 0 "
				sql = sql & " And Id_Segmento = 0 "
				sql = sql & " And Id_RangoTamano = 0 "
				sql = sql & " And id_Semana in( " & idSemana & ")"
			end if
			
			'response.write "<br>36 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			'response.write "<br>257 LLEGO" 
			'response.end
			if rsx1.eof then
				rsx1.close
				Valor = 0
				Valor = FormatNumber(Valor,2)
			else
				'response.write "<br>84 LLEGO"
				'response.end
				gDatos1 = rsx1.GetRows
				rsx1.close
				Valor = 0
				Cantidad = 0
				for iDat = 0 to ubound(gDatos1,2)
					Cantidad = Cantidad + gDatos1(0,iDat)
				next
				Valor = FormatNumber(Cantidad,0)
			end if
		Case 4 'CompAct
			if iFab <> 0 then
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Cantidad, "
				sql = sql & " Id_Consumo "
				sql = sql & " FROM PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria =  " & sCat
				sql = sql & " AND Id_Fabricante = " & iFab
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
				end if
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				sql = sql & " Cantidad, "
				sql = sql & " Id_Consumo "
			else
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Cantidad, "
				sql = sql & " Id_Consumo "
				sql = sql & " FROM PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria =  " & sCat
				sql = sql & " AND Id_Fabricante = 0 "
				sql = sql & " AND Id_Marca = 0"
				sql = sql & " AND Id_Segmento = 0 "
				sql = sql & " AND Id_RangoTamano  = 0"
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				sql = sql & " Cantidad, "
				sql = sql & " Id_Consumo "
			
			end if
			'response.write "<br>36 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			'response.write "<br>257 LLEGO" 
			'response.end
			if rsx1.eof then
				rsx1.close
				Valor = 0
				Valor = FormatNumber(Valor,2)
			else
				'response.write "<br>84 LLEGO"
				'response.end
				gDatos1 = rsx1.GetRows
				rsx1.close
				Valor = 0
				Cantidad = 0
				for iDat = 0 to ubound(gDatos1,2)
					Cantidad = Cantidad + 1
				next
				Valor = FormatNumber(Cantidad,0)
			end if
		Case 5 'PenNum
			'response.write "<br>84 LLEGO"
			'response.end
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Id_Hogar AS Total "
			sql = sql & " FROM "
			sql = sql & " PH_DataCrudaMensual "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria =  " & sCat
			sql = sql & " And Id_Fabricante = " & iFab
			if sMar <> "" then 
				sql = sql & " And Id_Marca = " & iMar 
			end if
			if sSeg <> "" then 
				sql = sql & " And Id_Segmento = " & iSeg 
			end if
			if sRan <> "" then 
				sql = sql & " And Id_RangoTamano = " & iRan
			end if
			sql = sql & " And id_Semana in( " & idSemana & ")"
			sql = sql & " GROUP BY "
			sql = sql & " PH_DataCrudaMensual.Id_Hogar "
			'response.write "<br>36 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			'response.write "<br>257 LLEGO" 
			'response.end
			if rsx1.eof then
				rsx1.close
				Valor = 0
				Valor = FormatNumber(Valor,2)
			else
				'response.write "<br>84 LLEGO"
				'response.end
				gDatos1 = rsx1.GetRows
				rsx1.close
				Valor = 0
				Cantidad = 0
				for iDat = 0 to ubound(gDatos1,2)
					'Cantidad = gDatos1(0,0)
					Cantidad = Cantidad + 1
				next
				Valor = FormatNumber(Cantidad,0)
				'response.write "<br> Calculo Indicador 5:= " & Valor
			end if
		
		Case 6 'Penetracion
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Id_Hogar AS Total "
			sql = sql & " FROM "
			sql = sql & " PH_DataCrudaMensual "
			sql = sql & " WHERE "
			'sql = sql & " Id_Categoria =  " & sCat
			sql = sql & " id_Semana in( " & idSemana & ")"
			if iAre <> 0 then
				sql = sql & " and Id_Area = " & iAre
				'response.write "<br>1431 Paso"
			end if
			sql = sql & " GROUP BY "
			sql = sql & " PH_DataCrudaMensual.Id_Hogar "
			'response.write "<br>970 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			'response.write "<br>257 LLEGO" 
			'response.end
			if rsx1.eof then
				rsx1.close
				Valor = 0
				Valor = FormatNumber(Valor,2)
			else
				'response.write "<br>84 LLEGO"
				'response.end
				gDatos1 = rsx1.GetRows
				rsx1.close
				Valor = 0
				Cantidad = 0
				for iDat = 0 to ubound(gDatos1,2)
					'Cantidad = gDatos1(0,0)
					Cantidad = Cantidad + 1
				next
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Id_Hogar AS Total "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Area = " & iAre
				if sFab <> "" then 
					sql = sql & " And Id_Fabricante = " & iFab 
				end if
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
				end if
				if sTam <> "" then 
					sql = sql & " And Id_Tamano = " & iTam
				end if
				sql = sql & " and Id_Categoria =  " & sCat
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				sql = sql & " PH_DataCrudaMensual.Id_Hogar "
				'response.write "<br>1013 sql:=" & sql & "<br>"
				'response.end
				rsx1.Open sql ,conexion
				if rsx1.eof then
					rsx1.close
					Total = 0
					Valor = 0
					Valor = FormatNumber(Valor,2)
				else
					'response.write "<br>84 LLEGO"
					'response.end
					gDatos1 = rsx1.GetRows
					rsx1.close
					Total = 0
					for iDat = 0 to ubound(gDatos1,2)
						Total = Total + 1
					next
					'response.write "<br>1030 Cantidad:" & Cantidad
					'response.write "<br> Total:" & Total & "<br>"
					Valor = FormatNumber(((Total*100)/Cantidad),2)
				end if
			end if
		

		Case 7 'PenPonVol 
			Valor = 0
			
		Case 8 'PenPonVal
			Valor = 0

		Case 9 'CompraMedHog
			if iAre <> 0 then
				'response.write "Paso1"
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Tamano, "
				sql = sql & " Cantidad "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria = " & sCat
				sql = sql & " And Id_Area = " & iAre
				if sFab <> "" then 
					sql = sql & " And Id_Fabricante = " & iFab 
				end if
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
				end if
				if sTam <> "" then 
					sql = sql & " And Id_Tamano = " & iTam
				end if
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				sql = sql & " Tamano, "
				sql = sql & " Cantidad, "
				sql = sql & " Id_Consumo, "
				sql = sql & " CodigoBarra "

				'response.write "<br>1095 sql:=" & sql
				paso = 1
			else
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Id_Consumo, "
				sql = sql & " Producto, "
				sql = sql & " Tamano, "
				sql = sql & " Cantidad "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria = " & sCat
				sql = sql & " And Id_Area = 0 " 
				if sFab <> "" then 
					sql = sql & " And Id_Fabricante = " & iFab 
				end if
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
				end if
				if sTam <> "" then 
					sql = sql & " And Id_Tamano = " & iTam
				end if
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				sql = sql & " Id_Consumo, "
				sql = sql & " Producto, "
				sql = sql & " Tamano, "
				sql = sql & " Cantidad "
				paso = 2
				'response.write "<br>1908 sql:=" & sql
				'response.end
			end if
			
			'response.end
			rsx1.Open sql ,conexion
			'response.write "<br>257 LLEGO" 
			'response.end
			if rsx1.eof then
				rsx1.close
				Valor = 0
				Valor = FormatNumber(Valor,2)
			else
				'response.write "<br>1141 LLEGO"
				'response.end
				gDatos1 = rsx1.GetRows
				rsx1.close
				Valor = 0
				for iDat = 0 to ubound(gDatos1,2)
					if paso = 1 Then 
						Valor = Valor + (cdbl(gDatos1(0,iDat)) *cdbl(gDatos1(1,iDat)))
					else
						Valor = Valor + (cdbl(gDatos1(2,iDat)) *cdbl(gDatos1(3,iDat)))
					end if
				next
				if cint(sCat) <> 124 then
					Indicador1 = FormatNumber((Valor)/1000,2)
				else
					Indicador1 = FormatNumber((Valor),2)
				end if
				'response.write "<br>1149 LLEGO"
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Id_Hogar AS Total "
				'31Mar2021-1
				'sql = sql & " Id_Consumo "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria =  "  & sCat
				sql = sql & " And Id_Area = " & iAre
				if sFab <> "" then 
					sql = sql & " And Id_Fabricante = " & iFab
				else
					'sql = sql & " And Id_Fabricante = 0 "
				end if
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
				else
					'sql = sql & " And Id_Marca = 0 "   
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
				else  
					'sql = sql & " And Id_Segmento = 0 " 
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
				else
					'sql = sql & " And Id_RangoTamano = 0 " 
				end if
				if sTam <> "" then 
					sql = sql & " And Id_Tamano = " & iTam
				end if
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				sql = sql & " Id_Hogar "  
				'31Mar2021-1  
				'sql = sql & " Id_Consumo "
				'response.write "<br>1973 sql:=" & sql
				'response.end
				rsx1.Open sql ,conexion
				'response.write "<br>257 LLEGO" 
				'response.end
				if rsx1.eof then
					rsx1.close
				else
					'response.write "<br>84 LLEGO"
					'response.end
					gDatos1 = rsx1.GetRows
					rsx1.close
				end if
				Indicador5 = 0
				for iDat = 0 to ubound(gDatos1,2)
					'Cantidad = gDatos1(0,0)
					Indicador5 = Indicador5 + 1
				next
				Valor = (cdbl(Indicador1) / cdbl(Indicador5))
				'response.write "<br><br>772 Indicador1=" & Indicador1
				'response.write "<br>773 Indicador5=" & Indicador5
				'response.write "<br>"
				Valor = FormatNumber(Valor,2)
			end if

		Case 10 'GastMedHog Calcular 
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Cantidad, "
				sql = sql & " Precio_Producto, "
				sql = sql & " Dolar, "
				sql = sql & " Precio_producto/Dolar AS PrecioDolar, "
				sql = sql & " (Precio_producto/Dolar)*Cantidad AS ComprasValor "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria =  " & sCat
				sql = sql & " And Id_Area = " & iAre 
				if sFab <> "" then 
					sql = sql & " And Id_Fabricante = " & iFab 
				end if
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
				end if
				if sTam <> "" then 
					sql = sql & " And Id_Tamano = " & iTam
				end if
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				sql = sql & " Cantidad, "
				sql = sql & " Precio_Producto, "
				sql = sql & " Dolar, "
				sql = sql & " Id_Consumo, "
				sql = sql & " Producto "
			'response.write "<br>36 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			'response.write "<br>257 LLEGO" 
			'response.end
			if rsx1.eof then
				rsx1.close
				Valor = 0
				Valor = FormatNumber(Valor,2)
			else
				'response.write "<br>84 LLEGO"
				'response.end
				gDatos1 = rsx1.GetRows
				rsx1.close
				Indicador2 = 0
				for iDat = 0 to ubound(gDatos1,2)
					Indicador2 = Indicador2 + cdbl(gDatos1(4,iDat))
				next
				
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Id_Hogar AS Total "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria =  "  & sCat
				sql = sql & " and Id_Area =  "  & iAre
				if sFab <> "" then 
					sql = sql & " And Id_Fabricante = " & iFab
				end if
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
				end if
				if sTam <> "" then 
					sql = sql & " And Id_Tamano = " & iTam
				end if
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				sql = sql & " PH_DataCrudaMensual.Id_Hogar "
				'response.write "<br>36 sql:=" & sql
				'response.end
				rsx1.Open sql ,conexion
				'response.write "<br>257 LLEGO" 
				'response.end
				if rsx1.eof then
					rsx1.close
				else
					'response.write "<br>84 LLEGO"
					'response.end
					gDatos1 = rsx1.GetRows
					rsx1.close
				end if
				Indicador5 = 0
				for iDat = 0 to ubound(gDatos1,2)
					'Cantidad = gDatos1(0,0)
					Indicador5 = Indicador5 + 1
				next
				Valor = (cdbl(Indicador2) / cdbl(Indicador5))
				'response.write "<br>36 Indicador2=" & Indicador2
				'response.write "<br>36 Indicador5=" & Indicador5
				Valor = FormatNumber(Valor,2) 
			end if

		Case 11 'UnidCompHog Calcular 
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Cantidad "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria =  " & sCat
				sql = sql & " And Id_Area =  " & iAre
				paso = 0
				if sFab <> "" then 
					sql = sql & " And Id_Fabricante = " & iFab 
					paso = 1
				end if
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
					paso = 1
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
					paso = 1
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
					paso = 1
				end if
				if sTam <> "" then 
					sql = sql & " And Id_Tamano = " & iTam
					paso = 1
				end if
				sql = sql & " And id_Semana in( " & idSemana & ")"
				'if paso = 0 then 
					sql = sql & " GROUP BY "
					sql = sql & " Cantidad, "
					sql = sql & " Id_Consumo, "
					sql = sql & " CodigoBarra "
				'end if
			'response.write "<br>36 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			'response.write "<br>257 LLEGO" 
			'response.end
			if rsx1.eof then
				rsx1.close
				Valor = 0
				Valor = FormatNumber(Valor,2)
			else
				'response.write "<br>84 LLEGO"
				'response.end
				gDatos1 = rsx1.GetRows
				rsx1.close
				Indicador3 = 0
				for iDat = 0 to ubound(gDatos1,2)
					Indicador3 = Indicador3 + cdbl(gDatos1(0,iDat))
				next
				
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Id_Hogar AS Total "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria = " & sCat
				sql = sql & " and Id_Area = " & iAre
				if sFab <> "" then 
					sql = sql & " And Id_Fabricante = " & iFab
				end if
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
				end if
				if sTam <> "" then 
					sql = sql & " And Id_Tamano = " & iTam
					paso = 1
				end if
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				sql = sql & " PH_DataCrudaMensual.Id_Hogar "
				'response.write "<br>36 sql:=" & sql
				'response.end
				rsx1.Open sql ,conexion
				'response.write "<br>257 LLEGO" 
				'response.end
				if rsx1.eof then
					rsx1.close
				else
					'response.write "<br>84 LLEGO"
					'response.end
					gDatos1 = rsx1.GetRows
					rsx1.close
				end if
				Indicador5 = 0
				for iDat = 0 to ubound(gDatos1,2)
					'Cantidad = gDatos1(0,0)
					Indicador5 = Indicador5 + 1
				next
				Valor = (cdbl(Indicador3) / cdbl(Indicador5))
				'response.write "<br>36 Indicador1=" & Indicador2
				'response.write "<br>36 Indicador5=" & Indicador5
				Valor = FormatNumber(Valor,2)
			end if

		Case 12 'ActCompHog
			sql = ""
			sql = sql & " SELECT "
			'sql = sql & " Cantidad, "
			sql = sql & " Id_Consumo "
			sql = sql & " FROM PH_DataCrudaMensual "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = " & sCat
			sql = sql & " and Id_Area = " & iAre
			if sFab <> "" then 
				sql = sql & " AND Id_Fabricante = " & iFab
			end if
			if sMar <> "" then 
				sql = sql & " And Id_Marca = " & iMar 
			end if
			if sSeg <> "" then 
				sql = sql & " And Id_Segmento = " & iSeg 
			end if
			if sRan <> "" then 
				sql = sql & " And Id_RangoTamano = " & iRan
			end if
			if sTam <> "" then 
				sql = sql & " And Id_Tamano = " & iTam
			end if
			sql = sql & " And id_Semana in( " & idSemana & ")"
			sql = sql & " GROUP BY "
			'sql = sql & " Cantidad, "
			sql = sql & " Id_Consumo "
			'response.write "<br><br>36 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			'response.write "<br>257 LLEGO" 
			'response.end
			if rsx1.eof then
				rsx1.close
				Valor = 0
				Valor = FormatNumber(Valor,2)
			else
				'response.write "<br>84 LLEGO"
				'response.end
				gDatos1 = rsx1.GetRows
				rsx1.close
				Indicador4 = 0
				for iDat = 0 to ubound(gDatos1,2)
					Indicador4 = Indicador4 + 1
				next
				
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Id_Hogar AS Total "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria = " & sCat
				sql = sql & " And Id_Area = " & iAre
				if sFab <> "" then 
					sql = sql & " And Id_Fabricante = " & iFab
				end if
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
				end if
				if sTam <> "" then 
					sql = sql & " And Id_Tamano = " & iTam
				end if
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				sql = sql & " PH_DataCrudaMensual.Id_Hogar "
				'response.write "<br>36 sql:=" & sql
				'response.end
				rsx1.Open sql ,conexion
				'response.write "<br>257 LLEGO" 
				'response.end
				if rsx1.eof then
					rsx1.close
				else
					'response.write "<br>84 LLEGO"
					'response.end
					gDatos1 = rsx1.GetRows
					rsx1.close
				end if
				Indicador5 = 0
				for iDat = 0 to ubound(gDatos1,2)
					'Cantidad = gDatos1(0,0)
					Indicador5 = Indicador5 + 1
				next
				Valor = (cdbl(Indicador4) / cdbl(Indicador5))
				'response.write "<br>36 Indicador1=" & Indicador2
				'response.write "<br>36 Indicador5=" & Indicador5
				Valor = FormatNumber(Valor,2)
			end if

		Case 13 'CicloCompra
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad, "
			sql = sql & " Id_Consumo "
			sql = sql & " FROM PH_DataCrudaMensual "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = " & sCat
			sql = sql & " AND Id_Fabricante = " & iFab
			if sMar <> "" then 
				sql = sql & " And Id_Marca = " & iMar 
			end if
			if sSeg <> "" then 
				sql = sql & " And Id_Segmento = " & iSeg 
			end if
			if sRan <> "" then 
				sql = sql & " And Id_RangoTamano = " & iRan
			end if
			sql = sql & " And id_Semana in( " & idSemana & ")"
			sql = sql & " GROUP BY "
			sql = sql & " Cantidad, "
			sql = sql & " Id_Consumo "
			'response.write "<br>36 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			'response.write "<br>257 LLEGO" 
			'response.end
			if rsx1.eof then
				rsx1.close
				Valor = 0
				Valor = FormatNumber(Valor,2)
			else
				'response.write "<br>84 LLEGO"
				'response.end
				gDatos1 = rsx1.GetRows
				rsx1.close
				Indicador4 = 0
				for iDat = 0 to ubound(gDatos1,2)
					Indicador4 = Indicador4 + 1
				next
				
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Id_Hogar AS Total "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria =  " & sCat
				sql = sql & " And Id_Fabricante = " & iFab
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
				end if
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				sql = sql & " PH_DataCrudaMensual.Id_Hogar "
				'response.write "<br>36 sql:=" & sql
				'response.end
				rsx1.Open sql ,conexion
				'response.write "<br>257 LLEGO" 
				'response.end
				if rsx1.eof then
					rsx1.close
				else
					'response.write "<br>84 LLEGO"
					'response.end
					gDatos1 = rsx1.GetRows
					rsx1.close
				end if
				Indicador5 = 0
				for iDat = 0 to ubound(gDatos1,2)
					'Cantidad = gDatos1(0,0)
					Indicador5 = Indicador5 + 1
				next
				'Valor = 7/(cdbl(Indicador4) / cdbl(Indicador5))
				Valor = TotalDias/(cdbl(Indicador4) / cdbl(Indicador5))
				'response.write "<br>36 Indicador1=" & Indicador2
				'response.write "<br>36 Indicador5=" & Indicador5
				Valor = FormatNumber(Valor,2)
			end if

		Case 14 'VolActoCompra
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Tamano, "
				sql = sql & " Cantidad "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria = " & sCat
				sql = sql & " And Id_Area = " & iAre
				paso = 0
				if sFab <> "" then 
					sql = sql & " And Id_Fabricante = " & iFab 
					paso = 1
				end if
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
					paso = 1
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
					paso = 1
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
					paso = 1
				end if
				if sTam <> "" then 
					sql = sql & " And Id_Tamano = " & iTam
					paso = 1
				end if
				sql = sql & " And id_Semana in( " & idSemana & ")"
				'if paso = 0 then 
					sql = sql & " GROUP BY "
					sql = sql & " Tamano, "
					sql = sql & " Cantidad, "
					sql = sql & " Id_Consumo, "
					sql = sql & " CodigoBarra "
				'end if
			'response.write "<br>2387 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			'response.write "<br>257 LLEGO" 
			'response.end
			if rsx1.eof then
				rsx1.close
				Valor = 0
				Valor = FormatNumber(Valor,2)
			else
				'response.write "<br>84 LLEGO"
				'response.end
				gDatos1 = rsx1.GetRows
				rsx1.close
				Valor = 0
				Valor = FormatNumber(Valor,2)
				Indicador1 = 0
				for iDat = 0 to ubound(gDatos1,2)
					Indicador1 = Indicador1 + (cdbl(gDatos1(0,iDat)) *cdbl(gDatos1(1,iDat)))
				next

				sql = ""
				sql = sql & " SELECT "
				'sql = sql & " Cantidad, "
				sql = sql & " Id_Consumo "
				sql = sql & " FROM PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria = " & sCat
				sql = sql & " And Id_Area = " & iAre
				if sFab <> "" then 
					sql = sql & " AND Id_Fabricante = " & iFab
				end if
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
				end if
				if sTam <> "" then 
					sql = sql & " And Id_Tamano = " & iTam
				end if
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				'sql = sql & " Cantidad, "
				sql = sql & " Id_Consumo "
				'response.write "<br>2434 sql:=" & sql
				'response.end
				rsx1.Open sql ,conexion
				'response.write "<br>257 LLEGO" 
				'response.end
				if rsx1.eof then
					rsx1.close
				else
					'response.write "<br>84 LLEGO"
					'response.end
					gDatos1 = rsx1.GetRows
					rsx1.close
				end if
				Indicador4 = 0
				for iDat = 0 to ubound(gDatos1,2)
					Indicador4 = Indicador4 + 1
				next
				
				if cint(sCat) <> 124 then
					Indicador14 = ((cdbl(Indicador1)/1000) / cdbl(Indicador4))
				else
					Indicador14 = ((cdbl(Indicador1)) / cdbl(Indicador4))
				end if

				Valor = FormatNumber(Indicador14,2)

			end if

		Case 15 'ValActoCompra
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Cantidad, "
				sql = sql & " Precio_Producto, "
				sql = sql & " Dolar, "
				sql = sql & " Precio_producto/Dolar AS PrecioDolar, "
				sql = sql & " (Precio_producto/Dolar)*Cantidad AS ComprasValor "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria = " & sCat
				sql = sql & " And Id_Area = " & iAre
				paso = 0
				if sFab <> "" then 
					sql = sql & " And Id_Fabricante = " & iFab 
					paso = 1
				end if
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
					paso = 1
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
					paso = 1
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
					paso = 1
				end if
				if sTam <> "" then 
					sql = sql & " And Id_Tamano = " & iTam
					paso = 1
				end if
				sql = sql & " And id_Semana in( " & idSemana & ")"
				'if paso = 0 then 
					sql = sql & " GROUP BY "
					sql = sql & " Cantidad, "
					sql = sql & " Precio_Producto, "
					sql = sql & " Dolar, "
					sql = sql & " Id_Consumo, "
					sql = sql & " CodigoBarra "
				'end if
				'response.write "<br>1804 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			'response.write "<br>257 LLEGO" 
			'response.end
			if rsx1.eof then
				rsx1.close
				Valor = 0
				Valor = FormatNumber(Valor,2)
			else
				'response.write "<br>84 LLEGO"
				'response.end
				gDatos1 = rsx1.GetRows
				rsx1.close
				Indicador2 = 0
				for iDat = 0 to ubound(gDatos1,2)
					Indicador2 = Indicador2 + cdbl(gDatos1(4,iDat))
				next
				sql = ""
				sql = sql & " SELECT "
				'sql = sql & " Cantidad, "
				sql = sql & " Id_Consumo "
				sql = sql & " FROM PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria = " & sCat
				sql = sql & " And Id_Area = " & iAre
				if sFab <> "" then 
					sql = sql & " AND Id_Fabricante = " & iFab
				end if
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
				end if
				if sTam <> "" then 
					sql = sql & " And Id_Tamano = " & iTam
					paso = 1
				end if
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				'sql = sql & " Cantidad, "
				sql = sql & " Id_Consumo "
				'response.write "<br>36 sql:=" & sql
				'response.end
				rsx1.Open sql ,conexion
				'response.write "<br>257 LLEGO" 
				'response.end
				if rsx1.eof then
					rsx1.close
				else
					'response.write "<br>84 LLEGO"
					'response.end
					gDatos1 = rsx1.GetRows
					rsx1.close
				end if
				Indicador4 = 0
				for iDat = 0 to ubound(gDatos1,2)
					Indicador4 = Indicador4 + 1
				next
				
				Indicador14 = ((cdbl(Indicador2)) / cdbl(Indicador4))
				'response.write "<br>84 Indicador2:= " & Indicador2
				'response.write "<br>84 Indicador4:= " & Indicador4

				Valor = FormatNumber(Indicador14,2)
			end if

		Case 16 'UnidActoCompra
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Cantidad "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria = " & sCat
				sql = sql & " and Id_Area = " & iAre
				paso = 0
				if sFab <> "" then 
					sql = sql & " And Id_Fabricante = " & iFab 
					paso = 1
				end if
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
					paso = 1
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
					paso = 1
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
					paso = 1
				end if
				if sTam <> "" then 
					sql = sql & " And Id_Tamano = " & iTam
					paso = 1
				end if
				sql = sql & " And id_Semana in( " & idSemana & ")"
				'if paso = 0 then 
					sql = sql & " GROUP BY "
					sql = sql & " Cantidad, "
					sql = sql & " Id_Consumo, "
					sql = sql & " CodigoBarra "
				'end if
			'response.write "<br>36 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			'response.write "<br>257 LLEGO" 
			'response.end
			if rsx1.eof then
				rsx1.close
				Valor = 0
				Valor = FormatNumber(Valor,2)
			else
				'response.write "<br>84 LLEGO"
				'response.end
				gDatos1 = rsx1.GetRows
				rsx1.close
				Valor = 0
				Valor = FormatNumber(Valor,2)
				Indicador3 = 0
				for iDat = 0 to ubound(gDatos1,2)
					Indicador3 = Indicador3 + cdbl(gDatos1(0,iDat))
				next

				sql = ""
				sql = sql & " SELECT "
				'sql = sql & " Cantidad, "
				sql = sql & " Id_Consumo "
				sql = sql & " FROM PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria = " & sCat
				sql = sql & " and Id_Area = " & iAre
				if sFab <> "" then 
					sql = sql & " AND Id_Fabricante = " & iFab
				end if
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
				end if
				if sTam <> "" then 
					sql = sql & " And Id_Tamano = " & iTam
					paso = 1
				end if
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				'sql = sql & " Cantidad, "
				sql = sql & " Id_Consumo "
				'response.write "<br>36 sql:=" & sql
				'response.end
				rsx1.Open sql ,conexion
				'response.write "<br>257 LLEGO" 
				'response.end
				if rsx1.eof then
					rsx1.close
				else
					'response.write "<br>84 LLEGO"
					'response.end
					gDatos1 = rsx1.GetRows
					rsx1.close
				end if
				Indicador4 = 0
				for iDat = 0 to ubound(gDatos1,2)
					Indicador4 = Indicador4 + 1
				next
				
				Indicador14 = ((cdbl(Indicador3)) / cdbl(Indicador4))

				Valor = FormatNumber(Indicador14,2)
			end if

		Case 17 'IndiceConsumoVol 
			Valor = 0
		
		Case 18 'IndiceConsumoVal
			Valor = 0
		
		Case 19 'RepeticionCompra (NO VA - Es Mensual)
			Valor = 0
		Case 20 'FidelidadVol (NO VA - Es Mensual)
			Valor = 0
		Case 21 'FidelidadVal (NO VA - Es Mensual)
			Valor = 0
		Case 22 'FidelidadActos (NO VA - Es Mensual)
			Valor = 0
		
		Case 23 'CuotaMerVol 
			Valor = 0

		Case 24 'PrecPromVol
			paso = 0
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad, "
			sql = sql & " Precio_Producto, "
			sql = sql & " Dolar, "
			sql = sql & " Precio_producto/Dolar AS PrecioDolar, "
			sql = sql & " (Precio_producto/Dolar)*Cantidad AS ComprasValor "
			sql = sql & " FROM "
			sql = sql & " PH_DataCrudaMensual "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = " & sCat
			sql = sql & " and Id_Area = " & iAre
			if sFab <> "" then 
				sql = sql & " And Id_Fabricante = " & iFab 
				paso = 1
			end if
			if sMar <> "" then 
				sql = sql & " And Id_Marca = " & iMar 
				paso = 1
			end if
			if sSeg <> "" then 
				sql = sql & " And Id_Segmento = " & iSeg 
				paso = 1
			end if
			if sRan <> "" then 
				sql = sql & " And Id_RangoTamano = " & iRan
				paso = 1
			end if
			if sTam <> "" then 
				sql = sql & " And Id_Tamano = " & iTam
				paso = 1
			end if
			sql = sql & " And id_Semana in( " & idSemana & ")"
			if paso = 0 then 
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Cantidad, "
				sql = sql & " Precio_Producto, "
				sql = sql & " Dolar, "
				sql = sql & " Precio_producto/Dolar AS PrecioDolar, "
				sql = sql & " (Precio_producto/Dolar)*Cantidad AS ComprasValor "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " GROUP BY "
				sql = sql & " Id_Consumo, "
				sql = sql & " Producto, "
				sql = sql & " Cantidad, "
				sql = sql & " Precio_Producto, "
				sql = sql & " Dolar, "
				sql = sql & " Id_Categoria, "
				sql = sql & " Id_Area, "
				sql = sql & " id_Semana "
				sql = sql & " HAVING "
				sql = sql & " Id_Categoria = " & sCat 
				sql = sql & " AND Id_Area = " & iAre 
				sql = sql & " and id_Semana in( " & idSemana & ")"
			end if
			'response.write "<br>36 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			'response.write "<br>257 LLEGO" 
			'response.end
			if rsx1.eof then
				rsx1.close
				Valor = 0
				Valor = FormatNumber(Valor,2)
			else
				'response.write "<br>84 LLEGO"
				'response.end
				gDatos1 = rsx1.GetRows
				rsx1.close
				Indicador2 = 0
				for iDat = 0 to ubound(gDatos1,2)
					Indicador2 = Indicador2 + cdbl(gDatos1(4,iDat))
				next
				paso = 0
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Tamano, "
				sql = sql & " Cantidad "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria = " & sCat
				sql = sql & " and Id_Area = " & iAre
				if sFab <> "" then 
					sql = sql & " And Id_Fabricante = " & iFab 
					paso = 1
				end if
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
					paso = 1
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
					paso = 1
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
					paso = 1
				end if
				if sTam <> "" then 
					sql = sql & " And Id_Tamano = " & iTam
					paso = 1
				end if
				sql = sql & " And id_Semana in( " & idSemana & ")"
				if paso = 0 then
					sql = ""
					sql = sql & " SELECT "
					sql = sql & " Tamano, "
					sql = sql & " Cantidad "
					sql = sql & " FROM "
					sql = sql & " PH_DataCrudaMensual "
					sql = sql & " GROUP BY "
					sql = sql & " Id_Consumo, "
					sql = sql & " CodigoBarra, "
					sql = sql & " Tamano, "
					sql = sql & " Cantidad, "
					sql = sql & " Id_Categoria, "
					sql = sql & " Id_Area, "
					sql = sql & " id_Semana "
					sql = sql & " HAVING "
					sql = sql & " Id_Categoria = " & sCat 
					sql = sql & " AND Id_Area = " & iAre
					sql = sql & " and id_Semana in( " & idSemana & ")"
				end if
				
				'response.write "<br>36 sql:=" & sql
				'response.end
				rsx1.Open sql ,conexion
				'response.write "<br>257 LLEGO" 
				'response.end
				if rsx1.eof then
					rsx1.close
				else
					'response.write "<br>84 LLEGO"
					'response.end
					gDatos1 = rsx1.GetRows
					rsx1.close
				end if
				Indicador1 = 0
				for iDat = 0 to ubound(gDatos1,2)
					Indicador1 = Indicador1 + (cdbl(gDatos1(0,iDat)) *cdbl(gDatos1(1,iDat)))
				next
				if cint(sCat) <> 124 then
					Indicador1 = Indicador1/1000
				else
					Indicador1 = Indicador1
				end if
				
				Valor = cdbl(Indicador2)/cdbl(Indicador1)
				Valor = FormatNumber(Valor,2)
			end if

		Case 25 'PrecPromUnid
			paso = 0
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad, "
			sql = sql & " Precio_Producto, "
			sql = sql & " Dolar, "
			sql = sql & " Precio_producto/Dolar AS PrecioDolar, "
			sql = sql & " (Precio_producto/Dolar)*Cantidad AS ComprasValor "
			sql = sql & " FROM "
			sql = sql & " PH_DataCrudaMensual "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = " & sCat
			sql = sql & " and Id_Area = " & iAre
			if sFab <> "" then 
				sql = sql & " And Id_Fabricante = " & iFab 
				paso = 1
			end if
			if sMar <> "" then 
				sql = sql & " And Id_Marca = " & iMar 
				paso = 1
			end if
			if sSeg <> "" then 
				sql = sql & " And Id_Segmento = " & iSeg 
				paso = 1
			end if
			if sRan <> "" then 
				sql = sql & " And Id_RangoTamano = " & iRan
				paso = 1
			end if
			if sTam <> "" then 
				sql = sql & " And Id_Tamano = " & iTam
				paso = 1
			end if
			sql = sql & " And id_Semana in( " & idSemana & ")"
			if paso = 0 then
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Cantidad, "
				sql = sql & " Precio_Producto, "
				sql = sql & " Dolar, "
				sql = sql & " Precio_producto/Dolar AS PrecioDolar, "
				sql = sql & " (Precio_producto/Dolar)*Cantidad AS ComprasValor "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " GROUP BY "
				sql = sql & " Id_Consumo, "
				sql = sql & " CodigoBarra, "
				sql = sql & " Cantidad, "
				sql = sql & " Precio_Producto, "
				sql = sql & " Dolar, "
				sql = sql & " Id_Categoria, "
				sql = sql & " Id_Area, "
				sql = sql & " id_Semana "
				sql = sql & " HAVING "
				sql = sql & " Id_Categoria = " & sCat
				sql = sql & " AND Id_Area = " & iAre
				sql = sql & " And id_Semana in( " & idSemana & ")"
			end if
			'response.write "<br>36 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			'response.write "<br>257 LLEGO" 
			'response.end
			if rsx1.eof then
				rsx1.close
				Valor = 0
				Valor = FormatNumber(Valor,2)
			else
				'response.write "<br>84 LLEGO"
				'response.end
				gDatos1 = rsx1.GetRows
				rsx1.close
				Indicador2 = 0
				for iDat = 0 to ubound(gDatos1,2)
					Indicador2 = Indicador2 + cdbl(gDatos1(4,iDat))
				next
				paso = 0
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Cantidad "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria = " & sCat
				sql = sql & " and Id_Area = " & iAre
				if sFab <> "" then 
					sql = sql & " and Id_Fabricante = " & iFab 
					paso = 1
				end if
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
					paso = 1
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
					paso = 1
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
					paso = 1
				end if
				if sTam <> "" then 
					sql = sql & " And Id_Tamano = " & iTam
					paso = 1
				end if
				sql = sql & " And id_Semana in( " & idSemana & ")"
				if paso = 0 then
					sql = ""
					sql = sql & " SELECT "
					sql = sql & " Cantidad "
					sql = sql & " FROM "
					sql = sql & " PH_DataCrudaMensual "
					sql = sql & " WHERE "
					sql = sql & " Id_Categoria = " & sCat 
					sql = sql & " AND Id_Area = " & iAre
					sql = sql & " And id_Semana in( " & idSemana & ")"
					sql = sql & " GROUP BY "
					sql = sql & " Id_Consumo, "
					sql = sql & " CodigoBarra, "
					sql = sql & " Cantidad, "
					sql = sql & " Precio_Producto "
				end if
				

				'response.write "<br>36 sql:=" & sql
				'response.end
				rsx1.Open sql ,conexion
				'response.write "<br>257 LLEGO" 
				'response.end
				if rsx1.eof then
					rsx1.close
				else
					'response.write "<br>84 LLEGO"
					'response.end
					gDatos1 = rsx1.GetRows
					rsx1.close
				end if
				Valor = 0
				Indicador3 = 0
				for iDat = 0 to ubound(gDatos1,2)
					Indicador3 = Indicador3 + cdbl(gDatos1(0,iDat))
				next
				
				Valor = cdbl(Indicador2)/cdbl(Indicador3)
				Valor = FormatNumber(Valor,2)
			end if

		Case 26 'MarcasHogar Calcular 
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Id_Hogar AS Total "
			sql = sql & " FROM "
			sql = sql & " PH_DataCrudaMensual "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = " & sCat
			sql = sql & " and Id_Area = " & iAre
			if sFab <> "" then 
				sql = sql & " And Id_Fabricante = " & iFab
			end if
			if sMar <> "" then 
				sql = sql & " And Id_Marca = " & iMar 
			end if
			if sSeg <> "" then 
				sql = sql & " And Id_Segmento = " & iSeg 
			end if
			if sRan <> "" then 
				sql = sql & " And Id_RangoTamano = " & iRan
			end if
			if sTam <> "" then 
				sql = sql & " And Id_Tamano = " & iTam
			end if
			sql = sql & " And id_Semana in( " & idSemana & ")"
			sql = sql & " GROUP BY "
			sql = sql & " PH_DataCrudaMensual.Id_Hogar "
			'response.write "<br>36 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			'response.write "<br>257 LLEGO" 
			'response.end
			if rsx1.eof then
				rsx1.close
				Valor = 0
				Valor = FormatNumber(Valor,2)
			else
				'response.write "<br>84 LLEGO"
				'response.end
				gDatos1 = rsx1.GetRows
				rsx1.close
				Valor = 0
				Cantidad = 0
				for iDat = 0 to ubound(gDatos1,2)
					'Cantidad = gDatos1(0,0)
					Cantidad = Cantidad + 1
				next
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Id_Hogar AS Total "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Area = " & iAre
				if sFab <> "" then 
					sql = sql & " And Id_Fabricante = " & iFab
				end if
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
				end if
				if sTam <> "" then 
					sql = sql & " And Id_Tamano = " & iTam
				end if
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				sql = sql & " PH_DataCrudaMensual.Id_Hogar "
				'response.write "<br>36 sql:=" & sql
				'response.end
				rsx1.Open sql ,conexion
				if rsx1.eof then
					rsx1.close
				else
					'response.write "<br>84 LLEGO"
					'response.end
					gDatos1 = rsx1.GetRows
					rsx1.close
				end if
				Total = 0
				for iDat = 0 to ubound(gDatos1,2)
					Total = Total + 1
				next
				Penetracion = (Cantidad/Total)*100

				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Marca, "
				sql = sql & " Id_Hogar AS Total "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria = "  & sCat
				sql = sql & " And Id_Area = "  & iAre
				sql = sql & " AND Id_Marca <> 0 "
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				sql = sql & " Marca, "
				sql = sql & " Id_Hogar "
				'response.write "<br>36 sql:=" & sql & "<br>"
				'response.end
				rsx1.Open sql ,conexion
				if rsx1.eof then
					rsx1.close
				else
					'response.write "<br>84 LLEGO"
					'response.end
					gDatos1 = rsx1.GetRows
					rsx1.close
				end if
				PenetracionMarcas = 0
				for iDat = 0 to ubound(gDatos1,2)
					PenetracionMarcas = PenetracionMarcas + 1
				next
				Valor = PenetracionMarcas / Cantidad
				Valor = FormatNumber(Valor,2)
				'response.write "<br> Penetracion%:" & Penetracion
				'response.write "<br> PenetracionMarcas:" & PenetracionMarcas
				'response.write "<br> Hog Ref:" & Cantidad
				'response.write "<br> Hog: " & Total
			end if
			if iFab <> 0 then 
				Valor = "N/A"
			end if
			if iMar <> 0 then 
				Valor = "N/A"
			end if
			if iSeg <> 0 then 
				Valor = "N/A"
			end if
			if iRan <> 0 then 
				Valor = "N/A"
			end if
			if iTam <> 0 then 
				Valor = "N/A"
			end if
			'response.write "<br> Valor:" & Valor
			
		Case 27 'CadenasProm
			Valor = 0
		
		Case 28 'CuotaMercVol
			'response.write "<br>Revisando CuotaMercVol"
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Tamano, "
			sql = sql & " Cantidad "
			sql = sql & " FROM "
			sql = sql & " PH_DataCrudaMensual "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = " & sCat
			'20May2021
			'sql = sql & " and Id_Area = 0" '& iAre
			if sFab = "" and sMar = "" and sSeg = "" and sRan = "" and sTam <> "" then
				sql = sql & " and Id_Area = " & iAre
			else
				if sFab = "" and sMar = "" and sSeg = "" and sRan = "" then
					sql = sql & " and Id_Area = 0 "
				else
					sql = sql & " and Id_Area = " & iAre
				end if
			end if
			sql = sql & " And Id_Fabricante = 0 "
			sql = sql & " And Id_Marca = 0"
			sql = sql & " And Id_Segmento = 0"
			sql = sql & " And Id_RangoTamano = 0"
			sql = sql & " And id_Semana in( " & idSemana & ")"
			'response.write "<br>5499 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			'response.write "<br>257 LLEGO" 
			'response.end
			if rsx1.eof then
				rsx1.close
				Valor = 0
				Valor = FormatNumber(Valor,2)
			else
				'response.write "<br>84 LLEGO"
				'response.end
				gDatos1 = rsx1.GetRows
				rsx1.close
				TotalVolumen = 0
				for iDat = 0 to ubound(gDatos1,2)
					TotalVolumen = TotalVolumen + (cdbl(gDatos1(0,iDat)) *cdbl(gDatos1(1,iDat)))
				next
				TotalVolumen = (TotalVolumen)/1000
				
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Tamano, "
				sql = sql & " Cantidad "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria = " & sCat
				sql = sql & " and Id_Area = " & iAre
				if sFab = "" and sMar = "" and sSeg <> "" and sRan = "" then
					if iAre = 0 then
						sql = sql & " And Id_Fabricante =  0 "  
						sql = sql & " And Id_Marca = 0 "  
						sql = sql & " And Id_Segmento = " & iSeg 
						'sql = sql & " And Id_RangoTamano =  0 "   
						'response.write "<br>3243 LLEGO" 
					else
						sql = sql & " And Id_Segmento = " & iSeg 
					end if
				else
				if sFab = "" and sMar = "" and sSeg = "" and sRan = "" then
					'if sAre <> "" then 
					'	sql = sql & " and Id_Area = " & sAre
						'response.write "<br>3243 LLEGO2"
					'end if
					'if iAre = 0 then
						sql = sql & " And Id_Fabricante =  0 "  
						sql = sql & " And Id_Marca = 0 "  
						sql = sql & " And Id_Segmento = 0 "
						sql = sql & " And Id_RangoTamano =  0 "   
						'response.write "<br>3243 LLEGO2" 
					'end if
				end if
				if sFab = "" and sMar = "" and sSeg = "" and sRan <> "" then
					if iAre = 0 then
						sql = sql & " And Id_Fabricante =  0 "  
						sql = sql & " And Id_Marca = 0 "  
						sql = sql & " And Id_Segmento = 0 "
						'sql = sql & " And Id_RangoTamano =  0 "   
						'response.write "<br>3243 LLEGO3" 
					end if
				end if
					if sFab <> "" then 
						sql = sql & " And Id_Fabricante = " & iFab 
					else
						if iAre = 0 then
							sql = sql & " And Id_Fabricante =  0 "  
						end if
					end if
					if sMar <> "" then 
						sql = sql & " And Id_Marca = " & iMar 
					else
						'if iAre = 0 then
						'	sql = sql & " And Id_Marca = 0 "  
						'end if
					end if
					if sSeg <> "" then 
						sql = sql & " And Id_Segmento = " & iSeg 
					else
						'if iAre = 0 then
						'	sql = sql & " And Id_Segmento = 0 "  
						'end if
					end if
					if sRan <> "" then 
						sql = sql & " And Id_RangoTamano = " & iRan
					else
						'if iAre = 0 then
						'	sql = sql & " And Id_RangoTamano =  0 "   
						'end if
					end if
				
				end if
				if sTam <> "" then 
					sql = sql & " And Id_Tamano = " & iTam
				end if
				 if sFab = "" and sMar = "" and sSeg <> "" and sRan <> "" and sTam = "" then
					if sAre = "" then sAre = 0
					'response.write "<br>5596 Pasoooooo" & sAre
					'response.end 
					
					if sAre = 0 then
						sql = sql & " And Id_Marca = 0 "  
					else
						'sql = sql & " And Id_Marca = 0 "  
						'sql = sql & " And Id_Fabricante = 0 "  
						sql = replace(sql," And Id_Marca = 0 And Id_Fabricante = 0"," ")
					end if
					' sql = ""
					' sql = sql & " SELECT "
					' sql = sql & " Tamano, "
					' sql = sql & " Cantidad "
					' sql = sql & " FROM "
					' sql = sql & " PH_DataCrudaMensual "
					' sql = sql & " GROUP BY "
					' sql = sql & " Tamano, "
					' sql = sql & " Cantidad, "
					' sql = sql & " Id_Consumo, "
					' sql = sql & " Id_Categoria, "
					' sql = sql & " Id_Area, "
					' sql = sql & " Id_Fabricante, "
					' sql = sql & " Id_Segmento, "
					' sql = sql & " Id_RangoTamano, "
					' sql = sql & " id_Semana "
					' sql = sql & " HAVING "
					' sql = sql & " Id_Categoria = " & sCat
					' sql = sql & " AND Id_Area = " & iAre
					' sql = sql & " AND Id_Fabricante = 0 "
					' sql = sql & " AND Id_Segmento = " & iSeg
					' sql = sql & " AND Id_RangoTamano = " & iRan
				 end if 
				sql = sql & " And id_Semana in( " & idSemana & ")"
				'response.write "<br>5595 sql:=" & sql & "<br>"
				'response.end
				rsx1.Open sql ,conexion
				'response.write "<br>2357 LLEGO" 
				'response.end
				if rsx1.eof then
					rsx1.close
					Valor = 0
					Valor = FormatNumber(Valor,2)
				else
					'response.write "<br>84 LLEGO"
					'response.end
					gDatos1 = rsx1.GetRows
					rsx1.close
					TotalFiltro = 0
					for iDat = 0 to ubound(gDatos1,2)
						TotalFiltro = TotalFiltro + (cdbl(gDatos1(0,iDat)) *cdbl(gDatos1(1,iDat)))
					next
					TotalFiltro = TotalFiltro/1000
					'if iFab = 0 and iMar = 0 and iSeg = 0 and iTam = 0 then TotalFiltro=TotalVolumen: response.write "<br>*****"
					Valor = (TotalFiltro/TotalVolumen)*100
					Valor = FormatNumber(Valor,2)
					'response.write "<br> Total Volumen:= " & TotalVolumen
					'response.write "<br> Total Filtro:= " & TotalFiltro
				end if
				'response.end
			end if

		Case 29 'CuotaMercVal
			'response.write "<br>Revisando CuoMerVal"
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad, "
			sql = sql & " Precio_Producto, "
			sql = sql & " Dolar, "
			sql = sql & " Precio_producto/Dolar AS PrecioDolar, "
			sql = sql & " (Precio_producto/Dolar)*Cantidad AS ComprasValor "
			sql = sql & " FROM "
			sql = sql & " PH_DataCrudaMensual "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = " & sCat
			'20May2021
			'sql = sql & " and Id_Area = 0" '& iAre
			if sFab = "" and sMar = "" and sSeg = "" and sRan = "" and sTam <> "" then
				sql = sql & " and Id_Area = " & iAre
			else
				if sFab = "" and sMar = "" and sSeg = "" and sRan = "" then
					sql = sql & " and Id_Area = 0 "
				else
					sql = sql & " and Id_Area = " & iAre
				end if
			end if
			sql = sql & " And Id_Fabricante = 0 "
			sql = sql & " And Id_Marca = 0 "
			sql = sql & " And Id_Segmento = 0 "
			sql = sql & " And Id_RangoTamano = 0 "
			'sql = sql & " And Id_Tamano = 0 "
			sql = sql & " And id_Semana in( " & idSemana & ")"
			'response.write "<br>36 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			'response.write "<br>257 LLEGO" 
			'response.end
			if rsx1.eof then
				rsx1.close
				Valor = 0
				Valor = FormatNumber(Valor,2)
			else
				'response.write "<br>84 LLEGO"
				'response.end
				gDatos1 = rsx1.GetRows
				rsx1.close
				TotalValor = 0
				for iDat = 0 to ubound(gDatos1,2)
					TotalValor = TotalValor + FormatNumber(cdbl(gDatos1(4,iDat)))
				next

				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Cantidad, "
				sql = sql & " Precio_Producto, "
				sql = sql & " Dolar, "
				sql = sql & " Precio_producto/Dolar AS PrecioDolar, "
				sql = sql & " (Precio_producto/Dolar)*Cantidad AS ComprasValor "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria = " & sCat
				sql = sql & " and Id_Area = " & iAre
				if sFab = "" and sMar = "" and sSeg <> "" and sRan = "" then
					if iAre = 0 then
						sql = sql & " And Id_Fabricante =  0 "  
						sql = sql & " And Id_Marca = 0 "  
						sql = sql & " And Id_Segmento = " & iSeg 
						'sql = sql & " And Id_RangoTamano =  0 "   
						'response.write "<br>3243 LLEGO" 
					else
						sql = sql & " And Id_Segmento = " & iSeg 
					end if
				else
				if sFab = "" and sMar = "" and sSeg = "" and sRan = "" then
					'if iAre = 0 then
						sql = sql & " And Id_Fabricante =  0 "  
						sql = sql & " And Id_Marca = 0 "  
						sql = sql & " And Id_Segmento = 0 "
						sql = sql & " And Id_RangoTamano =  0 "   
						'response.write "<br>3243 LLEGO" 
					'end if
				end if
				if sFab = "" and sMar = "" and sSeg = "" and sRan <> "" then
					if iAre = 0 then
						sql = sql & " And Id_Fabricante =  0 "  
						sql = sql & " And Id_Marca = 0 "  
						sql = sql & " And Id_Segmento = 0 "
						'sql = sql & " And Id_RangoTamano =  0 "   
						'response.write "<br>3243 LLEGO" 
					end if
				end if
					if sFab <> "" then 
						sql = sql & " And Id_Fabricante = " & iFab 
					else
						if iAre = 0 then
							sql = sql & " And Id_Fabricante =  0 "  
						end if
					end if
					if sMar <> "" then 
						sql = sql & " And Id_Marca = " & iMar 
					else
						'if iAre = 0 then
						'	sql = sql & " And Id_Marca = 0 "  
						'end if
					end if
					if sSeg <> "" then 
						sql = sql & " And Id_Segmento = " & iSeg 
					else
						'if iAre = 0 then
						'	sql = sql & " And Id_Segmento =  0 "   
						'end if
					end if
					if sRan <> "" then 
						sql = sql & " And Id_RangoTamano = " & iRan
					end if
				end if
				if sTam <> "" then 
					sql = sql & " And Id_Tamano = " & iTam
				end if
				 if sFab = "" and sMar = "" and sSeg <> "" and sRan <> "" and sTam = "" then
					if sAre = "" then sAre = 0
					'response.write "<br>5596 Pasoooooo" & sAre
					'response.end 
					
					if sAre = 0 then
						sql = sql & " And Id_Marca = 0 "  
					else
						'sql = sql & " And Id_Marca = 0 "  
						'sql = sql & " And Id_Fabricante = 0 "  
						sql = replace(sql," And Id_Marca = 0 And Id_Fabricante = 0"," ")
					end if
					' sql = ""
					' sql = sql & " SELECT "
					' sql = sql & " Tamano, "
					' sql = sql & " Cantidad "
					' sql = sql & " FROM "
					' sql = sql & " PH_DataCrudaMensual "
					' sql = sql & " GROUP BY "
					' sql = sql & " Tamano, "
					' sql = sql & " Cantidad, "
					' sql = sql & " Id_Consumo, "
					' sql = sql & " Id_Categoria, "
					' sql = sql & " Id_Area, "
					' sql = sql & " Id_Fabricante, "
					' sql = sql & " Id_Segmento, "
					' sql = sql & " Id_RangoTamano, "
					' sql = sql & " id_Semana "
					' sql = sql & " HAVING "
					' sql = sql & " Id_Categoria = " & sCat
					' sql = sql & " AND Id_Area = " & iAre
					' sql = sql & " AND Id_Fabricante = 0 "
					' sql = sql & " AND Id_Segmento = " & iSeg
					' sql = sql & " AND Id_RangoTamano = " & iRan
				 end if 
				sql = sql & " And id_Semana in( " & idSemana & ")"
				'response.write "<br>36 sql:=" & sql
				'response.end
				rsx1.Open sql ,conexion
				'response.write "<br>257 LLEGO" 
				'response.end
				if rsx1.eof then
					rsx1.close
					Valor = 0
					Valor = FormatNumber(Valor,2)
				else
					'response.write "<br>84 LLEGO"
					'response.end
					gDatos1 = rsx1.GetRows
					rsx1.close
					TotalFiltro = 0
					for iDat = 0 to ubound(gDatos1,2)
						TotalFiltro = TotalFiltro + FormatNumber(cdbl(gDatos1(4,iDat)),2)
					next
					
					'if iFab = 0 and iMar = 0 and iSeg = 0 and iTam = 0 then TotalFiltro = TotalValor
					Valor = (TotalFiltro/TotalValor)*100
					Valor = FormatNumber(Valor,2)
					'response.write "<br> TotalValor:= " & TotalValor
					'response.write "<br> Total Filtro:= " & TotalFiltro
				end if
			end if 


		Case 30 'CuotaMercUnid
			'response.write "<br>Revisando CuotaMercUnid"
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad "
			sql = sql & " FROM "
			sql = sql & " PH_DataCrudaMensual "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = " & sCat
			'20May2021
			'sql = sql & " and Id_Area = 0" '& iAre
			if sFab = "" and sMar = "" and sSeg = "" and sRan = "" and sTam <> "" then
				sql = sql & " and Id_Area = " & iAre
			else
				if sFab = "" and sMar = "" and sSeg = "" and sRan = "" then
					sql = sql & " and Id_Area = 0 "
				else
					sql = sql & " and Id_Area = " & iAre
				end if
			end if
			sql = sql & " And Id_Fabricante = 0 "
			sql = sql & " And Id_Marca = 0 "
			sql = sql & " And Id_Segmento = 0 "
			sql = sql & " And Id_RangoTamano = 0 "
			sql = sql & " And id_Semana in( " & idSemana & ")"
			'response.write "<br>36 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			'response.write "<br>257 LLEGO" 
			'response.end
			if rsx1.eof then
				rsx1.close
				Valor = 0
				Valor = FormatNumber(Valor,2)
			else
				'response.write "<br>84 LLEGO"
				'response.end
				gDatos1 = rsx1.GetRows
				rsx1.close
				TotalUnidades = 0
				for iDat = 0 to ubound(gDatos1,2)
					TotalUnidades = TotalUnidades + cdbl(gDatos1(0,iDat))
				next
				
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Cantidad "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria = " & sCat
				sql = sql & " and Id_Area = " & iAre
				if sFab = "" and sMar = "" and sSeg <> "" and sRan = "" then
					if iAre = 0 then
						sql = sql & " And Id_Fabricante =  0 "  
						sql = sql & " And Id_Marca = 0 "  
						sql = sql & " And Id_Segmento = " & iSeg 
						'sql = sql & " And Id_RangoTamano =  0 "   
						'response.write "<br>3243 LLEGO" 
					else
						sql = sql & " And Id_Segmento = " & iSeg 
					end if
				else
				if sFab = "" and sMar = "" and sSeg = "" and sRan = "" then
					'if iAre = 0 then
						sql = sql & " And Id_Fabricante =  0 "  
						sql = sql & " And Id_Marca = 0 "  
						sql = sql & " And Id_Segmento = 0 "
						sql = sql & " And Id_RangoTamano =  0 "   
						'response.write "<br>3243 LLEGO" 
					'end if
				end if
				if sFab = "" and sMar = "" and sSeg = "" and sRan <> "" then
					if iAre = 0 then
						sql = sql & " And Id_Fabricante =  0 "  
						sql = sql & " And Id_Marca = 0 "  
						sql = sql & " And Id_Segmento = 0 "
						'sql = sql & " And Id_RangoTamano =  0 "   
						'response.write "<br>3243 LLEGO" 
					end if
				end if
					if sFab <> "" then 
						sql = sql & " And Id_Fabricante = " & iFab 
					else
						if iAre = 0 then
							sql = sql & " And Id_Fabricante =  0 "  
						end if
					end if
					if sMar <> "" then 
						sql = sql & " And Id_Marca = " & iMar 
					else
						'if iAre = 0 then
						'	sql = sql & " And Id_Marca = 0 "  
						'end if
					end if
					if sSeg <> "" then 
						sql = sql & " And Id_Segmento = " & iSeg 
					else
						'sql = sql & " And Id_Segmento =  0 "   
					end if
					if sRan <> "" then 
						sql = sql & " And Id_RangoTamano = " & iRan
					end if
				end if
				if sTam <> "" then 
					sql = sql & " And Id_Tamano = " & iTam
				end if
				 if sFab = "" and sMar = "" and sSeg <> "" and sRan <> "" and sTam = "" then
					if sAre = "" then sAre = 0
					'response.write "<br>5596 Pasoooooo" & sAre
					'response.end 
					
					if sAre = 0 then
						sql = sql & " And Id_Marca = 0 "  
					else
						'sql = sql & " And Id_Marca = 0 "  
						'sql = sql & " And Id_Fabricante = 0 "  
						sql = replace(sql," And Id_Marca = 0 And Id_Fabricante = 0"," ")
					end if
					' sql = ""
					' sql = sql & " SELECT "
					' sql = sql & " Tamano, "
					' sql = sql & " Cantidad "
					' sql = sql & " FROM "
					' sql = sql & " PH_DataCrudaMensual "
					' sql = sql & " GROUP BY "
					' sql = sql & " Tamano, "
					' sql = sql & " Cantidad, "
					' sql = sql & " Id_Consumo, "
					' sql = sql & " Id_Categoria, "
					' sql = sql & " Id_Area, "
					' sql = sql & " Id_Fabricante, "
					' sql = sql & " Id_Segmento, "
					' sql = sql & " Id_RangoTamano, "
					' sql = sql & " id_Semana "
					' sql = sql & " HAVING "
					' sql = sql & " Id_Categoria = " & sCat
					' sql = sql & " AND Id_Area = " & iAre
					' sql = sql & " AND Id_Fabricante = 0 "
					' sql = sql & " AND Id_Segmento = " & iSeg
					' sql = sql & " AND Id_RangoTamano = " & iRan
				 end if 
				sql = sql & " And id_Semana in( " & idSemana & ")"
				'response.write "<br>36 sql:=" & sql
				'response.end
				rsx1.Open sql ,conexion
				'response.write "<br>257 LLEGO" 
				'response.end
				if rsx1.eof then
					rsx1.close
					Valor = 0
					Valor = FormatNumber(Valor,2)
				else
					'response.write "<br>84 LLEGO"
					'response.end
					gDatos1 = rsx1.GetRows
					rsx1.close
					TotalFiltro = 0
					for iDat = 0 to ubound(gDatos1,2)
						TotalFiltro = TotalFiltro + cdbl(gDatos1(0,iDat))
					next
					
					'if iFab = 0 and iMar = 0 and iSeg = 0 and iTam = 0 then TotalFiltro = TotalUnidades
					Valor = (TotalFiltro/TotalUnidades)*100
					Valor = FormatNumber(Valor,2)
				end if
			end if

		Case 31 'CuoMerAct
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad, "
			sql = sql & " Id_Consumo "
			sql = sql & " FROM PH_DataCrudaMensual "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = " & sCat
			sql = sql & " AND Id_Fabricante = 0 "
			sql = sql & " AND Id_Marca = 0"
			sql = sql & " AND Id_Segmento = 0 "
			sql = sql & " AND Id_RangoTamano  = 0"
			sql = sql & " And id_Semana in( " & idSemana & ")"
			sql = sql & " GROUP BY "
			sql = sql & " Cantidad, "
			sql = sql & " Id_Consumo "
			'response.write "<br>36 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			'response.write "<br>257 LLEGO" 
			'response.end
			if rsx1.eof then
				rsx1.close
				'response.write "<br>257 LLEGO" 
				Valor = 0
			else
				'response.write "<br>84 LLEGO"
				'response.end
				gDatos1 = rsx1.GetRows
				rsx1.close
				TotalActos = 0
				for iDat = 0 to ubound(gDatos1,2)
					TotalActos = TotalActos + 1
				next
				

				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Cantidad, "
				sql = sql & " Id_Consumo "
				sql = sql & " FROM PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria = " & sCat
				sql = sql & " AND Id_Fabricante = " & iFab
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
				end if
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				sql = sql & " Cantidad, "
				sql = sql & " Id_Consumo "
				'response.write "<br>36 sql:=" & sql
				'response.end
				rsx1.Open sql ,conexion
				'response.write "<br>257 LLEGO" 
				'response.end
				if rsx1.eof then
					rsx1.close
					Valor = 0
				else
					'response.write "<br>84 LLEGO"
					'response.end
					gDatos1 = rsx1.GetRows
					rsx1.close
				TotalFiltro = 0
				for iDat = 0 to ubound(gDatos1,2)
					TotalFiltro = TotalFiltro + 1
				next

				if iFab = 0 then TotalFiltro = TotalActos
				Valor = (TotalFiltro/TotalActos)*100
				end if
				Valor = FormatNumber(Valor,2)
			end if
			
		
		Case 32 'PenetRelativa
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Id_Hogar AS Total "
			sql = sql & " FROM "
			sql = sql & " PH_DataCrudaMensual "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = " & sCat
			sql = sql & " And Id_Area = " & iAre
			sql = sql & " And id_Semana in( " & idSemana & ")"
			sql = sql & " GROUP BY "
			sql = sql & " PH_DataCrudaMensual.Id_Hogar "
			'response.write "<br><br>2522 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			'response.write "<br>257 LLEGO" 
			'response.end
			if rsx1.eof then
				rsx1.close
				Valor = 0
				Valor = FormatNumber(Valor,2)
			else
				'response.write "<br>84 LLEGO"
				'response.end
				gDatos1 = rsx1.GetRows
				rsx1.close
				Valor = 0
				Cantidad = 0
				for iDat = 0 to ubound(gDatos1,2)
					Cantidad = Cantidad + 1
				next
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Id_Hogar AS Total "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Area = " & iAre
				if sFab <> "" then 
					sql = sql & " and Id_Fabricante = " & iFab
				end if
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
				end if
				if sTam <> "" then 
					sql = sql & " And Id_Tamano = " & iTam
				end if
				sql = sql & " and Id_Categoria = " & sCat
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				sql = sql & " PH_DataCrudaMensual.Id_Hogar "
				'response.write "<br>1994 sql:=" & sql & "<br>"
				'response.end
				rsx1.Open sql ,conexion
				if rsx1.eof then
					rsx1.close
					Total = 0
					Valor = 0
					Valor = FormatNumber(Valor,2)
				else
					'response.write "<br>84 LLEGO"
					'response.end
					gDatos1 = rsx1.GetRows
					rsx1.close
					Total = 0
					for iDat = 0 to ubound(gDatos1,2)
						Total = Total + 1
					next
					'response.write "<br> Cantidad (bien):" & Cantidad
					'response.write "<br> Total:" & Total & "<br>"
					Valor = FormatNumber(((Total*100)/Cantidad),2)
				end if
			end if

		Case 33 'CompRel  
			Valor = 0
		
		Case 34 'PenAcum
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Id_Hogar AS Total "
			sql = sql & " FROM "
			sql = sql & " PH_DataCrudaMensual "
			sql = sql & " WHERE "
			'sql = sql & " Id_Categoria = " & sCat
			sql = sql & " id_Semana in( " & idSemana & ")"
			sql = sql & " GROUP BY "
			sql = sql & " PH_DataCrudaMensual.Id_Hogar "
			'response.write "<br>36 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			'response.write "<br>257 LLEGO" 
			'response.end
			if rsx1.eof then
				rsx1.close
				Valor = 0
				Valor = FormatNumber(Valor,2)
			else
				'response.write "<br>84 LLEGO"
				'response.end
				gDatos1 = rsx1.GetRows
				rsx1.close
				Valor = 0
				Cantidad = 0
				for iDat = 0 to ubound(gDatos1,2)
					'Cantidad = gDatos1(0,0)
					Cantidad = Cantidad + 1
				next
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Id_Hogar AS Total "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				
				sql = sql & " Id_Fabricante = " & iFab
				if sMar <> "" then 
					sql = sql & " And Id_Marca = " & iMar 
				end if
				if sSeg <> "" then 
					sql = sql & " And Id_Segmento = " & iSeg 
				end if
				if sRan <> "" then 
					sql = sql & " And Id_RangoTamano = " & iRan
				end if
				sql = sql & " and Id_Categoria =  " & sCat
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				sql = sql & " PH_DataCrudaMensual.Id_Hogar "
				'response.write "<br>2072 sql:=" & sql & "<br>"
				'response.end
				rsx1.Open sql ,conexion
				if rsx1.eof then
					rsx1.close
					Total = 0
					Valor = 0
					Valor = FormatNumber(Valor,2)
				else
					'response.write "<br>84 LLEGO"
					'response.end
					gDatos1 = rsx1.GetRows
					rsx1.close
					Total = 0
					for iDat = 0 to ubound(gDatos1,2)
						Total = Total + 1
					next
					'response.write "<br> Cantidad:" & Cantidad
					'response.write "<br> Total:" & Total & "<br>"
					Valor = FormatNumber(((Total*100)/Cantidad),2)
				end if
			end if
			
		Case 35 'HogRecomp
			if idSemana = "16,17,18,19,20,21,22,23,24,25,26,27,28" or idSemana = "20,21,22,23,24,25,26,27,28,29,30,31,32" or idSemana = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32" or idSemana = "24,25,26,27,28,29,30,31,32,33,34,35,36" then
				Valor = "N/A"
			else
				isw = 0
				if idSemana = "16,17,18,19" then 
					isw = 0
				end if 
				if idSemana = "20,21,22,23" then 
					isw = 1
					idSemana1 = "16,17,18,19"
					idSemana2 = "20,21,22,23"
				end if 
				if idSemana = "24,25,26,27,28" then 
					isw = 2
					idSemana1 = "20,21,22,23"
					idSemana2 = "24,25,26,27,28"
				end if 
				if idSemana = "29,30,31,32" then 
					isw = 2
					idSemana1 = "24,25,26,27,28"
					idSemana2 = "29,30,31,32"
				end if 
				if idSemana = "33,34,35,36" then 
					isw = 2
					idSemana1 = "29,30,31,32"
					idSemana2 = "33,34,35,36"
				end if 
				if idSemana = "37,38,39,40" then 
					isw = 2
					idSemana1 = "33,34,35,36"
					idSemana2 = "37,38,39,40"
				end if 
				if idSemana = "41,42,43,44,45" then 
					isw = 2
					idSemana1 = "37,38,39,40"
					idSemana2 = "41,42,43,44,45"
				end if 
				if idSemana = "46,47,48,49" then 
					isw = 2
					idSemana1 = "41,42,43,44,45"
					idSemana2 = "46,47,48,49"
				end if 
				if idSemana = "41,42,43,44,45,46,47,48,49,50,51,52,53,54" then 
					isw = 1
					'Abr-May-Jun
					idSemana1 = "29,30,31,32,33,34,35,36,37,38,39,40" 
					'Jul-Ago-Sep
					idSemana2 = "41,42,43,44,45,46,47,48,49,50,51,52,53,54" 
					'response.write "<br> Paso2222222222222222222"
					'response.end 
				end if 
				if idSemana = "50,51,52,53,54" then 
					isw = 2
					idSemana1 = "46,47,48,49"
					idSemana2 = "50,51,52,53,54"
				end if 
				if idSemana = "55,56,57,58" then 
					isw = 2
					idSemana1 = "50,51,52,53,54"
					idSemana2 = "55,56,57,58"
				end if 
				if idSemana = "59,60,61,62" then 
					isw = 2
					idSemana1 = "55,56,57,58"
					idSemana2 = "59,60,61,62"
				end if 
				if idSemana = "63,64,65,66,67" then 
					isw = 2
					idSemana1 = "59,60,61,62"
					idSemana2 = "63,64,65,66,67"
				end if 
				if idSemana = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67" then 
					isw = 2
					idSemana1 = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
					idSemana2 = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
				end if 
				if idSemana = "68,69,70,71" then 
					isw = 2
					idSemana1 = "63,64,65,66,67"
					idSemana2 = "68,69,70,71"
				end if 
				if idSemana = "55,56,57,58,59,60,61,62,63,64,65,66,67" then 
					isw = 2
					idSemana1 = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
					idSemana2 = "55,56,57,58,59,60,61,62,63,64,65,66,67"
				end if 
				if idSemana = "41,42,43,44,45,46,47,48,49,50,51,52,53,54" then 
					isw = 2
					idSemana1 = "29,30,31,32,33,34,35,36,37,38,39,40"
					idSemana2 = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
				end if 
				if idSemana = "29,30,31,32,33,34,35,36,37,38,39,40" then 
					isw = 2
					idSemana1 = "16,17,18,19,20,21,22,23,24,25,26,27,28"
					idSemana2 = "29,30,31,32,33,34,35,36,37,38,39,40"
				end if 
				
				'response.write "<br> idSemana:= " & idSemana
				'response.write "<br> iMes:= " & iMes
				if isw <> 0 Then 
					'response.write "<br> PASO 3643 <br>"
					'Mes Anterior
					dim vHogares35(10000,2)
					
					if iAre <> 0 then
						'response.write "Paso1"
						sql = ""
						sql = sql & " SELECT "
						sql = sql & " Id_Hogar "
						sql = sql & " FROM "
						sql = sql & " PH_DataCrudaMensual "
						sql = sql & " WHERE "
						sql = sql & " Id_Categoria = " & sCat
						sql = sql & " And Id_Area = " & iAre
						if sFab <> "" then 
							sql = sql & " And Id_Fabricante = " & iFab 
						end if
						if sMar <> "" then 
							sql = sql & " And Id_Marca = " & iMar 
						end if
						if sSeg <> "" then 
							sql = sql & " And Id_Segmento = " & iSeg 
						end if
						if sRan <> "" then 
							sql = sql & " And Id_RangoTamano = " & iRan
						end if
						if sTam <> "" then 
							sql = sql & " And Id_Tamano = " & iTam
						end if
						sql = sql & " And id_Semana in( " & idSemana1 & ")"
						sql = sql & " GROUP BY "
						sql = sql & " Id_Hogar "
						'response.write "<br>1095 sql:=" & sql
					else
						sql = ""
						sql = sql & " SELECT "
						sql = sql & " Id_Hogar "
						sql = sql & " FROM "
						sql = sql & " PH_DataCrudaMensual "
						sql = sql & " WHERE "
						sql = sql & " Id_Categoria = " & sCat
						sql = sql & " And Id_Area = 0 " 
						if sFab <> "" then 
							sql = sql & " And Id_Fabricante = " & iFab 
						end if
						if sMar <> "" then 
							sql = sql & " And Id_Marca = " & iMar 
						end if
						if sSeg <> "" then 
							sql = sql & " And Id_Segmento = " & iSeg 
						end if
						if sRan <> "" then 
							sql = sql & " And Id_RangoTamano = " & iRan
						end if
						if sTam <> "" then 
							sql = sql & " And Id_Tamano = " & iTam
						end if
						sql = sql & " And id_Semana in( " & idSemana1 & ")"
						sql = sql & " GROUP BY "
						sql = sql & " Id_Hogar "
						'response.write "<br>1893 sql:=" & sql
						'response.end
					end if
					'response.end
					rsx1.Open sql ,conexion
					'response.write "<br>257 LLEGO" 
					'response.end
					if rsx1.eof then
						rsx1.close
						Valor = 0
						Valor = FormatNumber(Valor,2)
					else
						'response.write "<br>3725 LLEGO"
						'response.end
						gDatos1 = rsx1.GetRows
						rsx1.close
						Valor = 0
						
						for iDat = 0 to ubound(gDatos1,2)
							Hogar = gDatos1(0,iDat)
							vHogares35(Hogar,1) = 1
						next
						
						'Mes Actual
						if iAre <> 0 then
							'response.write "Paso1"
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " Id_Hogar "
							sql = sql & " FROM "
							sql = sql & " PH_DataCrudaMensual "
							sql = sql & " WHERE "
							sql = sql & " Id_Categoria = " & sCat
							sql = sql & " And Id_Area = " & iAre
							if sFab <> "" then 
								sql = sql & " And Id_Fabricante = " & iFab 
							end if
							if sMar <> "" then 
								sql = sql & " And Id_Marca = " & iMar 
							end if
							if sSeg <> "" then 
								sql = sql & " And Id_Segmento = " & iSeg 
							end if
							if sRan <> "" then 
								sql = sql & " And Id_RangoTamano = " & iRan
							end if
							if sTam <> "" then 
								sql = sql & " And Id_Tamano = " & iTam
							end if
							sql = sql & " And id_Semana in( " & idSemana2 & ")"
							sql = sql & " GROUP BY "
							sql = sql & " Id_Hogar "
							'response.write "<br>1095 sql:=" & sql
						else
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " Id_Hogar "
							sql = sql & " FROM "
							sql = sql & " PH_DataCrudaMensual "
							sql = sql & " WHERE "
							sql = sql & " Id_Categoria = " & sCat
							sql = sql & " And Id_Area = 0 " 
							if sFab <> "" then 
								sql = sql & " And Id_Fabricante = " & iFab 
							end if
							if sMar <> "" then 
								sql = sql & " And Id_Marca = " & iMar 
							end if
							if sSeg <> "" then 
								sql = sql & " And Id_Segmento = " & iSeg 
							end if
							if sRan <> "" then 
								sql = sql & " And Id_RangoTamano = " & iRan
							end if
							if sTam <> "" then 
								sql = sql & " And Id_Tamano = " & iTam
							end if
							sql = sql & " And id_Semana in( " & idSemana2 & ")"
							sql = sql & " GROUP BY "
							sql = sql & " Id_Hogar "
							'response.write "<br>1893 sql:=" & sql
							'response.end
						end if
						'response.end
						rsx1.Open sql ,conexion
						'response.write "<br>257 LLEGO" 
						'response.end
						if rsx1.eof then
							rsx1.close
							Valor = 0
							Valor = FormatNumber(Valor,2)
						else
							'response.write "<br>3725 LLEGO"
							'response.end
							gDatos1 = rsx1.GetRows
							rsx1.close
							for iDat = 0 to ubound(gDatos1,2)
								Hogar = gDatos1(0,iDat)
								vHogares35(Hogar,2) = 1
							next
							Valor = 0
							Total = 0
							Repite = 0
							for iReg = 1 to 10000
								if vHogares35(iReg,2) = 1 then Total = Total  + 1
								if vHogares35(iReg,1) = 1 and vHogares35(iReg,2) = 1 then Repite = Repite  + 1
							next 

							Valor = ((cdbl(Repite) * 100) / cdbl(Total))
							'response.write "<br><br>772 Total=" & Total
							'response.write "<br>773 Repite=" & Repite
							'response.write "<br>"
							Valor = FormatNumber(Valor,2)
						end if
					end if
				else
					Valor = "N/A"
					'Valor = FormatNumber(Valor,2)
				end if
			end if
		
		Case 36 'HogNuevos
			if idSemana = "16,17,18,19,20,21,22,23,24,25,26,27,28" or idSemana = "20,21,22,23,24,25,26,27,28,29,30,31,32" or idSemana = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32" or idSemana = "24,25,26,27,28,29,30,31,32,33,34,35,36" then
				Valor = "N/A"
			else
				isw = 0
				if idSemana = "16,17,18,19" then 
					isw = 0
				end if 
				if idSemana = "20,21,22,23" then 
					isw = 1
					idSemana1 = "16,17,18,19"
					idSemana2 = "20,21,22,23"
				end if 
				if idSemana = "24,25,26,27,28" then 
					isw = 2
					idSemana1 = "20,21,22,23"
					idSemana2 = "24,25,26,27,28"
				end if 
				if idSemana = "29,30,31,32" then 
					isw = 2
					idSemana1 = "24,25,26,27,28"
					idSemana2 = "29,30,31,32"
				end if 
				if idSemana = "33,34,35,36" then 
					isw = 2
					idSemana1 = "29,30,31,32"
					idSemana2 = "33,34,35,36"
				end if 
				if idSemana = "37,38,39,40" then 
					isw = 2
					idSemana1 = "33,34,35,36"
					idSemana2 = "37,38,39,40"
				end if 
				if idSemana = "41,42,43,44,45" then 
					isw = 2
					idSemana1 = "37,38,39,40"
					idSemana2 = "41,42,43,44,45"
				end if 
				if idSemana = "46,47,48,49" then 
					isw = 2
					idSemana1 = "41,42,43,44,45"
					idSemana2 = "46,47,48,49"
				end if 
				if idSemana = "41,42,43,44,45,46,47,48,49,50,51,52,53,54" then 
					isw = 1
					'Abr-May-Jun
					idSemana1 = "29,30,31,32,33,34,35,36,37,38,39,40" 
					'Jul-Ago-Sep
					idSemana2 = "41,42,43,44,45,46,47,48,49,50,51,52,53,54" 
					'response.write "<br> Paso2222222222222222222"
					'response.end 
				end if 
				if idSemana = "50,51,52,53,54" then 
					isw = 2
					idSemana1 = "46,47,48,49"
					idSemana2 = "50,51,52,53,54"
				end if 
				if idSemana = "55,56,57,58" then 
					isw = 2
					idSemana1 = "50,51,52,53,54"
					idSemana2 = "55,56,57,58"
				end if 
				if idSemana = "59,60,61,62" then 
					isw = 2
					idSemana1 = "55,56,57,58"
					idSemana2 = "59,60,61,62"
				end if 
				if idSemana = "63,64,65,66,67" then 
					isw = 2
					idSemana1 = "59,60,61,62"
					idSemana2 = "63,64,65,66,67"
				end if 
				if idSemana = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67" then 
					isw = 2
					idSemana1 = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
					idSemana2 = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
				end if 
				if idSemana = "68,69,70,71" then 
					isw = 2
					idSemana1 = "63,64,65,66,67"
					idSemana2 = "68,69,70,71"
				end if 
				if idSemana = "55,56,57,58,59,60,61,62,63,64,65,66,67" then 
					isw = 2
					idSemana1 = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
					idSemana2 = "55,56,57,58,59,60,61,62,63,64,65,66,67"
				end if 
				if idSemana = "41,42,43,44,45,46,47,48,49,50,51,52,53,54" then 
					isw = 2
					idSemana1 = "29,30,31,32,33,34,35,36,37,38,39,40"
					idSemana2 = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
				end if 
				if idSemana = "29,30,31,32,33,34,35,36,37,38,39,40" then 
					isw = 2
					idSemana1 = "16,17,18,19,20,21,22,23,24,25,26,27,28"
					idSemana2 = "29,30,31,32,33,34,35,36,37,38,39,40"
				end if 
				
				
				'response.write "<br> idSemana:= " & idSemana
				'response.write "<br> iMes:= " & iMes
				if isw <> 0 Then 
					
					'response.write "<br> PASO 3643 <br>"
					'Mes Anterior
					dim vHogares36(10000,2)
					
					if iAre <> 0 then
						'response.write "Paso1"
						sql = ""
						sql = sql & " SELECT "
						sql = sql & " Id_Hogar "
						sql = sql & " FROM "
						sql = sql & " PH_DataCrudaMensual "
						sql = sql & " WHERE "
						sql = sql & " Id_Categoria = " & sCat
						sql = sql & " And Id_Area = " & iAre
						if sFab <> "" then 
							sql = sql & " And Id_Fabricante = " & iFab 
						end if
						if sMar <> "" then 
							sql = sql & " And Id_Marca = " & iMar 
						end if
						if sSeg <> "" then 
							sql = sql & " And Id_Segmento = " & iSeg 
						end if
						if sRan <> "" then 
							sql = sql & " And Id_RangoTamano = " & iRan
						end if
						if sTam <> "" then 
							sql = sql & " And Id_Tamano = " & iTam
						end if
						sql = sql & " And id_Semana in( " & idSemana1 & ")"
						sql = sql & " GROUP BY "
						sql = sql & " Id_Hogar "
						'response.write "<br>1095 sql:=" & sql
					else
						sql = ""
						sql = sql & " SELECT "
						sql = sql & " Id_Hogar "
						sql = sql & " FROM "
						sql = sql & " PH_DataCrudaMensual "
						sql = sql & " WHERE "
						sql = sql & " Id_Categoria = " & sCat
						sql = sql & " And Id_Area = 0 " 
						if sFab <> "" then 
							sql = sql & " And Id_Fabricante = " & iFab 
						end if
						if sMar <> "" then 
							sql = sql & " And Id_Marca = " & iMar 
						end if
						if sSeg <> "" then 
							sql = sql & " And Id_Segmento = " & iSeg 
						end if
						if sRan <> "" then 
							sql = sql & " And Id_RangoTamano = " & iRan
						end if
						if sTam <> "" then 
							sql = sql & " And Id_Tamano = " & iTam
						end if
						sql = sql & " And id_Semana in( " & idSemana1 & ")"
						sql = sql & " GROUP BY "
						sql = sql & " Id_Hogar "
						'response.write "<br>1893 sql:=" & sql
						'response.end
					end if
					'response.end
					rsx1.Open sql ,conexion
					'response.write "<br>257 LLEGO" 
					'response.end
					if rsx1.eof then
						rsx1.close
						Valor = 0
						Valor = FormatNumber(Valor,2)
					else
						'response.write "<br>3725 LLEGO"
						'response.end
						gDatos1 = rsx1.GetRows
						rsx1.close
						Valor = 0
						
						for iDat = 0 to ubound(gDatos1,2)
							Hogar = gDatos1(0,iDat)
							vHogares36(Hogar,1) = 1
						next
						
						'Mes Actual
						if iAre <> 0 then
							'response.write "Paso1"
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " Id_Hogar "
							sql = sql & " FROM "
							sql = sql & " PH_DataCrudaMensual "
							sql = sql & " WHERE "
							sql = sql & " Id_Categoria = " & sCat
							sql = sql & " And Id_Area = " & iAre
							if sFab <> "" then 
								sql = sql & " And Id_Fabricante = " & iFab 
							end if
							if sMar <> "" then 
								sql = sql & " And Id_Marca = " & iMar 
							end if
							if sSeg <> "" then 
								sql = sql & " And Id_Segmento = " & iSeg 
							end if
							if sRan <> "" then 
								sql = sql & " And Id_RangoTamano = " & iRan
							end if
							if sTam <> "" then 
								sql = sql & " And Id_Tamano = " & iTam
							end if
							sql = sql & " And id_Semana in( " & idSemana2 & ")"
							sql = sql & " GROUP BY "
							sql = sql & " Id_Hogar "
							'response.write "<br>1095 sql:=" & sql
						else
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " Id_Hogar "
							sql = sql & " FROM "
							sql = sql & " PH_DataCrudaMensual "
							sql = sql & " WHERE "
							sql = sql & " Id_Categoria = " & sCat
							sql = sql & " And Id_Area = 0 " 
							if sFab <> "" then 
								sql = sql & " And Id_Fabricante = " & iFab 
							end if
							if sMar <> "" then 
								sql = sql & " And Id_Marca = " & iMar 
							end if
							if sSeg <> "" then 
								sql = sql & " And Id_Segmento = " & iSeg 
							end if
							if sRan <> "" then 
								sql = sql & " And Id_RangoTamano = " & iRan
							end if
							if sTam <> "" then 
								sql = sql & " And Id_Tamano = " & iTam
							end if
							sql = sql & " And id_Semana in( " & idSemana2 & ")"
							sql = sql & " GROUP BY "
							sql = sql & " Id_Hogar "
							'response.write "<br>1893 sql:=" & sql
							'response.end
						end if
						'response.end
						rsx1.Open sql ,conexion
						'response.write "<br>257 LLEGO" 
						'response.end
						if rsx1.eof then
							rsx1.close
							Valor = 0
							Valor = FormatNumber(Valor,2)
						else
							'response.write "<br>3725 LLEGO"
							'response.end
							gDatos1 = rsx1.GetRows
							rsx1.close
							for iDat = 0 to ubound(gDatos1,2)
								Hogar = gDatos1(0,iDat)
								vHogares36(Hogar,2) = 1
							next
							Valor = 0
							Total = 0
							Nuevo = 0
							for iReg = 1 to 10000
								if vHogares36(iReg,2) = 1 then Total = Total  + 1
								if (vHogares36(iReg,1) = 0 or vHogares36(iReg,1) = "") and vHogares36(iReg,2) = 1 then Nuevo = Nuevo  + 1
							next 

							Valor = ((cdbl(Nuevo) * 100) / cdbl(Total))
							'response.write "<br><br>772 Total=" & Total
							'response.write "<br>773 Nuevo=" & Nuevo
							'response.write "<br>"
							Valor = FormatNumber(Valor,2)
						end if
					end if
				else
					Valor = "N/A"
					'Valor = 0
					'Valor = FormatNumber(Valor,2)
				end if
			end if

		Case 37 'HogNoRecomp
			if idSemana = "16,17,18,19,20,21,22,23,24,25,26,27,28" or idSemana = "20,21,22,23,24,25,26,27,28,29,30,31,32" or idSemana = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32" or idSemana = "24,25,26,27,28,29,30,31,32,33,34,35,36" then
				Valor = "N/A"
			else
				isw = 0
				if idSemana = "16,17,18,19" then 
					isw = 0
				end if 
				if idSemana = "20,21,22,23" then 
					isw = 1
					idSemana1 = "16,17,18,19"
					idSemana2 = "20,21,22,23"
				end if 
				if idSemana = "24,25,26,27,28" then 
					isw = 2
					idSemana1 = "20,21,22,23"
					idSemana2 = "24,25,26,27,28"
				end if 
				if idSemana = "29,30,31,32" then 
					isw = 2
					idSemana1 = "24,25,26,27,28"
					idSemana2 = "29,30,31,32"
				end if 
				if idSemana = "33,34,35,36" then 
					isw = 2
					idSemana1 = "29,30,31,32"
					idSemana2 = "33,34,35,36"
				end if 
				if idSemana = "37,38,39,40" then 
					isw = 2
					idSemana1 = "33,34,35,36"
					idSemana2 = "37,38,39,40"
				end if 
				if idSemana = "41,42,43,44,45" then 
					isw = 2
					idSemana1 = "37,38,39,40"
					idSemana2 = "41,42,43,44,45"
				end if 
				if idSemana = "46,47,48,49" then 
					isw = 2
					idSemana1 = "41,42,43,44,45"
					idSemana2 = "46,47,48,49"
				end if 
				if idSemana = "41,42,43,44,45,46,47,48,49,50,51,52,53,54" then 
					isw = 1
					'Abr-May-Jun
					idSemana1 = "29,30,31,32,33,34,35,36,37,38,39,40" 
					'Jul-Ago-Sep
					idSemana2 = "41,42,43,44,45,46,47,48,49,50,51,52,53,54" 
					'response.write "<br> Paso2222222222222222222"
					'response.end 
				end if 
				if idSemana = "50,51,52,53,54" then 
					isw = 2
					idSemana1 = "46,47,48,49"
					idSemana2 = "50,51,52,53,54"
				end if 
				if idSemana = "55,56,57,58" then 
					isw = 2
					idSemana1 = "50,51,52,53,54"
					idSemana2 = "55,56,57,58"
				end if 
				if idSemana = "59,60,61,62" then 
					isw = 2
					idSemana1 = "55,56,57,58"
					idSemana2 = "59,60,61,62"
				end if 
				if idSemana = "63,64,65,66,67" then 
					isw = 2
					idSemana1 = "59,60,61,62"
					idSemana2 = "63,64,65,66,67"
				end if 
				if idSemana = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67" then 
					isw = 2
					idSemana1 = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
					idSemana2 = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
				end if 
				if idSemana = "68,69,70,71" then 
					isw = 2
					idSemana1 = "63,64,65,66,67"
					idSemana2 = "68,69,70,71"
				end if 
				if idSemana = "55,56,57,58,59,60,61,62,63,64,65,66,67" then 
					isw = 2
					idSemana1 = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
					idSemana2 = "55,56,57,58,59,60,61,62,63,64,65,66,67"
				end if 
				if idSemana = "41,42,43,44,45,46,47,48,49,50,51,52,53,54" then 
					isw = 2
					idSemana1 = "29,30,31,32,33,34,35,36,37,38,39,40"
					idSemana2 = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
				end if 
				if idSemana = "29,30,31,32,33,34,35,36,37,38,39,40" then 
					isw = 2
					idSemana1 = "16,17,18,19,20,21,22,23,24,25,26,27,28"
					idSemana2 = "29,30,31,32,33,34,35,36,37,38,39,40"
				end if 
				
				'response.write "<br> idSemana:= " & idSemana
				'response.write "<br> iMes:= " & iMes
				if isw <> 0 Then 
					'response.write "<br> PASO 3643 <br>"
					'Mes Anterior
					dim vHogares37(10000,2)
					
					if iAre <> 0 then
						'response.write "Paso1"
						sql = ""
						sql = sql & " SELECT "
						sql = sql & " Id_Hogar "
						sql = sql & " FROM "
						sql = sql & " PH_DataCrudaMensual "
						sql = sql & " WHERE "
						sql = sql & " Id_Categoria = " & sCat
						sql = sql & " And Id_Area = " & iAre
						if sFab <> "" then 
							sql = sql & " And Id_Fabricante = " & iFab 
						end if
						if sMar <> "" then 
							sql = sql & " And Id_Marca = " & iMar 
						end if
						if sSeg <> "" then 
							sql = sql & " And Id_Segmento = " & iSeg 
						end if
						if sRan <> "" then 
							sql = sql & " And Id_RangoTamano = " & iRan
						end if
						if sTam <> "" then 
							sql = sql & " And Id_Tamano = " & iTam
						end if
						sql = sql & " And id_Semana in( " & idSemana1 & ")"
						sql = sql & " GROUP BY "
						sql = sql & " Id_Hogar "
						'response.write "<br>1095 sql:=" & sql
					else
						sql = ""
						sql = sql & " SELECT "
						sql = sql & " Id_Hogar "
						sql = sql & " FROM "
						sql = sql & " PH_DataCrudaMensual "
						sql = sql & " WHERE "
						sql = sql & " Id_Categoria = " & sCat
						sql = sql & " And Id_Area = 0 " 
						if sFab <> "" then 
							sql = sql & " And Id_Fabricante = " & iFab 
						end if
						if sMar <> "" then 
							sql = sql & " And Id_Marca = " & iMar 
						end if
						if sSeg <> "" then 
							sql = sql & " And Id_Segmento = " & iSeg 
						end if
						if sRan <> "" then 
							sql = sql & " And Id_RangoTamano = " & iRan
						end if
						if sTam <> "" then 
							sql = sql & " And Id_Tamano = " & iTam
						end if
						sql = sql & " And id_Semana in( " & idSemana1 & ")"
						sql = sql & " GROUP BY "
						sql = sql & " Id_Hogar "
						'response.write "<br>1893 sql:=" & sql
						'response.end
					end if
					'response.end
					rsx1.Open sql ,conexion
					'response.write "<br>257 LLEGO" 
					'response.end
					if rsx1.eof then
						rsx1.close
						Valor = 0
						Valor = FormatNumber(Valor,2)
					else
						'response.write "<br>3725 LLEGO"
						'response.end
						gDatos1 = rsx1.GetRows
						rsx1.close
						Valor = 0
						
						for iDat = 0 to ubound(gDatos1,2)
							Hogar = gDatos1(0,iDat)
							vHogares37(Hogar,1) = 1
						next
						
						'Mes Actual
						if iAre <> 0 then
							'response.write "Paso1"
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " Id_Hogar "
							sql = sql & " FROM "
							sql = sql & " PH_DataCrudaMensual "
							sql = sql & " WHERE "
							sql = sql & " Id_Categoria = " & sCat
							sql = sql & " And Id_Area = " & iAre
							if sFab <> "" then 
								sql = sql & " And Id_Fabricante = " & iFab 
							end if
							if sMar <> "" then 
								sql = sql & " And Id_Marca = " & iMar 
							end if
							if sSeg <> "" then 
								sql = sql & " And Id_Segmento = " & iSeg 
							end if
							if sRan <> "" then 
								sql = sql & " And Id_RangoTamano = " & iRan
							end if
							if sTam <> "" then 
								sql = sql & " And Id_Tamano = " & iTam
							end if
							sql = sql & " And id_Semana in( " & idSemana2 & ")"
							sql = sql & " GROUP BY "
							sql = sql & " Id_Hogar "
							'response.write "<br>1095 sql:=" & sql
						else
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " Id_Hogar "
							sql = sql & " FROM "
							sql = sql & " PH_DataCrudaMensual "
							sql = sql & " WHERE "
							sql = sql & " Id_Categoria = " & sCat
							sql = sql & " And Id_Area = 0 " 
							if sFab <> "" then 
								sql = sql & " And Id_Fabricante = " & iFab 
							end if
							if sMar <> "" then 
								sql = sql & " And Id_Marca = " & iMar 
							end if
							if sSeg <> "" then 
								sql = sql & " And Id_Segmento = " & iSeg 
							end if
							if sRan <> "" then 
								sql = sql & " And Id_RangoTamano = " & iRan
							end if
							if sTam <> "" then 
								sql = sql & " And Id_Tamano = " & iTam
							end if
							sql = sql & " And id_Semana in( " & idSemana2 & ")"
							sql = sql & " GROUP BY "
							sql = sql & " Id_Hogar "
							'response.write "<br>1893 sql:=" & sql
							'response.end
						end if
						'response.end
						rsx1.Open sql ,conexion
						'response.write "<br>257 LLEGO" 
						'response.end
						if rsx1.eof then
							rsx1.close
							Valor = 0
							Valor = FormatNumber(Valor,2)
						else
							'response.write "<br>3725 LLEGO"
							'response.end
							gDatos1 = rsx1.GetRows
							rsx1.close
							for iDat = 0 to ubound(gDatos1,2)
								Hogar = gDatos1(0,iDat)
								vHogares37(Hogar,2) = 1
							next
							Valor = 0
							Total = 0
							NoRepite = 0
							for iReg = 1 to 10000
								if vHogares37(iReg,1) = 1 then Total = Total  + 1
								if (vHogares37(iReg,2) = 0 or vHogares37(iReg,2) = "") and vHogares37(iReg,1) = 1 then NoRepite = NoRepite  + 1
							next 

							Valor = ((cdbl(NoRepite) * 100) / cdbl(Total))
							'response.write "<br><br>772 Total=" & Total
							'response.write "<br>773 NoRepite=" & NoRepite
							'response.write "<br>"
							Valor = FormatNumber(Valor,2)
						end if
					end if
				else
					Valor = "N/A"
					'Valor = 0
					'Valor = FormatNumber(Valor,2)
				end if
			end if

		Case 39 'HogRecompAnt
			if idSemana = "16,17,18,19,20,21,22,23,24,25,26,27,28" or idSemana = "20,21,22,23,24,25,26,27,28,29,30,31,32" or idSemana = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32" or idSemana = "24,25,26,27,28,29,30,31,32,33,34,35,36" then
				Valor = "N/A"
			else
				isw = 0
				if idSemana = "16,17,18,19" then 
					isw = 0
				end if 
				if idSemana = "20,21,22,23" then 
					isw = 1
					idSemana1 = "16,17,18,19"
					idSemana2 = "20,21,22,23"
				end if 
				if idSemana = "24,25,26,27,28" then 
					isw = 2
					idSemana1 = "20,21,22,23"
					idSemana2 = "24,25,26,27,28"
				end if 
				if idSemana = "29,30,31,32" then 
					isw = 2
					idSemana1 = "24,25,26,27,28"
					idSemana2 = "29,30,31,32"
				end if 
				if idSemana = "33,34,35,36" then 
					isw = 2
					idSemana1 = "29,30,31,32"
					idSemana2 = "33,34,35,36"
				end if 
				if idSemana = "37,38,39,40" then 
					isw = 2
					idSemana1 = "33,34,35,36"
					idSemana2 = "37,38,39,40"
				end if 
				if idSemana = "41,42,43,44,45" then 
					isw = 2
					idSemana1 = "37,38,39,40"
					idSemana2 = "41,42,43,44,45"
				end if 
				if idSemana = "46,47,48,49" then 
					isw = 2
					idSemana1 = "41,42,43,44,45"
					idSemana2 = "46,47,48,49"
				end if 
				if idSemana = "41,42,43,44,45,46,47,48,49,50,51,52,53,54" then 
					isw = 1
					'Abr-May-Jun
					idSemana1 = "29,30,31,32,33,34,35,36,37,38,39,40" 
					'Jul-Ago-Sep
					idSemana2 = "41,42,43,44,45,46,47,48,49,50,51,52,53,54" 
					'response.write "<br> Paso2222222222222222222"
					'response.end 
				end if 
				if idSemana = "50,51,52,53,54" then 
					isw = 2
					idSemana1 = "46,47,48,49"
					idSemana2 = "50,51,52,53,54"
				end if 
				if idSemana = "55,56,57,58" then 
					isw = 2
					idSemana1 = "50,51,52,53,54"
					idSemana2 = "55,56,57,58"
				end if 
				if idSemana = "59,60,61,62" then 
					isw = 2
					idSemana1 = "55,56,57,58"
					idSemana2 = "59,60,61,62"
				end if 
				if idSemana = "63,64,65,66,67" then 
					isw = 2
					idSemana1 = "59,60,61,62"
					idSemana2 = "63,64,65,66,67"
				end if  
				if idSemana = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67" then 
					isw = 2
					idSemana1 = "16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40"
					idSemana2 = "41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67"
				end if 
				if idSemana = "68,69,70,71" then 
					isw = 2
					idSemana1 = "63,64,65,66,67"
					idSemana2 = "68,69,70,71"
				end if 
				if idSemana = "55,56,57,58,59,60,61,62,63,64,65,66,67" then 
					isw = 2
					idSemana1 = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
					idSemana2 = "55,56,57,58,59,60,61,62,63,64,65,66,67"
				end if 
				if idSemana = "41,42,43,44,45,46,47,48,49,50,51,52,53,54" then 
					isw = 2
					idSemana1 = "29,30,31,32,33,34,35,36,37,38,39,40"
					idSemana2 = "41,42,43,44,45,46,47,48,49,50,51,52,53,54"
				end if 
				if idSemana = "29,30,31,32,33,34,35,36,37,38,39,40" then 
					isw = 2
					idSemana1 = "16,17,18,19,20,21,22,23,24,25,26,27,28"
					idSemana2 = "29,30,31,32,33,34,35,36,37,38,39,40"
				end if 
 				
				'response.write "<br> idSemana:= " & idSemana
				'response.write "<br> iMes:= " & iMes
				if isw <> 0 Then 
					'response.write "<br> PASO 3643 <br>"
					'Mes Anterior
					dim vHogares39(10000,2)
					
					if iAre <> 0 then
						'response.write "Paso1"
						sql = ""
						sql = sql & " SELECT "
						sql = sql & " Id_Hogar "
						sql = sql & " FROM "
						sql = sql & " PH_DataCrudaMensual "
						sql = sql & " WHERE "
						sql = sql & " Id_Categoria = " & sCat
						sql = sql & " And Id_Area = " & iAre
						if sFab <> "" then 
							sql = sql & " And Id_Fabricante = " & iFab 
						end if
						if sMar <> "" then 
							sql = sql & " And Id_Marca = " & iMar 
						end if
						if sSeg <> "" then 
							sql = sql & " And Id_Segmento = " & iSeg 
						end if
						if sRan <> "" then 
							sql = sql & " And Id_RangoTamano = " & iRan
						end if
						if sTam <> "" then 
							sql = sql & " And Id_Tamano = " & iTam
						end if
						sql = sql & " And id_Semana in( " & idSemana1 & ")"
						sql = sql & " GROUP BY "
						sql = sql & " Id_Hogar "
						'response.write "<br>1095 sql:=" & sql
					else
						sql = ""
						sql = sql & " SELECT "
						sql = sql & " Id_Hogar "
						sql = sql & " FROM "
						sql = sql & " PH_DataCrudaMensual "
						sql = sql & " WHERE "
						sql = sql & " Id_Categoria = " & sCat
						sql = sql & " And Id_Area = 0 " 
						if sFab <> "" then 
							sql = sql & " And Id_Fabricante = " & iFab 
						end if
						if sMar <> "" then 
							sql = sql & " And Id_Marca = " & iMar 
						end if
						if sSeg <> "" then 
							sql = sql & " And Id_Segmento = " & iSeg 
						end if
						if sRan <> "" then 
							sql = sql & " And Id_RangoTamano = " & iRan
						end if
						if sTam <> "" then 
							sql = sql & " And Id_Tamano = " & iTam
						end if
						sql = sql & " And id_Semana in( " & idSemana1 & ")"
						sql = sql & " GROUP BY "
						sql = sql & " Id_Hogar "
						'response.write "<br>1893 sql:=" & sql
						'response.end
					end if
					'response.end
					rsx1.Open sql ,conexion
					'response.write "<br>257 LLEGO" 
					'response.end
					if rsx1.eof then
						rsx1.close
						Valor = 0
						Valor = FormatNumber(Valor,2)
					else
						'response.write "<br>3725 LLEGO"
						'response.end
						gDatos1 = rsx1.GetRows
						rsx1.close
						Valor = 0
						
						for iDat = 0 to ubound(gDatos1,2)
							Hogar = gDatos1(0,iDat)
							vHogares39(Hogar,1) = 1
						next
						
						'Mes Actual
						if iAre <> 0 then
							'response.write "Paso1"
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " Id_Hogar "
							sql = sql & " FROM "
							sql = sql & " PH_DataCrudaMensual "
							sql = sql & " WHERE "
							sql = sql & " Id_Categoria = " & sCat
							sql = sql & " And Id_Area = " & iAre
							if sFab <> "" then 
								sql = sql & " And Id_Fabricante = " & iFab 
							end if
							if sMar <> "" then 
								sql = sql & " And Id_Marca = " & iMar 
							end if
							if sSeg <> "" then 
								sql = sql & " And Id_Segmento = " & iSeg 
							end if
							if sRan <> "" then 
								sql = sql & " And Id_RangoTamano = " & iRan
							end if
							if sTam <> "" then 
								sql = sql & " And Id_Tamano = " & iTam
							end if
							sql = sql & " And id_Semana in( " & idSemana2 & ")"
							sql = sql & " GROUP BY "
							sql = sql & " Id_Hogar "
							'response.write "<br>1095 sql:=" & sql
						else
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " Id_Hogar "
							sql = sql & " FROM "
							sql = sql & " PH_DataCrudaMensual "
							sql = sql & " WHERE "
							sql = sql & " Id_Categoria = " & sCat
							sql = sql & " And Id_Area = 0 " 
							if sFab <> "" then 
								sql = sql & " And Id_Fabricante = " & iFab 
							end if
							if sMar <> "" then 
								sql = sql & " And Id_Marca = " & iMar 
							end if
							if sSeg <> "" then 
								sql = sql & " And Id_Segmento = " & iSeg 
							end if
							if sRan <> "" then 
								sql = sql & " And Id_RangoTamano = " & iRan
							end if
							if sTam <> "" then 
								sql = sql & " And Id_Tamano = " & iTam
							end if
							sql = sql & " And id_Semana in( " & idSemana2 & ")"
							sql = sql & " GROUP BY "
							sql = sql & " Id_Hogar "
							'response.write "<br>1893 sql:=" & sql
							'response.end
						end if
						'response.end
						rsx1.Open sql ,conexion
						'response.write "<br>257 LLEGO" 
						'response.end
						if rsx1.eof then
							rsx1.close
							Valor = 0
							Valor = FormatNumber(Valor,2)
						else
							'response.write "<br>3725 LLEGO"
							'response.end
							gDatos1 = rsx1.GetRows
							rsx1.close
							for iDat = 0 to ubound(gDatos1,2)
								Hogar = gDatos1(0,iDat)
								vHogares39(Hogar,2) = 1
							next
							Valor = 0
							Total = 0
							Repite = 0
							for iReg = 1 to 10000
								if vHogares39(iReg,1) = 1 then Total = Total  + 1
								if vHogares39(iReg,1) = 1 and vHogares39(iReg,2) = 1 then Repite = Repite  + 1
							next 

							Valor = ((cdbl(Repite) * 100) / cdbl(Total))
							'response.write "<br><br>772 Total=" & Total
							'response.write "<br>773 Repite=" & Repite
							'response.write "<br>"
							Valor = FormatNumber(Valor,2)
						end if
					end if
				else
					Valor = "N/A"
					'Valor = 0
					'Valor = FormatNumber(Valor,2)
				end if
			end if

	end select 
end Sub



	'response.end
%>
