<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	'response.write "<br>84 LLEGO"
	'response.end
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
	dim gProductosTotal

	sCat=Request.QueryString("cat")
	if sCat = "" Then sCat = 1
	
	sAre=Request.QueryString("are")
	sFab=Request.QueryString("fab")
	sMar=Request.QueryString("mar")
	sSeg=Request.QueryString("seg")
	sRan=Request.QueryString("ran")
	sInd=Request.QueryString("ind")

	'26Ene2021-8
	TotalFabricante = "NO"
	if sFab <> "" then
		if Mid(sFab,1,1) = "0" then
			TotalFabricante = "SI"
			sFab = mid(sFab,2)
			if Mid(sFab,1,1) = "," then
				sFab = mid(sFab,2)
			end if
		end if
	end if
	
	if sSeg <> "" and sFab = "" and sMar = "" then 
		sFab = "0"
		sMar = "0"
	end if
	
	
	
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

	
	idSemana = 16
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " IdSemana, "
	sql = sql & " Semana "
	sql = sql & " FROM "
	sql = sql & " ss_Semana "
	sql = sql & " WHERE "
	sql = sql & " IdSemana = " & idSemana
	'response.write "<br>36 sql:=" & sql
	'response.end
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gDatos1 = rsx1.GetRows
		rsx1.close
		strSemana = gDatos1(1,0)
	end if
	strSemana1 = "(01) Del 04 Ene 2021 al 10 Ene 2021"
	strSemana2 = "(02) Del 11 Ene 2021 al 17 Ene 2021"
	strSemana3 = "(03) Del 18 Ene 2021 al 24 Ene 2021"
	strSemana4 = "(04) Del 25 Ene 2021 al 31 Ene 2021"
	strSemana5 = "(05) Del 01 Feb 2021 al 07 Feb 2021"
	strSemana6 = "(06) Del 08 Feb 2021 al 14 Feb 2021"
	strSemana7 = "(07) Del 15 Feb 2021 al 21 Feb 2021"
	strSemana8 = "(08) Del 22 Feb 2021 al 28 Feb 2021"
	strSemana9 = "(09) Del 01 Mar 2021 al 07 Mar 2021"
	strSemana10 = "(10) Del 08 Mar 2021 al 14 Mar 2021"
	strSemana11 = "(11) Del 15 Mar 2021 al 21 Mar 2021"
	strSemana12 = "(12) Del 22 Mar 2021 al 28 Mar 2021"
	strSemana13 = "(13) Del 29 Mar 2021 al 04 Abr 2021"
	strSemana14 = "(14) Del 05 Abr 2021 al 11 Abr 2021"
	strSemana15 = "(15) Del 12 Abr 2021 al 18 Abr 2021"
	strSemana16 = "(16) Del 19 Abr 2021 al 25 Abr 2021"
	strSemana17 = "(17) Del 26 Abr 2021 al 02 May 2021"
	strSemana18 = "(18) Del 03 May 2021 al 09 May 2021"
	strSemana19 = "(19) Del 10 May 2021 al 16 May 2021"
	strSemana20 = "(20) Del 17 May 2021 al 23 May 2021"
	strSemana21 = "(21) Del 24 May 2021 al 30 May 2021"
	strSemana22 = "(22) Del 31 May 2021 al 06 Jun 2021"
	strSemana23 = "(23) Del 07 Jun 2021 al 13 Jun 2021"
	strSemana24 = "(24) Del 14 Jun 2021 al 20 Jun 2021"
	strSemana25 = "(25) Del 21 Jun 2021 al 27 Jun 2021"
	strSemana26 = "(26) Del 28 Jun 2021 al 04 Jul 2021"
	strSemana27 = "(27) Del 05 Jul 2021 al 11 Jul 2021"
	strSemana28 = "(28) Del 12 Jul 2021 al 18 Jul 2021"
	strSemana29 = "(29) Del 19 Jul 2021 al 25 Jul 2021"
	strSemana30 = "(30) Del 26 Jul 2021 al 01 Ago 2021"
	strSemana31 = "(31) Del 02 Ago 2021 al 08 Ago 2021"
	strSemana32 = "(32) Del 09 Ago 2021 al 15 Ago 2021"
	strSemana33 = "(33) Del 16 Ago 2021 al 22 Ago 2021"
	strSemana34 = "(34) Del 23 Ago 2021 al 29 Ago 2021"
	strSemana35 = "(35) Del 30 Ago 2021 al 05 Sep 2021"
	strSemana36 = "(36) Del 06 Sep 2021 al 12 Sep 2021"
	strSemana37 = "(37) Del 13 Sep 2021 al 19 Sep 2021"
	strSemana38 = "(38) Del 20 Sep 2021 al 26 Sep 2021"
	strSemana39 = "(39) Del 27 Sep 2021 al 03 Oct 2021"
	strSemana40 = "(40) Del 04 Oct 2021 al 10 Oct 2021"
	strSemana41 = "(41) Del 11 Oct 2021 al 17 Oct 2021"
	strSemana42 = "(42) Del 18 Oct 2021 al 24 Oct 2021"
	strSemana43 = "(43) Del 25 Oct 2021 al 31 Oct 2021"	
	strSemana44 = "(44) Del 01 Nov 2021 al 07 Nov 2021"
	strSemana45 = "(45) Del 08 Nov 2021 al 14 Nov 2021"
	strSemana46 = "(46) Del 15 Nov 2021 al 21 Nov 2021"
	strSemana47 = "(47) Del 22 Nov 2021 al 28 Nov 2021"
	strSemana48 = "(48) Del 29 Nov 2021 al 05 Dic 2021"
	strSemana49 = "(49) Del 06 Dic 2021 al 12 Dic 2021"	
	strSemana50 = "(50) Del 13 Dic 2021 al 19 Dic 2021"	
	strSemana51 = "(51) Del 20 Dic 2021 al 26 Dic 2021"
	strSemana52 = "(52) Del 27 Dic 2021 al 02 Ene 2022"
	strSemana53 = "(01) Del 03 Ene 2022 al 09 Ene 2022"
	strSemana54 = "(02) Del 10 Ene 2022 al 16 Ene 2022"	
	strSemana55 = "(03) Del 17 Ene 2022 al 23 Ene 2022"
	strSemana56 = "(04) Del 24 Ene 2022 al 30 Ene 2022"
	strSemana57 = "(05) Del 31 Ene 2022 al 06 Feb 2022"
	strSemana58 = "(06) Del 07 Feb 2022 al 13 Feb 2022"
	'if Session("perusu") = 5 then
		'strSemana4 = "Acum Sem 1+2+3"
	'else
		
	'end if

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Indicador, "
	sql = sql & " Abreviatura, "
	sql = sql & " UnidadMedida "
	sql = sql & " FROM "
	sql = sql & " PH_Indicadores "
	sql = sql & " WHERE "
	if Session("perusu") = 5 then
		sql = sql & " Ind_Sem = 1 " 
	else
		sql = sql & " Ind_Activo = 1 " 
	end if
	if sInd <> "" then
		sql = sql & " And Id_Indicador in (" & sInd & ")"
	end if
	sql = sql & " ORDER BY "
	sql = sql & " Id_Indicador "
	'response.write "<br>372 Combo1:=" & sql
	'response.end 
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gIndicadores = rsx1.GetRows
		rsx1.close
	end if
	
	'26Ene2021-Todo Query
	sql = ""
    sql = sql & " SELECT "
	sql = sql & " Id_Fabricante, "
	sql = sql & " Fabricante, "
	sql = sql & " Id_Marca, "
	sql = sql & " Marca, "
	sql = sql & " Id_Segmento, "
	sql = sql & " Segmento "
	sql = sql & " FROM "
	sql = sql & " PH_DataCrudaMed"
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria = " & sCat
	sql = sql & " GROUP BY "
	sql = sql & " Id_Fabricante, "
	sql = sql & " Fabricante, "
	sql = sql & " Id_Marca, "
	sql = sql & " Marca, "
	sql = sql & " Id_Segmento, "
	sql = sql & " Segmento "
	sql = sql & " HAVING "
	sql = sql & " Id_Fabricante = 0 "
	sql = sql & " AND Id_Marca = 0 "
	sql = sql & " AND Id_Segmento = 0 "
	sql = sql & " ORDER BY "
	sql = sql & " Id_Fabricante "
	'response.write "<br>36 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	iExiste = 0
	'response.write "<br>84 LLEGO"
	'response.end
	if rsx1.eof then
		rsx1.close
	else
		gProductosTotal = rsx1.GetRows
		rsx1.close
	end if
	
	'response.write "<br>372 sFab:=" & sFab
	sql = ""
    sql = sql & " SELECT "
	sql = sql & " Id_Fabricante, "	'0
	sql = sql & " Fabricante "		'1
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

	sql = sql & " FROM PH_DataCrudaMed "

	sql = sql & " WHERE "
	sql = sql & " Id_Categoria = " & sCat

	sql = sql & " GROUP BY "
	sql = sql & " Id_Fabricante, "
	sql = sql & " Fabricante "
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
	isw = 0
	if sFab <> "" then
		if isw = 0 then
			sql = sql & " HAVING "
			isw = 1
		else
			sql = sql & " AND "
		end if
		sql = sql & " Id_Fabricante in (" & sFab & ")"
	else
		if isw = 0 then
			sql = sql & " HAVING "
			isw = 1
		else
			sql = sql & " AND "
		end if
		sql = sql & " Id_Fabricante <>0 "
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

	sql = sql & " ORDER BY "
	sql = sql & " Id_Fabricante "
	if sMar <> "" then
		sql = sql & " ,Id_Marca "
	end if
	if sSeg <> "" then
		sql = sql & " ,Id_Segmento "
	end if
	if sRan <> "" then
		sql = sql & " ,Id_RangoTamano "
	end if

	'response.write "<br>36 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
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
											<th class="cell100 column1 text-left">Fabricante</th>
											<th class="cell100 column2 text-center">Marca</th>
											<th class="cell100 column3 text-center">Segmento</th>
											<th class="cell100 column4 text-center"></th>
											<th class="cell100 column5 text-center">Indicador</th>
											<th class="cell100 column6 text-center">UniMed</th>
											<th class="cell100 column7 text-center"><%=strSemana1%></th>
											<th class="cell100 column8 text-center"><%=strSemana2%></th>									
											<th class="cell100 column9 text-center"><%=strSemana3%></th>									
											<th class="cell100 column10 text-center"><%=strSemana4%></th>
											<th class="cell100 column10 text-center"><%=strSemana5%></th>
											<th class="cell100 column10 text-center"><%=strSemana6%></th>
											<th class="cell100 column10 text-center"><%=strSemana7%></th>
											<th class="cell100 column10 text-center"><%=strSemana8%></th>
											<th class="cell100 column10 text-center"><%=strSemana9%></th>
											<th class="cell100 column10 text-center"><%=strSemana10%></th>
											<th class="cell100 column10 text-center"><%=strSemana11%></th>
											<th class="cell100 column10 text-center"><%=strSemana12%></th>
											<th class="cell100 column10 text-center"><%=strSemana13%></th>
											<th class="cell100 column10 text-center"><%=strSemana14%></th>
											<th class="cell100 column10 text-center"><%=strSemana15%></th>
											<th class="cell100 column10 text-center"><%=strSemana16%></th>
											<th class="cell100 column10 text-center"><%=strSemana17%></th>
											<th class="cell100 column10 text-center"><%=strSemana18%></th>
											<th class="cell100 column10 text-center"><%=strSemana19%></th>
											<th class="cell100 column10 text-center"><%=strSemana20%></th>
											<th class="cell100 column10 text-center"><%=strSemana21%></th>
											<th class="cell100 column10 text-center"><%=strSemana22%></th>
											<th class="cell100 column10 text-center"><%=strSemana23%></th>
											<th class="cell100 column10 text-center"><%=strSemana24%></th>
											<th class="cell100 column10 text-center"><%=strSemana25%></th>
											<th class="cell100 column10 text-center"><%=strSemana26%></th>
											<th class="cell100 column10 text-center"><%=strSemana27%></th>
											<th class="cell100 column10 text-center"><%=strSemana28%></th>
											<th class="cell100 column10 text-center"><%=strSemana29%></th>
											<th class="cell100 column10 text-center"><%=strSemana30%></th>
											<th class="cell100 column10 text-center"><%=strSemana31%></th>
											<th class="cell100 column10 text-center"><%=strSemana32%></th>
											<th class="cell100 column10 text-center"><%=strSemana33%></th>
											<th class="cell100 column10 text-center"><%=strSemana34%></th>
											<th class="cell100 column10 text-center"><%=strSemana35%></th>
											<th class="cell100 column10 text-center"><%=strSemana36%></th>
											<th class="cell100 column10 text-center"><%=strSemana37%></th>
											<th class="cell100 column10 text-center"><%=strSemana38%></th>
											<th class="cell100 column10 text-center"><%=strSemana39%></th>
											<th class="cell100 column10 text-center"><%=strSemana40%></th>
											<th class="cell100 column10 text-center"><%=strSemana41%></th>
											<th class="cell100 column10 text-center"><%=strSemana42%></th>
											<th class="cell100 column10 text-center"><%=strSemana43%></th>
											<th class="cell100 column10 text-center"><%=strSemana44%></th>
											<th class="cell100 column10 text-center"><%=strSemana45%></th>
											<th class="cell100 column10 text-center"><%=strSemana46%></th>
											<th class="cell100 column10 text-center"><%=strSemana47%></th>
											<th class="cell100 column10 text-center"><%=strSemana48%></th>
											<th class="cell100 column10 text-center"><%=strSemana49%></th>
											<th class="cell100 column10 text-center"><%=strSemana50%></th>
											<th class="cell100 column10 text-center"><%=strSemana51%></th>
											<th class="cell100 column10 text-center"><%=strSemana52%></th>
											<th class="cell100 column10 text-center"><%=strSemana53%></th>
											<th class="cell100 column10 text-center"><%=strSemana54%></th>
											<th class="cell100 column10 text-center"><%=strSemana55%></th>
											<th class="cell100 column10 text-center"><%=strSemana56%></th>
											<th class="cell100 column10 text-center"><%=strSemana57%></th>
											<th class="cell100 column10 text-center"><%=strSemana58%></th>
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
	Response.AddHeader "Content-disposition","attachment; filename=tem.xls"
	Response.ContentType = "application/vnd.ms-excel"
		%>
		<div class="limiter">
			
			<div class="container-table100">
			
				<div class="wrap-table100">
								
					<div class="table100 ver1 m-b-110">
						
							<div class="table100-head">
							
								<table border=0>
									<thead>
										<tr class="row100 head">
											<th class="cell100 column1 text-left">Fabricante</th>
											<th class="cell100 column2 text-center">Marca</th>
											<th class="cell100 column3 text-center">Segmento</th>
											<th class="cell100 column4 text-center"></th>
											<th class="cell100 column5 text-center">Indicador</th>
											<th class="cell100 column6 text-center">UniMed</th>
											<th class="cell100 column7 text-center"><%=strSemana1%></th>
											<th class="cell100 column8 text-center"><%=strSemana2%></th>									
											<th class="cell100 column9 text-center"><%=strSemana3%></th>									
											<th class="cell100 column10 text-center"><%=strSemana4%></th>
											<th class="cell100 column10 text-center"><%=strSemana5%></th>
											<th class="cell100 column10 text-center"><%=strSemana6%></th>
											<th class="cell100 column10 text-center"><%=strSemana7%></th>
											<th class="cell100 column10 text-center"><%=strSemana8%></th>
											<th class="cell100 column10 text-center"><%=strSemana9%></th>
											<th class="cell100 column10 text-center"><%=strSemana10%></th>
											<th class="cell100 column10 text-center"><%=strSemana11%></th>
											<th class="cell100 column10 text-center"><%=strSemana12%></th>
											<th class="cell100 column10 text-center"><%=strSemana13%></th>
											<th class="cell100 column10 text-center"><%=strSemana14%></th>
											<th class="cell100 column10 text-center"><%=strSemana15%></th>
											<th class="cell100 column10 text-center"><%=strSemana16%></th>
											<th class="cell100 column10 text-center"><%=strSemana17%></th>
											<th class="cell100 column10 text-center"><%=strSemana18%></th>
											<th class="cell100 column10 text-center"><%=strSemana19%></th>
											<th class="cell100 column10 text-center"><%=strSemana20%></th>
											<th class="cell100 column10 text-center"><%=strSemana21%></th>
											<th class="cell100 column10 text-center"><%=strSemana22%></th>
											<th class="cell100 column10 text-center"><%=strSemana23%></th>
											<th class="cell100 column10 text-center"><%=strSemana24%></th>
											<th class="cell100 column10 text-center"><%=strSemana25%></th>
											<th class="cell100 column10 text-center"><%=strSemana26%></th>
											<th class="cell100 column10 text-center"><%=strSemana27%></th>
											<th class="cell100 column10 text-center"><%=strSemana28%></th>
											<th class="cell100 column10 text-center"><%=strSemana29%></th>
											<th class="cell100 column10 text-center"><%=strSemana30%></th>
											<th class="cell100 column10 text-center"><%=strSemana31%></th>
											<th class="cell100 column10 text-center"><%=strSemana32%></th>
											<th class="cell100 column10 text-center"><%=strSemana33%></th>
											<th class="cell100 column10 text-center"><%=strSemana34%></th>
											<th class="cell100 column10 text-center"><%=strSemana35%></th>
											<th class="cell100 column10 text-center"><%=strSemana36%></th>
											<th class="cell100 column10 text-center"><%=strSemana37%></th>
											<th class="cell100 column10 text-center"><%=strSemana38%></th>
											<th class="cell100 column10 text-center"><%=strSemana39%></th>
											<th class="cell100 column10 text-center"><%=strSemana40%></th>
											<th class="cell100 column10 text-center"><%=strSemana41%></th>
											<th class="cell100 column10 text-center"><%=strSemana42%></th>
											<th class="cell100 column10 text-center"><%=strSemana43%></th>
											<th class="cell100 column10 text-center"><%=strSemana44%></th>
											<th class="cell100 column10 text-center"><%=strSemana45%></th>
											<th class="cell100 column10 text-center"><%=strSemana46%></th>
											<th class="cell100 column10 text-center"><%=strSemana47%></th>
											<th class="cell100 column10 text-center"><%=strSemana48%></th>
											<th class="cell100 column10 text-center"><%=strSemana49%></th>
											<th class="cell100 column10 text-center"><%=strSemana50%></th>
											<th class="cell100 column10 text-center"><%=strSemana51%></th>
											<th class="cell100 column10 text-center"><%=strSemana52%></th>
											<th class="cell100 column10 text-center"><%=strSemana53%></th>
											<th class="cell100 column10 text-center"><%=strSemana54%></th>
											<th class="cell100 column10 text-center"><%=strSemana55%></th>
											<th class="cell100 column10 text-center"><%=strSemana56%></th>
											<th class="cell100 column10 text-center"><%=strSemana57%></th>
											<th class="cell100 column10 text-center"><%=strSemana58%></th>
										</tr>
									</thead>
								</table>
								
							</div>
												
							<div class="table100-body js-pscroll">
								<table border=0>
									<tbody>					
										<% 
										'26Ene2021-Todo el IF
										if TotalFabricante = "SI" then 
											for iPro = 0 to  ubound(gProductosTotal,2)
												'Response.flush
												for iInd = 0 to  ubound(gIndicadores,2)
												response.write "<tr class='row100 body'>"
													'Fabricante
													response.write "<td width=15% class='cell100 column1'>"
														response.write gProductosTotal(1,iPro) 
													response.write "</td>"
													'Marca
													response.write "<td width=15% class='cell100 column1'>"
													response.write "</td>"
													'Segmento
													response.write "<td width=15% class='cell100 column1'>"
													response.write "</td>"
													'Rango
													response.write "<td width=15% class='cell100 column1'>"
													response.write "</td>"

													response.write "</td>"
														response.write "<td width=5% class='cell100 column5 text-center'>"
															response.write "<b>"
															'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
															response.write gIndicadores(1,iInd)
															response.write "</b>"
														response.write "</td>"
														response.write "<td width=5%  class='text-center'>"
															response.write "<b>"
															response.write gIndicadores(2,iInd)
															response.write "</b>"
														response.write "</td>"
														Indicador = gIndicadores(0,iInd)
														iFab = gProductosTotal(0,iPro)
														iMar = gProductosTotal(2,iPro)
														iSeg = gProductosTotal(4,iPro)
														'response.write "<br>Ind = " & Indicador
														idSemana = 16
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='cell100 column8 text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 17
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 18
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 19
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 20
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 21
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='cell100 column8 text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 22
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 23
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 24
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 25
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 26
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='cell100 column8 text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 27
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 28
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 29
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 30
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 31
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 32
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 33
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 34
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 35
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 36
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 37
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 38
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 39
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 40
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 41
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 42
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 43
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 44
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 45
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 46
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 47
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 48
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 49
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 50
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 51
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 52
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 53
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 54
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 55
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 56
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 57
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 58
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 59
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 60
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 61
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 62
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 63
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 64
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 65
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 66
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 67
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 68
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 69
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 70
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 71
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 72
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
														idSemana = 73
														TotalDias = 7
														CalcularIndicador
														response.write "<td width=10% class='text-right'>"
															response.write Valor
														response.write "</td>"
													response.write "</tr>"
												next
											next					
										end if
										if sFab = "" and sMar = "" and sRan = "" then
										
										else
										for iPro = 0 to  ubound(gProductos,2)
											'Response.flush
											for iInd = 0 to  ubound(gIndicadores,2)
											response.write "<tr class='row100 body'>"
												'Fabricante
												response.write "<td width=15% class='cell100 column1'>"
													response.write gProductos(1,iPro) 
												response.write "</td>"
												'Marca
												response.write "<td width=15% class='cell100 column2 text-center'>"
													if sMar <> "" then
														'if "TOTAL CATEGORIA" = trim(gProductos(1,iPro)) then
														'else
															response.write gProductos(3,iPro)
														'end if
													end if
												response.write "</td>"
												'Segmento
												response.write "<td width=10% class='cell100 column3 text-center'>"
													if sSeg <> "" then
														'if "TOTAL CATEGORIA" = trim(gProductos(1,iPro)) then
														'else
															response.write gProductos(5,iPro)
														'end if
													end if
												response.write "</td>"
												'Rango
												response.write "<td width=5% class='cell100 column4 text-center'>"
													if sRan <> "" then
														response.write gProductos(7,iPro)
													end if
												response.write "</td>"
											
													response.write "<td width=5% class='cell100 column5 text-center'>"
														response.write "<b>"
														'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
														response.write gIndicadores(1,iInd)
														response.write "</b>"
													response.write "</td>"
													response.write "<td width=5%  class='text-center'>"
														response.write "<b>"
														response.write gIndicadores(2,iInd) 
														response.write "</b>"
													response.write "</td>"
													Indicador = gIndicadores(0,iInd)
													iFab = gProductos(0,iPro)
													if sMar <> "" then
														iMar = gProductos(2,iPro)
													end if
													if sSeg <> "" then
														iSeg = gProductos(4,iPro)
													end if
													if sRan <> "" then
														iRan = gProductos(6,iPro)
													end if
													'response.write "<br>Ind = " & Indicador
													idSemana = 16
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='cell100 column8 text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 17
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 18
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 19
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 20
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 21
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='cell100 column8 text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 22
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 23
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 24
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 25
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 26
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='cell100 column8 text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 27
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 28
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 29
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 30
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 31
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 32
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 33
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 34
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 35
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 36
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 37
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 38
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 39
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 40
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 41
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 42
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 43
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 44
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 45
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 46
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 47
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 48
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 49
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 50
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 51
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 52
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 53
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 54
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 55
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 56
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 57
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 58
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 59
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 60
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 61
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 62
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 63
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 64
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 65
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 66
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 67
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 68
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 69
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 70
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 71
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 72
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
													idSemana = 73
													TotalDias = 7
													CalcularIndicador
													response.write "<td width=10% class='text-right'>"
														response.write Valor
													response.write "</td>"
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
				sql = sql & " PH_DataCrudaMed "
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
				sql = sql & " PH_DataCrudaMed "
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
				sql = sql & " PH_DataCrudaMed "
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
				sql = sql & " Cantidad, "
				sql = sql & " Precio_Producto, "
				sql = sql & " Dolar, "
				sql = sql & " Precio_producto/Dolar AS PrecioDolar, "
				sql = sql & " (Precio_producto/Dolar)*Cantidad AS ComprasValor "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMed "
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
				sql = sql & " PH_DataCrudaMed "
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
				sql = sql & " PH_DataCrudaMed "
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
				sql = sql & " FROM PH_DataCrudaMed "
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
				sql = sql & " FROM PH_DataCrudaMed "
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
			sql = sql & " PH_DataCrudaMed "
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
			sql = sql & " PH_DataCrudaMed.Id_Hogar "
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
		
		Case 6 'PenPor
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Id_Hogar AS Total "
			sql = sql & " FROM "
			sql = sql & " PH_DataCrudaMed "
			sql = sql & " WHERE "
			sql = sql & " id_Semana in( " & idSemana & ")"
			sql = sql & " GROUP BY "
			sql = sql & " PH_DataCrudaMed.Id_Hogar "
			
			sql = sql & "Union"

			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Id_Hogar AS Total "
			sql = sql & " FROM "
			sql = sql & " PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " id_Semana in( " & idSemana & ")"
			sql = sql & " GROUP BY "
			sql = sql & " PH_DataCruda.Id_Hogar "
			
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
				sql = sql & " PH_DataCrudaMed "
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
				sql = sql & " and Id_Categoria = " & sCat
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				sql = sql & " PH_DataCrudaMed.Id_Hogar "
				'response.write "<br>36 sql:=" & sql & "<br>"
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
		

		Case 7 'PenPonVol (NO VA - Es Mensual)
			Valor = 0
		Case 8 'PenPonVal  (NO VA - Es Mensual)
			Valor = 0

		Case 9 'CompraMedHog
			if iFab <> 0 then
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Tamano, "
				sql = sql & " Cantidad "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMed "
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
				'response.write "<br>933Paso 1" 
			else
				if iFab = 0 and iMar = 0 and iSeg <> 0 then
					sql = ""
					sql = sql & " SELECT "
					sql = sql & " Tamano, "
					sql = sql & " Cantidad "
					sql = sql & " FROM "
					sql = sql & " PH_DataCrudaMed "
					sql = sql & " WHERE "
					sql = sql & " Id_Categoria = " & sCat
					sql = sql & " And Id_Fabricante = 0 "
					sql = sql & " And Id_Marca =  0 "
					sql = sql & " And Id_Segmento = " & iSeg 
					sql = sql & " And id_Semana in( " & idSemana & ")"
				else
					sql = ""
					sql = sql & " SELECT "
					sql = sql & " Tamano, "
					sql = sql & " Cantidad "
					sql = sql & " FROM "
					sql = sql & " PH_DataCrudaMed "
					sql = sql & " WHERE "
					sql = sql & " Id_Categoria = " & sCat
					sql = sql & " And Id_Fabricante = 0 "
					sql = sql & " And Id_Marca = 0"
					sql = sql & " And Id_Segmento = 0"
					sql = sql & " And Id_RangoTamano = 0"
					sql = sql & " And id_Semana in( " & idSemana & ")"
					'response.write "<br>933Paso 2" 
				end if
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
				Indicador1 = FormatNumber((Valor)/1000,2)
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Id_Hogar AS Total "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMed "
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
				sql = sql & " GROUP BY "
				sql = sql & " PH_DataCrudaMed.Id_Hogar "
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
				Valor = (cdbl(Indicador1) / cdbl(Indicador5))
				'response.write "<br>772 Indicador1=" & Indicador1
				'response.write "<br>773 Indicador5=" & Indicador5
				Valor = FormatNumber(Valor,2)
			end if

				

		Case 10 'GastMedHog
			if iFab <> 0 then
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Cantidad, "
				sql = sql & " Precio_Producto, "
				sql = sql & " Dolar, "
				sql = sql & " Precio_producto/Dolar AS PrecioDolar, "
				sql = sql & " (Precio_producto/Dolar)*Cantidad AS ComprasValor "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMed "
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
				if iFab = 0 and iMar = 0 and iSeg <> 0 then
					sql = ""
					sql = sql & " SELECT "
					sql = sql & " Cantidad, "
					sql = sql & " Precio_Producto, "
					sql = sql & " Dolar, "
					sql = sql & " Precio_producto/Dolar AS PrecioDolar, "
					sql = sql & " (Precio_producto/Dolar)*Cantidad AS ComprasValor "
					sql = sql & " FROM "
					sql = sql & " PH_DataCrudaMed "
					sql = sql & " WHERE "
					sql = sql & " Id_Categoria = " & sCat
					sql = sql & " And Id_Fabricante = 0 "
					sql = sql & " And Id_Marca = 0 "
					sql = sql & " And Id_Segmento =  " & iSeg
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
					sql = sql & " PH_DataCrudaMed "
					sql = sql & " WHERE "
					sql = sql & " Id_Categoria = " & sCat
					sql = sql & " And Id_Fabricante = 0 "
					sql = sql & " And Id_Marca = 0 "
					sql = sql & " And Id_Segmento = 0 "
					sql = sql & " And Id_RangoTamano = 0 "
					sql = sql & " And id_Semana in( " & idSemana & ")"
				end if
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
				
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Id_Hogar AS Total "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMed "
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
				sql = sql & " GROUP BY "
				sql = sql & " PH_DataCrudaMed.Id_Hogar "
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
				'response.write "<br>36 Indicador1=" & Indicador2
				'response.write "<br>36 Indicador5=" & Indicador5
				Valor = FormatNumber(Valor,2) 
			end if

		Case 11 'UnidCompHog
			if iFab <> 0 then
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Cantidad "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMed "
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
				if iFab = 0 and iMar = 0 and iSeg <> 0 then
					sql = ""
					sql = sql & " SELECT "
					sql = sql & " Cantidad "
					sql = sql & " FROM "
					sql = sql & " PH_DataCrudaMed "
					sql = sql & " WHERE "
					sql = sql & " Id_Categoria = " & sCat
					sql = sql & " And Id_Fabricante = 0 "
					sql = sql & " And Id_Marca = 0 "
					sql = sql & " And Id_Segmento =  " & iSeg
					sql = sql & " And id_Semana in( " & idSemana & ")"
				else
					sql = ""
					sql = sql & " SELECT "
					sql = sql & " Cantidad "
					sql = sql & " FROM "
					sql = sql & " PH_DataCrudaMed "
					sql = sql & " WHERE "
					sql = sql & " Id_Categoria = " & sCat
					sql = sql & " And Id_Fabricante = 0 "
					sql = sql & " And Id_Marca = 0 "
					sql = sql & " And Id_Segmento = 0 "
					sql = sql & " And Id_RangoTamano = 0 "
					sql = sql & " And id_Semana in( " & idSemana & ")"
				end if
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
				Indicador3 = 0
				for iDat = 0 to ubound(gDatos1,2)
					Indicador3 = Indicador3 + gDatos1(0,iDat)
				next
				
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Id_Hogar AS Total "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMed "
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
				sql = sql & " GROUP BY "
				sql = sql & " PH_DataCrudaMed.Id_Hogar "
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

		Case 12 'ActComp
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad, "
			sql = sql & " Id_Consumo "
			sql = sql & " FROM PH_DataCrudaMed "
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
				sql = sql & " PH_DataCrudaMed "
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
				sql = sql & " GROUP BY "
				sql = sql & " PH_DataCrudaMed.Id_Hogar "
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

		Case 13 'CiCloComp
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad, "
			sql = sql & " Id_Consumo "
			sql = sql & " FROM PH_DataCrudaMed "
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
				sql = sql & " PH_DataCrudaMed "
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
				sql = sql & " GROUP BY "
				sql = sql & " PH_DataCrudaMed.Id_Hogar "
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
			if iFab <> 0 then
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Tamano, "
				sql = sql & " Cantidad "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMed "
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
				if iFab = 0 and iMar = 0 and iSeg <> 0 then
					sql = ""
					sql = sql & " SELECT "
					sql = sql & " Tamano, "
					sql = sql & " Cantidad "
					sql = sql & " FROM "
					sql = sql & " PH_DataCrudaMed "
					sql = sql & " WHERE "
					sql = sql & " Id_Categoria = " & sCat
					sql = sql & " And Id_Fabricante = 0 "
					sql = sql & " And Id_Marca = 0"
					sql = sql & " And Id_Segmento = " & iSeg
					sql = sql & " And id_Semana in( " & idSemana & ")"
				else
					sql = ""
					sql = sql & " SELECT "
					sql = sql & " Tamano, "
					sql = sql & " Cantidad "
					sql = sql & " FROM "
					sql = sql & " PH_DataCrudaMed "
					sql = sql & " WHERE "
					sql = sql & " Id_Categoria = " & sCat
					sql = sql & " And Id_Fabricante = 0 "
					sql = sql & " And Id_Marca = 0"
					sql = sql & " And Id_Segmento = 0"
					sql = sql & " And Id_RangoTamano = 0"
					sql = sql & " And id_Semana in( " & idSemana & ")"
				end if
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
				Valor = FormatNumber(Valor,2)
				Indicador1 = 0
				for iDat = 0 to ubound(gDatos1,2)
					Indicador1 = Indicador1 + (cdbl(gDatos1(0,iDat)) *cdbl(gDatos1(1,iDat)))
				next

				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Cantidad, "
				sql = sql & " Id_Consumo "
				sql = sql & " FROM PH_DataCrudaMed "
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
				
				Indicador14 = ((cdbl(Indicador1)/1000) / cdbl(Indicador4))

				Valor = FormatNumber(Indicador14,2)

			end if

		Case 15 'ValActoCompra
			if iFab <> 0 then
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Cantidad, "
				sql = sql & " Precio_Producto, "
				sql = sql & " Dolar, "
				sql = sql & " Precio_producto/Dolar AS PrecioDolar, "
				sql = sql & " (Precio_producto/Dolar)*Cantidad AS ComprasValor "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMed "
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
				if iFab = 0 and iMar = 0 and iSeg <> 0 then
					sql = ""
					sql = sql & " SELECT "
					sql = sql & " Cantidad, "
					sql = sql & " Precio_Producto, "
					sql = sql & " Dolar, "
					sql = sql & " Precio_producto/Dolar AS PrecioDolar, "
					sql = sql & " (Precio_producto/Dolar)*Cantidad AS ComprasValor "
					sql = sql & " FROM "
					sql = sql & " PH_DataCrudaMed "
					sql = sql & " WHERE "
					sql = sql & " Id_Categoria = " & sCat
					sql = sql & " And Id_Fabricante = 0 "
					sql = sql & " And Id_Marca = 0 "
					sql = sql & " And Id_Segmento = " & iSeg
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
					sql = sql & " PH_DataCrudaMed "
					sql = sql & " WHERE "
					sql = sql & " Id_Categoria = " & sCat
					sql = sql & " And Id_Fabricante = 0 "
					sql = sql & " And Id_Marca = 0 "
					sql = sql & " And Id_Segmento = 0 "
					sql = sql & " And Id_RangoTamano = 0 "
					sql = sql & " And id_Semana in( " & idSemana & ")"
				
				end if
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
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Cantidad, "
				sql = sql & " Id_Consumo "
				sql = sql & " FROM PH_DataCrudaMed "
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

				Valor = FormatNumber(Indicador14,2)
			end if

		Case 16 'UnidActoCompra
			if iFab <> 0 then
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Cantidad "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMed "
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
				if iFab = 0 and iMar = 0 and iSeg <> 0 then
					sql = ""
					sql = sql & " SELECT "
					sql = sql & " Cantidad "
					sql = sql & " FROM "
					sql = sql & " PH_DataCrudaMed "
					sql = sql & " WHERE "
					sql = sql & " Id_Categoria = " & sCat
					sql = sql & " And Id_Fabricante = 0 "
					sql = sql & " And Id_Marca = 0 "
					sql = sql & " And Id_Segmento = " & iSeg
					sql = sql & " And id_Semana in( " & idSemana & ")"
				else
					sql = ""
					sql = sql & " SELECT "
					sql = sql & " Cantidad "
					sql = sql & " FROM "
					sql = sql & " PH_DataCrudaMed "
					sql = sql & " WHERE "
					sql = sql & " Id_Categoria = " & sCat
					sql = sql & " And Id_Fabricante = 0 "
					sql = sql & " And Id_Marca = 0 "
					sql = sql & " And Id_Segmento = 0 "
					sql = sql & " And Id_RangoTamano = 0 "
					sql = sql & " And id_Semana in( " & idSemana & ")"
				end if
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
				Valor = FormatNumber(Valor,2)
				Indicador3 = 0
				for iDat = 0 to ubound(gDatos1,2)
					Indicador3 = Indicador3 + gDatos1(0,iDat)
				next

				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Cantidad, "
				sql = sql & " Id_Consumo "
				sql = sql & " FROM PH_DataCrudaMed "
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

		Case 17 'IndiceConsumoVol (NO VA - Es Mensual)
			Valor = 0
		Case 18 'IndiceConsumoVal (NO VA - Es Mensual)
			Valor = 0
		Case 19 'RepeticionCompra (NO VA - Es Mensual)
			Valor = 0
		Case 20 'FidelidadVol (NO VA - Es Mensual)
			Valor = 0
		Case 21 'FidelidadVal (NO VA - Es Mensual)
			Valor = 0
		Case 22 'FidelidadActos (NO VA - Es Mensual)
			Valor = 0
		Case 23 'CuotaMerVol (NO VA - Es Mensual)
			Valor = 0

		Case 24 'PrecProVol (NO VA - Es Mensual)
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad, "
			sql = sql & " Precio_Producto, "
			sql = sql & " Dolar, "
			sql = sql & " Precio_producto/Dolar AS PrecioDolar, "
			sql = sql & " (Precio_producto/Dolar)*Cantidad AS ComprasValor "
			sql = sql & " FROM "
			sql = sql & " PH_DataCrudaMed "
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
				sql = sql & " Tamano, "
				sql = sql & " Cantidad "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMed "
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
				Indicador1 = Indicador1/1000
				
				Valor = cdbl(Indicador2)/cdbl(Indicador1)
				Valor = FormatNumber(Valor,2)
			end if

		Case 25 'PrecProVal
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad, "
			sql = sql & " Precio_Producto, "
			sql = sql & " Dolar, "
			sql = sql & " Precio_producto/Dolar AS PrecioDolar, "
			sql = sql & " (Precio_producto/Dolar)*Cantidad AS ComprasValor "
			sql = sql & " FROM "
			sql = sql & " PH_DataCrudaMed "
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
				sql = sql & " Cantidad "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMed "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria = " & sCat
				sql = sql & " and Id_Fabricante = " & iFab 
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
					Indicador3 = Indicador3 + gDatos1(0,iDat)
				next
				
				Valor = cdbl(Indicador2)/cdbl(Indicador3)
				Valor = FormatNumber(Valor,2)
			end if

		Case 26 'MarcasHogar
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Id_Hogar AS Total "
			sql = sql & " FROM "
			sql = sql & " PH_DataCrudaMed "
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
			sql = sql & " GROUP BY "
			sql = sql & " PH_DataCrudaMed.Id_Hogar "
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
				sql = sql & " PH_DataCrudaMed "
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
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				sql = sql & " PH_DataCrudaMed.Id_Hogar "
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
				sql = sql & " PH_DataCrudaMed "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria = " & sCat
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
			'response.write "<br> Valor:" & Valor
			
		Case 27 'CadenasProm  (NO VA - Es Mensual)
			Valor = 0
		
		Case 28 'CuotaMercVol
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Tamano, "
			sql = sql & " Cantidad "
			sql = sql & " FROM "
			sql = sql & " PH_DataCrudaMed "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = " & sCat
			sql = sql & " And Id_Fabricante = 0 "
			sql = sql & " And Id_Marca = 0"
			sql = sql & " And Id_Segmento = 0"
			sql = sql & " And Id_RangoTamano = 0"
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
				sql = sql & " PH_DataCrudaMed "
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
						TotalFiltro = TotalFiltro + (cdbl(gDatos1(0,iDat)) *cdbl(gDatos1(1,iDat)))
					next
					TotalFiltro = TotalFiltro/1000
					if iFab = 0 and iMar = 0 and iSeg = 0 then TotalFiltro=TotalVolumen
					Valor = (TotalFiltro/TotalVolumen)*100
					Valor = FormatNumber(Valor,2)
					'response.write "<br> Total Volumen:= " & TotalVolumen
					'response.write "<br> Total Filtro:= " & TotalFiltro
				end if
			end if

		Case 29 'CuoMerVal
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad, "
			sql = sql & " Precio_Producto, "
			sql = sql & " Dolar, "
			sql = sql & " Precio_producto/Dolar AS PrecioDolar, "
			sql = sql & " (Precio_producto/Dolar)*Cantidad AS ComprasValor "
			sql = sql & " FROM "
			sql = sql & " PH_DataCrudaMed "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = " & sCat
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
				TotalValor = 0
				for iDat = 0 to ubound(gDatos1,2)
					TotalValor = TotalValor + cdbl(gDatos1(4,iDat))
				next

				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Cantidad, "
				sql = sql & " Precio_Producto, "
				sql = sql & " Dolar, "
				sql = sql & " Precio_producto/Dolar AS PrecioDolar, "
				sql = sql & " (Precio_producto/Dolar)*Cantidad AS ComprasValor "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMed "
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
						TotalFiltro = TotalFiltro + cdbl(gDatos1(4,iDat))
					next
					
					if iFab = 0 and iMar = 0 and iSeg = 0 then TotalFiltro = TotalValor
					Valor = (TotalFiltro/TotalValor)*100
					Valor = FormatNumber(Valor,2)
				end if
			end if


		Case 30 'CuoMerUni
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad "
			sql = sql & " FROM "
			sql = sql & " PH_DataCrudaMed "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = " & sCat
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
					TotalUnidades = TotalUnidades + gDatos1(0,iDat)
				next
				
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Cantidad "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMed "
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
						TotalFiltro = TotalFiltro + gDatos1(0,iDat)
					next
					
					if iFab = 0 and iMar = 0 and iSeg = 0 then TotalFiltro = TotalUnidades
					Valor = (TotalFiltro/TotalUnidades)*100
					Valor = FormatNumber(Valor,2)
				end if
			end if


		Case 31 'CuoMerAct
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad, "
			sql = sql & " Id_Consumo "
			sql = sql & " FROM PH_DataCrudaMed "
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
				sql = sql & " FROM PH_DataCrudaMed "
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
			
		
		Case 32 'PenRel
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Id_Hogar AS Total "
			sql = sql & " FROM "
			sql = sql & " PH_DataCrudaMed "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = " & sCat
			sql = sql & " And id_Semana in( " & idSemana & ")"
			sql = sql & " GROUP BY "
			sql = sql & " PH_DataCrudaMed.Id_Hogar "
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
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Id_Hogar AS Total "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMed "
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
				sql = sql & " and Id_Categoria = " & sCat
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				sql = sql & " PH_DataCrudaMed.Id_Hogar "
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

		Case 33 'CompRel  (NO VA - Es Mensual)
			Valor = 0
		Case 34 'PenAcum
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Id_Hogar AS Total "
			sql = sql & " FROM "
			sql = sql & " PH_DataCrudaMed "
			sql = sql & " WHERE "
			'sql = sql & " Id_Categoria = " & sCat
			sql = sql & " id_Semana in( " & idSemana & ")"
			sql = sql & " GROUP BY "
			sql = sql & " PH_DataCrudaMed.Id_Hogar "
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
				sql = sql & " PH_DataCrudaMed "
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
				sql = sql & " and Id_Categoria = " & sCat
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				sql = sql & " PH_DataCrudaMed.Id_Hogar "
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
			
		Case 35 'HogRecomp (NO VA - Es Mensual)
			Valor = 0
		Case 36 'HogNuevos (NO VA - Es Mensual)
			Valor = 0
		Case 37 'HogNoRecomp (NO VA - Es Mensual)
			Valor = 0

		Case 40 'Penetracion Med
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Id_Hogar AS Total "
			sql = sql & " FROM "
			sql = sql & " PH_DataCrudaMed "
			sql = sql & " WHERE "
			'sql = sql & " Id_Categoria =  " & sCat
			sql = sql & " id_Semana in( " & idSemana & ")"
			sql = sql & " GROUP BY "
			sql = sql & " PH_DataCrudaMed.Id_Hogar "
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
				sql = sql & " PH_DataCrudaMed "
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
				sql = sql & " and Id_Categoria = " & sCat
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				sql = sql & " PH_DataCrudaMed.Id_Hogar "
				'response.write "<br>36 sql:=" & sql & "<br>"
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

	end select 
end Sub



	'response.end
%>
