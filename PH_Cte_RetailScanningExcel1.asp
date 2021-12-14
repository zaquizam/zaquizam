<%@language=vbscript%>
<!--#include file="conexionRS.asp"-->
<%
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	'response.write "<br>84 LLEGO"
	'response.end
	Response.Buffer = true
	dim sCat
	dim sAre
	dim sZon
	dim sCan
	dim sFab
	dim sMar
	dim sSeg
	dim sRan
	dim sTam
	dim sPro
	dim sInd
	dim sSem
	dim sSemanas
	dim sSemAcum
	dim iAre
	dim iZon
	dim iCan
	dim iFab
	dim iMar
	dim iSeg
	dim iRan
	dim iTam
	dim iPro
	dim iInd
	dim idSemana
	dim gSemanas
	dim gSemanasAcum
	dim TotalSemAcum

	'sCat=Request.QueryString("cat")
	sCat=Request.form("cat")
	if sCat = "" Then response.end

	 'sAre=Request.QueryString("are")
	 'sZon=Request.QueryString("zon")
	 'sCan=Request.QueryString("can")
	 'sFab=Request.QueryString("fab")
	 'sMar=Request.QueryString("mar")
	 'sSeg=Request.QueryString("seg")
	 'sRan=Request.QueryString("ran")
	 'sTam=Request.QueryString("tam")
	 'sPro=Request.QueryString("pro")
	 'sInd=Request.QueryString("ind")
	 'sSem=Request.QueryString("sem")
	 'sSemAcum=Request.QueryString("semacum")

	sAre=Request.form("are")
	sZon=Request.form("zon")
	sCan=Request.form("can")
	sFab=Request.form("fab")
	sMar=Request.form("mar")
	sSeg=Request.form("seg")
	sRan=Request.form("ran")
	sTam=Request.form("tam")
	sPro=Request.form("pro")
	sInd=Request.form("ind")
	sSem=Request.form("sem")
	sSemAcum=Request.form("semacum")

	sSemanas = sSem
	'response.write "<br>84 Sem:=" & sSem
	'response.end
	
	iCol = 1
	do
		ix = instr(sSem,",")
		if ix <> 0 then
			'response.write "<br>Semana:= " & mid(sSem,iCol,ix-1) & "=" & ix
			ix = ix + 1
			sSem = mid(sSem,ix)
		end if
	loop until ix = 0
	'response.write "<br>Semana:= " & sSem

	if sAre = "" then
		sAre = "0"
	end if
	if sZon = "" then
		sZon = "0"
	end if
	if sCan = "" then
		sCan = "0"
	end if
	if sFab = "" then
		sFab = "0"
	end if
	if sMar = "" then
		sMar = "0"
	end if
	if sSeg = "" then
		sSeg = "0"
	end if
	if sTam = "" then
		sTam = "0"
	end if
	if sPro = "" then
		sPro = ""
	end if
	'response.write "<br>84 LLEGO" & sFab
	'response.end
	
	dim gProductosTotal
	dim gIndicadores
	dim Indicador
	dim Valor
	
	dim gDatos1
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 1 'adLockOptimistic 

	'Semanas
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " IdSemana, "
	sql = sql & " Semana "
	sql = sql & " FROM "
	sql = sql & " ss_Semana "
	sql = sql & " WHERE "
	'sql = sql & " IdSemana in ( " & sSemanas & ")"
	sql = sql & " IdSemana >23"
	sql = sql & " and IdSemana <57"
	sql = sql & " Order By "
	sql = sql & " IdSemana "
	'response.write "<br>36 sql:=" & sql
	'response.end
	rsx1.Open sql ,conexionRS
	if rsx1.eof then
		rsx1.close
	else
		gSemanas = rsx1.GetRows
		rsx1.close
	end if
	strSemana1 = ""
	strSemana2 = ""
	strSemana3 = ""
	strSemana4 = ""
	strSemana5 = ""
	for iSem = 0 to  ubound(gSemanas,2)
		if iSem = 0 then strSemana1 = gSemanas(1,iSem)
		if iSem = 1 then strSemana2 = gSemanas(1,iSem)
		if iSem = 2 then strSemana3 = gSemanas(1,iSem)
		if iSem = 3 then strSemana4 = gSemanas(1,iSem)
		if iSem = 4 then strSemana5 = gSemanas(1,iSem)
	next
	idCliente = Session("idCliente")
	'response.write "<br> Cliente:= " & idCliente 
	'if idCliente = 1 then
	'	response.write "<br> Cliente:= " & idCliente 
	'	
	'	erase gSemanas
	'	redim gSemanas(1,10)
	'	gSemanas(1,0) = 24
	'	gSemanas(1,1) = 25
	'	gSemanas(1,2) = 26
	'	gSemanas(1,3) = 27
	'	gSemanas(1,4) = 28
	'	gSemanas(1,5) = 29
	'	gSemanas(1,6) = 30
	'	gSemanas(1,7) = 31
	'	gSemanas(1,8) = 32
	'	gSemanas(1,9) = 33
	'	gSemanas(1,10) = 24
	'end if
	if sSemAcum <> "" then
		'Semanas Acumuladas
		sql = ""
		sql = sql & " SELECT "
		sql = sql & " IdSemana, "
		sql = sql & " Semana "
		sql = sql & " FROM "
		sql = sql & " ss_Semana "
		sql = sql & " WHERE "
		sql = sql & " IdSemana in ( " & sSemAcum & ")"
		sql = sql & " Order By "
		sql = sql & " IdSemana "
		'response.write "<br>36 sql:=" & sql
		'response.end
		isw = 0
		rsx1.Open sql ,conexionRS
		if rsx1.eof then
			'response.write "<br>152 Paso" 
			rsx1.close
			isw = 0
		else
			gSemanasAcum = rsx1.GetRows
			'response.write "<br>157 Paso" 
			rsx1.close
			isw = 1
			strSemana5 = "Acum. "
			TotalSemAcum = 0
			for iSem = 0 to  ubound(gSemanasAcum,2)
				strSemana5 = strSemana5 & mid(gSemanasAcum(1,iSem),1,5)
				TotalSemAcum = TotalSemAcum + 1
			next
		end if
		if isw = 1 then
			'response.write "<br>164 Paso" 
		end if 
	end if
	'response.write "<br>173 Paso" 
	'response.end

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Indicador, "
	sql = sql & " Abreviatura, "
	sql = sql & " UnidadMedida "
	sql = sql & " FROM "
	sql = sql & " RS_Indicadores "
	sql = sql & " WHERE "
	sql = sql & " Ind_Activo = 1 " 
	if sInd <> "" then
		sql = sql & " And Id_Indicador in (" & sInd & ")"
	end if
	sql = sql & " ORDER BY "
	sql = sql & " Id_Indicador "
	'response.write "<br>372 sql:=" & sql
	'response.end 
	rsx1.Open sql ,conexionRS
	if rsx1.eof then
		rsx1.close
	else
		gIndicadores = rsx1.GetRows
		rsx1.close
	end if
	'response.write "<br>203 Paso" 
	'response.end
	
	'Query
	sql = ""
    sql = sql & " SELECT "
	sql = sql & " Id_Area, "
	sql = sql & " Area, "
	sql = sql & " Id_Zona, "
	sql = sql & " Zona, "
	sql = sql & " Id_Canal, "
	sql = sql & " Canal, "
	sql = sql & " Id_Fabricante, "
	sql = sql & " Fabricante, "
	sql = sql & " Id_Marca, "
	sql = sql & " Marca, "
	sql = sql & " Id_Segmento, "
	sql = sql & " Segmento, "
	sql = sql & " Id_Tamano, "
	sql = sql & " Tamano, "
	sql = sql & " CodigoBarra, "
	sql = sql & " Descripcion, "
	sql = sql & " UnidadMedida, "
	sql = sql & " VentasUni, "			'17
	sql = sql & " VentasVal, "			'18
	sql = sql & " VentasUniMed, "		'19
	sql = sql & " VentasNo, "			'20
	sql = sql & " DistribucionNum, "	'21
	sql = sql & " DistribucionPon, "	'22
	sql = sql & " DistribucionEfe, "	'23
	sql = sql & " ShareUni, "			'24
	sql = sql & " ShareVol, "			'25
	sql = sql & " ShareVal, "			'26
	sql = sql & " PrecioPro, "			'27
	sql = sql & " PrecioMax, "			'28
	sql = sql & " PrecioMin, "			'29
	sql = sql & " PrecioUni, "			'30
	sql = sql & " PrecioUniMed, "		'31
	sql = sql & " id_Semana "			'32
	sql = sql & " FROM "
	sql = sql & " RS_DataProcSem "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria = " & sCat
	sql = sql & " And Id_Semana <54 "
	'sql = sql & " And Id_Semana in ( " & sSemanas & ")"
	sql = sql & " And Id_Area in (" & sAre & ")"
	sql = sql & " And Id_Zona in (" & sZon & ")"
	sql = sql & " And Id_Canal in (" & sCan & ")"
	sql = sql & " And Id_Fabricante in (" & sFab & ")"
	sql = sql & " And Id_Marca in (" & sMar & ")"
	sql = sql & " And Id_Segmento in (" & sSeg & ")"
	if len(sTam) > 1 then
		sql = sql & " And Id_Tamano in (" & sTam & ")"
	else
		sql = sql & " And Id_Tamano = 0 "
	end if
	if sPro <> "" then
		sPro = replace(sPro,",","','")
		sql = sql & " And CodigoBarra in ('" & sPro & "')"
	else
		if sPro = "" then
			sql = sql & " And CodigoBarra = ''"
		end if
	end if
    sql = sql & " ORDER BY "
	sql = sql & " Id_Area, "
	sql = sql & " Id_Zona, "
	sql = sql & " Id_Canal, "
	sql = sql & " Id_Fabricante, "
	sql = sql & " Id_Marca, "
	sql = sql & " Id_Segmento, "
	sql = sql & " Id_Tamano, "
	sql = sql & " CodigoBarra, "
	sql = sql & " Descripcion, "
	sql = sql & " id_Semana "
	if sAre = "0" and sZon = "0" and sCan = "0" and sFab = "0" and sMar = "0" and sSeg = "0" and sTam = "0" and sPro <> "" then
		sql = replace(sql,"And Id_Tamano = 0","")
	else
	end if
	'response.write "<br>258 sql:= " & sql
	'response.end
    rsx1.Open sql ,conexionRS
	iExiste = 0
	if rsx1.eof then
		iExiste = 0
		rsx1.close
	else
		iExiste = 1
		gProductosTotal = rsx1.GetRows
		rsx1.close
	end if
	'response.write "<br>328 Paso" 
	'response.end
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-disposition","attachment; filename=tem.xls"
	
	if iExiste = 0 then
		
		%>
		<center>
		<h2>No hay Data para Mostrar</h2>
		</center>
		<%
			Response.end
		%>
		<div class="limiter">
			<div class="container-table100">
				<div class="wrap-table100">
					<div class="table100 ver1 m-b-110">
							<div class="table100-head">
								<table>
									<thead>
										<tr class="row100 head">
											<th class="cell100 column1 text-center">Area</th>
											<th class="cell100 column2 text-center">Zona</th>
											<th class="cell100 column3 text-center">Canal</th>
											<th class="cell100 column4 text-center">Fabricante</th>
											<th class="cell100 column5 text-center">Marca</th>
											<th class="cell100 column6 text-center">Segmento</th>
											<th class="cell100 column7 text-center">Tamano</th>
											<th class="cell100 column8 text-center">Producto</th>
											<th class="cell100 column9 text-center">Indicador</th>
											<th class="cell100 column10 text-center">UniMed</th>
											<th class="cell100 column11 text-center"><%=strSemana1%></th>
											<th class="cell100 column12 text-center"><%=strSemana2%></th>									
											<th class="cell100 column13 text-center"><%=strSemana3%></th>									
											<th class="cell100 column14 text-center"><%=strSemana4%></th>
											<th class="cell100 column15 text-center"><%=strSemana5%></th>
										</tr>
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
		'response.write "<br>354 LLEGO"
		'response.end
		
		%>
		<div class="limiter">
			
			<div class="container-table100">
			
				<div class="wrap-table100">
								
					<div class="table100 ver1 m-b-110">
						
							<div class="table100-head">
							
								<table border=0>
									<thead>
										<tr class="row100 head">
											<th class="cell100 column1 text-center">Area</th>
											<th class="cell100 column2 text-center">Zona</th>
											<th class="cell100 column3 text-center">Canal</th>
											<th class="cell100 column4 text-center">Fabricante</th>
											<th class="cell100 column5 text-center">Marca</th>
											<th class="cell100 column6 text-center">Segmento</th>
											<th class="cell100 column7 text-center">Tamano</th>
											<th class="cell100 column8 text-center">Producto</th>
											<th class="cell100 column9 text-center">Indicador</th>
											<th class="cell100 column10 text-center">UniMed</th>
											<th class="cell100 column11 text-center">
											<%
												for iSem = 0 to  ubound(gSemanas,2)
													response.write gSemanas(1,iSem) & "</td>"
													response.write "<td>" 
												next 
											%></th>
										</tr>
									</thead>
								</table>
								
							</div>
							<div class="table100-body js-pscroll">
								<table border=0>
									<tbody>					
										<% 
										'response.write "<br>397 LLEGO:= " & ubound(gProductosTotal,2)
										'response.end
										for iPro = 0 to  ubound(gProductosTotal,2)
											'response.write "<br>354 LLEGO:= " & iPro
											iPro2 = iPro
											isw = 0
											for iInd = 0 to  ubound(gIndicadores,2)
											
											'response.write "<br>404 LLEGO:= " & iPro
											response.write "<tr class='row100 body'>"
												'Area
												response.write "<td width=6% class='cell100 column1'>"
													response.write gProductosTotal(1,iPro) 
												response.write "</td>"
												'Zona
												response.write "<td width=6% class='cell100 column2'>"
													response.write gProductosTotal(3,iPro) 
												response.write "</td>"
												'Canal
												response.write "<td width=6% class='cell100 column3'>"
													response.write gProductosTotal(5,iPro) 
												response.write "</td>"
												'Fabricante
												response.write "<td width=6% class='cell100 column4'>"
													response.write gProductosTotal(7,iPro) 
												response.write "</td>"
												'Marca
												response.write "<td width=6% class='cell100 column5'>"
													response.write gProductosTotal(9,iPro) 
												response.write "</td>"
												'Segmento
												response.write "<td width=6% class='cell100 column6'>"
													response.write gProductosTotal(11,iPro) 
												response.write "</td>"
												'Tamaño
												response.write "<td width=6% class='cell100 column7'>"
													'response.write gProductosTotal(13,iPro) 
													if gProductosTotal(12,iPro) <> 0 then
														Valor = gProductosTotal(13,iPro)
														Valor = replace(Valor,".",",")
														response.write formatnumber(Valor,2) 
													else
													
													end if
													
												response.write "</td>"
												'Producto
												response.write "<td width=6% class='cell100 column8'>"
													response.write gProductosTotal(14,iPro) & "-" & gProductosTotal(15,iPro)
												response.write "</td>"
												'response.write "<td width=6% colspan=7 class='cell100 column9' >"
												'response.write "</td>"
											'response.write "</tr>"
												iPro1 = iPro
												'response.write "<br>354 LLEGO:= " & iPro1

													if isw = 0 then
														isw = 1
													else
														'response.write "<tr>"
														'response.write "<td width=6% colspan=8 class='cell100 column9' >"
														'response.write "</td>"
													end if
													response.write "<td width=6% class='cell100 column9 text-center'>"
														response.write "<b>"
														'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
														response.write gIndicadores(1,iInd)
														response.write "</b>"
													response.write "</td>"
													response.write "<td width=6% class='cell100 column10 text-center'>"
														response.write "<b>"
														response.write gIndicadores(2,iInd)
														response.write "</b>"
													response.write "</td>"
													Indicador = gIndicadores(0,iInd)
													Columna = Indicador + 16
													Menos = 0
													'response.write "<br>iPro1:=" & iPro1 & ""
													'response.write "<br>496 "
													'response.end
													for iSem = 0 to  ubound(gSemanas,2)
														response.write "<td width=6% class='cell100 column11 text-right'>"
															'response.write "iPro:=" & iPro & "=>"
															iSem1 = gSemanas(0,iSem)
															if cint(iSem1) = cint(gProductosTotal(32,iPro1)) then 
																Valor = gProductosTotal(Columna,iPro1)
																if Valor <> "" then
																	Valor = FormatNumber(Valor,2)
																else
																	Valor = "IND"
																end if
																iPro1 = iPro1 + 1
															else
																Valor = 0
																Valor = FormatNumber(Valor,2)
																Menos = Menos + 1
															end if
															response.write Valor
														response.write "</td>"
													next					
													'response.write "<br>397 LLEGO:= " & ubound(gProductosTotal,2)
													'response.end
													ix = cint(ubound(gSemanas,2))
													iy = 4 - ix
													Menos = 0
													if iy <> 0 then  
														for ia = 1 to iy
															response.write "<td width=6% class='cell100 column15 text-left'>"
															
															response.write "</td>"
														next 
													end if
												response.write "</tr>"
												'response.write "<br>497 LLEGO:= " & iPro
												'response.end
											next
											iPro = iPro2 + iPro1 - 1
											iPro = iPro1 - 1
											'response.write "<br>iPro:=" & iPro & ""
											'response.write "<br>508 LLEGO:= " & ubound(gProductosTotal,2)
											'response.end
											if iPro = 0 then exit for
										next					
										
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
	


	'response.end
%>
