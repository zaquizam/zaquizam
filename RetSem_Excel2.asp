<%@language=vbscript%>
<!--#include file="conexionRS.asp"-->
<%
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	'response.write "<br>84 LLEGO"
	'response.end
	Dim sCat
	Dim sAre
	Dim sZon
	Dim sCan
	Dim sFab
	Dim sMar
	Dim sSeg
	Dim sRan
	Dim sTam
	Dim sPro
	Dim sInd
	Dim sSem
	Dim sSemanas
	Dim sSemAcum
	Dim iAre
	Dim iZon
	Dim iCan
	Dim iFab
	Dim iMar
	Dim iSeg
	Dim iRan
	Dim iTam
	Dim iPro
	Dim iInd
	Dim idSemana
	Dim gSemanas
	Dim gSemanasAcum
	Dim TotalSemAcum
	'	
	sCat=Request.Form("cat")
	sAre=Request.Form("are")
	sZon=Request.Form("zon")
	sCan=Request.Form("can")
	sFab=Request.Form("fab")
	sMar=Request.Form("mar")
	sSeg=Request.Form("seg")
	sRan=Request.Form("ran")
	sTam=Request.Form("tam")
	sPro=Request.Form("pro")
	sInd=Request.Form("ind")
	sSem=Request.Form("sem")
	'sSemAcum=Request.Form("semacum")

	sSemanas = sSem
	'response.write "<br>84 Sem:=" & sSem
	'response.end
	
	if sPro = 0 then sPro = "" end if
	if sInd = 0 then sInd = "" end if
	
	' iCol = 1
	' do
		' ix = instr(sSem,",")
		' if ix <> 0 then
			' 'response.write "<br>Semana:= " & mid(sSem,iCol,ix-1) & "=" & ix
			' ix = ix + 1
			' sSem = mid(sSem,ix)
		' end if
	' loop until ix = 0
	'response.write "<br>Semana:= " & sSem
	'response.write "<br>84 LLEGO" & sFab
	'response.end
	
	Dim gProductosTotal
	Dim gIndicadores
	Dim Indicador
	Dim Valor
	
	Dim gDatos1
	Dim rsx1
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
	sql = sql & " IdSemana in ( " & sSemanas & ")"
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
	'	
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Indicador, "
	sql = sql & " Abreviatura, "
	sql = sql & " UnidadMedida "
	sql = sql & " FROM "
	sql = sql & " RS_Indicadores "
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
	sql = sql & " And Id_Semana in ( " & sSemanas & ")"
	sql = sql & " And Id_Area in (" & sAre & ")"
	sql = sql & " And Id_Zona in (" & sZon & ")"
	sql = sql & " And Id_Canal in (" & sCan & ")"
	sql = sql & " And Id_Fabricante in (" & sFab & ")"
	sql = sql & " And Id_Marca in (" & sMar & ")"
	sql = sql & " And Id_Segmento in (" & sSeg & ")"
	if sTam <> "" and sTam <> "0" then
		sql = sql & " And Id_Tamano in (" & sTam & ")"
	else
		if sPro <> "" then
		else
			sql = sql & " And Id_Tamano = 0 "
		end if
	end if
	if sPro <> "" then
		sPro = replace(sPro,",","','")
		sql = sql & " And CodigoBarra in ('" & sPro & "')"
	else
		sql = sql & " And CodigoBarra = ''"
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
	''
	'response.write "<br>258 sql:= " & sql
	'response.end
	''
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
	'response.write "<br>271 Paso" 
	'response.end
	'response.write "<br>" & sql
	'response.end

	if iExiste = 0 then
		
		%>
		<center>
		<h2>No hay Data para Mostrar</h2>
		<hr>
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
		'response.write "<br>84 LLEGO"
		'response.end
		
		%>
		<div class="">
			
			<div class="">
			
				<div class="">
								
					<div class="">
						
							<div class="">
							
								<table border=0>
									<thead>
										<!--<tr class="class="container-fluid"">-->
										<tr>
											<th class="text-center">Area</th>
											<th class="text-center">Zona</th>
											<th class="text-center">Canal</th>
											<th class="text-center">Fabricante</th>
											<th class="text-center">Marca</th>
											<th class="text-center">Segmento</th>
											<th class="text-center">Tamano</th>
											<th class="text-center">Producto</th>
											<th class="text-center">Indicador</th>
											<th class="text-center">UniMed</th>
											<th class="text-center"><%=strSemana1%></th>
											<th class="text-center"><%=strSemana2%></th>									
											<th class="text-center"><%=strSemana3%></th>									
											<th class="text-center"><%=strSemana4%></th>
											<th class="text-center"><%=strSemana5%></th>
										</tr>
									</thead>
								</table>
								
							</div>
							<br>
							<div class="">
								<table border=1>
									<tbody>					
										<% 
										'response.write "<br>354 LLEGO:= " & ubound(gProductosTotal,2)
										'response.end
										TotalReg = ubound(gProductosTotal,2)
										for iPro = 0 to  ubound(gProductosTotal,2)
											'response.write "<br>354 LLEGO:= " & iPro
											response.write "<tr class='row100 body'>"
												'Area
												response.write "<td class=''>"
													response.write gProductosTotal(1,iPro) 
												response.write "</td>"
												'Zona
												response.write "<td class=''>"
													response.write gProductosTotal(3,iPro) 
												response.write "</td>"
												'Canal
												response.write "<td class=''>"
													response.write gProductosTotal(5,iPro) 
												response.write "</td>"
												'Fabricante
												response.write "<td class=''>"
													response.write gProductosTotal(7,iPro) 
												response.write "</td>"
												'Marca
												response.write "<td class=''>"
													response.write gProductosTotal(9,iPro) 
												response.write "</td>"
												'Segmento
												response.write "<td class=''>"
													response.write gProductosTotal(11,iPro) 
												response.write "</td>"
												'Tamaño
												response.write "<td class=''>"
													'response.write gProductosTotal(13,iPro) 
													'response.end
													if gProductosTotal(12,iPro) <> 0 then
														Valor = gProductosTotal(13,iPro)
														Valor = replace(Valor,".",",")
														response.write formatnumber(Valor,2) 
													else
													
													end if
													
												response.write "</td>"
												'Producto
												
												response.write "<td class=''>"
													response.write gProductosTotal(14,iPro) & "-" & gProductosTotal(15,iPro)
												response.write "</td>"
												'response.write "<td colspan=7 class='cell100 column9' >"
												'response.write "</td>"
											'response.write "</tr>"
											iPro2 = iPro
											isw = 0
											for iInd = 0 to  ubound(gIndicadores,2)
												iPro1 = iPro
												'response.write "<br>354 LLEGO:= " & iPro1
													if isw = 0 then
														isw = 1
													else
														response.write "<tr>"
														response.write "<td colspan=8 class='' >"
														response.write "</td>"
													end if
													response.write "<td class='text-center'>"
														response.write "<b>"
														'response.write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
														response.write gIndicadores(1,iInd)
														response.write "</b>"
													response.write "</td>"
													response.write "<td class='text-center'>"
														response.write "<b>"
														response.write gIndicadores(2,iInd)
														response.write "</b>"
													response.write "</td>"
													Indicador = gIndicadores(0,iInd)
													Columna = Indicador + 16
													Menos = 0
													'response.write "<br>iPro1:=" & iPro1 & ""
													sw=0
													for iSem = 0 to  ubound(gSemanas,2)
														response.write "<td class='text-right'>"
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
															if iPro1 > TotalReg then 
																'response.write "menos: " & Menos
																sw=sw+1
																exit for
															end if
														response.write "</td>"
														
													next					
													
													ix = cint(ubound(gSemanas,2))
													iy = 4 - ix
													if sw <> 0 then 
														iy = 4- Menos
														'response.write "<br>Paso:=" & iy
													end if
													
													Menos = 0
													if iy <> 0 then  
														for ia = 1 to iy
															response.write "<td class='text-left'>"
																'Valor = 0
																'Valor = FormatNumber(Valor,2)
																'response.write Valor
															response.write "</td>"
														next 
													end if
												response.write "</tr>"
											next
											'if iPro > TotalReg then response.end
											iPro = iPro2 + iPro1 - 1
											iPro = iPro1 - 1
											'response.write "<br>iPro:=" & iPro & ""
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
