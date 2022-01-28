<%@language=vbscript%>
<!--#include file="conexionRS.asp"-->
<!-- RetSem_Datos.asp - 12oct21 - 27ene22 -->
<%
	' Variables y Constantes
	'Response.Write(Server.ScriptTimeout)
	Server.ScriptTimeout = 30000
	Response.Buffer = True	
	Session.lcid = 1034
	Response.ContentType = "text/html"	
	Response.CodePage = 65001
	Response.CharSet = "UTF-8"	
	'
	StartTime = Timer
	'				
	Dim sCat
	Dim sAre
	Dim sZon
	Dim sCan
	Dim sFab
	Dim sMar
	Dim sSeg
	Dim sTam
	Dim sPro
	Dim sInd
	Dim sSem
	Dim sSemanas
	Dim sSemAcum
	Dim gSemanas
	Dim gSemanasAcum
	Dim TotalSemAcum
	Dim idCliente
	Dim gMeses
	Dim gValor
	Dim iValor
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
	sMes=Request.Form("mes")
	idCliente = Session("idCliente")
	'		
	' Response.Write "<br>sMes " & sMes & "<br>"
	' Response.End
	'	
	sSemanas = sSem
	'Response.Write "<br>84 Sem:=" & sSem	
	
	if Len(sPro) = 0 then sPro = "" end if
	if Len(sInd) = 0 then sInd = "" end if	
		
	Dim gProductosTotal
	Dim gIndicadores
	Dim Indicador
	Dim Valor
	Dim sSemMes
	
	Dim gDatos1
	Dim rSx1
	set rSx1 = CreateObject("ADODB.Recordset")
	'rSx1.CursorType = adOpenKeyset 
	'rSx1.LockType = 1 'adLockOptimistic 

	'Semanas	
	sQl = vbNullstring
	sQl = " SELECT IdSemana, SemanaCorta FROM ss_Semana WHERE IdSemana in ( " & sSemanas & ") Order By IdSemana "
	'Response.Write "<br>151 sQl:=" & sQl & "<br>"
	'Response.End
	rSx1.Open sQl,conexionRS,0,1
	if rSx1.Eof then
		rSx1.Close
	else
		gSemanas = rSx1.GetRows()
		rSx1.Close
	end if
	'	
	'Meses
	'
	sQl = vbNullstring
	sQl = " SELECT IdPeriodo, PeriodoCorto, Semanas FROM ss_Periodo WHERE Semanas IS NOT NULL AND IdPeriodo in ( " & sMes & ") Order By idPeriodo ASC "
	'Response.Write "<br>151 sQl:=" & sQl & "<br>"
	'Response.End
	rSx1.Open sQl,conexionRS,0,1
	if rSx1.Eof then
		rSx1.Close
	else
		gMeses = rSx1.GetRows()
		rSx1.Close
	end if
	sSemMes = ""
	if IsArray(gMeses) then
		for iMes = 0 to ubound(gMeses,2)
			sSemMes = sSemMes & gMeses(2,iMes) & ","
		next									
		sSemMes = Left(sSemMes,Len(sSemMes)-1)
	end if
	'	
	sQl = vbNullstring
	sQl = " SELECT Id_Indicador, Abreviatura, UnidadMedida FROM RS_Indicadores WHERE Ind_Activo = 1 " 
	'
	if (CInt(idCliente) = 1) then
		sQl = sQl & " AND Ind_atenas = 1 " 		
	else
		sQl = sQl & " AND Ind_men = 1 " 		
	end if
	if sInd <> "" then
		sQl = sQl & " And Id_Indicador in (" & sInd & ")"
	end if
	'
	if (sCat > 126 and sCat < 146) or (sCat = 41 or sCat = 18 or sCat = 54) then
		sQl = sQl & " AND ( Id_Indicador <> 3 and Id_Indicador <> 15 and Id_Indicador <> 9 ) "
	end if
	'	
	sQl = sQl & " ORDER BY Id_Indicador "
	'Response.Write "<br>191 sQl:=" & sQl & "<br>"
	''	
	'Response.End 
	rSx1.Open sQl,conexionRS,0,1
	if rSx1.Eof then
		rSx1.Close
	else
		gIndicadores = rSx1.GetRows()
		rSx1.Close
	end if
	'Response.Write "<br>203 Paso" 
	'Response.End	
	''
	'Query
	if Len(sSemanas) > 1 then
		sql = vbNullstring
		sql = " SELECT Id_Area, Area, Id_Zona, Zona, Id_Canal, Canal, Id_Fabricante, Fabricante, Id_Marca, Marca, Id_Segmento, Segmento, Id_Tamano, Tamano, CodigoBarra,"
		sql = sql & " Descripcion, UnidadMedida, VentasUni, VentasVal, VentasUniMed, VentasNo, DistribucionNum, DistribucionPon, DistribucionEfe, ShareUni,"			'24
		sql = sql & " ShareVol, ShareVal, PrecioPro, PrecioMax, PrecioMin, PrecioUni, PrecioUniMed, id_Semana FROM RS_DataProcSem WHERE"
		sql = sql & " Id_Categoria = " & sCat & " And Id_Semana in ( " & sSemanas & ") And Id_Area in (" & sAre & ")  And Id_Zona in (" & sZon & ") And Id_Canal in (" & sCan & ") And Id_Fabricante in (" & sFab & ")"
		sql = sql & " And Id_Marca in (" & sMar & ") And Id_Segmento in (" & sSeg & ")"
		if Len(sTam) > 1 then
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
		
		sql = sql & " ORDER BY Id_Area, Id_Zona, Id_Canal, Id_Fabricante, Id_Marca, Id_Segmento, Id_Tamano, CodigoBarra, Descripcion, id_Semana "
		'
		if sAre = "0" and sZon = "0" and sCan = "0" and sFab = "0" and sMar = "0" and sSeg = "0" and sTam = "0" and sPro <> "" then
			sql = replace(sql,"And Id_Tamano = 0","")		
		end if
		'
	else
		'
		sql = vbNullstring
		sql = " SELECT Id_Area, Area, Id_Zona, Zona, Id_Canal, Canal, Id_Fabricante, Fabricante, Id_Marca, Marca, Id_Segmento, Segmento, Id_Tamano, Tamano, CodigoBarra, "
		sql = sql & " Descripcion, UnidadMedida FROM RS_DataProcSem WHERE "
		sql = sql & " Id_Categoria = " & sCat & " And Id_Semana in ( " & sSemMes & ") And Id_Area in (" & sAre & ") And Id_Zona in (" & sZon & ") And Id_Canal in (" & sCan & ") And Id_Fabricante in (" & sFab & ")"
		sql = sql & " And Id_Marca in (" & sMar & ") And Id_Segmento in (" & sSeg & ")"
		if Len(sTam) > 1 then
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
		'
		sql = sql & " GROUP BY Id_Area, Area, Id_Zona, Zona, Id_Canal, Canal, Id_Fabricante, Fabricante, Id_Marca, Marca, Id_Segmento, Segmento, Id_Tamano, Tamano, CodigoBarra, Descripcion, UnidadMedida "
		sql = sql & " ORDER BY Id_Area, Id_Zona, Id_Canal, Id_Fabricante, Id_Marca, Id_Segmento, Id_Tamano,  CodigoBarra, Descripcion "
		'
		if sAre = "0" and sZon = "0" and sCan = "0" and sFab = "0" and sMar = "0" and sSeg = "0" and sTam = "0" and sPro <> "" then
			sql = replace(sql,"And Id_Tamano = 0","")		
		end if
		
	end if
	if sPro <> "" then
		sql = replace(sql,"And Id_Tamano = 0","")		
	end if
	' 
	'Response.Write "<br>276 sQl:=" & sQl & "<br>"
	'Response.End
	'
    rSx1.Open sQl,conexionRS,0,1
	iExiste = 0
	if rSx1.Eof then
		iExiste = 0
		rSx1.Close
	else
		iExiste = 1
		gProductosTotal = rSx1.GetRows()
		rSx1.Close
	end if
	'Response.Write "<br>271 Paso" 
	'Response.End
	'Response.Write "<br>" & sQl
	'Response.End
	'	
	if iExiste = 0 then
		Response.Write "<center><h2 class='text-danger'>No hay datos para mostrar..!</h2><hr></center>"
		Response.End
	else
		'Response.Write "<br>84 LLEGO"
		'Response.End		
		
		Response.Write "<div class='container-fluid'>"
		
    		Response.Write "<div class='card'>"
			
				Response.Write "<div class='card-header'>"
				
					Response.Write "<div class='row'>"
						Response.Write "<div class='col col-md-12 text-danger text-left'><strong><h4><span class='label label-primary'><i class='fas fa-print'></i>&nbsp;Reporte Mensual</span></h4></strong></div>"
					Response.Write "</div>"
					
				Response.Write "</div>"
				
    			Response.Write "<div class='card-body'>"	
					
					'Response.Write "<div class='table-responsive-md'>"
					Response.Write "<div class='mygrid-wrapper-div'>"

						Response.Write "<table id='tbl_Datos' class='table table-condensed'>"
									
							Response.Write "<thead>"
								
								Response.Write "<tr>"
								
									Response.Write "<th class='text-center'>Area</th>"
									Response.Write "<th class='text-center'>Zona</th>"
									Response.Write "<th class='text-center'>Canal</th>"
									Response.Write "<th class='text-center'>Fabricante</th>"
									Response.Write "<th class='text-center'>Marca</th>"
									Response.Write "<th class='text-center'>Segmento</th>"
									Response.Write "<th class='text-center'>Tama&ntilde;o</th>"
									Response.Write "<th class='text-center'>Producto</th>"
									Response.Write "<th class='text-center'>Indicador</th>"
									Response.Write "<th class='text-center'>UniMed</th>"
									
									if IsArray(gSemanas) then
										for iSem = 0 to ubound(gSemanas,2)										
											Response.Write "<th class='text-center'>" & Trim(gSemanas(1,iSem)) & "</th>"
										next
									end if

									if IsArray(gMeses) then
										for iMes = 0 to ubound(gMeses,2)
											Response.Write "<th class='text-center'>" & Trim(gMeses(1,iMes)) & "</th>"
										next									
									
									end if
 									
									
								Response.Write "</tr>"
								
							Response.Write "</thead>"
						
							Response.Write "<tbody>"
								
								TotalReg = ubound(gProductosTotal,2)
								
								FOR iPro = 0 TO  ubound(gProductosTotal,2)
									'Response.Write "<br>354 LLEGO:= " & iPro
									Response.Write "<tr>"									
										Response.Write "<td>"
											'Area
											Response.Write gProductosTotal(1,iPro) 
										Response.Write "</td>"									
										Response.Write "<td>"
											'Zona
											Response.Write gProductosTotal(3,iPro) 
										Response.Write "</td>"									
										Response.Write "<td>"
											'Canal
											Response.Write gProductosTotal(5,iPro) 
										Response.Write "</td>"									
										Response.Write "<td>"
											'Fabricante
											Response.Write gProductosTotal(7,iPro) 
										Response.Write "</td>"									
										Response.Write "<td>"
											'Marca
											Response.Write gProductosTotal(9,iPro) 
										Response.Write "</td>"									
										Response.Write "<td>"
											'Segmento
											Response.Write gProductosTotal(11,iPro) 
										Response.Write "</td>"									
										Response.Write "<td>"
											'Tamaño
											if gProductosTotal(12,iPro) <> 0 then
												Valor = gProductosTotal(13,iPro)
												Valor = replace(Valor,".",",")
												Response.Write formatnumber(Valor,2) 											
											end if										
										Response.Write "</td>"
																			
										Response.Write "<td class='text-left'>"
											'Producto
											Response.Write gProductosTotal(14,iPro) & "-" & gProductosTotal(15,iPro)
										Response.Write "</td>"
										iPro2 = iPro
										isw = 0										
										Response.Flush										
										for iInd = 0 to  ubound(gIndicadores,2)
											iPro1 = iPro
											'Response.Write "<br>354 LLEGO:= " & iPro1
												if isw = 0 then
													isw = 1
												else
													Response.Write "<tr>"
													Response.Write "<td colspan=8>"
													Response.Write "</td>"													
												end if
												Response.Write "<td class='text-center'>"
													Response.Write "<b>"
													'Response.Write gIndicadores(0,iInd) & ".-" & gIndicadores(1,iInd)
													Response.Write gIndicadores(1,iInd)
													Response.Write "</b>"
												Response.Write "</td>"
												Response.Write "<td class='text-center'>"
													Response.Write "<b>"
													Response.Write gIndicadores(2,iInd)
													Response.Write "</b>"
												Response.Write "</td>"
												Indicador = gIndicadores(0,iInd)
												Columna = Indicador + 16
												Menos = 0
												'Response.Write "<br>iPro1:=" & iPro1 & ""
												sw=0
												if IsArray(gSemanas) then
													for iSem = 0 to ubound(gSemanas,2)
														Response.Write "<td class='text-right'>"
															'Response.Write "iPro:=" & iPro & "=>"
															iSem1 = gSemanas(0,iSem)
															if CInt(iSem1) = CInt(gProductosTotal(32,iPro1)) then 
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
															Response.Write Valor
															if iPro1 > TotalReg then 
																'Response.Write "menos: " & Menos
																sw=sw+1
																exit for
															end if
														Response.Write "</td>"
													next
													ix = CInt(ubound(gSemanas,2))
													iy = 4 - ix
													if sw <> 0 then 
														iy = 4- Menos
														'Response.Write "<br>Paso:=" & iy
													end if
												end if
												if IsArray(gMeses) then
													for iMes = 0 to ubound(gMeses,2)										
														'Response.Write "<th class='text-center'>" & gMeses(2,iMes) & "-" & Columna & "</th>"
														'Response.End
														ValorMes = gMeses(2,iMes)
														sAre = gProductosTotal(0,iPro)
														sZon = gProductosTotal(2,iPro)
														sCan = gProductosTotal(4,iPro)
														sFab = gProductosTotal(6,iPro)
														sMar = gProductosTotal(8,iPro)
														sSeg = gProductosTotal(10,iPro)
														sTam = gProductosTotal(12,iPro)
														sPro = gProductosTotal(14,iPro)
														Select Case Columna
															
															Case 17	'1VentasUni
																sql = vbNullstring
																sql = " SELECT sum(VentasUni) FROM RS_DataProcSem WHERE Id_Categoria = " & sCat
																sql = sql & " And Id_Semana in ( " & gMeses(2,iMes) & ") And Id_Area in (" & sAre & ") And Id_Zona in (" & sZon & ") And Id_Canal in (" & sCan & ")"
																sql = sql & " And Id_Fabricante in (" & sFab & ") And Id_Marca in (" & sMar & ") And Id_Segmento in (" & sSeg & ") "
																if Len(sTam) > 1 then
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
																if sAre = "0" and sZon = "0" and sCan = "0" and sFab = "0" and sMar = "0" and sSeg = "0" and sTam = "0" and sPro <> "" then
																	sql = replace(sql,"And Id_Tamano = 0","")																
																end if
																'
																'Response.Write "<br>276 sQl:=" & sQl & "<br>"
																'Response.End
																'
																rSx1.Open sQl,conexionRS,0,1
																'iExiste = 0
																iValor = 0
																if rSx1.Eof then
																	'iExiste = 0
																	rSx1.Close
																	iValor = 0
																else
																	'iExiste = 1
																	gValor = rSx1.GetRows()
																	rSx1.Close
																	iValor = gValor(0,0)
																end if
																if isNull(iValor) then
																	iValor = 0
																end if
																iValor = FormatNumber(iValor,2)
																Response.Write "<td class='text-right'>"
																	Response.Write iValor
																Response.Write "</td>"
															
															Case 18	'2VentasVal
																sql = vbNullstring
																sql = " SELECT sum(VentasVal) FROM RS_DataProcSem WHERE Id_Categoria = " & sCat & " And Id_Semana in ( " & gMeses(2,iMes) & ") And Id_Area in (" & sAre & ")"
																sql = sql & " And Id_Zona in (" & sZon & ") And Id_Canal in (" & sCan & ") And Id_Fabricante in (" & sFab & ") And Id_Marca in (" & sMar & ") And Id_Segmento in (" & sSeg & ")"
																if Len(sTam) > 1 then
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
																if sAre = "0" and sZon = "0" and sCan = "0" and sFab = "0" and sMar = "0" and sSeg = "0" and sTam = "0" and sPro <> "" then
																	sql = replace(sql,"And Id_Tamano = 0","")
																else
																end if
																'
																'Response.Write "<br>276 sQl:=" & sQl & "<br>"
																'Response.End
																'
																rSx1.Open sQl,conexionRS,0,1
																'iExiste = 0
																iValor = 0
																if rSx1.Eof then
																	'iExiste = 0
																	rSx1.Close
																	iValor = 0
																else
																	'iExiste = 1
																	gValor = rSx1.GetRows()
																	rSx1.Close
																	iValor = gValor(0,0)
																end if
																if isNull(iValor) then
																	iValor = 0
																end if
																iValor = FormatNumber(iValor,2)
																Response.Write "<td class='text-right'>"
																	Response.Write iValor
																Response.Write "</td>"
															
															Case 19	'3VentasVol
																sql = vbNullstring
																sql = " SELECT sum(VentasUniMed) FROM RS_DataProcSem WHERE Id_Categoria = " & sCat & " And Id_Semana in ( " & gMeses(2,iMes) & ") And Id_Area in (" & sAre & ") And Id_Zona in (" & sZon & ")"
																sql = sql & " And Id_Canal in (" & sCan & ") And Id_Fabricante in (" & sFab & ") And Id_Marca in (" & sMar & ") And Id_Segmento in (" & sSeg & ") "
																if Len(sTam) > 1 then
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
																if sAre = "0" and sZon = "0" and sCan = "0" and sFab = "0" and sMar = "0" and sSeg = "0" and sTam = "0" and sPro <> "" then
																	sql = replace(sql,"And Id_Tamano = 0","")
																end if
																'
																'Response.Write "<br>276 sQl:=" & sQl & "<br>"
																'Response.End
																'
																rSx1.Open sQl,conexionRS,0,1
																'iExiste = 0
																iValor = 0
																if rSx1.Eof then
																	'iExiste = 0
																	rSx1.Close
																	iValor = 0
																else
																	'iExiste = 1
																	gValor = rSx1.GetRows()
																	rSx1.Close
																	iValor = gValor(0,0)
																end if
																if isNull(iValor) then
																	iValor = 0
																end if
																iValor = FormatNumber(iValor,2)
																Response.Write "<td class='text-right'>"
																	Response.Write iValor
																Response.Write "</td>"
															
															Case 20	'4VentasNo
																Response.Write "<td class='text-right'>"
																	Response.Write "IND"
																Response.Write "</td>"
															
															Case 21	'5DistNum
																sql = vbNullstring
																sql = " SELECT Max(DistribucionNum) FROM RS_DataProcSem WHERE Id_Categoria = " & sCat & " And Id_Semana in ( " & gMeses(2,iMes) & ") And Id_Area in (" & sAre & ") And Id_Zona in (" & sZon & ")"
																sql = sql & " And Id_Canal in (" & sCan & ") And Id_Fabricante in (" & sFab & ") And Id_Marca in (" & sMar & ") And Id_Segmento in (" & sSeg & ")"
																if Len(sTam) > 1 then
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
																if sAre = "0" and sZon = "0" and sCan = "0" and sFab = "0" and sMar = "0" and sSeg = "0" and sTam = "0" and sPro <> "" then
																	sql = replace(sql,"And Id_Tamano = 0","")
																end if
																'
																'Response.Write "<br>276 sQl:=" & sQl & "<br>"
																'Response.End
																'
																rSx1.Open sQl,conexionRS,0,1
																'iExiste = 0
																iValor = 0
																if rSx1.Eof then
																	'iExiste = 0
																	rSx1.Close
																	iValor = 0
																else
																	'iExiste = 1
																	gValor = rSx1.GetRows()
																	iValor = gValor(0,0)
																	'Response.Write "<br>659 Query:=" & gValor(0,0) & " Convertido:= " & cDbl(gValor(0,0)) &  "<br>"
																	if isNull(iValor) then
																		iValor = 0
																	else
																		iValor = cDbl(gValor(0,0))
																	end if
																	rSx1.Close
																end if
																if isNull(iValor) then
																	iValor = 0
																end if
																iValor = FormatNumber(iValor,2)
																Response.Write "<td class='text-right'>"
																	Response.Write iValor
																Response.Write "</td>"
															
															Case 22	'6DistPon
																sql = vbNullstring
																sql = " SELECT Max(DistribucionNum) FROM RS_DataProcSem WHERE Id_Categoria = " & sCat & " And Id_Semana in ( " & gMeses(2,iMes) & ") And Id_Area in (" & sAre & ") And Id_Zona in (" & sZon & ")"
																sql = sql & " And Id_Canal in (" & sCan & ") And Id_Fabricante in (" & sFab & ") And Id_Marca in (" & sMar & ") And Id_Segmento in (" & sSeg & ")"
																if Len(sTam) > 1 then
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
																if sAre = "0" and sZon = "0" and sCan = "0" and sFab = "0" and sMar = "0" and sSeg = "0" and sTam = "0" and sPro <> "" then
																	sql = replace(sql,"And Id_Tamano = 0","")
																end if
																'
																'Response.Write "<br>276 sQl:=" & sQl & "<br>"
																'Response.End
																'
																rSx1.Open sQl,conexionRS,0,1
																'iExiste = 0
																iValorDN = 0
																if rSx1.Eof then
																	'iExiste = 0
																	rSx1.Close
																	iValorDN = 0
																else
																	'iExiste = 1
																	gValor = rSx1.GetRows()
																	rSx1.Close
																	'Response.Write "<br>659 Query:=" & gValor(0,0) & " Convertido:= " & cDbl(gValor(0,0)) &  "<br>"
																	iValorDN = gValor(0,0)
																end if
																if isNull(iValorDN) then
																	iValorDN = 0
																end if
																iValorDN = replace(iValorDN,",",".")
																
																sql = vbNullstring
																sql = " SELECT DistribucionPon FROM RS_DataProcSem WHERE Id_Categoria = " & sCat & " And DistribucionNum = '" & iValorDN & "' And Id_Semana in ( " & gMeses(2,iMes) & ") And Id_Area in (" & sAre & ")"
																sql = sql & " And Id_Zona in (" & sZon & ") And Id_Canal in (" & sCan & ") And Id_Fabricante in (" & sFab & ") And Id_Marca in (" & sMar & ") And Id_Segmento in (" & sSeg & ")"
																if Len(sTam) > 1 then
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
																if sAre = "0" and sZon = "0" and sCan = "0" and sFab = "0" and sMar = "0" and sSeg = "0" and sTam = "0" and sPro <> "" then
																	sql = replace(sql,"And Id_Tamano = 0","")															
																end if
																'
																'Response.Write "<br>276 sQl:=" & sQl & "<br>"
																'Response.End
																'
																rSx1.Open sQl,conexionRS,0,1
																'iExiste = 0
																iValor = 0
																if rSx1.Eof then
																	'iExiste = 0
																	rSx1.Close
																	iValor = 0
																else
																	'iExiste = 1
																	gValor = rSx1.GetRows()
																	rSx1.Close
																	'Response.Write "<br>659 Query:=" & gValor(0,0) & " Convertido:= " & cDbl(gValor(0,0)) &  "<br>"
																	iValor = cDbl(gValor(0,0))
																end if
																if isNull(iValor) then
																	iValor = 0
																end if
																iValor = FormatNumber(iValor,2)
																Response.Write "<td class='text-right'>"
																	Response.Write iValor
																Response.Write "</td>"
																
															
															Case 23	'7EfecDist
																Response.Write "<td class='text-right'>"
																	Response.Write "IND"
																Response.Write "</td>"
															
															Case 24	'8ShareUni
																sql = vbNullstring
																sql = " SELECT sum(VentasUni) FROM RS_DataProcSem WHERE Id_Categoria = " & sCat & " And Id_Semana in ( " & gMeses(2,iMes) & ") And Id_Area in (" & sAre & ") And Id_Zona in (" & sZon & ")"
																sql = sql & " And Id_Canal in (" & sCan & ") And Id_Fabricante in (" & sFab & ") And Id_Marca in (" & sMar & ") And Id_Segmento in (" & sSeg & ")"
																if Len(sTam) > 1 then
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
																if sAre = "0" and sZon = "0" and sCan = "0" and sFab = "0" and sMar = "0" and sSeg = "0" and sTam = "0" and sPro <> "" then
																	sql = replace(sql,"And Id_Tamano = 0","")
																end if
																'
																'Response.Write "<br>276 sQl:=" & sQl & "<br>"
																'Response.End
																'
																rSx1.Open sQl,conexionRS,0,1
																'iExiste = 0
																iValor = 0
																if rSx1.Eof then
																	'iExiste = 0
																	rSx1.Close
																	iValor = 0
																else
																	'iExiste = 1
																	gValor = rSx1.GetRows()
																	rSx1.Close
																	'19nov21 - uFev
																	if isNull(gValor(0,0)) then
																		iValor = 0
																	else
																		iValor = cDbl(gValor(0,0)) * 100
																	end if																	
																end if
																if isNull(iValor) then
																	iValor = 0
																end if
																sql = vbNullstring
																sql = " SELECT sum(VentasUni) FROM RS_DataProcSemVentas WHERE Id_Semana in ( " & gMeses(2,iMes) & ") AND Id_Categoria = " & sCat
																'
																'Response.Write "<br>276 sQl:=" & sQl & "<br>"
																'Response.End
																'
																rSx1.Open sQl,conexionRS,0,1
																'iExiste = 0
																iValorTotal = 0
																if rSx1.Eof then
																	'iExiste = 0
																	rSx1.Close
																	iValorTotal = 0
																else
																	'iExiste = 1
																	gValor = rSx1.GetRows()
																	rSx1.Close
																	iValorTotal = gValor(0,0)
																end if
																if isNull(iValorTotal) then
																	iValorTotal = 0
																end if																
																'19nov21 - uFev
																if cDbl(iValorTotal) > 0 then
																	iValor = cDbl(iValor) / cDbl(iValorTotal)																
																	iValor = FormatNumber(iValor,2)
																else
																	iValor = 0
																end if
																Response.Write "<td class='text-right'>"
																	Response.Write iValor
																Response.Write "</td>"

															Case 25	'9ShareVol
																sql = vbNullstring
																sql = " SELECT sum(VentasUniMed) FROM RS_DataProcSem WHERE Id_Categoria = " & sCat & " And Id_Semana in ( " & gMeses(2,iMes) & ") And Id_Area in (" & sAre & ") And Id_Zona in (" & sZon & ")"
																sql = sql & " And Id_Canal in (" & sCan & ") And Id_Fabricante in (" & sFab & ") And Id_Marca in (" & sMar & ") And Id_Segmento in (" & sSeg & ")"
																if Len(sTam) > 1 then
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
																if sAre = "0" and sZon = "0" and sCan = "0" and sFab = "0" and sMar = "0" and sSeg = "0" and sTam = "0" and sPro <> "" then
																	sql = replace(sql,"And Id_Tamano = 0","")																
																end if
																'
																'Response.Write "<br>276 sQl:=" & sQl & "<br>"
																'Response.End
																'
																rSx1.Open sQl,conexionRS,0,1
																'iExiste = 0
																iValor = 0
																if rSx1.Eof then
																	'iExiste = 0
																	rSx1.Close
																	iValor = 0
																else
																	'iExiste = 1
																	gValor = rSx1.GetRows()
																	rSx1.Close
																	'19nov21 - uFev
																	if isNull(gValor(0,0)) then
																		iValor = 0
																	else
																		iValor = cDbl(gValor(0,0)) * 100
																	end if																	
																end if
																if isNull(iValor) then
																	iValor = 0
																end if
																sql = vbNullstring
																sql = " SELECT sum(VentasUniMed) FROM RS_DataProcSemVentas WHERE Id_Semana in ( " & gMeses(2,iMes) & ") AND Id_Categoria = " & sCat
																'
																'Response.Write "<br>276 sQl:=" & sQl & "<br>"
																'Response.End
																'
																rSx1.Open sQl,conexionRS,0,1
																'iExiste = 0
																iValorTotal = 0
																if rSx1.Eof then
																	'iExiste = 0
																	rSx1.Close
																	iValorTotal = 0
																else
																	'iExiste = 1
																	gValor = rSx1.GetRows()
																	rSx1.Close
																	iValorTotal = gValor(0,0)
																end if
																if isNull(iValorTotal) then
																	iValorTotal = 0
																end if
																iValor = cDbl(iValor) / cDbl(iValorTotal)
																iValor = FormatNumber(iValor,2)
																Response.Write "<td class='text-right'>"
																	Response.Write iValor
																Response.Write "</td>"
															
															Case 26	'10ShareVal
																sql = vbNullstring
																sql = " SELECT sum(VentasVal) FROM RS_DataProcSem WHERE Id_Categoria = " & sCat & " And Id_Semana in ( " & gMeses(2,iMes) & ") And Id_Area in (" & sAre & ") And Id_Zona in (" & sZon & ")"
																sql = sql & " And Id_Canal in (" & sCan & ") And Id_Fabricante in (" & sFab & ") And Id_Marca in (" & sMar & ") And Id_Segmento in (" & sSeg & ")"
																if Len(sTam) > 1 then
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
																if sAre = "0" and sZon = "0" and sCan = "0" and sFab = "0" and sMar = "0" and sSeg = "0" and sTam = "0" and sPro <> "" then
																	sql = replace(sql,"And Id_Tamano = 0","")																
																end if
																'
																'Response.Write "<br>276 sQl:=" & sQl & "<br>"
																'Response.End
																'
																rSx1.Open sQl,conexionRS,0,1
																'iExiste = 0
																iValor = 0
																if rSx1.Eof then
																	'iExiste = 0
																	rSx1.Close
																	iValor = 0
																else
																	'iExiste = 1
																	gValor = rSx1.GetRows()
																	rSx1.Close
																	'19nov21 - uFev
																	if isNull(gValor(0,0)) then
																		iValor = 0
																	else
																		iValor = cDbl(gValor(0,0)) * 100
																	end if																	
																end if
																if isNull(iValor) then
																	iValor = 0
																end if
																sql = vbNullstring
																sql = " SELECT sum(VentasVal) FROM RS_DataProcSemVentas WHERE Id_Semana in ( " & gMeses(2,iMes) & ") AND Id_Categoria = " & sCat
																'
																'Response.Write "<br>276 sQl:=" & sQl & "<br>"
																'Response.End
																'
																rSx1.Open sQl,conexionRS,0,1
																'iExiste = 0
																iValorTotal = 0
																if rSx1.Eof then
																	'iExiste = 0
																	rSx1.Close
																	iValorTotal = 0
																else
																	'iExiste = 1
																	gValor = rSx1.GetRows()
																	rSx1.Close
																	iValorTotal = gValor(0,0)
																end if
																if isNull(iValorTotal) then
																	iValorTotal = 0
																end if
																iValor = cDbl(iValor) / cDbl(iValorTotal)
																iValor = FormatNumber(iValor,2)
																Response.Write "<td class='text-right'>"
																	Response.Write iValor
																Response.Write "</td>"
															
															Case 27	'11PrecioPro
																Response.Write "<td class='text-right'>"
																	Response.Write "IND"
																Response.Write "</td>"
															
															Case 28	'12PrecioMax
																sql = vbNullstring
																sql = " SELECT Max(PrecioMax) FROM RS_DataProcSem WHERE Id_Categoria = " & sCat & " And Id_Semana in ( " & gMeses(2,iMes) & ") And Id_Area in (" & sAre & ") And Id_Zona in (" & sZon & ")"
																sql = sql & " And Id_Canal in (" & sCan & ") And Id_Fabricante in (" & sFab & ") And Id_Marca in (" & sMar & ") And Id_Segmento in (" & sSeg & ")"
																if Len(sTam) > 1 then
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
																if sAre = "0" and sZon = "0" and sCan = "0" and sFab = "0" and sMar = "0" and sSeg = "0" and sTam = "0" and sPro <> "" then
																	sql = replace(sql,"And Id_Tamano = 0","")
																end if
																'
																'Response.Write "<br>276 sQl:=" & sQl & "<br>"
																'Response.End
																'
																rSx1.Open sQl,conexionRS,0,1
																'iExiste = 0
																iValor = 0
																if rSx1.Eof then
																	'iExiste = 0
																	rSx1.Close
																	iValor = 0
																else
																	'iExiste = 1
																	gValor = rSx1.GetRows()
																	rSx1.Close
																	iValor = gValor(0,0)
																end if
																if isNull(iValor) then
																	iValor = 0
																end if
																iValor = FormatNumber(iValor,2)
																Response.Write "<td class='text-right'>"
																	Response.Write iValor
																Response.Write "</td>"

															Case 29	'13PrecioMin
																sql = vbNullstring
																sql = " SELECT Min(PrecioMin) FROM RS_DataProcSem WHERE Id_Categoria = " & sCat & " And Id_Semana in ( " & gMeses(2,iMes) & ") And Id_Area in (" & sAre & ") And Id_Zona in (" & sZon & ")"
																sql = sql & " And Id_Canal in (" & sCan & ") And Id_Fabricante in (" & sFab & ") And Id_Marca in (" & sMar & ") And Id_Segmento in (" & sSeg & ") "
																if Len(sTam) > 1 then
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
																if sAre = "0" and sZon = "0" and sCan = "0" and sFab = "0" and sMar = "0" and sSeg = "0" and sTam = "0" and sPro <> "" then
																	sql = replace(sql,"And Id_Tamano = 0","")
																end if
																'
																'Response.Write "<br>276 sQl:=" & sQl & "<br>"
																'Response.End
																'
																rSx1.Open sQl,conexionRS,0,1
																'iExiste = 0
																iValor = 0
																if rSx1.Eof then
																	'iExiste = 0
																	rSx1.Close
																	iValor = 0
																else
																	'iExiste = 1
																	gValor = rSx1.GetRows()
																	rSx1.Close
																	iValor = gValor(0,0)
																end if
																if isNull(iValor) then
																	iValor = 0
																end if
																iValor = FormatNumber(iValor,2)
																Response.Write "<td class='text-right'>"
																	Response.Write iValor
																Response.Write "</td>"
															
															Case 30	'14PrecioUni
																sql = vbNullstring
																sql = " SELECT sum(VentasVal)/sum(VentasUni) FROM RS_DataProcSem WHERE Id_Categoria = " & sCat & " And Id_Semana in ( " & gMeses(2,iMes) & ") And Id_Area in (" & sAre & ") And Id_Zona in (" & sZon & ")"
																sql = sql & " And Id_Canal in (" & sCan & ") And Id_Fabricante in (" & sFab & ") And Id_Marca in (" & sMar & ") And Id_Segmento in (" & sSeg & ")"
																if Len(sTam) > 1 then
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
																if sAre = "0" and sZon = "0" and sCan = "0" and sFab = "0" and sMar = "0" and sSeg = "0" and sTam = "0" and sPro <> "" then
																	sql = replace(sql,"And Id_Tamano = 0","")
																end if
																'
																'Response.Write "<br>276 sQl:=" & sQl & "<br>"
																'Response.End
																'
																rSx1.Open sQl,conexionRS,0,1
																'iExiste = 0
																iValor = 0
																if rSx1.Eof then
																	'iExiste = 0
																	rSx1.Close
																	iValor = 0
																else
																	'iExiste = 1
																	gValor = rSx1.GetRows()
																	rSx1.Close
																	iValor = gValor(0,0)
																end if
																if isNull(iValor) then
																	iValor = 0
																end if
																iValor = FormatNumber(iValor,2)
																Response.Write "<td class='text-right'>"
																	Response.Write iValor
																Response.Write "</td>"
															
															Case 31	'15PrecioUniMed
																sql = vbNullstring
																sql = " SELECT sum(VentasVal)/sum(VentasUniMed) FROM RS_DataProcSem WHERE Id_Categoria = " & sCat & " And Id_Semana in ( " & gMeses(2,iMes) & ")  And Id_Area in (" & sAre & ") And Id_Zona in (" & sZon & ")"
																sql = sql & " And Id_Canal in (" & sCan & ") And Id_Fabricante in (" & sFab & ") And Id_Marca in (" & sMar & ") And Id_Segmento in (" & sSeg & ")"
																if Len(sTam) > 1 then
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
																if sAre = "0" and sZon = "0" and sCan = "0" and sFab = "0" and sMar = "0" and sSeg = "0" and sTam = "0" and sPro <> "" then
																	sql = replace(sql,"And Id_Tamano = 0","")																
																end if
																'
																'Response.Write "<br>276 sQl:=" & sQl & "<br>"
																'Response.End
																'
																rSx1.Open sQl,conexionRS,0,1
																'iExiste = 0
																iValor = 0
																if rSx1.Eof then
																	'iExiste = 0
																	rSx1.Close
																	iValor = 0
																else
																	'iExiste = 1
																	gValor = rSx1.GetRows()
																	rSx1.Close
																	iValor = gValor(0,0)
																end if
																if isNull(iValor) then
																	iValor = 0
																end if
																iValor = FormatNumber(iValor,2)
																Response.Write "<td class='text-right'>"
																	Response.Write iValor
																Response.Write "</td>"
														end Select 
														''
														Response.Flush
														''
													next									
													'Response.End
												end if
												''
												Response.Flush
												''
											Response.Write "</tr>"
										next
									Response.Write "</tr>"
									Response.Flush
									if IsArray(gSemanas) then
										iPro = iPro2 + iPro1 - 1
										iPro = iPro1 - 1
									end if
									'
								NEXT		
								'	
								'ElapsedTime = Timer - StartTime
								'Response.Write "<br>Proceso tardo: " & Cstr(ElapsedTime) & " Segundos."
								
							Response.Write "</tbody>"							
							
						Response.Write "</table>"
						
    			    Response.Write "</div>"
    			Response.Write "</div>"
    		Response.Write "</div>"
    	Response.Write "</div>"

	end if
	'Response.End
	'
	' Cerrar conexiones
	'	
	conexionRS.Close : Set conexionRS = Nothing	
	Set rSx1 = Nothing	
%>
