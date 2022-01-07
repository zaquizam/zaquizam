<%@language=vbscript%>
<!--#include file="conexionRS.asp"-->
<!-- RetSem_Excel.asp - 12oct21 - 04ene22 -->
<%
	' Variables y Constantes
	Response.buffer = True	
	Session.lcid = 1034
	Response.ContentType = "text/html"	
	Response.CodePage = 65001
	Response.CharSet = "UTF-8"
			
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
	' Response.Write "Pro " & sPro & "<br>"
	' response.end
	'	
	sSemanas = sSem
	'Response.Write "<br>84 Sem:=" & sSem	
	
	if len(sPro) = 0 then sPro = "" end if
	if len(sInd) = 0 then sInd = "" end if	
	if len(sMes) = 0 then sMes = "" end if
		
	Dim gProductosTotal
	Dim gIndicadores
	Dim Indicador
	Dim Valor
	
	Dim gDatos1
	Dim rSx1
	set rSx1 = CreateObject("ADODB.Recordset")
	rSx1.CursorType = adOpenKeyset 
	rSx1.LockType = 1 'adLockOptimistic 

	'Semanas	
	sQl = vbnullstring
	sQl = sQl & " SELECT "
	sQl = sQl & " IdSemana, "
	sQl = sQl & " SemanaCorta "
	sQl = sQl & " FROM "
	sQl = sQl & " ss_Semana "
	sQl = sQl & " WHERE "
	sQl = sQl & " IdSemana in ( " & sSemanas & ")"	
	sQl = sQl & " Order By "
	sQl = sQl & " IdSemana "
	'Response.Write "<br>151 sQl:=" & sQl & "<br>"
	'response.end
	rSx1.Open sQl ,conexionRS
	if rSx1.eof then
		rSx1.close
	else
		gSemanas = rSx1.GetRows
		rSx1.close
	end if
	'	
	'Meses
	'
	if sMes <> "" then
	
		sQl = vbnullstring
		sQl = sQl & " SELECT "
		sQl = sQl & " IdPeriodo, "
		sQl = sQl & " PeriodoCorto "
		sQl = sQl & " FROM "
		sQl = sQl & " ss_Periodo"
		sQl = sQl & " WHERE Semanas IS NOT NULL"
		sQl = sQl & " AND IdPeriodo in ( " & sMes & ")"	
		sQl = sQl & " Order By "
		sQl = sQl & " idPeriodo ASC "
		'Response.Write "<br>151 sQl:=" & sQl & "<br>"
		'response.end
		rSx1.Open sQl ,conexionRS
		if rSx1.eof then
			rSx1.close
		else
			gMeses = rSx1.GetRows
			rSx1.close
		end if
	'	
	end if
	sQl = vbnullstring
	sQl = sQl & " SELECT "
	sQl = sQl & " Id_Indicador, "
	sQl = sQl & " Abreviatura, "
	sQl = sQl & " UnidadMedida "
	sQl = sQl & " FROM "
	sQl = sQl & " RS_Indicadores "
	sQl = sQl & " WHERE "	'
	sQl = sQl & " Ind_Activo = 1 " 
	'
	if (CInt(idCliente) = 1) then
		sQl = sQl & " AND Ind_atenas = 1 " 		
	else
		sQl = sQl & " AND Ind_men = 1 " 		
	end if
	if sInd <> "" then
		sQl = sQl & " And Id_Indicador in (" & sInd & ")"
	end if
	if (sCat >= 127 and sCat <= 145) or sCat = 41 or sCat = 18 then
		sql = sql & " and (Id_Indicador <> 3 and Id_Indicador <> 15 and Id_Indicador <> 9) "
	end if	
	sQl = sQl & " ORDER BY "
	sQl = sQl & " Id_Indicador "
	'Response.Write "<br>191 sQl:=" & sQl & "<br>"
	''	
	'response.end 
	rSx1.Open sQl ,conexionRS
	if rSx1.eof then
		rSx1.close
	else
		gIndicadores = rSx1.GetRows
		rSx1.close
	end if
	'Response.Write "<br>203 Paso" 
	'response.end	
	''
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
	'
	'Response.Write "<br>276 sQl:=" & sQl & "<br>"
	'response.end
	'
    rSx1.Open sQl ,conexionRS
	iExiste = 0
	if rSx1.eof then
		iExiste = 0
		rSx1.close
	else
		iExiste = 1
		gProductosTotal = rSx1.GetRows
		rSx1.close
	end if
	'Response.Write "<br>271 Paso" 
	'response.end
	'Response.Write "<br>" & sQl
	'response.end
	'
	if iExiste = 0 then			
		Response.Write "<center>"
		Response.Write "<h2 class='text-danger'>No hay datos para mostrar..!</h2>"
		Response.Write "<hr>"
		Response.Write "</center>"
		Response.end		
	else
		'Response.Write "<br>84 LLEGO"
		'response.end		
			
		Response.Write "<div class='container-fluid'>"
		
    		Response.Write "<div class='card'>"
			
				Response.Write "<div class='card-header'>"
				
					Response.Write "<div class='row'>"
						Response.Write "<div class='col col-md-12 text-danger text-left'><strong><h4><span class='label label-primary'><i class='fas fa-print'></i>&nbsp;Reporte Semanal</span></h4></strong></div>"
					Response.Write "</div>"
					
				Response.Write "</div>"
				
    			Response.Write "<div class='card-body'>"	
					
					'Response.Write "<div class='table-responsive-md'>"
					Response.Write "<div class='mygrid-wrapper-div'>"

						Response.Write "<table id='tbl_datos' class='table table-condensed'>"
									
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
									
									for iSem = 0 to ubound(gSemanas,2)										
										Response.Write "<th class='text-center'>" & Trim(gSemanas(1,iSem)) & "</th>"
									next

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
											'Response.Write gProductosTotal(13,iPro) 
											'response.end
											if gProductosTotal(12,iPro) <> 0 then
												Valor = gProductosTotal(13,iPro)
												Valor = replace(Valor,".",",")
												Response.Write formatnumber(Valor,2) 											
											end if										
										Response.Write "</td>"
																			
										Response.Write "<td>"
											'Producto
											Response.Write gProductosTotal(14,iPro) & "-" & gProductosTotal(15,iPro)
										Response.Write "</td>"
										'Response.Write "<td width=6% colspan=7 class='cell100 column9' >"
										'Response.Write "</td>"
										'Response.Write "</tr>"
										iPro2 = iPro
										isw = 0
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
												
												Menos = 0
												'04nov41 eliminada rutina por no hacer nada
												' if iy <> 0 then  
													' for ia = 1 to iy
														' 'Response.Write "<td width=6% class='cell100 column15 text-left'>"
															' 'Valor = 0
															' 'Valor = FormatNumber(Valor,2)
														' ''	Response.Write "Demas"
														' 'Response.Write "</td>"
													' next 
												' end if
											Response.Write "</tr>"
										next
										'if iPro > TotalReg then response.end
										iPro = iPro2 + iPro1 - 1
										iPro = iPro1 - 1
										'Response.Write "<br>iPro:=" & iPro & ""
									Response.Write "</tr>"
									Response.flush
								NEXT								
								
							Response.Write "</tbody>"							
							
						Response.Write "</table>"
						
    			    Response.Write "</div>"
    			Response.Write "</div>"
    		Response.Write "</div>"
    	Response.Write "</div>"

	end if
	'response.end
%>
