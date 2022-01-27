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
	' Response.End
	'	
	sSemanas = sSem
	'Response.Write "<br>84 Sem:=" & sSem	
	
	if Len(sPro) = 0 then sPro = "" end if
	if Len(sInd) = 0 then sInd = "" end if	
	if Len(sMes) = 0 then sMes = "" end if
		
	Dim gProductosTotal
	Dim gIndicadores
	Dim Indicador
	Dim Valor
	
	Dim gDatos1
	Dim rSx1
	set rSx1 = CreateObject("ADODB.Recordset")
	rSx1.CursorType = adOpenKeyset 
	rSx1.LockType = 1 'adLockOptimistic 
	'
	'Semanas	
	'
	sQl = vbNullstring
	sQl = " SELECT IdSemana, SemanaCorta FROM ss_Semana WHERE IdSemana in ( " & sSemanas & ") Order By IdSemana"
	'nse.Write "<br>151 sQl:=" & sQl & "<br>"
	'Response.End
	rSx1.Open sQl ,conexionRS
	if rSx1.Eof then
		rSx1.Close
	else
		gSemanas = rSx1.GetRows
		rSx1.Close
	end if
	'	
	'Meses
	'
	if sMes <> "" then	
		sQl = vbNullstring
		sQl = " SELECT IdPeriodo, PeriodoCorto FROM ss_Periodo WHERE Semanas IS NOT NULL AND IdPeriodo in ( " & sMes & ") Order By idPeriodo ASC "
		'Response.Write "<br>151 sQl:=" & sQl & "<br>"
		'Response.End
		rSx1.Open sQl ,conexionRS
		if rSx1.Eof then
			rSx1.Close
		else
			gMeses = rSx1.GetRows
			rSx1.Close
		end if
	'	
	end if
	'	
	' Indicadores
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
	'if (sCat >= 127 and sCat <= 145) or sCat = 41 or sCat = 18 then
	'	sQl = sQl & " and (Id_Indicador <> 3 and Id_Indicador <> 15 and Id_Indicador <> 9) "
	'end if	
	if (sCat > 126 and sCat < 146) or (sCat = 41 or sCat = 18 or sCat = 54) then
		sQl = sQl & " AND ( Id_Indicador <> 3 and Id_Indicador <> 15 and Id_Indicador <> 9 ) "
	end if
	sQl = sQl & " ORDER BY Id_Indicador "
	'
	'Response.Write "<br>191 sQl:=" & sQl & "<br>"
	'	
	'Response.End 
	'
	rSx1.Open sQl ,conexionRS
	if rSx1.Eof then
		rSx1.Close
	else
		gIndicadores = rSx1.GetRows
		rSx1.Close
	end if
	'Response.Write "<br>203 Paso" 
	'Response.End	
	''
	sQl = vbNullstring
    sQl = sQl & " SELECT "
	sQl = sQl & " Id_Area, "
	sQl = sQl & " Area, "
	sQl = sQl & " Id_Zona, "
	sQl = sQl & " Zona, "
	sQl = sQl & " Id_Canal, "
	sQl = sQl & " Canal, "
	sQl = sQl & " Id_Fabricante, "
	sQl = sQl & " Fabricante, "
	sQl = sQl & " Id_Marca, "
	sQl = sQl & " Marca, "
	sQl = sQl & " Id_Segmento, "
	sQl = sQl & " Segmento, "
	sQl = sQl & " Id_Tamano, "
	sQl = sQl & " Tamano, "
	sQl = sQl & " CodigoBarra, "
	sQl = sQl & " Descripcion, "
	sQl = sQl & " UnidadMedida, "
	sQl = sQl & " VentasUni, "			'17
	sQl = sQl & " VentasVal, "			'18
	sQl = sQl & " VentasUniMed, "		'19
	sQl = sQl & " VentasNo, "			'20
	sQl = sQl & " DistribucionNum, "	'21
	sQl = sQl & " DistribucionPon, "	'22
	sQl = sQl & " DistribucionEfe, "	'23
	sQl = sQl & " ShareUni, "			'24
	sQl = sQl & " ShareVol, "			'25
	sQl = sQl & " ShareVal, "			'26
	sQl = sQl & " PrecioPro, "			'27
	sQl = sQl & " PrecioMax, "			'28
	sQl = sQl & " PrecioMin, "			'29
	sQl = sQl & " PrecioUni, "			'30
	sQl = sQl & " PrecioUniMed, "		'31
	sQl = sQl & " id_Semana "			'32
	sQl = sQl & " FROM RS_DataProcSem WHERE "
	sQl = sQl & " Id_Categoria = " & sCat
	sQl = sQl & " And Id_Semana in ( " & sSemanas & ")"
	sQl = sQl & " And Id_Area in (" & sAre & ")"
	sQl = sQl & " And Id_Zona in (" & sZon & ")"
	sQl = sQl & " And Id_Canal in (" & sCan & ")"
	sQl = sQl & " And Id_Fabricante in (" & sFab & ")"
	sQl = sQl & " And Id_Marca in (" & sMar & ")"
	sQl = sQl & " And Id_Segmento in (" & sSeg & ")"
	if Len(sTam) > 1 then
		sQl = sQl & " And Id_Tamano in (" & sTam & ")"
	else
		sQl = sQl & " And Id_Tamano = 0 "
	end if
	if sPro <> "" then
		sPro = Replace(sPro,",","','")
		sQl = sQl & " And CodigoBarra in ('" & sPro & "')"
	else
		if sPro = "" then
			sQl = sQl & " And CodigoBarra = ''"
		end if
	end if
    sQl = sQl & " ORDER BY Id_Area, Id_Zona, Id_Canal, Id_Fabricante, Id_Marca, Id_Segmento, Id_Tamano, CodigoBarra, Descripcion, id_Semana "
	if sAre = "0" and sZon = "0" and sCan = "0" and sFab = "0" and sMar = "0" and sSeg = "0" and sTam = "0" and sPro <> "" then
		sQl = Replace(sQl,"And Id_Tamano = 0","")
	else
	end if
	'
	'Response.Write "<br>276 sQl:=" & sQl & "<br>"
	'Response.End
	'
    rSx1.Open sQl ,conexionRS
	iExiste = 0
	if rSx1.Eof then
		iExiste = 0
		rSx1.Close
	else
		iExiste = 1
		gProductosTotal = rSx1.GetRows
		rSx1.Close
	end if
	'Response.Write "<br>271 Paso" 
	'Response.End
	'Response.Write "<br>" & sQl
	'Response.End
	'
	if iExiste = 0 then			
		Response.Write "<center>"
		Response.Write "<h2 class='text-danger'>No hay datos para mostrar..!</h2>"
		Response.Write "<hr>"
		Response.Write "</center>"
		Response.End		
	else
		'Response.Write "<br>84 LLEGO"
		'Response.End		
			
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
											'Response.End
											if gProductosTotal(12,iPro) <> 0 then
												Valor = gProductosTotal(13,iPro)
												Valor = Replace(Valor,".",",")
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
												'04nov21 eliminada rutina por no hacer nada
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
										'if iPro > TotalReg then Response.End
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
	'Response.End
	'
	' Cerrar conexiones
	'	
	conexionRS.Close : Set conexionRS = Nothing	
	Set rSx1 = Nothing	
	'	
%>
