<%@language=vbscript%>
<!--#include file="conexionRS.asp"-->
<!-- RetSem_ExcelFull.asp - 12oct21 - 29oct21 -->
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
	
	Dim iPro
	Dim iInd
	Dim idSemana
	Dim gSemanas
	Dim gSemanasAcum
	Dim TotalSemAcum
	Dim idCliente
	Dim strSemana
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
	'	
	' sCat=Request.QueryString("cat")
	' sAre=Request.QueryString("are")
	' sZon=Request.QueryString("zon")
	' sCan=Request.QueryString("can")
	' sFab=Request.QueryString("fab")
	' sMar=Request.QueryString("mar")
	' sSeg=Request.QueryString("seg")
	' sRan=Request.QueryString("ran")
	' sTam=Request.QueryString("tam")
	' sPro=Request.QueryString("pro")
	' sInd=Request.QueryString("ind")
	' sSem=Request.QueryString("sem")
	idCliente = Session("idCliente")
	strSemana = vbnullstring
	'	
	'sSemAcum=Request.Form("semacum")

	sSemanas = sSem
	'response.write "<br>84 Sem:=" & sSem
	
	
	if len(sPro) = 0 then sPro = 0 end if
	if len(sInd) = 0 then sInd = 0 end if
	''		
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
	sQl = sQl & " SELECT"
	sQl = sQl & " IdSemana,"
	sQl = sQl & " SemanaCorta"
	sQl = sQl & " FROM"
	sQl = sQl & " ss_Semana"
	sQl = sQl & " WHERE"
	sQl = sQl & " IdSemana >23"
	sQl = sQl & " and IdSemana <58"	
	sQl = sQl & " Order By"
	sQl = sQl & " IdSemana" 
	''
	'response.write "<br>151 sQl:=" & sQl & "<br>"
	'response.end
	rSx1.Open sQl , conexionRS
	if rSx1.eof then
		rSx1.close
	else
		gSemanas = rSx1.GetRows
		rSx1.close
	end if
	''	
	'	
	sQl = vbnullstring
	sQl = sQl & " SELECT "
	sQl = sQl & " Id_Indicador, "
	sQl = sQl & " Abreviatura, "
	sQl = sQl & " UnidadMedida "
	sQl = sQl & " FROM "
	sQl = sQl & " RS_Indicadores "
	sQl = sQl & " WHERE "	
	sQl = sQl & " Ind_Activo = 1 " 	
	
	if sInd <> 0 then
		sQl = sQl & " And Id_Indicador in (" & sInd & ")"
	end if
	sQl = sQl & " ORDER BY "
	sQl = sQl & " Id_Indicador "
	'response.write "<br>191 sQl:=" & sQl & "<br>"
	'response.end 
	rSx1.Open sQl ,conexionRS
	if rSx1.eof then
		rSx1.close
	else
		gIndicadores = rSx1.GetRows
		rSx1.close
	end if
	'response.write "<br>203 Paso" 
	'response.end	
	
	'Query
	sQl = vbnullstring
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
	sQl = sQl & " FROM "
	sQl = sQl & " RS_DataProcSem "
	sQl = sQl & " WHERE "
	sQl = sQl & " Id_Categoria = " & sCat
	sQl = sQl & " And Id_Semana < 58 "
	'sQl = sQl & " And Id_Semana in ( " & sSemanas & ")"
	sQl = sQl & " And Id_Area in (" & sAre & ")"
	sQl = sQl & " And Id_Zona in (" & sZon & ")"
	sQl = sQl & " And Id_Canal in (" & sCan & ")"
	sQl = sQl & " And Id_Fabricante in (" & sFab & ")"
	sQl = sQl & " And Id_Marca in (" & sMar & ")"
	sQl = sQl & " And Id_Segmento in (" & sSeg & ")"
	if len(sTam) > 1 then
		sQl = sQl & " And Id_Tamano in (" & sTam & ")"
	else
		sQl = sQl & " And Id_Tamano = 0 "
	end if
	'if len(sPro) <> "" then
	if len(sPro) > 1 then
		sPro = replace(sPro,",","','")
		sQl = sQl & " And CodigoBarra in ('" & sPro & "')"
	else
		if len(sPro) = 1 then
			sQl = sQl & " And CodigoBarra = ''"
		end if
	end if
    sQl = sQl & " ORDER BY "
	sQl = sQl & " Id_Area, "
	sQl = sQl & " Id_Zona, "
	sQl = sQl & " Id_Canal, "
	sQl = sQl & " Id_Fabricante, "
	sQl = sQl & " Id_Marca, "
	sQl = sQl & " Id_Segmento, "
	sQl = sQl & " Id_Tamano, "
	sQl = sQl & " CodigoBarra, "
	sQl = sQl & " Descripcion, "
	sQl = sQl & " id_Semana "
	if sAre = "0" and sZon = "0" and sCan = "0" and sFab = "0" and sMar = "0" and sSeg = "0" and sTam = "0" and sPro <> "" then
		sQl = replace(sQl,"And Id_Tamano = 0","")
	else
	end if
	
	'response.write "<br>276 sQl:=" & sQl & "<br>"	
	'response.end
	''
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
	'response.write "<br>271 Paso" 
	'response.end
	'response.write "<br>" & sQl
	'response.end

	if iExiste = 0 then
		
		Response.write "<center>"
		Response.write "<h2>No hay Data para Mostrar</h2>"
		Response.write "<hr>"
		Response.write "</center>"
		Response.end
		
	else
		'response.write "<br>84 LLEGO"
		'response.end
		Response.write "<div class='container-fluid'>"
		
		Response.write "<div style='height: 100px;'>"
		
    		Response.write "<div class='card'>"
			
				Response.write "<div class='card-header'>"
				
					Response.write "<div class='row'>"
						Response.write "<div class='col col-md-12 text-danger text-left'><strong><h4><span class='label label-default'><i class='fas fa-print'></i>&nbsp;Reporte Histórico Semanal</span></h4></strong></div>"
					Response.write "</div>"
					
				Response.write "</div>"
				
    			Response.write "<div class='card-body'>"	
					
					'Response.write "<div class='col-12 col-sm-12 col-lg-12 table-responsive'>"
					'Response.write "<div style='overflow-x:auto;'>"
					
					Response.write "<div class='mygrid-wrapper-div'>"
										
						Response.write "<table id='tbl_exportar_to_xls' class='table table-condensed'>"
									
							Response.write "<thead>"
								
								Response.write "<tr>"
								
									Response.write "<th class='text-center'>Area</th>"
									Response.write "<th class='text-center'>Zona</th>"
									Response.write "<th class='text-center'>Canal</th>"
									Response.write "<th class='text-center'>Fabricante</th>"
									Response.write "<th class='text-center'>Marca</th>"
									Response.write "<th class='text-center'>Segmento</th>"
									Response.write "<th class='text-center'>Tama&ntilde;o</th>"
									Response.write "<th class='text-center'>Producto</th>"
									Response.write "<th class='text-center'>Indicador</th>"
									Response.write "<th class='text-center'>UniMed</th>"
									
									for iSem = 0 to ubound(gSemanas,2)
										'strSemana = gSemanas(1,iSem)
										Response.write "<th class='text-center'>" & Trim(gSemanas(1,iSem)) & "</th>"
									next 
									
								Response.write "</tr>"
								
							Response.write "</thead>"
						
							Response.write "<tbody>"
								
								TotalReg = ubound(gProductosTotal,2)
								
								FOR iPro = 0 TO  ubound(gProductosTotal,2)
								
									'response.write "<br>354 LLEGO:= " & iPro
									
									response.write "<tr>"									
										response.write "<td>"
											'Area
											response.write gProductosTotal(1,iPro) 
										response.write "</td>"									
										response.write "<td>"
											'Zona
											response.write gProductosTotal(3,iPro) 
										response.write "</td>"									
										response.write "<td>"
											'Canal
											response.write gProductosTotal(5,iPro) 
										response.write "</td>"									
										response.write "<td>"
											'Fabricante
											response.write gProductosTotal(7,iPro) 
										response.write "</td>"									
										response.write "<td>"
											'Marca
											response.write gProductosTotal(9,iPro) 
										response.write "</td>"									
										response.write "<td>"
											'Segmento
											response.write gProductosTotal(11,iPro) 
										response.write "</td>"									
										response.write "<td>"
											'Tamaño
											'response.write gProductosTotal(13,iPro) 
											'response.end
											if gProductosTotal(12,iPro) <> 0 then
												Valor = gProductosTotal(13,iPro)
												Valor = replace(Valor,".",",")
												response.write formatnumber(Valor,2) 
											else										
											end if										
										response.write "</td>"
																			
										response.write "<td>"
											'Producto
											response.write gProductosTotal(14,iPro) & "-" & gProductosTotal(15,iPro)
										response.write "</td>"
										'response.write "<td width=6% colspan=7 class='cell100 column9' >"
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
													response.write "<td colspan=8>"
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
												for iSem = 0 to ubound(gSemanas,2)
													response.write "<td class='text-right'>"
														'response.write "iPro:=" & iPro & "=>"
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
														response.write Valor
														if iPro1 > TotalReg then 
															'response.write "menos: " & Menos
															sw=sw+1
															exit for
														end if
													response.write "</td>"												
												next					
												
												ix = CInt(ubound(gSemanas,2))
												iy = 4 - ix
												if sw <> 0 then 
													iy = 4- Menos
													'response.write "<br>Paso:=" & iy
												end if
												
												Menos = 0
												if iy <> 0 then  
													for ia = 1 to iy
														'response.write "<td width=6% class='cell100 column15 text-left'>"
															'Valor = 0
															'Valor = FormatNumber(Valor,2)
														''	response.write "Demas"
														'response.write "</td>"
													next 
												end if
											response.write "</tr>"
										next
										'if iPro > TotalReg then response.end
										iPro = iPro2 + iPro1 - 1
										iPro = iPro1 - 1
										'response.write "<br>iPro:=" & iPro & ""
									response.write "</tr>"
								NEXT								
								
							Response.write "</tbody>"
							
						Response.write "</table>"
    			    Response.write "</div>"
    			Response.write "</div>"
    		Response.write "</div>"
    	Response.write "</div>"
		Response.write "</div>"
		
		''
	end if
	
	'response.end
%>
