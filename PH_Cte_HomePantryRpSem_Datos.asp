<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<!-- PH_Cte_HomePantryRpSem_Datos.asp - 07feb22 - 07feb22 -->
<%
	' Variables y Constantes
	Server.ScriptTimeout = 10000	
	Session.lcid = 1034
	Response.Buffer = True	
	Response.CodePage = 65001
	Response.ContentType = "text/html"	
	Response.CharSet = "UTF-8"	
	
	Dim sCat, sFab, sMar, sSeg, sInd, sSem, gSemanas, idCliente
	Dim gProductosTotal, gIndicadores, Indicador, Valor	
	Dim gDatos1, rSx1, iSem, iPro, iPro1, iPro2, TotalReg, iColumna
	'	
	sCat=Request.Form("cat")	
	sFab=Request.Form("fab")
	sMar=Request.Form("mar")
	sSeg=Request.Form("seg")	
	sInd=Request.Form("ind")
	sSem=Request.Form("sem")
	'idCliente = Session("idCliente")
	idCliente =Request.Form("cli")
	'		
	'Response.Write "<br>" & idCliente & "<br>"
	'Response.Write "<br>" & Request.Form("cli") idCliente & "<br>"
	'Response.End 
	'		
	'Response.Write "<br>84 Sem:=" & sSem	
	' Response.End
	if Len(sInd) = 0 then sInd = "" end if				
	''
	set rSx1 = CreateObject("ADODB.Recordset")	
	'
	'Semanas		
	sQl = vbNullstring
	sQl = " SELECT IdSemana, SemanaCorta FROM ss_Semana WHERE IdSemana in ( " & sSem & ") Order By IdSemana"
	rSx1.Open sQl, conexion, 0, 1
	if rSx1.Eof then
		rSx1.Close
	else
		gSemanas = rSx1.GetRows
		rSx1.Close
	end if	
	'	
	' Indicadores
	sQl = vbNullstring
	sQl = " SELECT Id_Indicador, Abreviatura, UnidadMedida FROM PH_Indicadores WHERE Ind_Activo = 1 AND Ind_sem = 1 " 	
	' if CInt(idCliente) = 1 then
		' sQl = sQl & " AND Ind_atenas = 1 " 		
	' else
		'sQl = sQl & " AND Ind_sem = 1 " 		
	' end if
	if sInd <> "" then
		sQl = sQl & " And Id_Indicador in (" & sInd & ")"
		end if		
	sQl = sQl & " ORDER BY Id_Indicador "
	'
	'Response.Write "<br>" & sQl & "<br>"
	'Response.End 
	'
	rSx1.Open sQl, conexion, 0, 1
	if rSx1.Eof then
		rSx1.Close
	else
		gIndicadores = rSx1.GetRows
		rSx1.Close
	end if
	'Response.Write "<br>203 Paso" 
	'Response.End	
	'
	'07feb22
	'
	sQl = vbNullstring	
	sQl = sQl & " SELECT "
	sQl = sQl & " Id_Area, " 		'0
	sQl = sQl & " Area, "				'1
	sQl = sQl & " Id_Fabricante, "	'2
	sQl = sQl & " Fabricante, "		'3
	sQl = sQl & " Id_Marca, " 		'4
	sQl = sQl & " Marca, "			'5
	sQl = sQl & " Id_Segmento, "		'6
	sQl = sQl & " Segmento, "		'7
	sQl = sQl & " Id_Tamano, "		'8
	sQl = sQl & " Tamano, " 			'9
	'
	sQl = sQl & " ComprasVol, " 		'10
	sQl = sQl & " ComprasVal, "		'11
	sQl = sQl & " ComprasUnid, "		'12
	sQl = sQl & " NroActosComp, "	'13
	sQl = sQl & " PenetracionNum, " '14
	sQl = sQl & " Penetracion, "		'15
	sQl = sQl & " PenPonVol, "		'16
	sQl = sQl & " PenPonVal, "		'17
	sQl = sQl & " CompraMedHog, "	'18
	sQl = sQl & " GastMedHog, " 		'19
	sQl = sQl & " UnidCompHog, "		'20
	sQl = sQl & " ActCompHog, " 		'21
	sQl = sQl & " CicloCompra, " 	'22
	sQl = sQl & " VolActoCompra, "	'23
	sQl = sQl & " ValActoCompra, " 	'24
	sQl = sQl & " UnidActoCompra, "		'25
	sQl = sQl & " IndiceConsumoVol, " 	'26
	sQl = sQl & " IndiceConsumoVal, " 	'27
	sQl = sQl & " RepeticionCompra, "	'28
	sQl = sQl & " FidelidadVol, "		'29
	sQl = sQl & " FidelidadVal, "		'30
	sQl = sQl & " FidelidadActos, " 	'31
	sQl = sQl & " CuotaMerPonVol, " 	'32
	sQl = sQl & " PrecPromVol, "			'33
	sQl = sQl & " PrecPromUnid, "		'34
	sQl = sQl & " MarcasHogar, "			'35
	sQl = sQl & " CadenasProm, " 		'36
	sQl = sQl & " CuotaMercVol, " 		'37
	sQl = sQl & " CuotaMercVal, "		'38
	sQl = sQl & " CuotaMercUnid, "		'39
	sQl = sQl & " CuotaMercActos, "		'40
	sQl = sQl & " PenetRelativa, "		'41
	sQl = sQl & " CompRel, "				'42
	sQl = sQl & " PenetracionAcum, "	'43
	sQl = sQl & " HogRecomp, "			'44
	sQl = sQl & " HogNuevos, "			'45
	sQl = sQl & " HogNoRecomp, " 		'46
	sQl = sQl & " AporPenet, "   		'47
	sQl = sQl & " HogRecompAnt, "		'48
	sQl = sQl & " PenetracionMed, "		'49
	sQl = sQl & " id_Semana "			'50
	''
	sQl = sQl & " FROM PH_DataProcesadaSem WHERE "
	sQl = sQl & " Id_Categoria = " & sCat
	sQl = sQl & " And Id_Semana in ( " & sSem & ")"	
	sQl = sQl & " And Id_Fabricante in (" & sFab & ")"
	sQl = sQl & " And Id_Marca in (" & sMar & ")"
	sQl = sQl & " And Id_Segmento in (" & sSeg & ")"
	sQl = sQl & " ORDER BY Id_Fabricante, Id_Marca, Id_Segmento, id_Semana "	
	'
    rSx1.Open sQl, conexion, 0, 1
	if rSx1.Eof then
		rSx1.Close
	else
		gProductosTotal = rSx1.GetRows()
		rSx1.Close
	end if
	'	
	'Response.Write "<br>" & sQl
	'Response.End
	'
	if Not IsArray(gProductosTotal) then		
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
				
					'Response.Write "<div class='row'>"
					'	Response.Write "<div class='col col-md-12 text-danger text-left'><strong><h4><span class='label label-primary'><i class='fas fa-print'></i>&nbsp;Reporte Semanal</span></h4></strong></div>"
					'Response.Write "</div>"
					
				Response.Write "</div>"
				
    			Response.Write "<div class='card-body'>"	
					
					'Response.Write "<div class='table-responsive-md'>"
					Response.Write "<div class='mygrid-wrapper-div'>"

						Response.Write "<table id='tbl_datos' class='table table-condensed'>"
									
							Response.Write "<thead>"
								
								Response.Write "<tr>"
																	
									Response.Write "<th class='text-center' style='vertical-align:middle;'>Fabricante</th>"
									Response.Write "<th class='text-center' style='vertical-align:middle;'>Marca</th>"
									Response.Write "<th class='text-center' style='vertical-align:middle;'>Segmento</th>"								
									Response.Write "<th class='text-center' style='vertical-align:middle;'>Indicador</th>"									
									Response.Write "<th class='text-center' style='vertical-align:middle;'>UniMed</th>"									
									
									for iSem = 0 to ubound(gSemanas,2)										
										Response.Write "<th class='text-right' style='vertical-align:middle;'>" & Trim(gSemanas(1,iSem)) & "</th>"
									next
									
								Response.Write "</tr>"
								
							Response.Write "</thead>"
						
							Response.Write "<tbody>"
								
								TotalReg = ubound(gProductosTotal,2)
								
								FOR iPro = 0 TO ubound(gProductosTotal,2)
									'Response.Write "<br>354 LLEGO:= " & iPro
									Response.Write "<tr>"									
										Response.Write "<td>"
											'fab
											Response.Write gProductosTotal(3,iPro) 
										Response.Write "</td>"
										Response.Write "<td>"
											'Marca
											Response.Write gProductosTotal(5,iPro) 
										Response.Write "</td>"									
										Response.Write "<td>"
											'Segmento
											Response.Write gProductosTotal(7,iPro) 
										Response.Write "</td>"									
										'
										iPro2 = iPro
										isw = 0
										for iInd = 0 to ubound(gIndicadores,2)
											iPro1 = iPro
											'Response.Write "<br>354 LLEGO:= " & iPro1
												if isw = 0 then
													isw = 1
												else
													Response.Write "<tr>"
													Response.Write "<td colspan=3>"
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
												Columna = Indicador + 9
												Menos = 0
												'Response.Write "<br>iPro1:=" & iPro1 & ""
												sw=0												
												for iSem = 0 to ubound(gSemanas,2)
													Response.Write "<td class='text-right'>"
														'Response.Write "iPro:=" & iPro & "=>"
														iSem1 = gSemanas(0,iSem)
														if CInt(iSem1) = CInt(gProductosTotal(50,iPro1)) then 
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
													iy = 4 - Menos
													'Response.Write "<br>Paso:=" & iy
												end if
												'
												Menos = 0
												'
											Response.Write "</tr>"
										next
										'if iPro > TotalReg then Response.End
										iPro = iPro2 + iPro1 - 1
										iPro = iPro1 - 1
										'Response.Write "<br>iPro:=" & iPro & ""
									
									Response.Write "</tr>"
									'Separador
									'Response.Write "<tr style='height:2px;background-color:#5787C2;' >"
									Response.Write "<tr>"
										iColumna = 6 + CInt(ubound(gSemanas,2)) 
										'Response.Write "<td colspan=" & iColumna & " class='separador'>"
										Response.Write "<td colspan=" & iColumna & " style='padding:1px; font-size:1px; line-height: 0px; background-color:#afc8f4;'>"
										
										Response.Write " - "													
										Response.Write "</td>"													
									Response.Write "</tr>"
									
									Response.Flush
								NEXT								
								' Response.Write "<tr>"
									' Response.Write sql
								' Response.Write "</tr>"
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
	conexion.Close : Set conexion = Nothing	
	Set rSx1 = Nothing	
	'	
%>
