<%@language=vbscript%>

<!--#include file="Conexion.asp"-->
 
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

	sCat=Request.QueryString("cat")
	if sCat = "" Then sCat = 1
	
	sAre=Request.QueryString("are")
	sFab=Request.QueryString("fab")
	sMar=Request.QueryString("mar")
	sSeg=Request.QueryString("seg")
	sRan=Request.QueryString("ran")
	sInd=Request.QueryString("ind")
	
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
	rsx1.LockType = 2 'adLockOptimistic 

	
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

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Indicador, "
	sql = sql & " Abreviatura, "
	sql = sql & " UnidadMedida "
	sql = sql & " FROM "
	sql = sql & " PH_Indicadores "
	sql = sql & " WHERE "
	sql = sql & " Ind_Activo = 1 " 
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
	
	sql = ""
    sql = sql & " SELECT "
	sql = sql & " Id_Fabricante, "	'0
	sql = sql & " Fabricante, "		'1
	sql = sql & " Id_Marca, "		'2
	sql = sql & " Marca, "			'3
	sql = sql & " Id_Segmento, "	'4
	sql = sql & " Segmento, "		'5
	sql = sql & " Id_RangoTamano, "	'6
	sql = sql & " RangoTamano "		'7
	sql = sql & " FROM PH_DataCruda "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria = " & sCat
	sql = sql & " GROUP BY "
	sql = sql & " Id_Fabricante, "
	sql = sql & " Fabricante, "
	sql = sql & " Id_Marca, "
	sql = sql & " Marca, "
	sql = sql & " Id_Segmento, "
	sql = sql & " Segmento, "
	sql = sql & " Id_RangoTamano, "
	sql = sql & " RangoTamano "
	isw = 0
	if sFab <> "" then
		if isw = 0 then
			sql = sql & " HAVING "
			isw = 1
		else
			sql = sql & " AND "
		end if
		sql = sql & " Id_Fabricante in (" & sFab & ")"
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
	sql = sql & " Id_Fabricante, "
	sql = sql & " Id_Marca, "
	sql = sql & " Id_Segmento, "
	sql = sql & " Id_RangoTamano "
	'response.write "<br>36 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	iExiste = 0
	'response.write "<br>84 LLEGO"
	'response.end
	if rsx1.eof then
		rsx1.close
	else
		'response.write "<br>84 LLEGO"
		'response.end
		gProductos = rsx1.GetRows
		rsx1.close
		%>
		<div id="DivHomePartySem">
			<div class="ex1">
				<table class="w3-table w3-striped w3-bordered w3-card-4 w3-small" style="width:1200px; margin-left:auto; margin-right:auto;margin-top:10px ">
					<thead>
						<tr class="w3-blue">
							<th>Fabricante</th>
							<th>Marca</th>
							<th>Segmento</th>
							<th>Rango Tama&ntilde;o</th>
							<th>Indicador</th>
							<th>UniMed</th>
							<th><%=strSemana%></th>
						</tr>
					</thead>
					<%
					for iPro = 0 to  ubound(gProductos,2)
						
						response.write "<tr>"
							'Fabricante
							response.write "<td>"
								response.write gProductos(1,iPro)
							response.write "</td>"
							'Marca
							response.write "<td>"
								response.write gProductos(3,iPro)
							response.write "</td>"
							'Sepmento
							response.write "<td>"
								response.write gProductos(5,iPro)
							response.write "</td>"
							'Rango
							response.write "<td>"
								response.write gProductos(7,iPro)
							response.write "</td>"
							response.write "<td colspan=2>"
							response.write "</td>"
						response.write "</tr>"
						for iInd = 0 to  ubound(gIndicadores,2)
							response.write "<tr>"
								response.write "<td colspan=4>"
								response.write "</td>"
								response.write "<td>"
									response.write "<b>"
									response.write gIndicadores(0,iInd) & "-" & gIndicadores(1,iInd)
									response.write "</b>"
								response.write "</td>"
								response.write "<td>"
									response.write "<b>"
									response.write gIndicadores(2,iInd)
									response.write "</b>"
								response.write "</td>"
								Indicador = gIndicadores(0,iInd)
								iFab = gProductos(0,iPro)
								iMar = gProductos(2,iPro)
								iSeg = gProductos(4,iPro)
								iRan = gProductos(6,iPro)
								'response.write "<br>Ind = " & Indicador
								CalcularIndicador
								response.write "<td style='text-align:right'>"
									response.write Valor
								response.write "</td>"
							response.write "</tr>"
						next
					next
					
					%>
				</table>
			</div>
		</div>
		<%
	end if
	
Sub CalcularIndicador
	Select Case Indicador
		Case 1 'CompVol 
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Tamano, "
			sql = sql & " Cantidad "
			sql = sql & " FROM "
			sql = sql & " PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = " & sCat
			sql = sql & " And Id_Fabricante = " & iFab 
			sql = sql & " And Id_Marca = " & iMar 
			sql = sql & " And Id_Segmento = " & iSeg 
			sql = sql & " And Id_RangoTamano = " & iRan
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
			for iDat = 0 to ubound(gDatos1,2)
				Valor = Valor + (cdbl(gDatos1(0,iDat)) *cdbl(gDatos1(1,iDat)))
			next
			Valor = FormatNumber((Valor)/1000,2)
		Case 2 'CompVal
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad, "
			sql = sql & " Precio_Producto, "
			sql = sql & " Dolar, "
			sql = sql & " Precio_producto/Dolar AS PrecioDolar, "
			sql = sql & " (Precio_producto/Dolar)*Cantidad AS ComprasValor "
			sql = sql & " FROM "
			sql = sql & " PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = 1 "
			sql = sql & " And Id_Fabricante = " & iFab 
			sql = sql & " And Id_Marca = " & iMar 
			sql = sql & " And Id_Segmento = " & iSeg 
			sql = sql & " And Id_RangoTamano = " & iRan
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
			for iDat = 0 to ubound(gDatos1,2)
				Valor = Valor + cdbl(gDatos1(4,iDat))
			next
			Valor = FormatNumber(Valor,2) 
			
		Case 3 'CompUni
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad "
			sql = sql & " FROM "
			sql = sql & " PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = 1 "
			sql = sql & " and Id_Fabricante = " & iFab 
			sql = sql & " And Id_Marca = " & iMar 
			sql = sql & " And Id_Segmento = " & iSeg 
			sql = sql & " And Id_RangoTamano = " & iRan
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
			Cantidad = 0
			for iDat = 0 to ubound(gDatos1,2)
				Cantidad = Cantidad + gDatos1(0,iDat)
			next
			Valor = FormatNumber(Cantidad,0)
		Case 4 'CompAct
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad, "
			sql = sql & " Id_Consumo "
			sql = sql & " FROM PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = 1 "
			sql = sql & " AND Id_Fabricante = " & iFab
			sql = sql & " AND Id_Marca = " & iMar 
			sql = sql & " AND Id_Segmento = " & iSeg
			sql = sql & " AND Id_RangoTamano  = " & iRan
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
			Valor = 0
			Cantidad = 0
			for iDat = 0 to ubound(gDatos1,2)
				Cantidad = Cantidad + 1
			next
			Valor = FormatNumber(Cantidad,0)
		Case 5 'PenNum
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Id_Hogar AS Total "
			sql = sql & " FROM "
			sql = sql & " PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = 1 "
			sql = sql & " And Id_Fabricante = " & iFab
			sql = sql & " AND Id_Marca = " & iMar 
			sql = sql & " AND Id_Segmento = " & iSeg
			sql = sql & " AND Id_RangoTamano = " & iRan 
			sql = sql & " GROUP BY "
			sql = sql & " PH_DataCruda.Id_Hogar "
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
			Cantidad = 0
			for iDat = 0 to ubound(gDatos1,2)
				'Cantidad = gDatos1(0,0)
				Cantidad = Cantidad + 1
			next
			Valor = FormatNumber(Cantidad,0)
		Case 6 'PenPor
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Id_Hogar AS Total "
			sql = sql & " FROM "
			sql = sql & " PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = 1 "
			sql = sql & " And Id_Fabricante = " & iFab
			sql = sql & " AND Id_Marca = " & iMar 
			sql = sql & " AND Id_Segmento = " & iSeg
			sql = sql & " AND Id_RangoTamano = " & iRan 
			sql = sql & " GROUP BY "
			sql = sql & " PH_DataCruda.Id_Hogar "
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
			Cantidad = 0
			for iDat = 0 to ubound(gDatos1,2)
				'Cantidad = gDatos1(0,0)
				Cantidad = Cantidad + 1
			next
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Id_Hogar AS Total "
			sql = sql & " FROM "
			sql = sql & " PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Fabricante = " & iFab
			sql = sql & " AND Id_Marca = " & iMar 
			sql = sql & " AND Id_Segmento = " & iSeg
			sql = sql & " AND Id_RangoTamano = " & iRan 
			sql = sql & " GROUP BY "
			sql = sql & " PH_DataCruda.Id_Hogar "
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
			Valor = FormatNumber((Cantidad/Total)*100,2)
		'Case 7 'PenPonVol
		'	sql = ""
		'	sql = sql & " SELECT "
		'	sql = sql & " Tamano "
		'	sql = sql & " FROM "
		'	sql = sql & " PH_DataCruda "
		'	sql = sql & " WHERE "
		'	sql = sql & " Id_Fabricante = " & iFab 
		'	sql = sql & " And Id_Marca = " & iMar 
		'	sql = sql & " And Id_Segmento = " & iSeg 
		'	sql = sql & " And Id_RangoTamano = " & iRan
		'	'response.write "<br>36 sql:=" & sql
		'	'response.end
		'	rsx1.Open sql ,conexion
		'	'response.write "<br>500 LLEGO" 
		'	'response.end
		'	if rsx1.eof then
		'		rsx1.close
		'	else
		'		'response.write "<br>84 LLEGO"
		'		'response.end
		'		gDatos1 = rsx1.GetRows
		'		rsx1.close
		'	end if
		'	Valor = 0
		'	Cantidad = 0
		'	for iDat = 0 to ubound(gDatos1,2)
		'		Cantidad = Cantidad + cdbl(gDatos1(0,iDat))
		'	next
		'	'response.write "<br>515 LLEGO"
		'	'response.end
		'	sql = ""
		'	sql = sql & " SELECT "
		'	sql = sql & " Tamano "
		'	sql = sql & " FROM "
		'	sql = sql & " PH_DataCruda "
		'	sql = sql & " WHERE "
		'	sql = sql & " Id_Fabricante = " & iFab
		'	'response.write "<br>36 sql:=" & sql
		'	'response.end
		'	rsx1.Open sql ,conexion
		'	'response.write "<br>257 LLEGO" 
		'	'response.end
		'	if rsx1.eof then
		'		rsx1.close
		'	else
		'		'response.write "<br>84 LLEGO"
		'		'response.end
		'		gDatos1 = rsx1.GetRows
		'		rsx1.close
		'	end if
		'	Total = 0
		'	for iDat = 0 to ubound(gDatos1,2)
		'		Total = Total + cdbl(gDatos1(0,iDat))
		'	next
		'	Valor = FormatNumber((Cantidad/Total)*100,2) 
		Case 7 'PenPonVol (NO VA - Es Mensual)
			Valor = 0
		Case 8 'PenPonVal  (NO VA - Es Mensual)
			Valor = 0

		Case 9 'CompMed
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Tamano, "
			sql = sql & " Cantidad "
			sql = sql & " FROM "
			sql = sql & " PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = " & sCat
			sql = sql & " And Id_Fabricante = " & iFab 
			sql = sql & " And Id_Marca = " & iMar 
			sql = sql & " And Id_Segmento = " & iSeg 
			sql = sql & " And Id_RangoTamano = " & iRan
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
			
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Id_Hogar AS Total "
			sql = sql & " FROM "
			sql = sql & " PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = 1 "
			sql = sql & " And Id_Fabricante = " & iFab
			sql = sql & " AND Id_Marca = " & iMar 
			sql = sql & " AND Id_Segmento = " & iSeg
			sql = sql & " AND Id_RangoTamano = " & iRan 
			sql = sql & " GROUP BY "
			sql = sql & " PH_DataCruda.Id_Hogar "
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
			Valor = (cdbl(Indicador1) / cdbl(Indicador5)/1000)
			'response.write "<br>36 Indicador1=" & Indicador1
			'response.write "<br>36 Indicador5=" & Indicador5
			Valor = FormatNumber(Valor,2)

		Case 10 'GastMed
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad, "
			sql = sql & " Precio_Producto, "
			sql = sql & " Dolar, "
			sql = sql & " Precio_producto/Dolar AS PrecioDolar, "
			sql = sql & " (Precio_producto/Dolar)*Cantidad AS ComprasValor "
			sql = sql & " FROM "
			sql = sql & " PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = 1 "
			sql = sql & " And Id_Fabricante = " & iFab 
			sql = sql & " And Id_Marca = " & iMar 
			sql = sql & " And Id_Segmento = " & iSeg 
			sql = sql & " And Id_RangoTamano = " & iRan
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
			Indicador2 = 0
			for iDat = 0 to ubound(gDatos1,2)
				Indicador2 = Indicador2 + cdbl(gDatos1(4,iDat))
			next
			
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Id_Hogar AS Total "
			sql = sql & " FROM "
			sql = sql & " PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = 1 "
			sql = sql & " And Id_Fabricante = " & iFab
			sql = sql & " AND Id_Marca = " & iMar 
			sql = sql & " AND Id_Segmento = " & iSeg
			sql = sql & " AND Id_RangoTamano = " & iRan 
			sql = sql & " GROUP BY "
			sql = sql & " PH_DataCruda.Id_Hogar "
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

		Case 11 'UniComp
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad "
			sql = sql & " FROM "
			sql = sql & " PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = 1 "
			sql = sql & " and Id_Fabricante = " & iFab 
			sql = sql & " And Id_Marca = " & iMar 
			sql = sql & " And Id_Segmento = " & iSeg 
			sql = sql & " And Id_RangoTamano = " & iRan
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
			Indicador3 = 0
			for iDat = 0 to ubound(gDatos1,2)
				Indicador3 = Indicador3 + gDatos1(0,iDat)
			next
			
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Id_Hogar AS Total "
			sql = sql & " FROM "
			sql = sql & " PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = 1 "
			sql = sql & " And Id_Fabricante = " & iFab
			sql = sql & " AND Id_Marca = " & iMar 
			sql = sql & " AND Id_Segmento = " & iSeg
			sql = sql & " AND Id_RangoTamano = " & iRan 
			sql = sql & " GROUP BY "
			sql = sql & " PH_DataCruda.Id_Hogar "
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

		Case 12 'ActComp
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad, "
			sql = sql & " Id_Consumo "
			sql = sql & " FROM PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = 1 "
			sql = sql & " AND Id_Fabricante = " & iFab
			sql = sql & " AND Id_Marca = " & iMar 
			sql = sql & " AND Id_Segmento = " & iSeg
			sql = sql & " AND Id_RangoTamano  = " & iRan
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
			
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Id_Hogar AS Total "
			sql = sql & " FROM "
			sql = sql & " PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = 1 "
			sql = sql & " And Id_Fabricante = " & iFab
			sql = sql & " AND Id_Marca = " & iMar 
			sql = sql & " AND Id_Segmento = " & iSeg
			sql = sql & " AND Id_RangoTamano = " & iRan 
			sql = sql & " GROUP BY "
			sql = sql & " PH_DataCruda.Id_Hogar "
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

		Case 13 'CiCloComp
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad, "
			sql = sql & " Id_Consumo "
			sql = sql & " FROM PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = 1 "
			sql = sql & " AND Id_Fabricante = " & iFab
			sql = sql & " AND Id_Marca = " & iMar 
			sql = sql & " AND Id_Segmento = " & iSeg
			sql = sql & " AND Id_RangoTamano  = " & iRan
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
			
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Id_Hogar AS Total "
			sql = sql & " FROM "
			sql = sql & " PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = 1 "
			sql = sql & " And Id_Fabricante = " & iFab
			sql = sql & " AND Id_Marca = " & iMar 
			sql = sql & " AND Id_Segmento = " & iSeg
			sql = sql & " AND Id_RangoTamano = " & iRan 
			sql = sql & " GROUP BY "
			sql = sql & " PH_DataCruda.Id_Hogar "
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
			Valor = 7/(cdbl(Indicador4) / cdbl(Indicador5))
			'response.write "<br>36 Indicador1=" & Indicador2
			'response.write "<br>36 Indicador5=" & Indicador5
			Valor = FormatNumber(Valor,2)

		Case 14 'ActCompVol
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Tamano, "
			sql = sql & " Cantidad "
			sql = sql & " FROM "
			sql = sql & " PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = " & sCat
			sql = sql & " And Id_Fabricante = " & iFab 
			sql = sql & " And Id_Marca = " & iMar 
			sql = sql & " And Id_Segmento = " & iSeg 
			sql = sql & " And Id_RangoTamano = " & iRan
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

			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad, "
			sql = sql & " Id_Consumo "
			sql = sql & " FROM PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = 1 "
			sql = sql & " AND Id_Fabricante = " & iFab
			sql = sql & " AND Id_Marca = " & iMar 
			sql = sql & " AND Id_Segmento = " & iSeg
			sql = sql & " AND Id_RangoTamano  = " & iRan
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

		Case 15 'ActCompVal
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad, "
			sql = sql & " Precio_Producto, "
			sql = sql & " Dolar, "
			sql = sql & " Precio_producto/Dolar AS PrecioDolar, "
			sql = sql & " (Precio_producto/Dolar)*Cantidad AS ComprasValor "
			sql = sql & " FROM "
			sql = sql & " PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = 1 "
			sql = sql & " And Id_Fabricante = " & iFab 
			sql = sql & " And Id_Marca = " & iMar 
			sql = sql & " And Id_Segmento = " & iSeg 
			sql = sql & " And Id_RangoTamano = " & iRan
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
			Indicador2 = 0
			for iDat = 0 to ubound(gDatos1,2)
				Indicador2 = Indicador2 + cdbl(gDatos1(4,iDat))
			next
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad, "
			sql = sql & " Id_Consumo "
			sql = sql & " FROM PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = 1 "
			sql = sql & " AND Id_Fabricante = " & iFab
			sql = sql & " AND Id_Marca = " & iMar 
			sql = sql & " AND Id_Segmento = " & iSeg
			sql = sql & " AND Id_RangoTamano  = " & iRan
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

		Case 16 'ActCompUni
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad "
			sql = sql & " FROM "
			sql = sql & " PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = 1 "
			sql = sql & " and Id_Fabricante = " & iFab 
			sql = sql & " And Id_Marca = " & iMar 
			sql = sql & " And Id_Segmento = " & iSeg 
			sql = sql & " And Id_RangoTamano = " & iRan
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
			Indicador3 = 0
			for iDat = 0 to ubound(gDatos1,2)
				Indicador3 = Indicador3 + gDatos1(0,iDat)
			next

			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad, "
			sql = sql & " Id_Consumo "
			sql = sql & " FROM PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = 1 "
			sql = sql & " AND Id_Fabricante = " & iFab
			sql = sql & " AND Id_Marca = " & iMar 
			sql = sql & " AND Id_Segmento = " & iSeg
			sql = sql & " AND Id_RangoTamano  = " & iRan
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
			sql = sql & " PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = 1 "
			sql = sql & " And Id_Fabricante = " & iFab 
			sql = sql & " And Id_Marca = " & iMar 
			sql = sql & " And Id_Segmento = " & iSeg 
			sql = sql & " And Id_RangoTamano = " & iRan
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
			Indicador2 = 0
			for iDat = 0 to ubound(gDatos1,2)
				Indicador2 = Indicador2 + cdbl(gDatos1(4,iDat))
			next

			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Tamano, "
			sql = sql & " Cantidad "
			sql = sql & " FROM "
			sql = sql & " PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = " & sCat
			sql = sql & " And Id_Fabricante = " & iFab 
			sql = sql & " And Id_Marca = " & iMar 
			sql = sql & " And Id_Segmento = " & iSeg 
			sql = sql & " And Id_RangoTamano = " & iRan
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

		Case 25 'PrecProVal
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad, "
			sql = sql & " Precio_Producto, "
			sql = sql & " Dolar, "
			sql = sql & " Precio_producto/Dolar AS PrecioDolar, "
			sql = sql & " (Precio_producto/Dolar)*Cantidad AS ComprasValor "
			sql = sql & " FROM "
			sql = sql & " PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = 1 "
			sql = sql & " And Id_Fabricante = " & iFab 
			sql = sql & " And Id_Marca = " & iMar 
			sql = sql & " And Id_Segmento = " & iSeg 
			sql = sql & " And Id_RangoTamano = " & iRan
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
			Indicador2 = 0
			for iDat = 0 to ubound(gDatos1,2)
				Indicador2 = Indicador2 + cdbl(gDatos1(4,iDat))
			next

			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad "
			sql = sql & " FROM "
			sql = sql & " PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = 1 "
			sql = sql & " and Id_Fabricante = " & iFab 
			sql = sql & " And Id_Marca = " & iMar 
			sql = sql & " And Id_Segmento = " & iSeg 
			sql = sql & " And Id_RangoTamano = " & iRan
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

		Case 26 'MarcasHogar
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " PH_DataCruda.Marca, "
			sql = sql & " Count(Id_DataCruda) AS CantidadMarca "
			sql = sql & " FROM "
			sql = sql & " PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Marca<>0 "
			sql = sql & " GROUP BY "
			sql = sql & " Marca, "
			sql = sql & " Id_Categoria "
			sql = sql & " HAVING "
			sql = sql & " Id_Categoria = 1 "
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
			Total = 0
			for iDat = 0 to ubound(gDatos1,2)
				Total = Total + gDatos1(1,iDat)
			next
			PenetracionMarca = 0
			Suma = 0
			for iDat = 0 to ubound(gDatos1,2)
				PenetracionMarca = (gDatos1(1,iDat)*100) / Total
				Suma = Suma + PenetracionMarca 
			next
			Valor = Suma
			Valor = FormatNumber(Valor,2)
			
			
		Case 27 'CadenasProm  (NO VA - Es Mensual)
			Valor = 0
		Case 28 'CuotaMercVol
			Valor = 0
		Case 29 'CuoMerVal
			Valor = 0
		Case 30 'CuoMerUni
			Valor = 0
		Case 31 'CuoMerAct
			Valor = 0
		Case 32 'PenRel (NO VA - Es Mensual)
			Valor = 0
		Case 33 'CompRel
			Valor = 0
		Case 34 'PenAcum
			Valor = 0
		Case 35 'HogRecomp (NO VA - Es Mensual)
			Valor = 0
		Case 36 'HogNuevos (NO VA - Es Mensual)
			Valor = 0
		Case 37 'HogNoRecomp (NO VA - Es Mensual)
			Valor = 0

	end select 
end Sub



	'response.end
%>
