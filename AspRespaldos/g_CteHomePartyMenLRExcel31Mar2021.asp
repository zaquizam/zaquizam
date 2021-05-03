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
	dim TotalArea 
	dim gProductosTotal
	dim gProductosTotalNacional
	dim Contador
	Contador = 0

	sCat=Request.QueryString("cat")
	if sCat = "" Then sCat = 1
	
	sAre=Request.QueryString("are")
	sFab=Request.QueryString("fab")
	sMar=Request.QueryString("mar")
	sSeg=Request.QueryString("seg")
	sRan=Request.QueryString("ran")
	sInd=Request.QueryString("ind")
	
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
	TotalFabricante = "NO"
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
	rsx1.LockType = 2 'adLockOptimistic 

	
	strSemana1 = "Enero 2021"
	strSemana2 = "Febrero 2021"

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Indicador, "
	sql = sql & " Abreviatura, "
	sql = sql & " UnidadMedida "
	sql = sql & " FROM "
	sql = sql & " PH_Indicadores "
	sql = sql & " WHERE "
	'if Session("perusu") = 5 then
		sql = sql & " Ind_Men = 1 " 
	'else
	'	sql = sql & " Ind_Activo = 1 " 
	'end if
	if sInd <> "" then
		sql = sql & " And Id_Indicador in (" & sInd & ")"
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
	isw = 0
	if sAre <> "" then
		if isw = 0 then
			sql = sql & " HAVING "
			isw = 1
		else
			sql = sql & " AND "
		end if
		sql = sql & " Id_Area in (" & sAre & ")"
		'response.write "<br>222 Paso1"
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
			end if
			'sql = sql & " Id_Area <>0 "
		end if
		'response.write "<br>222 Paso2"
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

	'response.write "<br>313 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
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
		<table>
			<thead>
				<tr class="row100 head">
					<th class="cell100 column1 text-left">Área</th>
					<th class="cell100 column2 text-left">Fabricante</th>
					<th class="cell100 column3 text-center">Marca</th>
					<th class="cell100 column4 text-center">Segmento</th>
					<th class="cell100 column5 text-center">Rango Tamaño</th>
					<th class="cell100 column6 text-center">Tamaño</th>
					<th class="cell100 column7 text-center">Indicador</th>
					<th class="cell100 column8 text-center">UniMed</th>
					<th class="cell100 column9 text-center"><%=strSemana1%></th>
					<th class="cell100 column10 text-center"><%=strSemana2%></th>
				</tr>
			</thead>
		</table>
		<%
	else
		'response.write "<br>84 LLEGO"
		'response.end
		gProductos = rsx1.GetRows
		rsx1.close
	Response.AddHeader "Content-disposition","attachment; filename=tem.xls"
	Response.ContentType = "application/vnd.ms-excel"
		%>
		<table border=0>
			<thead>
				<tr class="row100 head">
					<th class="cell100 column1 text-left">Área</th>
					<th class="cell100 column2 text-left">Fabricante</th>
					<th class="cell100 column3 text-center">Marca</th>
					<th class="cell100 column4 text-center">Segmento</th>
					<th class="cell100 column5 text-center">Rango Tamaño</th>
					<th class="cell100 column6 text-center">Tamaño</th>
					<th class="cell100 column7 text-center">Indicador</th>
					<th class="cell100 column8 text-center">UniMed</th>
					<th class="cell100 column9 text-center"><%=strSemana1%></th>
					<th class="cell100 column10 text-center"><%=strSemana2%></th>
				</tr>
			</thead>
		</table>
								
		<table border=0>
			<% 
			if TotalArea = "SI" and sFab <> "" and sMar = "" and sRan = "" then 
				'response.write "<br>386 Total Area"
				sAre = "0"
				iAre = 1
				for iPro = 0 to  ubound(gProductosTotalNacional,2)
					for iInd = 0 to  ubound(gIndicadores,2)
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

							'Tamaño
							response.write "<td width=10% class='cell100 column6 text-center'>"
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
							idSemana = "16,17,18,19"
							TotalDias = 28
							CalcularIndicador
							response.write "<td width=10% class='cell100 column9 text-right'>"
								response.write Valor
							response.write "</td>"
							idSemana = "20,21,22,23"
							TotalDias = 28
							CalcularIndicador
							response.write "<td width=10% class='cell100 column9 text-right'>"
								response.write Valor
							response.write "</td>"
						response.write "</tr>"
					next
				next					
				response.end
			end if

			'09Feb2021
			if TotalArea = "SI" then 
				'response.write "<br>386 Total Area"
				for iPro = 0 to  ubound(gProductosTotalNacional,2)
					for iInd = 0 to  ubound(gIndicadores,2)
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

							'Tamaño
							response.write "<td width=10% class='cell100 column6 text-center'>"

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
							idSemana = "16,17,18,19"
							TotalDias = 28
							CalcularIndicador
							response.write "<td width=10% class='cell100 column9 text-right'>"
								response.write Valor
							response.write "</td>"
							idSemana = "20,21,22,23"
							TotalDias = 28
							CalcularIndicador
							response.write "<td width=10% class='cell100 column9 text-right'>"
								response.write Valor
							response.write "</td>"
						response.write "</tr>"
					next
				next					
			end if
			if sAre <> "" and sFab = "" and sMar = "" and sRan = "" then
				'response.end
				'response.write "<br>386 Todos Blanco"
				for iPro = 0 to  ubound(gProductos,2)
					'response.write "<br>428 LLEGO"
					'response.end
					for iInd = 0 to  ubound(gIndicadores,2)
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
							'Tamaño
							response.write "<td width=10% class='cell100 column6 text-center'>"
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
							idSemana = "16,17,18,19"
							TotalDias = 28
							CalcularIndicador
							response.write "<td width=10% class='cell100 column9 text-right'>"
								response.write Valor
							response.write "</td>"
							idSemana = "20,21,22,23"
							TotalDias = 28
							CalcularIndicador
							response.write "<td width=10% class='cell100 column9 text-right'>"
								response.write Valor
							response.write "</td>"
						response.write "</tr>"
					next
				next					
				
			else
			
			end if
			
			'26Ene2021-Todo el IF
			if TotalFabricante = "SI" then 
				for iPro = 0 to  ubound(gProductosTotal,2)
					for iInd = 0 to  ubound(gIndicadores,2)
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
							'Tamaño
							response.write "<td width=10% class='cell100 column6 text-center'>"

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
							idSemana = "16,17,18,19"
							TotalDias = 28
							CalcularIndicador
							response.write "<td width=10% class='cell100 column9 text-right'>"
								response.write Valor
							response.write "</td>"
							idSemana = "20,21,22,23"
							TotalDias = 28
							CalcularIndicador
							response.write "<td width=10% class='cell100 column9 text-right'>"
								response.write Valor
							response.write "</td>"
						response.write "</tr>"
					next
				next					
			end if
			
			if sFab = "" and sMar = "" and sRan = "" then
			
			else
			for iPro = 0 to  ubound(gProductos,2)
				for iInd = 0 to  ubound(gIndicadores,2)
				'response.write "<br>579 Paso"
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
						response.write "<td width=10% class='cell100 column3'>"
							if sMar <> "" then
								response.write gProductos(5,iPro)
							end if
						response.write "</td>"
						'Segmento
						response.write "<td width=10% class='cell100 column4 text-center'>"
							if sSeg <> "" then
									response.write gProductos(7,iPro)
							end if
						response.write "</td>"
						'Rango
						response.write "<td width=10% class='cell100 column5 text-center'>"
							if sRan <> "" then
								if sSeg = "" then resta = 2
								response.write gProductos(9 - resta,iPro) 
							end if
						response.write "</td>"
						'Tamaño
						response.write "<td width=10% class='cell100 column6 text-center'>"
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
						iFab = gProductos(2,iPro)
						if sMar <> "" then
							iMar = gProductos(4,iPro)
						end if
						if sSeg <> "" then
							iSeg = gProductos(6,iPro)
						end if
						if sRan <> "" then
							if sSeg = ""  then resta = 2
							iRan = gProductos(8 - resta,iPro)
						end if
						'response.write "<br>Ind = " & Indicador
						idSemana = "16,17,18,19"
						TotalDias = 28
						CalcularIndicador
						response.write "<td width=10% class='cell100 column9 text-right'>"
							response.write Valor
						response.write "</td>"
						idSemana = "20,21,22,23"
						TotalDias = 28
						CalcularIndicador
						response.write "<td width=10% class='cell100 column9 text-right'>"
							response.write Valor
						response.write "</td>"
					response.write "</td>"
				next
			next					
			end if
			%>
		</table>
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
		
		Case 6 'Penetracion Calcular - Listo
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

		Case 9 'CompraMedHog Calcular - Listo
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
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				sql = sql & " Tamano, "
				sql = sql & " Cantidad, "
				sql = sql & " Id_Consumo, "
				sql = sql & " CodigoBarra "

				'response.write "<br>1095 sql:=" & sql
				paso = 1
			else
				'response.write "Paso2"
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
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				sql = sql & " Id_Consumo, "
				sql = sql & " Producto, "
				sql = sql & " Tamano, "
				sql = sql & " Cantidad "
				paso = 2
				'response.write "<br>36 sql:=" & sql
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
				Indicador1 = FormatNumber((Valor)/1000,2)
				'response.write "<br>1149 LLEGO"
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Id_Hogar AS Total "
				sql = sql & " FROM "
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria =  "  & sCat
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
				sql = sql & " And id_Semana in( " & idSemana & ")"
				sql = sql & " GROUP BY "
				sql = sql & " PH_DataCrudaMensual.Id_Hogar "
				'response.write "<br>1173 sql:=" & sql
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
				Valor = FormatNumber(Valor,2)
			end if

		Case 10 'GastMedHog Calcular - Listo
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

		Case 11 'UnidCompHog Calcular - Listo
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
				sql = sql & " And id_Semana in( " & idSemana & ")"
				if paso = 0 then 
					sql = sql & " GROUP BY "
					sql = sql & " Cantidad, "
					sql = sql & " Id_Consumo, "
					sql = sql & " CodigoBarra "
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

		Case 12 'ActCompHog Calcular - Listo
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

		Case 13 'CiCloComp
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

		Case 14 'VolActoCompra Calcular - Listo
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
				sql = sql & " And id_Semana in( " & idSemana & ")"
				if paso = 0 then 
					sql = sql & " GROUP BY "
					sql = sql & " Tamano, "
					sql = sql & " Cantidad, "
					sql = sql & " Id_Consumo, "
					sql = sql & " CodigoBarra "
				end if
			'response.write "<br>1697 sql:=" & sql
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

		Case 15 'ValActoCompra Calcular - Listo
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
				sql = sql & " And id_Semana in( " & idSemana & ")"
				if paso = 0 then 
					sql = sql & " GROUP BY "
					sql = sql & " Cantidad, "
					sql = sql & " Precio_Producto, "
					sql = sql & " Dolar, "
					sql = sql & " Id_Consumo, "
					sql = sql & " CodigoBarra "
				end if
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
				sql = sql & " Cantidad, "
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

		Case 16 'UnidActoCompra Calcular - Listo
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
				sql = sql & " And id_Semana in( " & idSemana & ")"
				if paso = 0 then 
					sql = sql & " GROUP BY "
					sql = sql & " Cantidad, "
					sql = sql & " Id_Consumo, "
					sql = sql & " CodigoBarra "
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

		Case 24 'PrecPromVol Calcular - Listo
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

		Case 25 'PrecPromUnid Calcular - Listo
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
				sql = sql & " PH_DataCrudaMensual "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria = " & sCat
				sql = sql & " and Id_Area = " & iAre
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

		Case 26 'MarcasHogar Calcular - Listo
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
			'response.write "<br> Valor:" & Valor
			
		Case 27 'CadenasProm
			Valor = 0
		
		Case 28 'CuotaMercVol Calcular - Listo
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Tamano, "
			sql = sql & " Cantidad "
			sql = sql & " FROM "
			sql = sql & " PH_DataCrudaMensual "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = " & sCat
			sql = sql & " and Id_Area = " & iAre
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
					if iFab = 0 then TotalFiltro=TotalVolumen
					Valor = (TotalFiltro/TotalVolumen)*100
					Valor = FormatNumber(Valor,2)
					'response.write "<br> Total Volumen:= " & TotalVolumen
					'response.write "<br> Total Filtro:= " & TotalFiltro
				end if
			end if

		Case 29 'CuoMerVal Calcular - Listo
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
					
					if iFab = 0 then TotalFiltro = TotalValor
					Valor = (TotalFiltro/TotalValor)*100
					Valor = FormatNumber(Valor,2)
				end if
			end if


		Case 30 'CuoMerUni Calcular - Listo
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Cantidad "
			sql = sql & " FROM "
			sql = sql & " PH_DataCrudaMensual "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = " & sCat
			sql = sql & " and Id_Area = " & iAre
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
					
					if iFab = 0 then TotalFiltro = TotalUnidades
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
			
		
		Case 32 'PenetRelativa Calcular - Listo
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
