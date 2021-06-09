<%@language=vbscript%>
<head>
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
</head>

<!--#include file="Conexion.asp"-->

<%
Session.LCID = 8202 
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	yMes=cint(Request.QueryString("mes"))
	yCat=Request.QueryString("cat")
	yMar=Request.QueryString("mar")
	yMesAnt=ymes-1

	dim SemanasAct
	dim SemanasAnt
	dim HogaresAnt
	dim MarcaSeleccionada
	dim gMarcas
	dim gHogares
	dim Data(40000,2)

	CONST adOpenDynamic = 2
	CONST adUseClient = 3
	CONST adVarChar = 200
	CONST adDouble = 5
	CONST adDecimal  = 14

	dim RsCuadro1
	Set RsCuadro1 = Server.CreateObject("ADODB.Recordset")
	dim RsCuadro2
	Set RsCuadro2 = Server.CreateObject("ADODB.Recordset")
	
	RsCuadro1.CursorLocation = adUseClient
	RsCuadro1.CursorType = adOpenDynamic
	RsCuadro1.Fields.Append "Marcas", adVarChar, 500
	RsCuadro1.Fields.Append "Marca", adVarChar, 500
	RsCuadro1.Fields.Append "Hogares", adDouble
	RsCuadro1.Fields.Append "Porcentaje", adDecimal
	RsCuadro1.open
	
	RsCuadro2.CursorLocation = adUseClient
	RsCuadro2.CursorType = adOpenDynamic
	RsCuadro2.Fields.Append "Marca", adVarChar, 500
	RsCuadro2.Fields.Append "Hogares", adDouble
	RsCuadro2.Fields.Append "Porcentaje", adDecimal
	RsCuadro2.open


	dim gDatosSol
	dim gDatosSol1
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	'Buscar Semanas Mes Actual
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " semanas "
	sql = sql & " FROM "
	sql = sql & " ss_Periodo "
	sql = sql & " Where "
	sql = sql & " idPeriodo = " & yMes
	'response.write "<br>220 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	'response.write "<br> Linea 223 " &
	'response.end
	iExiste = 0
	if rsx1.eof then
		iExiste = 0
		rsx1.close
	else
		gDatosSol = rsx1.GetRows
		SemanasAct = gDatosSol(0,0) 
		rsx1.close
		iExiste = 1
	end if

	'Buscar Semanas Mes Anterior
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " semanas "
	sql = sql & " FROM "
	sql = sql & " ss_Periodo "
	sql = sql & " Where "
	sql = sql & " idPeriodo = " & yMesAnt
	'response.write "<br>220 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	'response.write "<br> Linea 223 " &
	'response.end
	iExiste = 0
	if rsx1.eof then
		iExiste = 0
		rsx1.close
	else
		gDatosSol = rsx1.GetRows
		SemanasAnt = gDatosSol(0,0) 
		rsx1.close
		iExiste = 1
	end if
	'response.write "<br>42 SemanasAnt:= " & SemanasAnt
	'esponse.end 

	'Buscar Marca Seleccionada
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Marca "
	sql = sql & " FROM "
	sql = sql & " PH_DataCruda "
	sql = sql & " WHERE "
	sql = sql & " Id_Semana In (" & SemanasAct & ")"
	sql = sql & " AND Id_Categoria = " & yCat
	sql = sql & " AND Id_Marca = " & yMar
	sql = sql & " GROUP BY "
	sql = sql & " PH_DataCruda.Marca "
	'response.write "<br>220 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	'response.write "<br> Linea 223 " &
	'response.end
	iExiste = 0
	if rsx1.eof then
		iExiste = 0
		rsx1.close
		response.write "<center>"
		response.write "Esta Marca No Fue Comprada en el Periodo Actual"
		response.write "</center>"
		response.end 
	else
		gDatosSol = rsx1.GetRows
		MarcaSeleccionada = trim(gDatosSol(0,0))
		rsx1.close
		iExiste = 1
	end if
	'response.write "<br>131 MarcaSeleccionada:= " & MarcaSeleccionada
	'response.end 

	'Buscar Marca en Mes Anterior
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Hogar "
	sql = sql & " FROM "
	sql = sql & " PH_DataCruda "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria = " & cint(yCat)
	sql = sql & " AND Id_Semana In (" & SemanasAnt & ")"
	sql = sql & " AND Id_Marca = " & yMar
	sql = sql & " GROUP BY "
	sql = sql & " Id_Hogar "
	sql = sql & " ORDER BY "
	sql = sql & " Id_Hogar "
	'response.write "<br>220 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	'response.write "<br> Linea 223 " &
	'response.end
	iExiste = 0
	if rsx1.eof then
		iExiste = 0
		rsx1.close
		response.write "<center>"
		response.write "Esta Marca No Fue Comprada en el Periodo Anterior"
		response.write "</center>"
		response.end 
	else
		gDatosSol = rsx1.GetRows
		rsx1.close
		iExiste = 1
	end if
	
	HogaresAnt = ""
	for iReg = 0 to ubound(gDatosSol,2)
		HogaresAnt = HogaresAnt & gDatosSol(0,iReg) & ","
	next
	HogaresAnt = mid(HogaresAnt,1, Len(HogaresAnt)-1)
	'response.write "<br>76 HogaresAnt;=" & HogaresAnt
	'response.end
	response.write "<center>"
	response.write "Total Hogares Mes Anterior = " & ubound(gDatosSol,2) + 1
	response.write "</center>"
	HogaresMesAnteriorCuadro1 = ubound(gDatosSol,2) + 1

	'Buscar Los Hogares en mes Actual
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Hogar, "
	sql = sql & " Id_Marca, "
	sql = sql & " Marca "
	sql = sql & " FROM "
	sql = sql & " PH_DataCruda "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria = " & yCat
	sql = sql & " AND Id_Semana In (" & SemanasAct & ")"
	sql = sql & " GROUP BY "
	sql = sql & " Id_Hogar, "
	sql = sql & " Id_Marca, "
	sql = sql & " Marca "
	sql = sql & " HAVING "
	sql = sql & " Id_Hogar In (" & HogaresAnt & ")"
	sql = sql & " AND Id_Marca<>0 "
	sql = sql & " ORDER BY "
	sql = sql & " Id_Marca "
	'response.write "<br>220 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	'response.write "<br> Linea 223 " &
	'response.end
	iExiste = 0
	if rsx1.eof then
		iExiste = 0
		rsx1.close
		response.write "<center>"
		response.write "Esta Marca No Fue Recomptada para este Periodo"
		response.write "</center>"
		response.end 
	else
		gDatosSol = rsx1.GetRows
		rsx1.close
		iExiste = 1
		response.write "<center>"
		response.write "Total Hogares Mes Actual = " & ubound(gDatosSol,2) + 1
		response.write "</center>"
	end if
	for iReg = 1 to 40000
		Data(iReg,1) = 0
	next
	
	for iReg = 0 to ubound(gDatosSol,2)
		idMarca = gDatosSol(1,iReg)
		Marca = gDatosSol(2,iReg)
		Data(idMarca,1) = Data(idMarca,1) + 1
		Data(idMarca,2) = Marca
	next
	Total = ubound(gDatosSol,2) + 1
	TotalReg = 0
	for iReg = 1 to 40000
		if Data(iReg,1) <> 0 then
			Marca = Data(iReg,2)
			HogaresMar = Data(iReg,1)
			Porcentaje = (HogaresMar * 100) / Total
			Porcentaje = formatnumber(Porcentaje,2)
			'response.write "<br> Marca = " & Marca & " (" & HogaresMar & ") " & "==>" & Porcentaje & "%"
			RsCuadro2.AddNew
			RsCuadro2.Fields("Marca") = Marca
			RsCuadro2.Fields("Hogares") = HogaresMar
			RsCuadro2.Fields("Porcentaje") = formatnumber(Porcentaje,2)
			RsCuadro2.update	
			TotalReg = TotalReg + 1
		end if
	next
	RsCuadro2.Sort = "Porcentaje Desc, Marca" 
	RsCuadro2.movefirst
	
	%>
	<div id="DivData"> 
		<div class="ex1">
			<table class="w3-table w3-striped w3-bordered w3-card-4 w3-small" style="width:700px; margin-left:auto; margin-right:auto;margin-top:10px ">
				<thead>
					<tr class="w3-blue">
						<th>Tu Marca==>Destino</th>
						<th>Marca</th>
						<th>Hogares</th>
						<th>Porcentaje</th>
					</tr>
				</thead>
				<%
				for ia = 1 to TotalReg
					response.write "<tr>"
						xMarca = trim(RsCuadro2("Marca"))
						aster = ""
						
						if MarcaSeleccionada = xMarca then
							response.write "<td style='color:#FF0000'>"
								response.write MarcaSeleccionada & "==>" & xMarca
							response.write "</td>"
							aster = " (Tu Marca)"
							response.write "<td style='color:#FF0000'>"
								response.write xMarca & aster
							response.write "</td>"
							response.write "<td style='color:#FF0000'>"
								response.write cint(RsCuadro2("Hogares"))
							response.write "</td>"
							response.write "<td style='color:#FF0000'>"
								Valor = (RsCuadro2("Hogares") * 100) / HogaresMesAnteriorCuadro1
								response.write formatnumber(Valor,2)
							response.write "</td>"
						else
							response.write "<td>"
								response.write MarcaSeleccionada & "==>" & xMarca
							response.write "</td>"
							response.write "<td>"
								response.write xMarca & aster
							response.write "</td>"
							response.write "<td>"
								response.write cint(RsCuadro2("Hogares"))
							response.write "</td>"
							response.write "<td>"
								Valor = (RsCuadro2("Hogares") * 100) / HogaresMesAnteriorCuadro1
								response.write formatnumber(Valor,2)
							response.write "</td>"
						end if
					response.write "</tr>"
					RsCuadro2.MoveNext
				next 
				RsCuadro2.close
				set RsCuadro2 = nothing
				%>
			</table>
			<br>
			<%
			'*********** Segundo Cuadro
			'Buscar Marca en Mes Anterior
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Id_Hogar "
			sql = sql & " FROM "
			sql = sql & " PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = " & cint(yCat)
			sql = sql & " AND Id_Semana In (" & SemanasAnt & ")"
			sql = sql & " AND Id_Marca <> " & yMar
			sql = sql & " GROUP BY "
			sql = sql & " Id_Hogar "
			sql = sql & " ORDER BY "
			sql = sql & " PH_DataCruda.Id_Hogar "
			'response.write "<br>220 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			'response.write "<br> Linea 223 " &
			'response.end
			iExiste = 0
			if rsx1.eof then
				iExiste = 0
				rsx1.close
			else
				gDatosSol1 = rsx1.GetRows
				rsx1.close
				iExiste = 1
			end if

			HogaresAnt = ""
			for iReg = 0 to ubound(gDatosSol1,2)
				HogaresAnt = HogaresAnt & gDatosSol1(0,iReg) & ","
			next
			HogaresAnt = mid(HogaresAnt,1, Len(HogaresAnt)-1)
			TotalHogaresMesAnterior = ubound(gDatosSol1,2) + 1
			'response.write "<br>76 HogaresAnt;=" & HogaresAnt
			'response.end
			response.write "<center>"
			response.write "Total Hogares Mes Anterior = " & TotalHogaresMesAnterior
			response.write "</center>"

			'Buscar Marcas
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Id_Marca, "
			sql = sql & " Marca "
			sql = sql & " FROM "
			sql = sql & " PH_DataCruda "
			sql = sql & " WHERE "
			sql = sql & " Id_Categoria = " & yCat
			sql = sql & " AND Id_Semana In (" & SemanasAnt & ")"
			sql = sql & " AND Id_Marca <> " & yMar
			sql = sql & " GROUP BY "
			sql = sql & " Id_Marca, "
			sql = sql & " Marca "
			sql = sql & " HAVING "
			sql = sql & " Id_Marca <> 0 "
			sql = sql & " ORDER BY "
			sql = sql & " Id_Marca "
			rsx1.Open sql ,conexion
			'response.write "<br> Linea 223 " &
			'response.end
			iExiste = 0
			if rsx1.eof then
				iExiste = 0
				rsx1.close
			else
				gMarcas = rsx1.GetRows
				rsx1.close
				iExiste = 1
			end if
			TotalCuadro = 0
			for iReg = 0 to ubound(gMarcas,2)
				idMarca = gMarcas(0,iReg)
				Marca = gMarcas(1,iReg)

				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Id_Hogar "
				sql = sql & " FROM "
				sql = sql & " PH_DataCruda "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria = " & yCat 
				sql = sql & " AND Id_Semana In (" & SemanasAnt & ")"
				sql = sql & " AND Id_Marca <> " & yMar
				sql = sql & " AND Id_Marca = " & idMarca
				sql = sql & " GROUP BY "
				sql = sql & " Id_Hogar "
				sql = sql & " ORDER BY "
				sql = sql & " Id_Hogar "
				rsx1.Open sql ,conexion
				'response.write "<br> Linea 223 " &
				'response.end
				iExiste = 0
				if rsx1.eof then
					iExiste = 0
					rsx1.close
				else
					gHogares = rsx1.GetRows
					rsx1.close
					iExiste = 1
				end if
				HogaresMar = ""
				for iReg1 = 0 to ubound(gHogares,2)
					HogaresMar = HogaresMar & gHogares(0,iReg1) & ","
				next
				HogaresMar = mid(HogaresMar,1, Len(HogaresMar)-1)

				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Id_Hogar, "
				sql = sql & " Id_Marca "
				sql = sql & " FROM "
				sql = sql & " PH_DataCruda "
				sql = sql & " WHERE "
				sql = sql & " Id_Categoria = " & yCat 
				sql = sql & " AND Id_Semana In (" & SemanasAnt & ")"
				sql = sql & " AND Id_Marca <> 0 "
				sql = sql & " GROUP BY "
				sql = sql & " Id_Hogar, "
				sql = sql & " Id_Marca "
				sql = sql & " HAVING "
				sql = sql & " Id_Hogar In (" & HogaresMar & ")"
				sql = sql & " AND Id_Marca = " & yMar
				sql = sql & " ORDER BY "
				sql = sql & " Id_Hogar, "
				sql = sql & " Id_Marca "
				rsx1.Open sql ,conexion
				'response.write "<br> Linea 223 " & sql
				'response.end
				iExiste = 0
				if rsx1.eof then
					iExiste = 0
					rsx1.close
					TotalHogaresMar = 0
				else
					gX = rsx1.GetRows
					rsx1.close
					iExiste = 1
					TotalHogaresMar = ubound(gX,2) + 1
					RsCuadro1.AddNew
					RsCuadro1.Fields("Marcas") = Marca & "==>" & MarcaSeleccionada
					RsCuadro1.Fields("Marca") = Marca
					RsCuadro1.Fields("Hogares") = cint(TotalHogaresMar)
					Valor = (TotalHogaresMar * 100) / TotalHogaresMesAnterior
					RsCuadro1.Fields("Porcentaje") = Valor
					RsCuadro1.update	
					TotalCuadro = TotalCuadro + 1
				end if
			next
			RsCuadro1.Sort = "Hogares Desc, Marca" 
			RsCuadro1.movefirst
				%>
			</table>
			<table class="w3-table w3-striped w3-bordered w3-card-4 w3-small" style="width:700px; margin-left:auto; margin-right:auto;margin-top:10px ">
				<thead>
					<tr class="w3-blue">
						<th>Otra Marca==>Tu Marca</th>
						<th>Otra Marca</th>
						<th>Hogares</th>
						<th>Porcentaje</th>
					</tr>
				</thead>
				<%
				for iReg = 1 to TotalCuadro
					response.write "<tr>"
						response.write "<td>"
							response.write RsCuadro1("Marcas")
						response.write "</td>" 
						response.write "<td>"
							response.write RsCuadro1("Marca")
						response.write "</td>"
						response.write "<td>"
							response.write RsCuadro1("Hogares")
						response.write "</td>"
						response.write "<td>"
							Valor = (cint(RsCuadro1("Hogares")) * 100) / TotalHogaresMesAnterior
							response.write formatnumber(Valor,2)
						response.write "</td>"
					response.write "</tr>"
					RsCuadro1.MoveNext
				next
				%>
			</table>
			<%
			RsCuadro1.close
			set RsCuadro1 = nothing

			%>
		</div>
	</div> 
	<%
	
%>