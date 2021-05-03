<!DOCTYPE HTML>
<html >
<head>
	<title>Sorteo 3</title>
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
	<!--<meta http-equiv="refresh" content="240" />-->
	<meta name="viewport" content="width=device-width, initial-scale=1">
	<meta charset="utf-8">
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<link href="w3.css" rel="stylesheet" type="text/css" media="screen" />	

<!--#include file="Conexion.asp"-->

<%
Session.LCID = 8202 
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	'ynum=Request.QueryString("num") 
	Dim gDatosSolSal
	Dim TotalHogares
	dim Consecutivo

	dim gDatosSol
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	dim rsx2
	set rsx2 = CreateObject("ADODB.Recordset")
	rsx2.CursorType = adOpenKeyset 
	rsx2.LockType = 2 'adLockOptimistic 
	
	dim rsx3
	set rsx3 = CreateObject("ADODB.Recordset")
	rsx3.CursorType = 0
	rsx3.LockType = 3

	'Borrar Tabla PH_SorteoParticipantes
	sql = ""
	sql = sql & " Delete "
	sql = sql & " From  "
	sql = sql & " PH_SorteoParticipantes "
	sql = sql & " WHERE  "
	sql = sql & " id_Sorteo = 3 "
	'response.write "<br>30 sql:=" & sql
	'response.end
    rsx2.Open sql ,conexion
	'response.write "<br>Borrar Participante"


	'Buscar Hogares
	dim gHogares
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " PH_PanelHogar.Id_PanelHogar, "
	sql = sql & " PH_PanelHogar.CodigoHogar, "
	sql = sql & " ss_Area.Area, "
	sql = sql & " ss_Estado.Estado, "
	sql = sql & " PH_Ciudad.Ciudad, "
	sql = sql & " PH_Panelistas.Nombre1, "
	sql = sql & " PH_Panelistas.Apellido1, "
	sql = sql & " PH_Panelistas.Celular "
	sql = sql & " FROM ((((PH_PanelHogar INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) INNER JOIN ss_AreaEstado ON ss_Estado.Id_Estado = ss_AreaEstado.Id_Estado) INNER JOIN ss_Area ON ss_AreaEstado.Id_Area = ss_Area.Id_Area) INNER JOIN PH_Ciudad ON PH_PanelHogar.Id_Ciudad = PH_Ciudad.Id_Ciudad) LEFT JOIN PH_Panelistas ON PH_PanelHogar.Id_PanelHogar = PH_Panelistas.Id_Hogar "
	sql = sql & " WHERE "
	sql = sql & " PH_Panelistas.ResponsablePanel = 1 "
	sql = sql & " AND PH_PanelHogar.Ind_Activo = 1 "
	sql = sql & " GROUP BY "
	sql = sql & " PH_PanelHogar.Id_PanelHogar, "
	sql = sql & " PH_PanelHogar.CodigoHogar, "
	sql = sql & " ss_Area.Area, "
	sql = sql & " ss_Estado.Estado, "
	sql = sql & " PH_Ciudad.Ciudad, "
	sql = sql & " PH_Panelistas.Nombre1, "
	sql = sql & " PH_Panelistas.Apellido1, "
	sql = sql & " PH_Panelistas.Celular "
	'sql = sql & " Having  PH_PanelHogar.Id_PanelHogar = 706 "
	sql = sql & " ORDER BY "
	sql = sql & " ss_Area.Area "
	
	'response.write "<br>220 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gHogares = rsx1.GetRows
		rsx1.close
	end if
	dim gSemanasConConsumo
	TotalHogares = ubound(gHogares,2)
	for iReg = 0 to ubound(gHogares,2)
		Hogar = cint(gHogares(0,iReg))
		sql = ""
		sql = sql & " SELECT "
		sql = sql & " Id_Hogar, "
		sql = sql & " Id_Semana, "
		sql = sql & " Count(Id_Consumo) AS CuentaDeId_Consumo "
		sql = sql & " FROM "
		sql = sql & " PH_Consumo "
		sql = sql & " WHERE "
		sql = sql & " id_TipoConsumo = 1 "
		sql = sql & " GROUP BY "
		sql = sql & " Id_Hogar, "
		sql = sql & " Id_Semana "
		sql = sql & " HAVING "
		sql = sql & " Id_Hogar = " & Hogar
		sql = sql & " AND (Id_Semana = 15 "
		sql = sql & " Or Id_Semana = 16 "
		sql = sql & " Or Id_Semana = 17 "
		sql = sql & " Or Id_Semana = 18 "
		sql = sql & " Or Id_Semana = 19) "
		'response.write "<br>220 sql:=" & sql
		'response.end
		Cantidad = 0
		rsx1.Open sql ,conexion
		if rsx1.eof then
			rsx1.close
		else
			gSemanasConConsumo = rsx1.GetRows
			rsx1.close
			Cantidad = ubound(gSemanasConConsumo,2) + 1
		end if
		'response.write "<br>paso1"
		sql = ""
		sql = sql & " SELECT "
		sql = sql & " Id_Sorteo, "
		sql = sql & " Id_Hogar, "
		sql = sql & " Ind_Ganador "
		sql = sql & " FROM "
		sql = sql & " PH_SorteoParticipantes "
		sql = sql & " WHERE "
		sql = sql & " Id_Hogar = " & Hogar 
		sql = sql & " AND Ind_Ganador = 0 "
		Ganador = 0
		rsx2.Open sql ,conexion
		if rsx2.eof then
			rsx2.close
			Ganador = 1
		else
			rsx2.close
			Ganador = 0
		end if
		'response.write "<br>paso2"
		if Ganador = 0 and Cantidad > 2 then
			'Grabar Participantes
			sql = ""
			sql = sql & " SELECT * "
			sql = sql & " From  "
			sql = sql & " PH_SorteoParticipantes "
			'response.write "<br>220 sql:=" & sql
			'response.end
			rsx3.Open sql ,conexion
			rsx3.addNew
			rsx3("Id_Sorteo") = 3
			rsx3("Id_Hogar") = gHogares(0,iReg)
			rsx3("CodigoHogar") = gHogares(1,iReg)
			rsx3("Nombre") = gHogares(5,iReg)
			rsx3("Apellido") = gHogares(6,iReg)
			rsx3("Celular") = gHogares(7,iReg)
			rsx3("Estado") = gHogares(3,iReg)
			rsx3("Ciudad") = gHogares(4,iReg)
			rsx3("Usr") = gHogares(2,iReg)
			rsx3("Ind_Ganador") = 0
			rsx3("OrdenSorteo") = 0
			rsx3.Update
			rsx3.Close 
			'set rsx3 = nothing 
		end if
	next 
	
	'response.write "FIN"
	'response.end
	'GANADORES Area I Capital
	'Buscar Maximo y Minimo id
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Max(Id_SorteoParticipante) AS MaxId, "
	sql = sql & " Min(Id_SorteoParticipante) AS MinId "
	sql = sql & " FROM "
	sql = sql & " PH_SorteoParticipantes "
	sql = sql & " Where "
	sql = sql & " usr =  'Area I Capital' "
	sql = sql & " and id_sorteo  =  3 "
	'response.write "<br>220 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gDatosSol = rsx1.GetRows
		rsx1.close
	end if
	ValorMaximo = gDatosSol(0,0)
	ValorMinimo = gDatosSol(1,0)
	'response.write "<br>220 ValorMaximo:=" & ValorMaximo
	'response.write "<br>220 ValorMinimo:=" & ValorMinimo
	iGanador = 1
	dim rsx4
	set rsx4 = CreateObject("ADODB.Recordset")
	rsx4.CursorType = 0
	rsx4.LockType = 3
	Do While iGanador < 5
		Randomize
		NumeroRandom = CLng((ValorMaximo-ValorMinimo+1)*rnd+ValorMinimo)
		sql = ""
		sql = sql & " SELECT "
		sql = sql & " Id_Hogar, "
		sql = sql & " CodigoHogar, "
		sql = sql & " Nombre, "
		sql = sql & " Apellido, "
		sql = sql & " Celular, "
		sql = sql & " Estado, "
		sql = sql & " Ciudad, "
		sql = sql & " Ind_Ganador, "
		sql = sql & " OrdenSorteo "
		sql = sql & " FROM "
		sql = sql & " PH_SorteoParticipantes "
		sql = sql & " WHERE "
		sql = sql & " Id_SorteoParticipante = " & NumeroRandom
		sql = sql & " and id_sorteo  =  3 "
		'response.write "<br>220 sql:=" & sql
		rsx1.Open sql ,conexion
		if rsx1.eof then
			rsx1.close
		else
			gDatosSolSal = rsx1.GetRows
			rsx1.close
		end if
		if gDatosSolSal(7,0) = 0 then
			sql = ""
			sql = sql & " SELECT * "
			sql = sql & " From  "
			sql = sql & " PH_SorteoParticipantes "
			sql = sql & " WHERE "
			sql = sql & " Id_SorteoParticipante = " & NumeroRandom
			sql = sql & " and id_sorteo  =  3 "
			rsx4.Open sql ,conexion
			'response.write "<br>220 sql:=" & sql
			'response.end
			rsx4("Ind_Ganador") = 1
			rsx4("OrdenSorteo") = iGanador
			rsx4.Update
			rsx4.Close 
			iGanador = iGanador + 1
		end if
	loop 
	
	Consecutivo = 0
	
	'GANADORES Area II Occidente
	'Buscar Maximo y Minimo id
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Max(Id_SorteoParticipante) AS MaxId, "
	sql = sql & " Min(Id_SorteoParticipante) AS MinId "
	sql = sql & " FROM "
	sql = sql & " PH_SorteoParticipantes "
	sql = sql & " Where "
	sql = sql & " usr =  'Area II Occidente' "
	sql = sql & " and id_sorteo  =  3 "
	'response.write "<br>220 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gDatosSol = rsx1.GetRows
		rsx1.close
	end if
	ValorMaximo = gDatosSol(0,0)
	ValorMinimo = gDatosSol(1,0)
	'response.write "<br>220 ValorMaximo:=" & ValorMaximo
	'response.write "<br>220 ValorMinimo:=" & ValorMinimo
	iGanador = 1
	Do While iGanador < 5
		Randomize
		NumeroRandom = CLng((ValorMaximo-ValorMinimo+1)*rnd+ValorMinimo)
		sql = ""
		sql = sql & " SELECT "
		sql = sql & " Id_Hogar, "
		sql = sql & " CodigoHogar, "
		sql = sql & " Nombre, "
		sql = sql & " Apellido, "
		sql = sql & " Celular, "
		sql = sql & " Estado, "
		sql = sql & " Ciudad, "
		sql = sql & " Ind_Ganador, "
		sql = sql & " OrdenSorteo "
		sql = sql & " FROM "
		sql = sql & " PH_SorteoParticipantes "
		sql = sql & " WHERE "
		sql = sql & " Id_SorteoParticipante = " & NumeroRandom
		sql = sql & " and id_sorteo  =  3 "
		'response.write "<br>220 sql:=" & sql
		rsx1.Open sql ,conexion
		if rsx1.eof then
			rsx1.close
		else
			gDatosSolSal = rsx1.GetRows
			rsx1.close
		end if
		if gDatosSolSal(7,0) = 0 then
			sql = ""
			sql = sql & " SELECT * "
			sql = sql & " From  "
			sql = sql & " PH_SorteoParticipantes "
			sql = sql & " WHERE "
			sql = sql & " Id_SorteoParticipante = " & NumeroRandom
			sql = sql & " and id_sorteo  =  3 "
			rsx4.Open sql ,conexion
			'response.write "<br>220 sql:=" & sql
			'response.end
			rsx4("Ind_Ganador") = 1
			rsx4("OrdenSorteo") = Consecutivo
			rsx4.Update
			rsx4.Close 
			iGanador = iGanador + 1
			Consecutivo = Consecutivo + 1
		end if
	loop 

	'GANADORES Area III Oriente
	'Buscar Maximo y Minimo id
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Max(Id_SorteoParticipante) AS MaxId, "
	sql = sql & " Min(Id_SorteoParticipante) AS MinId "
	sql = sql & " FROM "
	sql = sql & " PH_SorteoParticipantes "
	sql = sql & " Where "
	sql = sql & " usr =  'Area III Oriente' "
	sql = sql & " and id_sorteo  =  3 "
	'response.write "<br>220 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gDatosSol = rsx1.GetRows
		rsx1.close
	end if
	ValorMaximo = gDatosSol(0,0)
	ValorMinimo = gDatosSol(1,0)
	'response.write "<br>220 ValorMaximo:=" & ValorMaximo
	'response.write "<br>220 ValorMinimo:=" & ValorMinimo
	iGanador = 1
	Do While iGanador < 5
		Randomize
		NumeroRandom = CLng((ValorMaximo-ValorMinimo+1)*rnd+ValorMinimo)
		sql = ""
		sql = sql & " SELECT "
		sql = sql & " Id_Hogar, "
		sql = sql & " CodigoHogar, "
		sql = sql & " Nombre, "
		sql = sql & " Apellido, "
		sql = sql & " Celular, "
		sql = sql & " Estado, "
		sql = sql & " Ciudad, "
		sql = sql & " Ind_Ganador, "
		sql = sql & " OrdenSorteo "
		sql = sql & " FROM "
		sql = sql & " PH_SorteoParticipantes "
		sql = sql & " WHERE "
		sql = sql & " Id_SorteoParticipante = " & NumeroRandom
		sql = sql & " and id_sorteo  =  3 "
		'response.write "<br>220 sql:=" & sql
		rsx1.Open sql ,conexion
		if rsx1.eof then
			rsx1.close
		else
			gDatosSolSal = rsx1.GetRows
			rsx1.close
		end if
		if gDatosSolSal(7,0) = 0 then
			sql = ""
			sql = sql & " SELECT * "
			sql = sql & " From  "
			sql = sql & " PH_SorteoParticipantes "
			sql = sql & " WHERE "
			sql = sql & " Id_SorteoParticipante = " & NumeroRandom
			rsx4.Open sql ,conexion
			'response.write "<br>220 sql:=" & sql
			'response.end
			rsx4("Ind_Ganador") = 1
			rsx4("OrdenSorteo") = Consecutivo 
			rsx4.Update
			rsx4.Close 
			iGanador = iGanador + 1
			Consecutivo = Consecutivo + 1
		end if
	loop 

	'GANADORES Area IV Centro
	'Buscar Maximo y Minimo id
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Max(Id_SorteoParticipante) AS MaxId, "
	sql = sql & " Min(Id_SorteoParticipante) AS MinId "
	sql = sql & " FROM "
	sql = sql & " PH_SorteoParticipantes "
	sql = sql & " Where "
	sql = sql & " usr =  'Area IV Centro' "
	sql = sql & " and id_sorteo  =  3 "
	'response.write "<br>220 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gDatosSol = rsx1.GetRows
		rsx1.close
	end if
	ValorMaximo = gDatosSol(0,0)
	ValorMinimo = gDatosSol(1,0)
	'response.write "<br>220 ValorMaximo:=" & ValorMaximo
	'response.write "<br>220 ValorMinimo:=" & ValorMinimo
	iGanador = 1
	Do While iGanador < 5
		Randomize
		NumeroRandom = CLng((ValorMaximo-ValorMinimo+1)*rnd+ValorMinimo)
		sql = ""
		sql = sql & " SELECT "
		sql = sql & " Id_Hogar, "
		sql = sql & " CodigoHogar, "
		sql = sql & " Nombre, "
		sql = sql & " Apellido, "
		sql = sql & " Celular, "
		sql = sql & " Estado, "
		sql = sql & " Ciudad, "
		sql = sql & " Ind_Ganador, "
		sql = sql & " OrdenSorteo "
		sql = sql & " FROM "
		sql = sql & " PH_SorteoParticipantes "
		sql = sql & " WHERE "
		sql = sql & " Id_SorteoParticipante = " & NumeroRandom
		sql = sql & " and id_sorteo  =  3 "
		'response.write "<br>220 sql:=" & sql
		rsx1.Open sql ,conexion
		if rsx1.eof then
			rsx1.close
		else
			gDatosSolSal = rsx1.GetRows
			rsx1.close
		end if
		if gDatosSolSal(7,0) = 0 then
			sql = ""
			sql = sql & " SELECT * "
			sql = sql & " From  "
			sql = sql & " PH_SorteoParticipantes "
			sql = sql & " WHERE "
			sql = sql & " Id_SorteoParticipante = " & NumeroRandom
			sql = sql & " and id_sorteo  =  3 "
			rsx4.Open sql ,conexion
			'response.write "<br>220 sql:=" & sql
			'response.end
			rsx4("Ind_Ganador") = 1
			rsx4("OrdenSorteo") = Consecutivo 
			Consecutivo = Consecutivo + 1
			rsx4.Update
			rsx4.Close 
			iGanador = iGanador + 1
		end if
	loop 


	'GANADORES Area V Centro Occidente
	'Buscar Maximo y Minimo id
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Max(Id_SorteoParticipante) AS MaxId, "
	sql = sql & " Min(Id_SorteoParticipante) AS MinId "
	sql = sql & " FROM "
	sql = sql & " PH_SorteoParticipantes "
	sql = sql & " Where "
	sql = sql & " usr =  'Area V Centro Occidente' "
	sql = sql & " and id_sorteo  =  3 "
	'response.write "<br>220 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gDatosSol = rsx1.GetRows
		rsx1.close
	end if
	ValorMaximo = gDatosSol(0,0)
	ValorMinimo = gDatosSol(1,0)
	'response.write "<br>220 ValorMaximo:=" & ValorMaximo
	'response.write "<br>220 ValorMinimo:=" & ValorMinimo
	iGanador = 1
	Do While iGanador < 5
		Randomize
		NumeroRandom = CLng((ValorMaximo-ValorMinimo+1)*rnd+ValorMinimo)
		sql = ""
		sql = sql & " SELECT "
		sql = sql & " Id_Hogar, "
		sql = sql & " CodigoHogar, "
		sql = sql & " Nombre, "
		sql = sql & " Apellido, "
		sql = sql & " Celular, "
		sql = sql & " Estado, "
		sql = sql & " Ciudad, "
		sql = sql & " Ind_Ganador, "
		sql = sql & " OrdenSorteo "
		sql = sql & " FROM "
		sql = sql & " PH_SorteoParticipantes "
		sql = sql & " WHERE "
		sql = sql & " Id_SorteoParticipante = " & NumeroRandom
		sql = sql & " and id_sorteo  =  3 "
		'response.write "<br>220 sql:=" & sql
		rsx1.Open sql ,conexion
		if rsx1.eof then
			rsx1.close
		else
			gDatosSolSal = rsx1.GetRows
			rsx1.close
		end if
		if gDatosSolSal(7,0) = 0 then
			sql = ""
			sql = sql & " SELECT * "
			sql = sql & " From  "
			sql = sql & " PH_SorteoParticipantes "
			sql = sql & " WHERE "
			sql = sql & " Id_SorteoParticipante = " & NumeroRandom
			sql = sql & " and id_sorteo  =  3 "
			rsx4.Open sql ,conexion
			'response.write "<br>220 sql:=" & sql
			'response.end
			rsx4("Ind_Ganador") = 1
			rsx4("OrdenSorteo") = Consecutivo 
			rsx4.Update
			rsx4.Close 
			iGanador = iGanador + 1
			Consecutivo = Consecutivo + 1
		end if
	loop 
	
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Hogar, "
	sql = sql & " CodigoHogar, "
	sql = sql & " Nombre, "
	sql = sql & " Apellido, "
	sql = sql & " Celular, "
	sql = sql & " Estado, "
	sql = sql & " Ciudad, "
	sql = sql & " Ind_Ganador, "
	sql = sql & " OrdenSorteo "
	sql = sql & " FROM "
	sql = sql & " PH_SorteoParticipantes "
	sql = sql & " WHERE "
	sql = sql & " Id_Sorteo = 3 "
	sql = sql & " and Ind_Ganador = 1 "
	'response.write "<br>220 sql:=" & sql
	'response.end
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gDatosSolSal = rsx1.GetRows
		rsx1.close
	end if
	
	
		%>
		<center>
		<h1>Sorteo Panel de Hogares 04 Diciembre 2021</h1>
		<h2>Total Hogares Participantes: <%=TotalHogares %></h2>
		<center>
		<div id="DivBuscarPanelistas">
			<div class="ex1">
				<table class="w3-table w3-striped w3-bordered w3-card-4 w3-small" style="width:1000px; margin-left:auto; margin-right:auto;margin-top:10px ">
					<thead>
						<tr class="w3-blue">
							<th>#</th>
							<!--<th>Numero Aleatorio</th>-->
							<th>IdHogar</th>
							<th>Codigo</th>
							<th>Nombre</th>
							<th>Apellido</th>
							<th>Celular</th>
							<th>Estado</th>
							<th>Ciudad</th>
						</tr>
					</thead>
					<%
					for iReg = 0 to ubound(gDatosSolSal,2)
						Response.Write "<tr>"
							Response.Write "<td>" & iGanador & "</td>"
							'Response.Write "<td>" & NumeroRandom & "</td>"
							Response.Write "<td>" & gDatosSolSal(0,iReg) & "</td>"
							Response.Write "<td>" & gDatosSolSal(1,iReg) & "</td>"
							Response.Write "<td>" & gDatosSolSal(2,iReg) & "</td>"
							Response.Write "<td>" & gDatosSolSal(3,iReg) & "</td>"
							Response.Write "<td>" & gDatosSolSal(4,iReg) & "</td>"
							Response.Write "<td>" & gDatosSolSal(5,iReg) & "</td>"
							Response.Write "<td>" & gDatosSolSal(6,iReg) & "</td>"
						Response.Write "</tr>"
					next 
					%>
				</table>
			</div>
		</div>
		<%

	'response.write "<br>Grabar Participante"


	'response.write "<br>LLEGOFINAL" 
	response.end
	
	
%>
</body>
</html>
