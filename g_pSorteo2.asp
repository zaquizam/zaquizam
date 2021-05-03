<!DOCTYPE HTML>
<html >
<head>
	<title>Sorteo</title>
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
	'response.write "<br>30 sql:=" & sql
	'response.end
    rsx2.Open sql ,conexion
	'response.write "<br>Borrar Participante"

	'ALTER1
	'sql = ""
	'sql = sql & " ALTER TABLE PH_SorteoParticipantes DROP COLUMN  Id_Sorteo "
	'response.write "<br>30 sql:=" & sql
	'response.end
    'rsx2.Open sql ,conexion
	'response.write "<br>ALTER1"

	'ALTER2
	'sql = ""
	'sql = sql & " ALTER TABLE PH_SorteoParticipantes add Id_Sorteo INT IDENTITY (1,1)	 "
	'response.write "<br>30 sql:=" & sql
	'response.end
    'rsx2.Open sql ,conexion
	'response.write "<br>ALTER1"
	


	'Buscar Participantes
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " PH_PanelHogar.Id_PanelHogar, "
	sql = sql & " PH_PanelHogar.CodigoHogar, "
	sql = sql & " PH_Panelistas.Nombre1, "
	sql = sql & " PH_Panelistas.Apellido1, "
	sql = sql & " PH_Panelistas.Celular, "
	sql = sql & " ss_Estado.Estado, "
	sql = sql & " PH_Ciudad.Ciudad "
	sql = sql & " FROM (((PH_PanelHogar INNER JOIN PH_Panelistas ON PH_PanelHogar.Id_PanelHogar = PH_Panelistas.Id_Hogar) INNER JOIN PH_EncuestaHogar ON PH_PanelHogar.Id_PanelHogar = PH_EncuestaHogar.Id_Hogar) INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) INNER JOIN PH_Ciudad ON PH_PanelHogar.Id_Ciudad = PH_Ciudad.Id_Ciudad "
	sql = sql & " WHERE "
	sql = sql & " PH_PanelHogar.Ind_Activo = 1 "
	sql = sql & " AND PH_Panelistas.ResponsablePanel = 1 "
	sql = sql & " AND PH_EncuestaHogar.Ind_Realizada = 1 "
	sql = sql & " AND PH_PanelHogar.Id_PanelHogar > 1 "
	sql = sql & " GROUP BY "
	sql = sql & " PH_PanelHogar.Id_PanelHogar, "
	sql = sql & " PH_PanelHogar.CodigoHogar, "
	sql = sql & " PH_Panelistas.Nombre1, "
	sql = sql & " PH_Panelistas.Apellido1, "
	sql = sql & " PH_Panelistas.Celular, "
	sql = sql & " ss_Estado.Estado, "
	sql = sql & " PH_Ciudad.Ciudad "
	sql = sql & " HAVING "
	sql = sql & " Count(PH_EncuestaHogar.Id_EncuestaHogar) = 5 "
	'response.write "<br>220 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gDatosSol = rsx1.GetRows
		rsx1.close
	end if
	'response.write "<br>Buscar Participante"
	TotalHogares = ubound(gDatosSol,2) + 1

	'Grabar Participantes
	sql = ""
	sql = sql & " SELECT * "
	sql = sql & " From  "
	sql = sql & " PH_SorteoParticipantes "
	'response.write "<br>220 sql:=" & sql
	'response.end
    rsx3.Open sql ,conexion

	for iReg = 0 to ubound(gDatosSol,2)
		rsx3.addNew
		rsx3("Id_Hogar") = gDatosSol(0,iReg)
		rsx3("CodigoHogar") = gDatosSol(1,iReg)
		rsx3("Nombre") = gDatosSol(2,iReg)
		rsx3("Apellido") = gDatosSol(3,iReg)
		rsx3("Celular") = gDatosSol(4,iReg)
		rsx3("Estado") = gDatosSol(5,iReg)
		rsx3("Ciudad") = gDatosSol(6,iReg)
		rsx3("Ind_Ganador") = 0
		rsx3("OrdenSorteo") = 0
		rsx3.Update
	next
	rsx3.Close 
	set rsx3 = nothing 
	'response.write "<br>Grabar Participante"

	'Buscar Maximo y Minimo id
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Max(Id_Sorteo) AS MaxId, "
	sql = sql & " Min(Id_Sorteo) AS MinId "
	sql = sql & " FROM "
	sql = sql & " PH_SorteoParticipantes "
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
	'response.write "<br>Buscar Maximo;= " & ValorMaximo & " y Minimo:= " & ValorMinimo


		%>
		<center>
		<h1>Sorteo Panel de Hogares 23 Diciembre 2020</h1>
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
					iGanador = 1
					Do While iGanador < 11
						Randomize
						'NumeroRandom = Int((ValorMaximo * Rnd) + ValorMinimo)
						NumeroRandom = int((ValorMaximo-ValorMinimo+1)*rnd+ValorMinimo)
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
						sql = sql & " Id_Sorteo = " & NumeroRandom
						'response.write "<br>220 sql:=" & sql
						'response.end
						rsx1.Open sql ,conexion
						if rsx1.eof then
							rsx1.close
						else
							gDatosSolSal = rsx1.GetRows
							rsx1.close
						end if
						if gDatosSolSal(7,0) = 0 then
							Response.Write "<tr>"
								Response.Write "<td>" & iGanador & "</td>"
								'Response.Write "<td>" & NumeroRandom & "</td>"
								Response.Write "<td>" & gDatosSolSal(0,0) & "</td>"
								Response.Write "<td>" & gDatosSolSal(1,0) & "</td>"
								Response.Write "<td>" & gDatosSolSal(2,0) & "</td>"
								Response.Write "<td>" & gDatosSolSal(3,0) & "</td>"
								Response.Write "<td>" & gDatosSolSal(4,0) & "</td>"
								Response.Write "<td>" & gDatosSolSal(5,0) & "</td>"
								Response.Write "<td>" & gDatosSolSal(6,0) & "</td>"
							Response.Write "</tr>"
							dim rsx4
							set rsx4 = CreateObject("ADODB.Recordset")
							rsx4.CursorType = 0
							rsx4.LockType = 3
							sql = ""
							sql = sql & " SELECT * "
							sql = sql & " From  "
							sql = sql & " PH_SorteoParticipantes "
							sql = sql & " WHERE "
							sql = sql & " Id_Sorteo = " & NumeroRandom
							rsx4.Open sql ,conexion
							'response.write "<br>220 sql:=" & sql
							'response.end
							rsx4("Ind_Ganador") = 1
							rsx4("OrdenSorteo") = iGanador
							rsx4.Update
							rsx4.Close 
							set rsx4 = nothing 
							iGanador = iGanador + 1
						end if
					loop 
					%>
				</table>
			</div>
		</div>
		<%
		'Response.Write "<br><br>Numero " & iGanador & " Numero: "
		'Response.Write NumeroRandom


	

	'response.write "<br>LLEGOFINAL" 
	response.end
	
	
%>
</body>
</html>
