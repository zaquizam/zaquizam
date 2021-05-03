<!DOCTYPE HTML>
<html >
<head>
	<title>Penetracion x Categroria</title>
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
	<link href="sweetalert.css" rel="stylesheet" type="text/css" media="screen" />
	<link rel="icon" href="favicon.ico" type="image/x-icon"> 
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
</head>
<body topmargin="0">
<!--#include file="estiloscss.asp"-->
<!--#include file="encabezado.asp"-->
<!--#include file="nn_subN.asp"-->
<!--#include file="in_DataEN.asp"-->

<%

  
'==========================================================================================
' Variables y Constantes
'==========================================================================================


    Apertura

	Dim gCategorias
	Dim gMeses
	Dim gSemanas
	Dim iMesDesde
	Dim iMesHasta
	Dim iAnoDesde
	Dim iAnoHasta
	
	Fecha = date
	
	iMesDesde = 1
	iMesHasta = month(Fecha)
	iAnoDesde = 2021
	iAnoHasta = year(Fecha)
	
	'response.write "<br>45 iMesDesde:= " & iMesDesde
	'response.write "<br>46 iMesHasta:= " & iMesHasta
	'response.write "<br>47 iAnoDesde:= " & iAnoDesde
	'response.write "<br>48 iAnoHasta:= " & iAnoHasta
	
	dim gDatosSol1
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	dim gDatosSol2
	dim rsx2
	set rsx2 = CreateObject("ADODB.Recordset")
	rsx2.CursorType = adOpenKeyset 
	rsx2.LockType = 2 'adLockOptimistic 

Sub Combos
 
	'response.write "<br>372 Combo1:=" & ed_sPar(1,0)
	'response.write " Combo2:=" & ed_sPar(2,0)
	'response.write " Combo3:=" & ed_sPar(3,0)
	'response.write " Combo3:=" & ed_sPar(4,0)
	'response.write " Combo3:=" & ed_sPar(5,0)
    ed_iCombo = 1

	sql = ""
	sql = sql & " SELECT  "
	sql = sql & " Id_Categoria, "
	sql = sql & " Categoria "
	sql = sql & " FROM PH_Categoria "
	sql = sql & " Where ind_Activo = 1 "
	sql = sql & " Order By "
	sql = sql & " Categoria "
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(1,0)="Categoria"
    ed_sCombo(1,1)=sql 
    ed_sCombo(1,2)="Seleccionar"
	
End Sub

Sub Categoria
	sql = ""
	sql = sql & " SELECT  "
	sql = sql & " Id_Categoria, "
	sql = sql & " Categoria "
	sql = sql & " FROM PH_Categoria "
	sql = sql & " Where ind_Activo = 1 "
	sql = sql & " Order By "
	sql = sql & " Categoria "
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else 
		gCategorias = rsx1.GetRows
		rsx1.close
	end if

End Sub

Sub Semanas
	sql = ""
	sql = sql & " SELECT  "
	sql = sql & " IdSemana, "
	sql = sql & " SemanaCorta "
	sql = sql & " FROM ss_Semana "
	sql = sql & " WHERE "
	sql = sql & " IdMes >= " & iMesDesde
	sql = sql & " and IdMes <= " & iMesHasta
	sql = sql & " AND IdAno = " & iAnoDesde
	sql = sql & " ORDER BY "
	sql = sql & " IdSemana "
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else 
		gSemanas = rsx1.GetRows
		rsx1.close
	end if

End Sub

Sub Meses
	sql = ""
	sql = sql & " SELECT  "
	sql = sql & " ss_Meses.Mes, "
	sql = sql & " ss_Periodo.IdAno, "
	sql = sql & " ss_Periodo.Semanas "
	sql = sql & " FROM ss_Periodo INNER JOIN ss_Meses ON ss_Periodo.IdMes = ss_Meses.Id_Mes "
	sql = sql & " WHERE "
	sql = sql & " ss_Periodo.IdMes >= " & iMesDesde
	sql = sql & " and ss_Periodo.IdMes <= " & iMesHasta
	sql = sql & " AND ss_Periodo.IdAno = " & iAnoDesde
	sql = sql & " ORDER BY "
	sql = sql & " ss_Periodo.Id "
	'response.write "<br>140 sql:= " & sql
	'response.end
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else 
		gMeses = rsx1.GetRows
		rsx1.close
	end if
	'response.write "<br>147 LLEGO"
	'response.end
end sub	
%>
	<script>

		//**Inicio Generar PDF
		function GenerarExcel(){
			//alert("Bus:= "+ document.getElementById("Bus").value );
			//alert("Buscar:= "+ document.getElementById("Excel").value );
			var sBus = document.getElementById("Excel").value
			window.open('Sys_mUsuarioExcel.asp?bus='+sBus,'_blank');
		}	
		//**Fin Generar PDF

		</script>   
<%

   
'==========================================================================================
' Parámetros del Manteniemiento
'==========================================================================================
    LeePar
  
    if ed_iPas<>4 then 
        Encabezado
    end if    
	sExcel = request.Form("bus")

	'response.write "llego1"
	'response.end
	if ed_sPar(1,0) = "" or ed_sPar(1,0) = "Seleccionar" then ed_sPar(1,0) = 0
    Combos
	Categoria
	Semanas
	Meses
%>
		
	<br>
	<div style="width:98%">
	<%
	
	%></div></center>
	<center>
		<h1>Version 1 </h1
	</center>
	<table border="0" align="right">
		<tr>
			<td>
				<%
				ed_vCombo
				%>
			</td>
		</tr>
	</table>
	</br>
	</br>
	</br>
	</br>
	</br>
	<%
	if ed_sPar(1,0) = "Seleccionar" then
		idCategoria = 0 
	else
		idCategoria = ed_sPar(1,0)
	end if
	'response.write "<br> Combo1:=" & ed_sPar(1,0) & "==>" & idArea
	'response.write "<br> Combo2:=" & ed_sPar(2,0) & "==>" & idEstado
	'response.write "<br> Combo3:=" & ed_sPar(3,0) & "==>" & idSemana
	'response.write "<br> Combo4:=" & ed_sPar(4,0) & "==>" & idCategoria
	'hidden 
	%>
	<input type="hidden" name="Programa" id="Programa" align="right" size=50>
	<br>
	<br>
	<br>
	<div id="DivBuscarInformación">
		<div class="ex1">
			<table class="w3-table w3-striped w3-bordered w3-card-4 w3-small" style="width:1000px; margin-left:auto; margin-right:auto;margin-top:10px ">
				<thead>
					<tr class="w3-blue">
						<th>Semana</th>
						<th>Mes</th>
						<th>Ano</th>
						<th>Total Hogares</th>
						<th>Total Hogares con Categoria</th>
						<th>Penetracion</th>
					</tr>
				</thead>
				<%
				if idCategoria <> 0 then
					sql = ""
					sql = sql & " SELECT "
					sql = sql & " ss_Semana.IdSemana, "
					sql = sql & " ss_Semana.Semana, "
					sql = sql & " ss_Meses.Mes, "
					sql = sql & " ss_Semana.IdAno "
					sql = sql & " FROM ss_Semana INNER JOIN ss_Meses ON ss_Semana.IdMes = ss_Meses.Id_Mes "
					sql = sql & " WHERE "
					sql = sql & " ss_Semana.IdSemana > 14 "
					'response.write "<br>232 sql:= " & sql
					'response.end
					rsx2.Open sql ,conexion
					if rsx2.eof then
						rsx2.close
					else 
						gDatosSol2 = rsx2.GetRows
						rsx2.close
						for iReg1 = 0 to ubound(gDatosSol2,2)
							iSemana = gDatosSol2(0,iReg1)
							sSemana = gDatosSol2(1,iReg1)
							sMes = gDatosSol2(2,iReg1)
							sAno = gDatosSol2(3,iReg1)
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " PH_Consumo.Id_Hogar "
							sql = sql & " FROM ((((((PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) INNER JOIN ss_Semana ON PH_Consumo.Id_Semana = ss_Semana.IdSemana) INNER JOIN PH_Categoria ON PH_Consumo_Detalle_Productos.Id_Categoria = PH_Categoria.Id_Categoria) INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN PH_GAreaEstado ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) INNER JOIN PH_Canal ON PH_Consumo.Id_Canal = PH_Canal.Id_Canal "
							sql = sql & " WHERE "
							sql = sql & " PH_Consumo.Id_Semana = " & iSemana
							sql = sql & " AND PH_PanelHogar.Ind_Activo = 1 "
							sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
							sql = sql & " GROUP BY "
							sql = sql & " PH_Consumo.Id_Hogar "
							sql = sql & " HAVING "
							sql = sql & " PH_Consumo.Id_Hogar > 1 "
							rsx1.Open sql ,conexion
							if rsx1.eof then
								rsx1.close
							else 
								gDatosSol1 = rsx1.GetRows
								rsx1.close
								TotalHog = ubound(gDatosSol1,2) + 1 
							end if

							sql = ""
							sql = sql & " SELECT "
							sql = sql & " PH_Consumo.Id_Hogar "
							sql = sql & " FROM ((((((PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) INNER JOIN ss_Semana ON PH_Consumo.Id_Semana = ss_Semana.IdSemana) INNER JOIN PH_Categoria ON PH_Consumo_Detalle_Productos.Id_Categoria = PH_Categoria.Id_Categoria) INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN PH_GAreaEstado ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) INNER JOIN PH_Canal ON PH_Consumo.Id_Canal = PH_Canal.Id_Canal "
							sql = sql & " WHERE "
							sql = sql & " PH_Consumo.Id_Semana = " & iSemana
							sql = sql & " AND PH_PanelHogar.Ind_Activo = 1 "
							sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
							sql = sql & " AND PH_Categoria.Id_Categoria = " & idCategoria
							sql = sql & " GROUP BY "
							sql = sql & " PH_Consumo.Id_Hogar "
							sql = sql & " HAVING "
							sql = sql & " PH_Consumo.Id_Hogar > 1 "
							'response.write "<br>232 sql:= " & sql
							rsx1.Open sql ,conexion
							if rsx1.eof then
								rsx1.close
							else 
								gDatosSol1 = rsx1.GetRows
								rsx1.close
								TotalHogCat = ubound(gDatosSol1,2) + 1 
							end if
							Penetracion = (TotalHogCat * 100) / TotalHog
							response.write "<tr>"
								response.write "<td>" & sSemana & "</td>"
								response.write "<td>" & sMes & "</td>"
								response.write "<td>" & sAno & "</td>"
								response.write "<td>" & TotalHog & "</td>"
								response.write "<td>" & TotalHogCat & "</td>"
								response.write "<td>" & formatnumber(Penetracion,2) & "</td>"
							response.write "</tr>"
						next
						
					end if

				end if
				%>
			</table>
		</div>
	</div>
	<center>
		<h1>Version 2 </h1
	</center>	
	<div id="Reporte">
		<div class="ex1">
			<table class="w3-table w3-striped w3-bordered w3-card-4 w3-small" style="width:1200px; margin-left:auto; margin-right:auto;margin-top:10px ">
				<thead>
					<tr class="w3-blue">
						<th>Categoria</th>
						<th>Indicador</th>
						<%
						for iSem = 0 to 3 'ubound(gSemanas,2)
							response.write "<th>" & gSemanas(1,iSem) & "</th>"
						next 
						for iMes = 0 to 0 'ubound(gMeses,2)
							response.write "<th>" & gMeses(0,iMes) & " " & gMeses(1,iMes) & "</th>"
						next 
						for iSem = 4 to 7 'ubound(gSemanas,2)
							response.write "<th>" & gSemanas(1,iSem) & "</th>"
						next 
						for iMes = 1 to 1 'ubound(gMeses,2)
							response.write "<th>" & gMeses(0,iMes) & " " & gMeses(1,iMes) & "</th>"
						next 
						%>
					</tr>
				</thead>
				<%
				for iCat = 0 to ubound(gCategorias,2)
					response.write "<tr>" 
						iCategoria = gCategorias(0,iCat)
						response.write "<td>" & gCategorias(1,iCat) & "</td>"
						response.write "<td>Penetracion</td>"
						'Enero 2021 - Semana
						for iSem = 0 to 3 'ubound(gSemanas,2)
							iSemana = gSemanas(0,iSem)
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " PH_Consumo.Id_Hogar "
							sql = sql & " FROM ((((((PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) INNER JOIN ss_Semana ON PH_Consumo.Id_Semana = ss_Semana.IdSemana) INNER JOIN PH_Categoria ON PH_Consumo_Detalle_Productos.Id_Categoria = PH_Categoria.Id_Categoria) INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN PH_GAreaEstado ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) INNER JOIN PH_Canal ON PH_Consumo.Id_Canal = PH_Canal.Id_Canal "
							sql = sql & " WHERE "
							sql = sql & " PH_Consumo.Id_Semana = " & iSemana
							sql = sql & " AND PH_PanelHogar.Ind_Activo = 1 "
							sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
							sql = sql & " GROUP BY "
							sql = sql & " PH_Consumo.Id_Hogar "
							sql = sql & " HAVING "
							sql = sql & " PH_Consumo.Id_Hogar > 1 "
							rsx1.Open sql ,conexion
							if rsx1.eof then
								rsx1.close
							else 
								gDatosSol1 = rsx1.GetRows
								rsx1.close
								TotalHog = ubound(gDatosSol1,2) + 1 
							end if
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " PH_Consumo.Id_Hogar "
							sql = sql & " FROM ((((((PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) INNER JOIN ss_Semana ON PH_Consumo.Id_Semana = ss_Semana.IdSemana) INNER JOIN PH_Categoria ON PH_Consumo_Detalle_Productos.Id_Categoria = PH_Categoria.Id_Categoria) INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN PH_GAreaEstado ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) INNER JOIN PH_Canal ON PH_Consumo.Id_Canal = PH_Canal.Id_Canal "
							sql = sql & " WHERE "
							sql = sql & " PH_Consumo.Id_Semana = " & iSemana
							sql = sql & " AND PH_PanelHogar.Ind_Activo = 1 "
							sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
							sql = sql & " AND PH_Categoria.Id_Categoria = " & iCategoria
							sql = sql & " GROUP BY "
							sql = sql & " PH_Consumo.Id_Hogar "
							sql = sql & " HAVING "
							sql = sql & " PH_Consumo.Id_Hogar > 1 "
							'response.write "<br>232 sql:= " & sql
							rsx1.Open sql ,conexion
							if rsx1.eof then
								rsx1.close
							else 
								gDatosSol1 = rsx1.GetRows
								rsx1.close
								TotalHogCat = ubound(gDatosSol1,2) + 1 
							end if
							Penetracion = (TotalHogCat * 100) / TotalHog
							response.write "<td>" & formatnumber(Penetracion,2) & "</td>"
						next 
						'Enero 2021 - Mes
						for iMes = 0 to 0 'ubound(gMeses,2)
							iSemana = gMeses(2,iMes)
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " PH_Consumo.Id_Hogar "
							sql = sql & " FROM ((((((PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) INNER JOIN ss_Semana ON PH_Consumo.Id_Semana = ss_Semana.IdSemana) INNER JOIN PH_Categoria ON PH_Consumo_Detalle_Productos.Id_Categoria = PH_Categoria.Id_Categoria) INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN PH_GAreaEstado ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) INNER JOIN PH_Canal ON PH_Consumo.Id_Canal = PH_Canal.Id_Canal "
							sql = sql & " WHERE "
							sql = sql & " PH_Consumo.Id_Semana in ( " & iSemana & ")"
							sql = sql & " AND PH_PanelHogar.Ind_Activo = 1 "
							sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
							sql = sql & " GROUP BY "
							sql = sql & " PH_Consumo.Id_Hogar "
							sql = sql & " HAVING "
							sql = sql & " PH_Consumo.Id_Hogar > 1 "
							rsx1.Open sql ,conexion
							if rsx1.eof then
								rsx1.close
							else 
								gDatosSol1 = rsx1.GetRows
								rsx1.close
								TotalHog = ubound(gDatosSol1,2) + 1 
							end if
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " PH_Consumo.Id_Hogar "
							sql = sql & " FROM ((((((PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) INNER JOIN ss_Semana ON PH_Consumo.Id_Semana = ss_Semana.IdSemana) INNER JOIN PH_Categoria ON PH_Consumo_Detalle_Productos.Id_Categoria = PH_Categoria.Id_Categoria) INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN PH_GAreaEstado ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) INNER JOIN PH_Canal ON PH_Consumo.Id_Canal = PH_Canal.Id_Canal "
							sql = sql & " WHERE "
							sql = sql & " PH_Consumo.Id_Semana in ( " & iSemana & ")"
							sql = sql & " AND PH_PanelHogar.Ind_Activo = 1 "
							sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
							sql = sql & " AND PH_Categoria.Id_Categoria = " & iCategoria
							sql = sql & " GROUP BY "
							sql = sql & " PH_Consumo.Id_Hogar "
							sql = sql & " HAVING "
							sql = sql & " PH_Consumo.Id_Hogar > 1 "
							'response.write "<br>232 sql:= " & sql
							rsx1.Open sql ,conexion
							if rsx1.eof then
								rsx1.close
							else 
								gDatosSol1 = rsx1.GetRows
								rsx1.close
								TotalHogCat = ubound(gDatosSol1,2) + 1 
							end if
							Penetracion = (TotalHogCat * 100) / TotalHog
							response.write "<td>" & formatnumber(Penetracion,2) & "</td>"
						next 
						'Febrero 2021 - Semana
						for iSem = 4 to 7 'ubound(gSemanas,2)
							iSemana = gSemanas(0,iSem)
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " PH_Consumo.Id_Hogar "
							sql = sql & " FROM ((((((PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) INNER JOIN ss_Semana ON PH_Consumo.Id_Semana = ss_Semana.IdSemana) INNER JOIN PH_Categoria ON PH_Consumo_Detalle_Productos.Id_Categoria = PH_Categoria.Id_Categoria) INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN PH_GAreaEstado ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) INNER JOIN PH_Canal ON PH_Consumo.Id_Canal = PH_Canal.Id_Canal "
							sql = sql & " WHERE "
							sql = sql & " PH_Consumo.Id_Semana = " & iSemana
							sql = sql & " AND PH_PanelHogar.Ind_Activo = 1 "
							sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
							sql = sql & " GROUP BY "
							sql = sql & " PH_Consumo.Id_Hogar "
							sql = sql & " HAVING "
							sql = sql & " PH_Consumo.Id_Hogar > 1 "
							rsx1.Open sql ,conexion
							if rsx1.eof then
								rsx1.close
							else 
								gDatosSol1 = rsx1.GetRows
								rsx1.close
								TotalHog = ubound(gDatosSol1,2) + 1 
							end if
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " PH_Consumo.Id_Hogar "
							sql = sql & " FROM ((((((PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) INNER JOIN ss_Semana ON PH_Consumo.Id_Semana = ss_Semana.IdSemana) INNER JOIN PH_Categoria ON PH_Consumo_Detalle_Productos.Id_Categoria = PH_Categoria.Id_Categoria) INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN PH_GAreaEstado ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) INNER JOIN PH_Canal ON PH_Consumo.Id_Canal = PH_Canal.Id_Canal "
							sql = sql & " WHERE "
							sql = sql & " PH_Consumo.Id_Semana = " & iSemana
							sql = sql & " AND PH_PanelHogar.Ind_Activo = 1 "
							sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
							sql = sql & " AND PH_Categoria.Id_Categoria = " & iCategoria
							sql = sql & " GROUP BY "
							sql = sql & " PH_Consumo.Id_Hogar "
							sql = sql & " HAVING "
							sql = sql & " PH_Consumo.Id_Hogar > 1 "
							'response.write "<br>232 sql:= " & sql
							rsx1.Open sql ,conexion
							if rsx1.eof then
								rsx1.close
							else 
								gDatosSol1 = rsx1.GetRows
								rsx1.close
								TotalHogCat = ubound(gDatosSol1,2) + 1 
							end if
							Penetracion = (TotalHogCat * 100) / TotalHog
							response.write "<td>" & formatnumber(Penetracion,2) & "</td>"
						next 
						'Febrero 2021 - Mes
						for iMes = 1 to 1 'ubound(gMeses,2)
							iSemana = gMeses(2,iMes)
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " PH_Consumo.Id_Hogar "
							sql = sql & " FROM ((((((PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) INNER JOIN ss_Semana ON PH_Consumo.Id_Semana = ss_Semana.IdSemana) INNER JOIN PH_Categoria ON PH_Consumo_Detalle_Productos.Id_Categoria = PH_Categoria.Id_Categoria) INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN PH_GAreaEstado ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) INNER JOIN PH_Canal ON PH_Consumo.Id_Canal = PH_Canal.Id_Canal "
							sql = sql & " WHERE "
							sql = sql & " PH_Consumo.Id_Semana in ( " & iSemana & ")"
							sql = sql & " AND PH_PanelHogar.Ind_Activo = 1 "
							sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
							sql = sql & " GROUP BY "
							sql = sql & " PH_Consumo.Id_Hogar "
							sql = sql & " HAVING "
							sql = sql & " PH_Consumo.Id_Hogar > 1 "
							rsx1.Open sql ,conexion
							if rsx1.eof then
								rsx1.close
							else 
								gDatosSol1 = rsx1.GetRows
								rsx1.close
								TotalHog = ubound(gDatosSol1,2) + 1 
							end if
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " PH_Consumo.Id_Hogar "
							sql = sql & " FROM ((((((PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) INNER JOIN ss_Semana ON PH_Consumo.Id_Semana = ss_Semana.IdSemana) INNER JOIN PH_Categoria ON PH_Consumo_Detalle_Productos.Id_Categoria = PH_Categoria.Id_Categoria) INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN PH_GAreaEstado ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) INNER JOIN PH_Canal ON PH_Consumo.Id_Canal = PH_Canal.Id_Canal "
							sql = sql & " WHERE "
							sql = sql & " PH_Consumo.Id_Semana in ( " & iSemana & ")"
							sql = sql & " AND PH_PanelHogar.Ind_Activo = 1 "
							sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
							sql = sql & " AND PH_Categoria.Id_Categoria = " & iCategoria
							sql = sql & " GROUP BY "
							sql = sql & " PH_Consumo.Id_Hogar "
							sql = sql & " HAVING "
							sql = sql & " PH_Consumo.Id_Hogar > 1 "
							'response.write "<br>232 sql:= " & sql
							rsx1.Open sql ,conexion
							if rsx1.eof then
								rsx1.close
							else 
								gDatosSol1 = rsx1.GetRows
								rsx1.close
								TotalHogCat = ubound(gDatosSol1,2) + 1 
							end if
							Penetracion = (TotalHogCat * 100) / TotalHog
							response.write "<td>" & formatnumber(Penetracion,2) & "</td>"
						next 

					response.write "</tr>"
				next 
				%>
			</table>
		</div>
	</div>
	</br>
	</br>
	</br>
	</br>
	</br>
    <%conexion.close%>


</body>
</html>