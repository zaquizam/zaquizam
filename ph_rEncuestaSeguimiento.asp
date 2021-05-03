<!DOCTYPE HTML>
<html >
<head>
	<title>Encuesta Seguimiento</title>
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="mensaje.css" rel="stylesheet" type="text/css" media="screen" />
	<link rel="icon" href="favicon.ico" type="image/x-icon"> 
	<script type="text/javascript" src="js/sweetalert.min.js"></script>
	<script type="text/javascript" src="js/jquery-1.12.4.min.js"></script>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />

</head>
<body topmargin="0">
<!--#include file="estiloscss.asp"-->
<!--#include file="meta.asp"-->
<!--#include file="encabezado.asp"-->
<!--#include file="nn_subN.asp"-->
<!--#include file="in_DataEN.asp"-->

<%

  
'==========================================================================================
' Variables y Constantes
'==========================================================================================


    Apertura
	   
	dim envMensaje
	dim envCelular
%>
<script type="text/javascript">
	function GenerarExcel()
	{
		num = document.getElementById("Excel").value;
		//alert("Generar Excel:="+ num);
		window.open("ph_rEncuestasRealizadasExcel.asp?num=" +num,"_blank");
	}

	function GenerarExcel1()
	{
		num = document.getElementById("Excel").value;
		//alert("Generar Excel1:="+ num);
		//return;
		window.open("ph_rEncuestasTotalesExcel.asp?num=" +num,"_blank");
	}

	function GenerarExcelFaltantes()
	{
		num = document.getElementById("Excel").value;
		//alert("Generar Excel1:="+ num);
		//return;
		window.open("ph_rEncuestasHogaresFaltanteExcel.asp?num=" +num,"_blank");
	}
	
	function alerta(total) 
	{
		swal("Se Enviaron Encuestas " + total + " Hogares ","Enviado","success");
		//window.open("?edpas=1&smenu=?x=1&smenu=Envio%20SMS%20Bienvenida%20y%20Link","_parent");
	}
</script>
<%

Sub Combos
 
	'response.write "<br>372 Combo1:=" & ed_sPar(1,0)
	'response.write " Combo2:=" & ed_sPar(2,0)
	'response.write " Combo3:=" & ed_sPar(3,0)
	'response.write " Combo3:=" & ed_sPar(4,0)
	'response.write " Combo3:=" & ed_sPar(5,0)
    ed_iCombo = 1
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_EncuestaEspecial, "
	sql = sql & " EncuestaEspecial "
	sql = sql & " FROM PH_EncuestaEspecial "
	sql = sql & " WHERE "
	sql = sql & " Ind_Activo = 1 "
	sql = sql & " Order By "
	sql = sql & " EncuestaEspecial  Desc"
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(1,0)="Encuesta Especial"
    ed_sCombo(1,1)=sql 
    ed_sCombo(1,2)="Seleccionar"



End Sub

   
'==========================================================================================
' Parámetros del Manteniemiento
'==========================================================================================
    LeePar
  
    
    if ed_iPas<>4 then 
        Encabezado
    end if    

	'response.write "llego1"
	'response.end
	'ParDat
%>
		
	<br>
	<br>
	<br>
	<div style="width:98%"></div></center>
<%
	Combos
	'response.write "paso"

%>

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

<%
'==========================================================================================
' Variables y Constantes
'==========================================================================================
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

	dim gDatosSol3
	dim rsx3
	set rsx3 = CreateObject("ADODB.Recordset")
	rsx3.CursorType = adOpenKeyset 
	rsx3.LockType = 2 'adLockOptimistic 

	dim gDatosSol0
	dim rsx0
	set rsx0 = CreateObject("ADODB.Recordset")
	rsx0.CursorType = adOpenKeyset 
	rsx0.LockType = 2 'adLockOptimistic 
	'Buscar Area
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " PH_GAreaEstado.Id_Estado, "
	sql = sql & " PH_GArea.Area "
	sql = sql & " FROM PH_GAreaEstado INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area "
	sql = sql & " ORDER "
	sql = sql & " BY PH_GAreaEstado.Id_Estado "
	'response.write "<br>36 sql:=" & sql
	'response.end
	rsx0.Open sql ,conexion
	if rsx0.eof then
		rsx0.close
	else
		gDatosSol0 = rsx0.GetRows
		rsx0.close
	end if
	dim gArea(50)
	for iReg = 0 to ubound(gDatosSol0,2)
		iEstado = gDatosSol0(0,iReg)
		sArea = gDatosSol0(1,iReg)
		gArea(iEstado) = sArea
	next 
	if ed_sPar(1,0) <> "Seleccionar" then
		Encuesta = cint(ed_sPar(1,0))
		sExcel = Encuesta
		%>
			<div class="container-fluid">        
				<div class="row">
					<!--Contenido Generalhidden-->			
					<div class="container">
						<div class="col-md-8 col-sm-8 col-xs-12">
							<div class="pull-right">
								<img src="images/Excel.png"  style="margin-left:0px;" title="Generar Excel" alt="PDF" width="70px" onclick="GenerarExcel1()"/>
								<input type="hidden" name="Excel" id="Excel" align="right" size=0 value='<%=sExcel%>'>
							</div>
						</div>
					</div>
				</div>
			</div>
			<br>
		<div id="DivBuscarPanelistas">
			<div class="ex1">
				<table class="w3-table w3-striped w3-bordered w3-card-4 w3-small" style="width:1000px; margin-left:auto; margin-right:auto;margin-top:10px ">
					<thead>
						<tr class="w3-blue">
							<th>Estado</th>
							<th>Area</th>
							<th>Respondidas</th>
							<th>Rechazadas</th>
							<th>Pendientes</th>
							<th>% de Cumplimiento</th>
						</tr>
					</thead>
					<% 
					Response.write "<tr>"
						Response.write "<td>Total Venezuela</td>"
						Response.write "<td>Total Venezuela</td>"
						'Realizadas
						sql = ""
						sql = sql & " SELECT "
						sql = sql & " Count(PH_EncuestaHogar.Id_Hogar) AS CuentaDeId_Hogar "
						sql = sql & " FROM PH_EncuestaHogar INNER JOIN PH_PanelHogar ON PH_EncuestaHogar.Id_Hogar = PH_PanelHogar.Id_PanelHogar "
						sql = sql & " WHERE PH_EncuestaHogar.Ind_Realizada = 1 AND PH_PanelHogar.Ind_Activo = 1 "
						sql = sql & " GROUP BY PH_EncuestaHogar.Id_EncuestaEspecial "
						sql = sql & " HAVING PH_EncuestaHogar.Id_EncuestaEspecial = " & Encuesta 
						'response.write "<br>36 sql:=" & sql
						'response.end
						rsx1.Open sql ,conexion
						if rsx1.eof then
							Response.write "<td>0</td>"
							rsx1.close
						else
							gDatosSol1 = rsx1.GetRows
							rsx1.close
							Response.write "<td>" & gDatosSol1(0,0) & "</td>"
							Realizadas = cint(gDatosSol1(0,0))
						end if
						'Rechazadas
						sql = ""
						sql = sql & " SELECT "
						sql = sql & " Count(PH_EncuestaHogar.Id_Hogar) AS CuentaDeId_Hogar "
						sql = sql & " FROM PH_EncuestaHogar INNER JOIN PH_PanelHogar ON PH_EncuestaHogar.Id_Hogar = PH_PanelHogar.Id_PanelHogar "
						sql = sql & " WHERE PH_EncuestaHogar.Ind_Rechazada = 1 AND PH_PanelHogar.Ind_Activo = 1 "
						sql = sql & " GROUP BY PH_EncuestaHogar.Id_EncuestaEspecial "
						sql = sql & " HAVING PH_EncuestaHogar.Id_EncuestaEspecial = " & Encuesta 
						'response.write "<br>36 sql:=" & sql
						'response.end
						rsx1.Open sql ,conexion
						if rsx1.eof then
							Response.write "<td>0</td>"
							rsx1.close
						else
							gDatosSol1 = rsx1.GetRows
							rsx1.close
							Response.write "<td>" & gDatosSol1(0,0) & "</td>"
							Rechazadas = cint(gDatosSol1(0,0))
						end if
						'Pendientes
						sql = ""
						sql = sql & " SELECT "
						sql = sql & " Count(PH_EncuestaHogar.Id_Hogar) AS CuentaDeId_Hogar "
						sql = sql & " FROM PH_EncuestaHogar INNER JOIN PH_PanelHogar ON PH_EncuestaHogar.Id_Hogar = PH_PanelHogar.Id_PanelHogar "
						sql = sql & " WHERE PH_EncuestaHogar.Id_EncuestaEspecial = " & Encuesta
						sql = sql & " AND PH_EncuestaHogar.Ind_Rechazada =0  AND PH_EncuestaHogar.Ind_Realizada = 0  AND PH_PanelHogar.Ind_Activo = 1 "
						'response.write "<br>36 sql:=" & sql
						'response.end
						rsx1.Open sql ,conexion
						if rsx1.eof then
							Response.write "<td>0</td>"
							rsx1.close
						else
							gDatosSol1 = rsx1.GetRows
							rsx1.close
							Response.write "<td>" & gDatosSol1(0,0) & "</td>"
							Pendientes = cint(gDatosSol1(0,0))
						end if
						Total = Pendientes + Realizadas + Rechazadas
						'Cumplimiento = ((Realizadas) / Pendientes) * 100 
						Cumplimiento = (Realizadas * 100) / Total
						Response.write "<td>" & formatnumber(Cumplimiento) & "</td>"
					Response.write "</tr>"
					
					Response.write "<tr>"
						'response.write "<br>277 Paso<br>"
						'Pendientes
						sql = ""
						sql = sql & " SELECT "
						sql = sql & " ss_Estado.Estado, "
						sql = sql & " Count(PH_EncuestaHogar.Id_EncuestaHogar) AS CuentaDeId_EncuestaHogar, "
						sql = sql & " ss_Estado.Id_Estado "
						sql = sql & " FROM (PH_EncuestaHogar INNER JOIN PH_PanelHogar ON PH_EncuestaHogar.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado "
						sql = sql & " WHERE "
						sql = sql & " PH_EncuestaHogar.Id_EncuestaEspecial = " & Encuesta  
						sql = sql & " AND PH_PanelHogar.Ind_Activo = 1 "
						sql = sql & " AND PH_EncuestaHogar.Ind_Rechazada = 0 "
						sql = sql & " AND PH_EncuestaHogar.Ind_Realizada = 0 "
						sql = sql & " GROUP BY "
						sql = sql & " ss_Estado.Estado, "
						sql = sql & " ss_Estado.Id_Estado "
						sql = sql & " ORDER BY "
						sql = sql & " ss_Estado.Estado "
						'response.write "<br>36 sql:=" & sql
						'response.end
						rsx1.Open sql ,conexion
						if rsx1.eof then
							rsx1.close
							'response.write "<br>300 Paso<br>"
						else
							gDatosSol1 = rsx1.GetRows
							rsx1.close
							'response.write "<br>304 Paso<br>"
						end if
						for iReg = 0 to ubound(gDatosSol1,2)
							Response.write "<tr>" 
								Estado = gDatosSol1(0,iReg)
								idEstado = gDatosSol1(2,iReg)
								Response.write "<td>" & gDatosSol1(0,iReg) & "</td>"
								Response.write "<td>" & gArea(idEstado) & "</td>"
								'Respondidas
								sql = ""
								sql = sql & " SELECT "
								sql = sql & " Count(PH_EncuestaHogar.Id_EncuestaHogar) AS CuentaDeId_EncuestaHogar "
								sql = sql & " FROM (PH_EncuestaHogar INNER JOIN PH_PanelHogar ON PH_EncuestaHogar.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado "
								sql = sql & " WHERE "
								sql = sql & " PH_EncuestaHogar.Id_EncuestaEspecial = " & Encuesta
								sql = sql & " AND PH_PanelHogar.Ind_Activo = 1 "
								sql = sql & " AND ss_Estado.Estado = '" & Estado & "'" 
								sql = sql & " AND PH_EncuestaHogar.Ind_Realizada = 1 "
								'response.write "<br>36 sql:=" & sql
								'response.end
								rsx2.Open sql ,conexion
								if rsx2.eof then
									rsx2.close
									Response.write "<td></td>"
									Realizadas = 0
								else
									gDatosSol2 = rsx2.GetRows
									rsx2.close
									Response.write "<td>" & gDatosSol2(0,0) & "</td>"
									Realizadas = cint(gDatosSol2(0,0))
								end if
								'Rechazadas
								sql = ""
								sql = sql & " SELECT "
								sql = sql & " Count(PH_EncuestaHogar.Id_EncuestaHogar) AS CuentaDeId_EncuestaHogar "
								sql = sql & " FROM (PH_EncuestaHogar INNER JOIN PH_PanelHogar ON PH_EncuestaHogar.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado "
								sql = sql & " WHERE "
								sql = sql & " PH_EncuestaHogar.Id_EncuestaEspecial = " & Encuesta
								sql = sql & " AND PH_PanelHogar.Ind_Activo = 1 "
								sql = sql & " AND ss_Estado.Estado = '" & Estado & "'" 
								sql = sql & " AND PH_EncuestaHogar.Ind_Rechazada = 1 "
								'response.write "<br>36 sql:=" & sql
								'response.end
								rsx3.Open sql ,conexion
								if rsx3.eof then
									rsx3.close
									Response.write "<td></td>"
									Rechazadas = 0
								else
									gDatosSol3 = rsx3.GetRows
									rsx3.close
									Response.write "<td>" & gDatosSol3(0,0) & "</td>"
									Rechazadas = gDatosSol3(0,0)
								end if
								'Pendientes
								Response.write "<td>" & gDatosSol1(1,iReg) & "</td>"
								Pendientes = cint(gDatosSol1(1,iReg))
								Total = Pendientes + Realizadas + Rechazadas
								Cumplimiento = (Realizadas * 100) / Total
								'Cumplimiento = (Realizadas / Pendientes) * 100
								Response.write "<td>" & formatnumber(Cumplimiento) & "</td>"
							Response.write "</tr>"
						next
					
					
					Response.write "</tr>"
					%>
				</table>
			</div>
		</div>
		<%
		
		'********Hogares Rechazadas
		sql = ""
		sql = sql & " SELECT "
		sql = sql & " PH_EncuestaHogar.Id_Hogar, "
		sql = sql & " PH_PanelHogar.CodigoHogar, "
		sql = sql & " ss_Estado.Estado "
		sql = sql & " FROM (PH_EncuestaHogar INNER JOIN PH_PanelHogar ON PH_EncuestaHogar.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado "
		sql = sql & " WHERE "
		sql = sql & " PH_EncuestaHogar.Id_EncuestaEspecial = " & Encuesta 
		sql = sql & " AND PH_EncuestaHogar.Ind_Rechazada = 1 "
		'response.write "<br>36 sql:=" & sql
		'response.end
		rsx1.Open sql ,conexion
		if rsx1.eof then
			rsx1.close
			%>
			<div id="DivBuscarPanelistas">
				<h3>Rechazada</h3>
				<div class="ex1">
					<table class="w3-table w3-striped w3-bordered w3-card-4 w3-small" style="width:1000px; margin-left:auto; margin-right:auto;margin-top:10px ">
						<thead>
							<tr class="w3-blue">
								<th>Id Hogar</th>
								<th>Hogar</th>
								<th>Estado</th>
							</tr>
						</thead>
					</table>
				</div>
			</div>
			<%
		else
			gDatosSol1 = rsx1.GetRows
			rsx1.close
			'response.write "<br>ubound(gDatosSol1,2):= " & ubound(gDatosSol1,2)
			%>
			<div id="DivBuscarPanelistas">
				<h3>Rechazada</h3>
				<div class="ex1">
					<table class="w3-table w3-striped w3-bordered w3-card-4 w3-small" style="width:1000px; margin-left:auto; margin-right:auto;margin-top:10px ">
						<thead>
							<tr class="w3-blue">
								<th>Id Hogar</th>
								<th>Hogar</th>
								<th>Estado</th>
							</tr>
						</thead>
						<% 
						for iReg = 0 to ubound(gDatosSol1,2)
							Response.write "<tr>"
								for ib = 0 to 2
									Response.write "<td>" & gDatosSol1(ib,iReg) & "</td>"
								next
							Response.write "</tr>"
						next
						%>
					</table>
				</div>
			</div>
			<%
		end if

		'********Hogares Faltantes
		sql = ""
		sql = sql & " SELECT "
		sql = sql & " PH_GArea.Area, ss_Estado.Estado, "
		sql = sql & " PH_EncuestaHogar.Id_Hogar, "
		sql = sql & " PH_PanelHogar.CodigoHogar, "
		sql = sql & " PH_Panelistas.Nombre1, "
		sql = sql & " PH_Panelistas.Apellido1, "
		sql = sql & " PH_Panelistas.Celular "
		sql = sql & " FROM (((PH_EncuestaHogar INNER JOIN PH_PanelHogar ON PH_EncuestaHogar.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN PH_Panelistas ON PH_PanelHogar.Id_PanelHogar = PH_Panelistas.Id_Hogar) INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) INNER JOIN (PH_GArea INNER JOIN PH_GAreaEstado ON PH_GArea.Id_Area = PH_GAreaEstado.Id_Area) ON (ss_Estado.Id_Estado = PH_GAreaEstado.Id_Estado) AND (PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) "
		sql = sql & " WHERE "
		sql = sql & " PH_EncuestaHogar.Id_EncuestaEspecial = " & Encuesta
		sql = sql & " AND PH_EncuestaHogar.Ind_Realizada = 0 "
		sql = sql & " AND PH_Panelistas.ResponsablePanel = 1 " 
		sql = sql & " AND PH_PanelHogar.Ind_Activo = 1 "
		sql = sql & " ORDER BY "
		sql = sql & " PH_GArea.Area, "
		sql = sql & " ss_Estado.Estado, "
		sql = sql & " PH_EncuestaHogar.Id_Hogar "
		'response.write "<br>36 sql:=" & sql
		'response.end
		rsx1.Open sql ,conexion
		if rsx1.eof then
			rsx1.close
		else
			gDatosSol1 = rsx1.GetRows
			rsx1.close
			'response.write "<br>ubound(gDatosSol1,2):= " & ubound(gDatosSol1,2)
			sExcelFaltantes = Encuesta
			%>
			<div id="DivBuscarPanelistas">
				<h3>Hogares Faltantes (<%=ubound(gDatosSol1,2)+1 %>)</h3>
				<br>
				<div style="width:98%">
				<div class="container-fluid">        
					<div class="row">
						<!--Contenido Generalhidden-->			
						<div class="container">
							<div class="col-md-8 col-sm-8 col-xs-12">
								<div class="pull-right">
									<img src="images/Excel.png"  style="margin-left:0px;" title="Generar Excel" alt="PDF" width="70px" onclick="GenerarExcelFaltantes()"/>
									<input type="hidden" name="ExcelFaltantes" id="ExcelFaltantes" align="right" size=0 value='<%=sExcelFaltantes%>'>
								</div>
							</div>
						</div>
					</div>
				</div>
				<br>
				<div class="ex1">
					<table class="w3-table w3-striped w3-bordered w3-card-4 w3-small" style="width:1000px; margin-left:auto; margin-right:auto;margin-top:10px ">
						<thead>
							<tr class="w3-blue">
								<th>Area</th>
								<th>Estado</th>
								<th>Id Hogar</th>
								<th>Hogar</th>
								<th>Nombre</th>
								<th>Apellido</th>
								<th>Celular</th>
							</tr>
						</thead>
						<%
						for iReg = 0 to ubound(gDatosSol1,2)
							Response.write "<tr>"
								for ib = 0 to 6
									Response.write "<td>" & gDatosSol1(ib,iReg) & "</td>"
								next
							Response.write "</tr>"
						next
						
						%>
					</table>
				</div>
			</div>
			<%
		end if
	
		dim gDatosSol11
		'********Hogares Realizadas
		sql = ""
		sql = sql & " SELECT "
		sql = sql & " PH_EncuestaEspecialResultados.Id_Hogar, "
		sql = sql & " PH_PanelHogar.CodigoHogar, "
		sql = sql & " ss_Estado.Estado, "
		sql = sql & " PH_EncuestaEspecial.EncuestaEspecial, "
		sql = sql & " PH_EncuestaEspecialDet.Pregunta, "
		sql = sql & " PH_EncuestaEspecialResultados.Id_Respuesta, "
		sql = sql & " PH_EncuestaEspecialResultados.RespuestaTexto, "
		sql = sql & " ss_Estado.id_Estado, "
		sql = sql & " PH_PanelHogar.ClaseSocial "
		sql = sql & " FROM ((((PH_EncuestaEspecialResultados INNER JOIN PH_PanelHogar ON PH_EncuestaEspecialResultados.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN PH_EncuestaEspecial ON PH_EncuestaEspecialResultados.Id_EncuestaEspecial = PH_EncuestaEspecial.Id_EncuestaEspecial) INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) INNER JOIN PH_EncuestaEspecialDet ON (PH_EncuestaEspecialResultados.Id_Pregunta_Encuesta = PH_EncuestaEspecialDet.Id_EncuestaEspecialDet) AND (PH_EncuestaEspecialResultados.Id_EncuestaEspecial = PH_EncuestaEspecialDet.Id_EncuestaEspecial)) INNER JOIN PH_EncuestaHogar ON (PH_EncuestaHogar.Id_Hogar = PH_EncuestaEspecialResultados.Id_Hogar) AND (PH_EncuestaEspecial.Id_EncuestaEspecial = PH_EncuestaHogar.Id_EncuestaEspecial) "
		sql = sql & " WHERE "
		sql = sql & " PH_EncuestaEspecial.Id_EncuestaEspecial = " & Encuesta
		sql = sql & " AND PH_EncuestaHogar.Ind_Realizada = 1 "
		'sql = sql & " AND PH_EncuestaEspecialResultados.Id_Hogar =35 "
		sql = sql & " ORDER BY "
		sql = sql & " PH_EncuestaEspecialResultados.Id_Hogar, "
		sql = sql & " PH_EncuestaEspecialDet.Orden "
		'response.write "<br>36 sql:=" & sql
		'response.end
		rsx1.Open sql ,conexion
		if rsx1.eof then
			ExisteData = 0
		else
			gDatosSol11 = rsx1.GetRows
			rsx1.close
			ExisteData = 1
		end if
		
		sExcel = Encuesta
		Response.flush
		Response.write "<br><br><br>"
		'Response.end
		%>
		<div id="DivBuscarPanelistas">
			<center>
			<h2>Data</h2>
			</center>
			<br>
			<div style="width:98%">
			<div class="container-fluid">        
				<div class="row">
					<!--Contenido Generalhidden-->			
					<div class="container">
						<div class="col-md-8 col-sm-8 col-xs-12">
							<div class="pull-right">
								<img src="images/Excel.png"  style="margin-left:0px;" title="Generar Excel" alt="PDF" width="70px" onclick="GenerarExcel()"/>
								<input type="hidden" name="Excel" id="Excel" align="right" size=0 value='<%=sExcel%>'>
							</div>
						</div>
					</div>
				</div>
			</div>
			<br>
			<div class="ex1">
				<table class="w3-table w3-striped w3-bordered w3-card-4 w3-small" style="width:1000px; margin-left:auto; margin-right:auto;margin-top:10px ">
					<thead>
						<tr class="w3-blue">
							<th>Id Hogar</th>
							<th>Hogar</th>
							<th>Estado</th>
							<th>Area</th>
							<th>ClaseSocial</th>
							<th>Encuesta</th>
							<th>Pregunta</th>
							<th>IdRespuesta</th>
							<th>Respuesta</th>
						</tr>
					</thead>
					<% 
					'reponse.write str(ubound(gDatosSol1,2)-1)
					'reponse.end
					if ExisteData = 1 then
						for iReg = 0 to ubound(gDatosSol11,2)
							Response.write "<tr>"
								'for ib = 0 to 5
								'	Response.write "<td>" & gDatosSol1(ib,iReg) & "</td>"
								'next
								Response.write "<td>" & cint(gDatosSol11(0,iReg)) & "</td>"
								Response.write "<td>" & cstr(gDatosSol11(1,iReg)) & "</td>"
								Response.write "<td>" & gDatosSol11(2,iReg) & "</td>"
								idEstado = gDatosSol11(7,iReg)
								Response.write "<td>" & gArea(idEstado) & "</td>"
								Response.write "<td>" & gDatosSol11(8,iReg) & "</td>"
								Response.write "<td>" & gDatosSol11(3,iReg) & "</td>"
								Response.write "<td>" & gDatosSol11(4,iReg) & "</td>"
								Response.write "<td>" & gDatosSol11(5,iReg) & "</td>"
								sx = replace(gDatosSol11(6,iReg),"_"," ")
								Response.write "<td>" & sx & "</td>"
								Response.flush
							Response.write "</tr>"
						next
					end if
					%>
				</table>
			</div>
		</div>
		<%




	end if
	'response.end 
%>
<style>


</style>	

    <%conexion.close%>
	
</body>
</html>