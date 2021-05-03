<!DOCTYPE HTML>
<html >
<head>
	<title>Consumos Semanal</title>
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
%>
<script type="text/javascript">
	function GenerarExcel()
	{
		//alert("Generar Excel");
		num = document.getElementById("Excel").value;
		//alert("Generar Excel:="+ num);
		window.open("ph_rConsumosSemanalExcel.asp?" +num,"_blank");
	}
</script>
<%
	dim gDatosSol1
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	dim gSemanas
	dim gTipoConsumo
	dim gHogares
	dim gCantidad
	dim idArea
	dim idEstado
	dim idTipoConsumo
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
    ed_iCombo = 3
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Area, "
	sql = sql & " Area "
	sql = sql & " FROM PH_GArea "
	sql = sql & " Order By "
	sql = sql & " Id_Area "
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(1,0)="Area"
    ed_sCombo(1,1)=sql 
    'ed_sCombo(1,2)="Seleccionar"


	sql = ""
	sql = sql & " SELECT "
	sql = sql & " PH_GAreaEstado.Id_Estado, "
	sql = sql & " ss_Estado.Estado "
	sql = sql & " FROM PH_GAreaEstado INNER JOIN ss_Estado ON PH_GAreaEstado.Id_Estado = ss_Estado.Id_Estado "
	if ed_sPar(1,0) <> 0 then
		sql = sql & " WHERE PH_GAreaEstado.Id_Area = " & ed_sPar(1,0)
	end if
	sql = sql & " Order By "
	sql = sql & " ss_Estado.Estado "
	'response.write "<br>372 Combo2:=" & sql
    ed_sCombo(2,0)="Estado"
    ed_sCombo(2,1)=sql 
    ed_sCombo(2,2)="Seleccionar"

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_TipoConsumo, "
	sql = sql & " TipoConsumo "
	sql = sql & " FROM "
	sql = sql & " PH_TipoConsumo "
	sql = sql & " Where "
	sql = sql & " Ind_Activo = 1 "
	'sql = sql & " and Id_TipoConsumo = 1 "
	sql = sql & " ORDER BY "
	sql = sql & " Id_TipoConsumo "
	'response.write "<br>372 Combo2:=" & sql
    ed_sCombo(3,0)="Tipo de Consumo"
    ed_sCombo(3,1)=sql 
    'ed_sCombo(3,2)="Seleccionar"
	
End Sub
Sub Semanas
	'Tipos de Consumo 
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Top 4 "
	sql = sql & " IdSemana, "
	sql = sql & " Semana "
	sql = sql & " FROM "
	sql = sql & " ss_Semana "
	'sql = sql & " Where "
	'sql = sql & " IdSemana < 17 "
	sql = sql & " ORDER BY "
	sql = sql & " IdSemana DESC "
	'response.write "<br>36 sql:=" & sql
	'response.end
	rsx2.Open sql ,conexion
	if rsx2.eof then
		rsx2.close
	else 
		gSemanas = rsx2.GetRows
		rsx2.close
	end if

End Sub

Sub TipoConsumo
	if ed_sPar(3,0) = "" then ed_sPar(3,0) = 1
	if ed_sPar(3,0) > 10 then ed_sPar(3,0) = 1 
	'Tipos de Consumo 
	sql = "" 
	sql = sql & " SELECT "
	sql = sql & " Id_TipoConsumo, "
	sql = sql & " TipoConsumo "
	sql = sql & " FROM "
	sql = sql & " PH_TipoConsumo "
	sql = sql & " Where "
	sql = sql & " Ind_Activo = 1 "
	sql = sql & " and Id_TipoConsumo = " & ed_sPar(3,0)
	sql = sql & " ORDER BY " 
	sql = sql & " Id_TipoConsumo "
	'response.write "<br>36 sql:=" & sql
	'response.end
	rsx2.Open sql ,conexion 
	if rsx2.eof then 
		rsx2.close 
	else  
		gTipoConsumo = rsx2.GetRows
		rsx2.close
	end if
End Sub

Sub Hogares
	idArea = ed_sPar(1,0)
	idEstado = ed_sPar(2,0)
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " PH_PanelHogar.Id_PanelHogar, "
	sql = sql & " PH_PanelHogar.CodigoHogar, "
	sql = sql & " PH_GArea.Area, "
	sql = sql & " ss_Estado.Estado, "
	sql = sql & " PH_Panelistas.Nombre1, "
	sql = sql & " PH_Panelistas.Apellido1, "
	sql = sql & " PH_Panelistas.Celular "
	sql = sql & " FROM (((PH_PanelHogar INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) INNER JOIN PH_GAreaEstado ON ss_Estado.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) LEFT JOIN PH_Panelistas ON PH_PanelHogar.Id_PanelHogar = PH_Panelistas.Id_Hogar "
	sql = sql & " WHERE "
	sql = sql & " PH_PanelHogar.Ind_Activo = 1 "
	sql = sql & " AND PH_PanelHogar.Id_PanelHogar > 1 "
	sql = sql & " AND PH_Panelistas.ResponsablePanel = 1 "
	if ed_sPar(1,0) <> "0" and ed_sPar(1,0) <> "" and ed_sPar(1,0) <> "Seleccionar" then
		sql = sql & " AND PH_GAreaEstado.Id_Area = " & idArea
	end if
	if ed_sPar(2,0) <> "" and ed_sPar(2,0) <> "Seleccionar" then
		sql = sql & " AND  PH_PanelHogar.Id_Estado = " & idEstado
	end if
	'sql = sql & " and PH_PanelHogar.Id_PanelHogar = 874 "
	'response.write "<br>36 sql:=" & sql
	'response.end
	rsx2.Open sql ,conexion
	if rsx2.eof then
		rsx2.close
	else 
		gHogares = rsx2.GetRows
		rsx2.close
	end if
End Sub
	
   
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
	if ed_sPar(1,0) = "" or ed_sPar(1,0) = "Seleccionar" then ed_sPar(1,0) = 1
    Combos
	Semanas
	TipoConsumo
	Hogares
	
%>
		
	<br>
	<div style="width:98%">
	<%
	
	%></div></center>
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
	if ed_sPar(1,0) = "" then
		idArea = 1
	else
		idArea = ed_sPar(1,0)
	end if
	if ed_sPar(2,0) = "Seleccionar" then
		idEstado = 0 
	else
		idEstado = ed_sPar(2,0)
	end if
	if ed_sPar(3,0) = "" then
		idTipoConsumo = 1
	else
		idArea = ed_sPar(1,0)
	end if
	'response.write "<br> Combo1:=" & ed_sPar(1,0) & "==>" & idArea
	'response.write "<br> Combo2:=" & ed_sPar(2,0) & "==>" & idEstado
	'response.write "<br> Combo3:=" & ed_sPar(3,0) & "==>" & idTipoConsumo
	'response.write "<br>llego"
	'response.end
	'hidden 
	idTipoConsumo = ed_sPar(3,0)
	sExcel = "are=" & idArea & "&est=" & idEstado & "&tip=" & idTipoConsumo 
	
	%>
	<input type="hidden" name="Programa" id="Programa" align="right" size=50>
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

	<div id="DivBuscarInformación">
		<div class="ex1">
			<table class="w3-table w3-striped w3-bordered w3-card-4 w3-small" style="width:1000px; margin-left:auto; margin-right:auto;margin-top:10px ">
				<thead>
					<tr class="w3-blue">
						<th>IdHogar</th>
						<th>CodHogar</th>
						<th>Area</th>
						<th>Estado</th>
						<th>Nombre</th>
						<th>Apellido</th>
						<th>Celular</th>
						<th>Tipo Consumo</th>
						<%
						for iReg = 0 to ubound(gSemanas,2)
							response.write "<th>" & gSemanas(1,iReg) & "</th>"
						next 
						%>
					</tr>
				</thead>
				<% 
				for iReg = 0 to ubound(gHogares,2)
					Response.flush
					response.write "<tr>"
						idHogar = gHogares(0,iReg)
						for iCol = 0 to 6
							response.write "<td>" & gHogares(iCol,iReg) & "</td>"
						next
						isw = 1
						'for iReg2 = 0 to ubound(gTipoConsumo,2)
							'idTipoConsumo = gTipoConsumo(0,iReg2)
							idTipoConsumo = ed_sPar(3,0)
							'response.write "<br>309 idTipoConsumo:= " & idTipoConsumo
							if isw = 0 then
								response.write "<td></td>"
								response.write "<td></td>"
								response.write "<td></td>"
								response.write "<td></td>"
								response.write "<td></td>"
								response.write "<td></td>"
								response.write "<td></td>"
							end if
							'response.write "<td>" & gTipoConsumo(1,iReg2) & "</td>"
							response.write "<td>" & gTipoConsumo(1,0) & "</td>"
							
							
							for iReg3 = 0 to 3
								idSemana = gSemanas(0,iReg3)
								'response.write "<br>Semana:= " & idSemana
								'Consumos
								sql = ""
								sql = sql & " SELECT "
								sql = sql & " Count(Id_Consumo) AS Total "
								sql = sql & " FROM "
								sql = sql & " PH_Consumo "
								sql = sql & " WHERE "
								sql = sql & " Id_Semana = " & idSemana
								sql = sql & " AND Id_Hogar = " & idHogar
								sql = sql & " AND id_TipoConsumo = " & idTipoConsumo
								'response.write "<br>36 sql:=" & sql
								'response.end
								rsx1.Open sql ,conexion
								if rsx1.eof then
									rsx1.close
									Cantidad = 0
								else 
									gCantidad = rsx1.GetRows
									rsx1.close
									Cantidad = gCantidad(0,0)
								end if
								response.write "<td>" & Cantidad & "</td>"
							next 
							response.write "</tr>"
							isw = 0
						'next
					response.write "</tr>"
					'Response.flush
				next 

				%>
			</table>
		</div>
	</div>
	<div id="Reporte">
	</div>
	</br>
	</br>
	</br>
	</br>
	</br>
    <%conexion.close%>


</body>
</html>