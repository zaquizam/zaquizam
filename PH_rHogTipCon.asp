<!DOCTYPE HTML>
<html >
<head>
	<title>Hogares x Tipo Consumo</title>
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
		alert("Generar Excel aun no Disponible");
		return;
		num = document.getElementById("Excel").value;
		//alert("Generar Excel:="+ num);
		window.open("ph_rConsumosSemanalExcel.asp?num=" +num,"_blank");
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
	idArea = ed_sPar(1,0)
	idEstado = ed_sPar(2,0)
 
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
	sql = sql & " PH_PanelHogar.Id_PanelHogar, "
	sql = sql & " str(PH_PanelHogar.Id_PanelHogar) + '-' + "
	sql = sql & " PH_PanelHogar.CodigoHogar"
	sql = sql & " FROM PH_PanelHogar INNER JOIN PH_GAreaEstado ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado "
	sql = sql & " WHERE "
	sql = sql & " PH_PanelHogar.Ind_Activo = 1 "
	sql = sql & " and PH_GAreaEstado.Id_Area = " & idArea
	if ed_sPar(2,0) <> "" and ed_sPar(2,0) <> "Seleccionar" then
		sql = sql & " AND  PH_PanelHogar.Id_Estado = " & idEstado
	end if
	sql = sql & " ORDER BY "
	sql = sql & " PH_PanelHogar.Id_PanelHogar "
	'response.write "<br>372 Combo2:=" & sql
    ed_sCombo(3,0)="Hogares"
    ed_sCombo(3,1)=sql 
    ed_sCombo(3,2)="Seleccionar"
	
End Sub

Sub Hogares
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " PH_PanelHogar.Id_PanelHogar, "
	sql = sql & " PH_PanelHogar.CodigoHogar, "
	sql = sql & " PH_GArea.Area, ss_Estado.Estado, "
	sql = sql & " PH_Panelistas.Nombre1, "
	sql = sql & " PH_Panelistas.Apellido1, "
	sql = sql & " PH_Panelistas.Cedula, "
	sql = sql & " PH_Panelistas.Celular, "
	sql = sql & " ss_Semana.Semana, "
	sql = sql & " PH_TipoConsumo.TipoConsumo, "
	sql = sql & " Count(PH_Consumo.Id_Consumo) AS CuentaDeId_Consumo "
	sql = sql & " FROM ((((((PH_PanelHogar INNER JOIN PH_GAreaEstado ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) LEFT JOIN PH_Panelistas ON PH_PanelHogar.Id_PanelHogar = PH_Panelistas.Id_Hogar) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) INNER JOIN PH_Consumo ON PH_PanelHogar.Id_PanelHogar = PH_Consumo.Id_Hogar) INNER JOIN ss_Semana ON PH_Consumo.Id_Semana = ss_Semana.IdSemana) INNER JOIN PH_TipoConsumo ON PH_Consumo.id_TipoConsumo = PH_TipoConsumo.Id_TipoConsumo "
	sql = sql & " GROUP BY "
	sql = sql & " PH_PanelHogar.Id_PanelHogar, "
	sql = sql & " ss_Semana.IdSemana, "
	sql = sql & " PH_PanelHogar.CodigoHogar, "
	sql = sql & " PH_GArea.Area, "
	sql = sql & " ss_Estado.Estado, "
	sql = sql & " PH_Panelistas.Nombre1, "
	sql = sql & " PH_Panelistas.Apellido1, "
	sql = sql & " PH_Panelistas.Cedula, "
	sql = sql & " PH_Panelistas.Celular, "
	sql = sql & " ss_Semana.Semana, "
	sql = sql & " PH_TipoConsumo.TipoConsumo, "
	sql = sql & " PH_PanelHogar.Id_Estado, "
	sql = sql & " PH_GAreaEstado.Id_Area, "
	sql = sql & " PH_PanelHogar.Ind_Activo, "
	sql = sql & " PH_Panelistas.ResponsablePanel "
	sql = sql & " HAVING "
	sql = sql & " ss_Semana.IdSemana > 15 "
	sql = sql & " AND PH_PanelHogar.Ind_Activo = 1 "
	sql = sql & " AND PH_Panelistas.ResponsablePanel = 1 "
	if ed_sPar(1,0) <> "" and ed_sPar(1,0) <> "Seleccionar" then
		sql = sql & " AND PH_GAreaEstado.Id_Area = " & ed_sPar(1,0)
	end if
	if ed_sPar(2,0) <> "" and ed_sPar(2,0) <> "Seleccionar" then
		sql = sql & " AND PH_PanelHogar.Id_Estado = " & ed_sPar(2,0)
	end if
	if ed_sPar(3,0) <> "" and ed_sPar(3,0) <> "Seleccionar" then
		sql = sql & " AND PH_PanelHogar.Id_PanelHogar = " & ed_sPar(3,0)
	end if
	sql = sql & " ORDER BY "
	sql = sql & " PH_PanelHogar.Id_PanelHogar, "
	sql = sql & " ss_Semana.IdSemana "
	
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
	'Hogares
	
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
	'response.write "<br> Combo1:=" & ed_sPar(1,0) & "==>" & idArea
	'response.write "<br> Combo2:=" & ed_sPar(2,0) & "==>" & idEstado
	'response.write "<br> Combo3:=" & ed_sPar(3,0) & "==>" & idTipoConsumo
	'response.write "<br>llego"
	'response.end
	'hidden 
	
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
						<th>Cedula</th>
						<th>Celular</th>
						<th>Semana</th>
						<th>Tipo Consumo</th>
						<th>Cantidad</th>
						<%
						Hogares
						for iReg = 0 to ubound(gHogares,2)
							Response.flush
							response.write "<tr>"
								idHogar = gHogares(0,iReg)
								for iCol = 0 to 10
									response.write "<td>" & gHogares(iCol,iReg) & "</td>"
								next
							response.write "</tr>"
							'Response.flush
						next 
						
						%>
					</tr>
				</thead>
				<% 

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