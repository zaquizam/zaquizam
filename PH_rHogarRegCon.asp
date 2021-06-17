<!DOCTYPE HTML>
<html >
<head>
	<title>Hogar Reg x Cons</title>
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
	<link rel="icon" href="favicon.ico" type="image/x-icon"> 
	<script type="text/javascript" src="js/sweetalert.min.js"></script>
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
	
	dim idCantidadConsumos
	dim idMesConsulta
	dim idTipoConsumo
	dim SemanasConsultas
	dim gHogares
	dim idArea
	dim idEstado
	Server.ScriptTimeout=1000
	Response.buffer = true
	
%>
	<script>
	function GenerarExcel()
	{
		//alert("Generar Excel");
		num = document.getElementById("Excel").value;
		//alert("Generar Excel:="+ num);
		window.open("PH_rHogarRegConExcel.asp?"+num,"_blank");
	}
		

	</script>   
	
<%
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
	'response.write " Combo4:=" & ed_sPar(4,0)
	'response.write " Combo3:=" & ed_sPar(5,0)
	'response.end 
    
	ed_iCombo = 2
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " idPeriodo, "
	sql = sql & " Periodo "
	sql = sql & " FROM "
	sql = sql & " ss_Periodo "
	sql = sql & " WHERE "
	sql = sql & " IdAno = 2021 "
	sql = sql & " Order By "
	sql = sql & " idPeriodo Desc "
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(1,0)="Mes"
    ed_sCombo(1,1)=sql 
    ed_sCombo(1,2)="Seleccionar"

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_TipoConsumo, "
	sql = sql & " TipoConsumo "
	sql = sql & " FROM "
	sql = sql & " PH_TipoConsumo "
	sql = sql & " WHERE "
	sql = sql & " Ind_Activo = 1 "
	sql = sql & " ORDER BY "
	sql = sql & " Id_TipoConsumo "
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(2,0)="Tipo de Consumo"
    ed_sCombo(2,1)=sql 
    ed_sCombo(2,2)="Seleccionar"

End Sub

sub BuscarSemanas
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " IdSemana, "
	sql = sql & " Semana "
	sql = sql & " FROM "
	sql = sql & " ss_Semana "
	sql = sql & " WHERE "
	sql = sql & " Id_Periodo = " & idMesConsulta
	'response.write "<br>232 sql:= " & sql
	'response.end
	rsx2.Open sql ,conexion
	if rsx2.eof then
		rsx2.close
	else 
		SemanasConsultas = rsx2.GetRows
		rsx2.close
	end if
end sub	

sub BuscarHogares
	'if idArea = "" then exit sub
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " PH_PanelHogar.Id_PanelHogar, "
	sql = sql & " PH_PanelHogar.CodigoHogar, "
	sql = sql & " PH_GArea.Area, "
	sql = sql & " ss_Estado.Estado, "
	sql = sql & " PH_Panelistas.Nombre1, "
	sql = sql & " PH_Panelistas.Apellido1, "
	sql = sql & " PH_Panelistas.Titular, "
	sql = sql & " PH_Panelistas.CedulaTitular, "
	sql = sql & " PH_Banco.Banco AS Expr1, "
	sql = sql & " PH_Banco.Codigo AS Expr2, "
	sql = sql & " PH_Panelistas.NumeroCuenta "
	sql = sql & " FROM (((((PH_PanelHogar INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) INNER JOIN PH_GAreaEstado ON ss_Estado.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) LEFT JOIN PH_Panelistas ON PH_PanelHogar.Id_PanelHogar = PH_Panelistas.Id_Hogar) LEFT JOIN PH_Nacionalidad ON PH_Panelistas.Id_Nacionalidad = PH_Nacionalidad.Id_Nacionalidad) LEFT JOIN PH_Banco ON PH_Panelistas.Id_Banco = PH_Banco.Id_Banco "
	sql = sql & " WHERE "
	sql = sql & " PH_PanelHogar.Ind_Activo = 1 "
	sql = sql & " AND PH_Panelistas.ResponsablePanel = 1 "
	'sql = sql & " and PH_PanelHogar.Id_PanelHogar = 706 "
	sql = sql & " Order by "
	sql = sql & " PH_GArea.Area, "
	sql = sql & " ss_Estado.Estado "
	'response.write "<br>232 sql:= " & sql
	'response.end
	rsx2.Open sql ,conexion
	if rsx2.eof then
		rsx2.close
	else 
		gHogares = rsx2.GetRows
		rsx2.close
	end if
		'response.write "<br>241 LLEGO <br>"
		'response.end

end sub
   
    LeePar
	Combos
  
	if ed_sPar(1,0) = "Seleccionar" or ed_sPar(1,0) = "" then
		idMesConsulta = 0 
	else
		idMesConsulta = ed_sPar(1,0)
	end if
    
	if ed_sPar(2,0) = "Seleccionar" or ed_sPar(2,0) = "" then
		idTipoConsumo = 0 
	else
		idTipoConsumo = ed_sPar(2,0)
	end if
    if ed_iPas<>4 then 
        Encabezado
    end if    

	sExcel = "mes=" & idMesConsulta & "&tip=" & idTipoConsumo
    
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
	</br>
	</br>
	</br>
	<%
	'hidden
	
	'response.write "<br> Combo1:=" & ed_sPar(1,0) & "==>" & idMesConsulta
	'response.write "<br> Combo2:=" & ed_sPar(2,0) & "==>" & idTipoConsumo
	'response.write "<br> Combo3:=" & ed_sPar(3,0) & "==>" & idArea
	'response.write "llego1"
	'response.end
	'hidden 
	'response.write "<br>241 LLEGO <br>"
	if cint(idMesConsulta) <> 0 and cint(idTipoConsumo) <> 0 then
		
		BuscarSemanas
		%>
		<div id="DivBuscarInformación">
			<div class="ex1">
				<table class="w3-table w3-striped w3-bordered w3-card-4 w3-small" style=" margin-left:auto; margin-right:auto;margin-top:10px ">
					<thead>
						<tr class="w3-blue">
							<th>idHogar</th>
							<th>Hogar</th>
							<th>Area</th>
							<th>Estado</th>
							<th>Nombre Panelista</th>
							<th>Apellido Panelista</th>
							<%
							for iSem = 0 to ubound(SemanasConsultas,2)
								response.write "<th>" & SemanasConsultas(1,iSem)  & "</th>"
							next
							%>
						</tr>
					</thead>
					<%
					BuscarHogares
					for iHog = 0 to ubound(gHogares,2)
						idHogar = gHogares(0,iHog)
						Response.flush
						response.write "<tr>"
							response.write "<td>" & gHogares(0,iHog) & "</td>"
							response.write "<td>" & gHogares(1,iHog) & "</td>"
							response.write "<td>" & gHogares(2,iHog) & "</td>"
							response.write "<td>" & gHogares(3,iHog) & "</td>"
							response.write "<td>" & gHogares(4,iHog) & "</td>"
							response.write "<td>" & gHogares(5,iHog) & "</td>"
							for iSem = 0 to ubound(SemanasConsultas,2)
								iSemana = SemanasConsultas(0,iSem)
								sql = ""
								sql = sql & " SELECT "
								sql = sql & " Count(PH_Consumo.Id_Consumo) AS CuentaDeId_Consumo "
								sql = sql & " FROM PH_Consumo "
								sql = sql & " WHERE "
								sql = sql & " PH_Consumo.Id_Semana = " & iSemana 
								sql = sql & " AND PH_Consumo.Id_Hogar = " & idHogar
								sql = sql & " AND PH_Consumo.id_TipoConsumo = " & idTipoConsumo
								'response.write "<br>232 sql:= " & sql
								'response.end
								rsx2.Open sql ,conexion
								if rsx2.eof then
									rsx2.close
									iSemanaT = 0
								else 
									gDatosSol1 = rsx2.GetRows
									rsx2.close
									iSemanaT = gDatosSol1(0,0)  
								end if
								sql = ""
								sql = sql & " SELECT "
								sql = sql & " Count(PH_Consumo_Detalle_Productos.Id_Consumo_Detalle_Productos) AS CuentaDeId_Consumo_Detalle_Productos "
								sql = sql & " FROM PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON (PH_Consumo.Id_Hogar = PH_Consumo_Detalle_Productos.Id_Hogar) AND (PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) "
								sql = sql & " WHERE "
								sql = sql & " PH_Consumo.Id_Semana = " & iSemana
								sql = sql & " GROUP BY "
								sql = sql & " PH_Consumo.Id_Hogar, "
								sql = sql & " PH_Consumo.id_TipoConsumo "
								sql = sql & " HAVING "
								sql = sql & " PH_Consumo.Id_Hogar = " & idHogar 
								sql = sql & " AND PH_Consumo.id_TipoConsumo = " & idTipoConsumo
								'response.write "<br>232 sql:= " & sql
								'response.end
								rsx2.Open sql ,conexion
								if rsx2.eof then
									rsx2.close
									iSemanaTReg = 0
								else 
									gDatosSol1 = rsx2.GetRows
									rsx2.close
									iSemanaTReg = gDatosSol1(0,0)  
								end if
								response.write "<td>" & iSemanaT & "-" & iSemanaTReg & "</td>"
							
							next
						response.write "<tr>"
					next 
				%>
				</table>
			</div>
		</div>
	<%
	end if
	conexion.close
	%>
</body>
</html>