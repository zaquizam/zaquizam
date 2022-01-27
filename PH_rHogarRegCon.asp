<!DOCTYPE HTML>
<html >
<head>
	<title>| Hogar Reg x Cons |</title>
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
	<link rel="icon" href="favicon.ico" type="image/x-icon"> 
	<script type="text/javascript" src="js/sweetalert.min.js"></script>
	<script>
		function GenerarExcel()
		{
			//alert("Generar Excel");
			num = document.getElementById("Excel").value;
			//alert("Generar Excel:="+ num);
			window.open("PH_rHogarRegConExcel.asp?"+num,"_blank");
		}
	</script>   
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
	Server.ScriptTimeout=10000
	Response.Buffer = true
	
	Dim idCantidadConsumos
	Dim idMesConsulta
	Dim idTipoConsumo
	Dim SemanasConsultas
	Dim gHogares
	Dim idArea
	Dim idEstado
	Dim gDatosSol1
	Dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 1 'adLockOptimistic 

	Dim gDatosSol2
	Dim rsx2
	set rsx2 = CreateObject("ADODB.Recordset")
	rsx2.CursorType = adOpenKeyset 
	rsx2.LockType = 1 'adLockOptimistic 


Sub Combos
 
	'Response.Write "<br>372 Combo1:=" & ed_sPar(1,0)
	'Response.Write " Combo2:=" & ed_sPar(2,0)
	'Response.Write " Combo3:=" & ed_sPar(3,0)
	'Response.Write " Combo4:=" & ed_sPar(4,0)
	'Response.Write " Combo3:=" & ed_sPar(5,0)
	'response.end 
    
	ed_iCombo = 2
	sql = vbnullstring
	sql = sql & " SELECT idPeriodo, Periodo FROM  ss_Periodo WHERE IdAno = 2021 or IdAno = 2022 Order By idPeriodo Desc "
	'Response.Write "<br>372 Combo1:=" & sql
    ed_sCombo(1,0)="Mes"
    ed_sCombo(1,1)=sql 
    ed_sCombo(1,2)="Seleccionar"

	sql = vbnullstring
	sql = sql & " SELECT Id_TipoConsumo, TipoConsumo FROM PH_TipoConsumo WHERE Ind_Activo = 1 ORDER BY Id_TipoConsumo "
	'Response.Write "<br>372 Combo1:=" & sql
    ed_sCombo(2,0)="Tipo de Consumo"
    ed_sCombo(2,1)=sql 
    ed_sCombo(2,2)="Seleccionar"

End Sub

sub BuscarSemanas
	sql = vbnullstring
	sql = sql & " SELECT IdSemana, Semana FROM ss_Semana WHERE Id_Periodo = " & idMesConsulta
	'Response.Write "<br>232 sql:= " & sql
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
	sql = vbnullstring
	sql = sql & " SELECT PH_PanelHogar.Id_PanelHogar, PH_PanelHogar.CodigoHogar, PH_GArea.Area, ss_Estado.Estado, PH_Panelistas.Nombre1, PH_Panelistas.Apellido1,"
	sql = sql & " PH_Panelistas.Titular, PH_Panelistas.CedulaTitular, PH_Banco.Banco AS Expr1, PH_Banco.Codigo AS Expr2, PH_Panelistas.NumeroCuenta "
	sql = sql & " FROM (((((PH_PanelHogar INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) INNER JOIN PH_GAreaEstado ON ss_Estado.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) LEFT JOIN PH_Panelistas ON PH_PanelHogar.Id_PanelHogar = PH_Panelistas.Id_Hogar) LEFT JOIN PH_Nacionalidad ON PH_Panelistas.Id_Nacionalidad = PH_Nacionalidad.Id_Nacionalidad) LEFT JOIN PH_Banco ON PH_Panelistas.Id_Banco = PH_Banco.Id_Banco "
	sql = sql & " WHERE PH_PanelHogar.Ind_Activo = 1 AND PH_Panelistas.ResponsablePanel = 1 "
	'sql = sql & " and PH_PanelHogar.Id_PanelHogar = 706 "
	sql = sql & " Order by PH_GArea.Area, ss_Estado.Estado "
	'Response.Write "<br>232 sql:= " & sql
	'response.end
	rsx2.Open sql ,conexion
	if rsx2.eof then
		rsx2.close
	else 
		gHogares = rsx2.GetRows
		rsx2.close
	end if
		'Response.Write "<br>241 LLEGO <br>"
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
	<!--<div style="width:98%"></div>-->
	
	</center>
	<table border="0" align="right">
		<tr>
			<td>
				<%ed_vCombo%>
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
	
	'Response.Write "<br> Combo1:=" & ed_sPar(1,0) & "==>" & idMesConsulta
	'Response.Write "<br> Combo2:=" & ed_sPar(2,0) & "==>" & idTipoConsumo
	'Response.Write "<br> Combo3:=" & ed_sPar(3,0) & "==>" & idArea
	'Response.Write "llego1"
	'response.end
	'hidden 
	'Response.Write "<br>241 LLEGO <br>"
	if CInt(idMesConsulta) <> 0 and CInt(idTipoConsumo) <> 0 then
		
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
								if IsArray(SemanasConsultas) then
									for iSem = 0 to ubound(SemanasConsultas,2)
										Response.Write "<th>" & SemanasConsultas(1,iSem)  & "</th>"
									next
								end if						
							%>
						</tr>
					</thead>
					<%
					BuscarHogares
					FOR iHog = 0 TO ubound(gHogares,2)
						idHogar = gHogares(0,iHog)						
						Response.Flush
						Response.Write "<tr>"
							Response.Write "<td>" & gHogares(0,iHog) & "</td>"
							Response.Write "<td>" & gHogares(1,iHog) & "</td>"
							Response.Write "<td>" & gHogares(2,iHog) & "</td>"
							Response.Write "<td>" & gHogares(3,iHog) & "</td>"
							Response.Write "<td>" & gHogares(4,iHog) & "</td>"
							Response.Write "<td>" & gHogares(5,iHog) & "</td>"
							for iSem = 0 to ubound(SemanasConsultas,2)
								iSemana = SemanasConsultas(0,iSem)
								sql = vbnullstring
								sql = sql & " SELECT Count(PH_Consumo.Id_Consumo) AS CuentaDeId_Consumo FROM PH_Consumo "
								sql = sql & " WHERE PH_Consumo.Id_Semana = " & iSemana 
								sql = sql & " AND PH_Consumo.Id_Hogar = " & idHogar
								sql = sql & " AND PH_Consumo.id_TipoConsumo = " & idTipoConsumo
								'Response.Write "<br>232 sql:= " & sql
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
								sql = vbnullstring
								sql = sql & " SELECT "
								sql = sql & " Count(PH_Consumo_Detalle_Productos.Id_Consumo_Detalle_Productos) AS CuentaDeId_Consumo_Detalle_Productos "
								sql = sql & " FROM PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON (PH_Consumo.Id_Hogar = PH_Consumo_Detalle_Productos.Id_Hogar) AND (PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) "
								sql = sql & " WHERE "
								sql = sql & " PH_Consumo.Id_Semana = " & iSemana
								sql = sql & " GROUP BY PH_Consumo.Id_Hogar, PH_Consumo.id_TipoConsumo "
								sql = sql & " HAVING "
								sql = sql & " PH_Consumo.Id_Hogar = " & idHogar 
								sql = sql & " AND PH_Consumo.id_TipoConsumo = " & idTipoConsumo
								'Response.Write "<br>232 sql:= " & sql
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
								
								Response.Write "<td>" & iSemanaT & "-" & iSemanaTReg & "</td>"
							
							next
						Response.Write "<tr>"
						Response.Flush						
					NEXT 
					Response.Flush
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