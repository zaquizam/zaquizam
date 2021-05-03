<!DOCTYPE HTML>
<html >
<head>
	<title>Consumos</title>
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
    ed_iCombo = 4
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
    ed_sCombo(1,2)="Seleccionar"


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
	sql = sql & " SELECT top 6 "
	sql = sql & " IdSemana, "
	sql = sql & " Semana "
	sql = sql & " FROM ss_Semana "
	sql = sql & " Order By "
	sql = sql & " IdSemana Desc "
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(3,0)="Semana"
    ed_sCombo(3,1)=sql 
    'ed_sCombo(1,2)="Seleccionar"

	sql = ""
	sql = sql & " SELECT  "
	sql = sql & " Id_Categoria, "
	sql = sql & " Categoria "
	sql = sql & " FROM PH_CB_Categoria "
	sql = sql & " Order By "
	sql = sql & " Categoria "
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(4,0)="Categoria"
    ed_sCombo(4,1)=sql 
    ed_sCombo(4,2)="Seleccionar"
	
End Sub
	
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
	if ed_sPar(1,0) = "Seleccionar" then
		idArea = 0 
	else
		idArea = ed_sPar(1,0)
	end if
	if ed_sPar(2,0) = "Seleccionar" then
		idEstado = 0 
	else
		idEstado = ed_sPar(2,0)
	end if
	idSemana = ed_sPar(3,0)
	if ed_sPar(4,0) = "Seleccionar" then
		idCategoria = 0 
	else
		idCategoria = ed_sPar(4,0)
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
						<th>Hogar</th>
						<th>Area</th>
						<th>Estado</th>
						<th>Cant. Prod. Compradas</th>
						<th>Merc. Repos.</th>
						<th>Medicamento</th>
					</tr>
				</thead>
				<%
				if idCategoria <> 0 then
					sql = ""
					sql = sql & " SELECT "
					sql = sql & " PH_Consumo.Id_Hogar, "
					sql = sql & " PH_PanelHogar.CodigoHogar, "
					sql = sql & " PH_GArea.Area, "
					sql = sql & " ss_Estado.Estado, "
					sql = sql & " Sum(PH_Consumo_Detalle_Productos.Cantidad) AS SumaDeCantidad "
					sql = sql & " FROM ((((PH_Consumo INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN (PH_GArea INNER JOIN PH_GAreaEstado ON PH_GArea.Id_Area = PH_GAreaEstado.Id_Area) ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) INNER JOIN PH_Consumo_Detalle_Productos ON (PH_Consumo.Id_Hogar = PH_Consumo_Detalle_Productos.Id_Hogar) AND (PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo)) INNER JOIN PH_CB_Producto ON PH_Consumo_Detalle_Productos.Numero_codigo_barras = PH_CB_Producto.CodigoBarra "
					sql = sql & " WHERE "
					sql = sql & " PH_PanelHogar.Ind_Activo = 1 "
					sql = sql & " AND PH_Consumo.Id_Semana = " & idSemana
					sql = sql & " AND PH_CB_Producto.Id_Categoria = " & idCategoria
					if idArea <> 0 then
						sql = sql & " AND PH_GAreaEstado.Id_Area = " & idArea
					end if
					if idEstado <> 0 then
						sql = sql & " and PH_PanelHogar.Id_Estado = " & idEstado
					end if
					sql = sql & " GROUP BY "
					sql = sql & " PH_Consumo.Id_Hogar, "
					sql = sql & " PH_PanelHogar.CodigoHogar, "
					sql = sql & " PH_GArea.Area, "
					sql = sql & " ss_Estado.Estado "
					sql = sql & " HAVING "
					sql = sql & " PH_Consumo.Id_Hogar<>1 "
					'response.write "<br>36 sql:=" & sql
					'response.end
					rsx1.Open sql ,conexion
					if rsx1.eof then
						rsx1.close
					else 
						gDatosSol1 = rsx1.GetRows
						rsx1.close
					end if
					Total = ubound(gDatosSol1,2) + 1
					Response.write "<h1>"
						Response.write "<center>"
							Response.write "Total Hogares:= " & Total
						Response.write "</center>"
					Response.write "</h1>"
					for iReg = 0 to ubound(gDatosSol1,2)
						Response.write "<tr>"
							Response.write "<td>" 
								Response.write gDatosSol1(0,iReg) & " - " & gDatosSol1(1,iReg)
							Response.write "</td>"
							Response.write "<td>" 
								Response.write gDatosSol1(2,iReg)
							Response.write "</td>"
							Response.write "<td>" 
								Response.write gDatosSol1(3,iReg)
							Response.write "</td>"
							Response.write "<td>" 
								Response.write gDatosSol1(4,iReg)
							Response.write "</td>"
							Response.write "<td>" 
								sql = ""
								sql = sql & " SELECT "
								sql = sql & " Count(PH_Consumo.Id_Consumo) AS CuentaDeId_Consumo "
								sql = sql & " FROM ((((PH_Consumo INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN (PH_GArea INNER JOIN PH_GAreaEstado ON PH_GArea.Id_Area = PH_GAreaEstado.Id_Area) ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) INNER JOIN PH_Consumo_Detalle_Productos ON (PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) AND (PH_Consumo.Id_Hogar = PH_Consumo_Detalle_Productos.Id_Hogar)) INNER JOIN PH_CB_Producto ON PH_Consumo_Detalle_Productos.Numero_codigo_barras = PH_CB_Producto.CodigoBarra "
								sql = sql & " WHERE "
								sql = sql & " PH_Consumo.Id_Hogar = " & gDatosSol1(0,iReg)
								sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
								sql = sql & " AND PH_PanelHogar.Ind_Activo = 1 "
								sql = sql & " AND PH_Consumo.Id_Semana = " & idSemana
								sql = sql & " AND PH_CB_Producto.Id_Categoria = " & idCategoria
								'response.write "<br>36 sql:=" & sql
								'response.end
								rsx2.Open sql ,conexion
								idTotalTipo = 0
								if rsx2.eof then
									rsx2.close
								else 
									gDatosSol2 = rsx2.GetRows
									rsx2.close
									idTotalTipo = gDatosSol2(0,0)
								end if
								Response.write idTotalTipo
							Response.write "</td>"
							Response.write "<td>" 
								sql = ""
								sql = sql & " SELECT "
								sql = sql & " Count(PH_Consumo.Id_Consumo) AS CuentaDeId_Consumo "
								sql = sql & " FROM ((((PH_Consumo INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN (PH_GArea INNER JOIN PH_GAreaEstado ON PH_GArea.Id_Area = PH_GAreaEstado.Id_Area) ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) INNER JOIN PH_Consumo_Detalle_Productos ON (PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) AND (PH_Consumo.Id_Hogar = PH_Consumo_Detalle_Productos.Id_Hogar)) INNER JOIN PH_CB_Producto ON PH_Consumo_Detalle_Productos.Numero_codigo_barras = PH_CB_Producto.CodigoBarra "
								sql = sql & " WHERE "
								sql = sql & " PH_Consumo.Id_Hogar = " & gDatosSol1(0,iReg)
								sql = sql & " AND PH_Consumo.id_TipoConsumo = 8 "
								sql = sql & " AND PH_PanelHogar.Ind_Activo = 1 "
								sql = sql & " AND PH_Consumo.Id_Semana = " & idSemana
								sql = sql & " AND PH_CB_Producto.Id_Categoria = " & idCategoria
								'response.write "<br>36 sql:=" & sql
								'response.end
								rsx2.Open sql ,conexion
								idTotalTipo = 0
								if rsx2.eof then
									rsx2.close
								else 
									gDatosSol2 = rsx2.GetRows
									rsx2.close
									idTotalTipo = gDatosSol2(0,0)
								end if
								Response.write idTotalTipo
							Response.write "</td>"
						Response.write "</tr>"
					next 
				end if
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