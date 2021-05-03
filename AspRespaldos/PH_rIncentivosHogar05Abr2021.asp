<!DOCTYPE HTML>
<html >
<head>
	<title>Pago de Incentivos al Hogar</title>
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
	dim idMesPago
	dim idSemanasPago
	dim gHogares
	dim idArea
	
%>
	<script>
	function GenerarExcel()
	{
		//alert("Generar Excel");
		num = document.getElementById("Excel").value;
		//alert("Generar Excel:="+ num);
		window.open("ph_rIncentivosHogarExcel.asp?"+num,"_blank");
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
	'response.write " Combo3:=" & ed_sPar(4,0)
	'response.write " Combo3:=" & ed_sPar(5,0)
    ed_iCombo = 3
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id, "
	sql = sql & " Periodo "
	sql = sql & " FROM "
	sql = sql & " ss_Periodo "
	sql = sql & " WHERE "
	sql = sql & " IdAno = 2021 "
	sql = sql & " AND IdMes > 1 "
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(1,0)="Mes"
    ed_sCombo(1,1)=sql 
    ed_sCombo(1,2)="Seleccionar"

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Cantidad, "
	sql = sql & " Cantidad "
	sql = sql & " FROM "
	sql = sql & " ss_Cantidad "
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(2,0)="Cantidad Semanas"
    ed_sCombo(2,1)=sql 
    ed_sCombo(2,2)="Seleccionar"

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Area, "
	sql = sql & " Area "
	sql = sql & " FROM PH_GArea "
	sql = sql & " Order By "
	sql = sql & " Id_Area "
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(3,0)="Area"
    ed_sCombo(3,1)=sql 
    ed_sCombo(3,2)="Seleccionar"

	
End Sub

sub BuscarSemanas
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Semanas "
	sql = sql & " FROM "
	sql = sql & " ss_Periodo "
	sql = sql & " WHERE "
	sql = sql & " Id = " & idMesPago
	'response.write "<br>232 sql:= " & sql
	'response.end
	rsx2.Open sql ,conexion
	if rsx2.eof then
		rsx2.close
	else 
		gDatosSol2 = rsx2.GetRows
		rsx2.close
		idSemanasPago = gDatosSol2(0,0)
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
	sql = sql & " PH_Nacionalidad.Abreviatura, "
	sql = sql & " PH_Panelistas.Cedula, "
	sql = sql & " PH_Panelistas.NumeroCuenta "
	sql = sql & " FROM ((((PH_PanelHogar INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) INNER JOIN PH_GAreaEstado ON ss_Estado.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) LEFT JOIN PH_Panelistas ON PH_PanelHogar.Id_PanelHogar = PH_Panelistas.Id_Hogar) LEFT JOIN PH_Nacionalidad ON PH_Panelistas.Id_Nacionalidad = PH_Nacionalidad.Id_Nacionalidad "
	sql = sql & " WHERE "
	sql = sql & " PH_PanelHogar.Ind_Activo = 1 "
	sql = sql & " AND PH_Panelistas.ResponsablePanel = 1 "
	sql = sql & " and PH_GArea.id_Area = "  & idArea
	'sql = sql & " and PH_PanelHogar.Id_PanelHogar = 123 "
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

end sub
   
    LeePar
	Combos
	
  
	if ed_sPar(1,0) = "Seleccionar" then
		idMesPago = 0 
	else
		idMesPago = ed_sPar(1,0)
	end if
	if ed_sPar(2,0) = "Seleccionar" then
		idCantidadConsumos = 0 
	else
		idCantidadConsumos = ed_sPar(2,0)
	end if

	if ed_sPar(3,0) = "Seleccionar" then
		idArea = 0 
	else
		idArea = ed_sPar(3,0)
	end if
    
    if ed_iPas<>4 then 
        Encabezado
    end if    

	sExcel = "mes=" & idMesPago & "&can=" & idCantidadConsumos & "&are=" & idArea
    
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
	
	'response.write "<br> Combo1:=" & ed_sPar(1,0) & "==>" & idMesPago
	'response.write "<br> Combo2:=" & ed_sPar(2,0) & "==>" & idCantidadConsumos
	'response.write "<br> Combo3:=" & ed_sPar(3,0) & "==>" & idArea
	'response.write "llego1"
	if idMesPago = "" then idMesPago = 0
	if idCantidadConsumos = "" then idCantidadConsumos = 0
	if idArea = "" then idArea = 0 
	'response.end
	'hidden 

	if cint(idMesPago) <> 0 and cint(idCantidadConsumos) <> 0 and cint(idArea) <> 0 then
		'BuscarSemanas
		BuscarHogares
		'response.write "<br>140 Semanas:= " & idSemanasPago
		%>
		<div id="DivBuscarInformación">
			<div class="ex1">
				<table class="w3-table w3-striped w3-bordered w3-card-4 w3-small" style="width:1400px; margin-left:auto; margin-right:auto;margin-top:10px ">
					<thead>
						<tr class="w3-blue">
							<th>idHogar</th>
							<th>Hogar</th>
							<th>Area</th>
							<th>Estado</th>
							<th>Nombre</th>
							<th>Apellido</th>
							<th>Cedula</th>
							<th>Cuenta</th>
							<th>(05)<br>Del 01 Feb 2021<br>al 07 Feb 2021</th>
							<th>(06)<br>Del 08 Feb 2021<br>al 14 Feb 2021</th>
							<th>(07)<br>Del 15 Feb 2021<br>al 21 Feb 2021</th>
							<th>(08)<br>Del 22 Feb 2021<br>al 28 Feb 2021</th>
							<th>Encuesta</th>
							<th>Pagar Incentivo</th>
							<th>Pagar Encuesta</th>
						</tr>
					</thead>
					<%
					for iHog = 0 to ubound(gHogares,2)
						Response.flush
						response.write "<tr>"
							idHogar = gHogares(0,iHog)
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " Count(PH_Consumo.Id_Consumo) AS CuentaDeId_Consumo "
							sql = sql & " FROM PH_Consumo "
							sql = sql & " WHERE "
							sql = sql & " PH_Consumo.Id_Semana = 20 "
							sql = sql & " AND PH_Consumo.Id_Hogar = " & idHogar
							sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
							'response.write "<br>232 sql:= " & sql
							'response.end
							rsx2.Open sql ,conexion
							if rsx2.eof then
								rsx2.close
								iSemana1 = 0
							else 
								gDatosSol1 = rsx2.GetRows
								rsx2.close
								iSemana1 = gDatosSol1(0,0)  
							end if
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " Count(PH_Consumo_Detalle_Productos.Id_Consumo_Detalle_Productos) AS CuentaDeId_Consumo_Detalle_Productos "
							sql = sql & " FROM PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON (PH_Consumo.Id_Hogar = PH_Consumo_Detalle_Productos.Id_Hogar) AND (PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) "
							sql = sql & " WHERE "
							sql = sql & " PH_Consumo.Id_Semana = 20 "
							sql = sql & " GROUP BY "
							sql = sql & " PH_Consumo.Id_Hogar, "
							sql = sql & " PH_Consumo.id_TipoConsumo "
							sql = sql & " HAVING "
							sql = sql & " PH_Consumo.Id_Hogar = " & idHogar 
							sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
							'response.write "<br>232 sql:= " & sql
							'response.end
							rsx2.Open sql ,conexion
							if rsx2.eof then
								rsx2.close
								iSemana1Reg = 0
							else 
								gDatosSol1 = rsx2.GetRows
								rsx2.close
								iSemana1Reg = gDatosSol1(0,0)  
							end if

							sql = ""
							sql = sql & " SELECT "
							sql = sql & " Count(PH_Consumo.Id_Consumo) AS CuentaDeId_Consumo "
							sql = sql & " FROM PH_Consumo "
							sql = sql & " WHERE "
							sql = sql & " PH_Consumo.Id_Semana = 21 "
							sql = sql & " AND PH_Consumo.Id_Hogar = " & idHogar
							sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
							'response.write "<br>232 sql:= " & sql
							'response.end
							rsx2.Open sql ,conexion
							if rsx2.eof then
								rsx2.close
								iSemana2 = 0
							else 
								gDatosSol1 = rsx2.GetRows
								rsx2.close
								iSemana2 = gDatosSol1(0,0)  
							end if
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " Count(PH_Consumo_Detalle_Productos.Id_Consumo_Detalle_Productos) AS CuentaDeId_Consumo_Detalle_Productos "
							sql = sql & " FROM PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON (PH_Consumo.Id_Hogar = PH_Consumo_Detalle_Productos.Id_Hogar) AND (PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) "
							sql = sql & " WHERE "
							sql = sql & " PH_Consumo.Id_Semana = 21 "
							sql = sql & " GROUP BY "
							sql = sql & " PH_Consumo.Id_Hogar, "
							sql = sql & " PH_Consumo.id_TipoConsumo "
							sql = sql & " HAVING "
							sql = sql & " PH_Consumo.Id_Hogar = " & idHogar 
							sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
							'response.write "<br>232 sql:= " & sql
							'response.end
							rsx2.Open sql ,conexion
							if rsx2.eof then
								rsx2.close
								iSemana2Reg = 0
							else 
								gDatosSol1 = rsx2.GetRows
								rsx2.close
								iSemana2Reg = gDatosSol1(0,0)  
							end if
							

							sql = ""
							sql = sql & " SELECT "
							sql = sql & " Count(PH_Consumo.Id_Consumo) AS CuentaDeId_Consumo "
							sql = sql & " FROM PH_Consumo "
							sql = sql & " WHERE "
							sql = sql & " PH_Consumo.Id_Semana = 22 "
							sql = sql & " AND PH_Consumo.Id_Hogar = " & idHogar
							sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
							'response.write "<br>232 sql:= " & sql
							'response.end
							rsx2.Open sql ,conexion
							if rsx2.eof then
								rsx2.close
								iSemana3 = 0
							else 
								gDatosSol1 = rsx2.GetRows
								rsx2.close
								iSemana3 = gDatosSol1(0,0)  
							end if
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " Count(PH_Consumo_Detalle_Productos.Id_Consumo_Detalle_Productos) AS CuentaDeId_Consumo_Detalle_Productos "
							sql = sql & " FROM PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON (PH_Consumo.Id_Hogar = PH_Consumo_Detalle_Productos.Id_Hogar) AND (PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) "
							sql = sql & " WHERE "
							sql = sql & " PH_Consumo.Id_Semana = 22 "
							sql = sql & " GROUP BY "
							sql = sql & " PH_Consumo.Id_Hogar, "
							sql = sql & " PH_Consumo.id_TipoConsumo "
							sql = sql & " HAVING "
							sql = sql & " PH_Consumo.Id_Hogar = " & idHogar 
							sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
							'response.write "<br>232 sql:= " & sql
							'response.end
							rsx2.Open sql ,conexion
							if rsx2.eof then
								rsx2.close
								iSemana3Reg = 0
							else 
								gDatosSol1 = rsx2.GetRows
								rsx2.close
								iSemana3Reg = gDatosSol1(0,0)  
							end if
							
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " Count(PH_Consumo.Id_Consumo) AS CuentaDeId_Consumo "
							sql = sql & " FROM PH_Consumo "
							sql = sql & " WHERE "
							sql = sql & " PH_Consumo.Id_Semana = 23 "
							sql = sql & " AND PH_Consumo.Id_Hogar = " & idHogar
							sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
							'response.write "<br>232 sql:= " & sql
							'response.end
							rsx2.Open sql ,conexion
							if rsx2.eof then
								rsx2.close
								iSemana4 = 0
							else 
								gDatosSol1 = rsx2.GetRows
								rsx2.close
								iSemana4 = gDatosSol1(0,0)  
							end if
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " Count(PH_Consumo_Detalle_Productos.Id_Consumo_Detalle_Productos) AS CuentaDeId_Consumo_Detalle_Productos "
							sql = sql & " FROM PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON (PH_Consumo.Id_Hogar = PH_Consumo_Detalle_Productos.Id_Hogar) AND (PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) "
							sql = sql & " WHERE "
							sql = sql & " PH_Consumo.Id_Semana = 23 "
							sql = sql & " GROUP BY "
							sql = sql & " PH_Consumo.Id_Hogar, "
							sql = sql & " PH_Consumo.id_TipoConsumo "
							sql = sql & " HAVING "
							sql = sql & " PH_Consumo.Id_Hogar = " & idHogar 
							sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
							'response.write "<br>232 sql:= " & sql
							'response.end
							rsx2.Open sql ,conexion
							if rsx2.eof then
								rsx2.close
								iSemana4Reg = 0
							else 
								gDatosSol1 = rsx2.GetRows
								rsx2.close
								iSemana4Reg = gDatosSol1(0,0)  
							end if

							sql = ""
							sql = sql & " SELECT "
							sql = sql & " Id_EncuestaEspecial, "
							sql = sql & " Id_Hogar, "
							sql = sql & " Ind_Realizada "
							sql = sql & " FROM PH_EncuestaHogar "
							sql = sql & " WHERE "
							sql = sql & " Id_EncuestaEspecial = 27 "
							sql = sql & " and Id_Hogar = " & idHogar
							sql = sql & " AND Ind_Realizada =  1 "
							'response.write "<br>232 sql:= " & sql
							'response.end
							rsx2.Open sql ,conexion
							if rsx2.eof then
								rsx2.close
								iEncuesta = 0
							else 
								gDatosSol1 = rsx2.GetRows
								rsx2.close
								iEncuesta = 1
							end if

							iSemana = 0
							if iSemana1 > 0 then iSemana = iSemana + 1
							if iSemana2 > 0 then iSemana = iSemana + 1
							if iSemana3 > 0 then iSemana = iSemana + 1
							if iSemana4 > 0 then iSemana = iSemana + 1
							
							if cint(iSemana) >= cint(idCantidadConsumos)  or iEncuesta = 1 then

								response.write "<td>" & gHogares(0,iHog) & "</td>"
								response.write "<td>" & gHogares(1,iHog) & "</td>"
								response.write "<td>" & gHogares(2,iHog) & "</td>"
								response.write "<td>" & gHogares(3,iHog) & "</td>"
								response.write "<td>" & gHogares(4,iHog) & "</td>"
								response.write "<td>" & gHogares(5,iHog) & "</td>"
								iLen = len(gHogares(7,iHog))
								Cedula = gHogares(6,iHog)
								if iLen = 7 then 
									Cedula = Cedula & "0" & gHogares(7,iHog)
								else
									Cedula = Cedula & gHogares(7,iHog)
								end if
								response.write "<td>" & Cedula & "</td>"
								response.write "<td>" & gHogares(8,iHog) & "</td>"
								
								
								
								response.write "<td>" & iSemana1 & "-" & iSemana1Reg & "</td>"
								response.write "<td>" & iSemana2 & "-" & iSemana2Reg &"</td>"
								response.write "<td>" & iSemana3 & "-" & iSemana3Reg &"</td>"
								response.write "<td>" & iSemana4 & "-" & iSemana4Reg &"</td>"
								response.write "<td>" & iEncuesta & "</td>"
								
								if cint(iSemana) >= cint(idCantidadConsumos) then
									response.write "<td>Si</td>"
								else
									response.write "<td>No</td>"
								end if
								
								if iEncuesta = 1 then
									response.write "<td>Si</td>"
								else
									response.write "<td>No</td>"
								end if
								
								
							end if
								
						response.write "</tr>"
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