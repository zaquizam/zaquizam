<!DOCTYPE HTML>
<html >
<head>
	<title>Tipo Consumo x Dia</title>
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
	
	dim vHogar(10000)
	
	
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
	
End Sub
	
%>
	<script>

		//**Inicio Generar PDF
		function GenerarExcel(){
			//alert("Bus:= "+ document.getElementById("Bus").value );
			//alert("Buscar:= "+ document.getElementById("Excel").value );
			var sArea = document.getElementById("Area").value
			var sEstado = document.getElementById("Estado").value
			var sSemana = document.getElementById("Semana").value
			var sTodo = "are=" + sArea 
			sTodo = sTodo + "&est=" + sEstado
			sTodo = sTodo + "&sem=" + sSemana
			window.open('PH_rTipConDiaExcel.asp?'+sTodo,'_blank');
			//window.open("ph_rHomePantryPenCatExcel.asp","_blank");
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
	'response.write "<br> Combo1:=" & ed_sPar(1,0) & "==>" & idArea
	'response.write "<br> Combo2:=" & ed_sPar(2,0) & "==>" & idEstado
	'response.write "<br> Combo3:=" & ed_sPar(3,0) & "==>" & idSemana
	'hidden 
	%>
	<input type="hidden" name="Area" id="Area" align="right" size=5 value='<%=idArea%>'>
	<input type="hidden" name="Estado" id="Estado" align="right" size=5 value='<%=idEstado%>'>
	<input type="hidden" name="Semana" id="Semana" align="right" size=5 value='<%=idSemana%>'>
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
						<th>Dia</th>
						<th>Area</th>
						<th>Estado</th>
						<th>Tipo de Consumo</th>
						<th># Hogares que Reportaron</th>
					</tr>
				</thead>
				<%
				'Response.write "<br>176 idSemana:" & idSemana
				'Response.write "<br>176 idArea:" & idArea
				'Response.write "<br>176 idEstado:" & idEstado
				if idArea <>0 Then
					sql = ""
					sql = sql & " SELECT "
					sql = sql & " PH_Consumo.Fecha_Creacion, "
					sql = sql & " PH_GAreaEstado.Id_Area, "
					sql = sql & " PH_GArea.Area, "
					sql = sql & " PH_PanelHogar.Id_Estado, "
					sql = sql & " ss_Estado.Estado, "
					sql = sql & " PH_Consumo.id_TipoConsumo, "
					sql = sql & " PH_TipoConsumo.TipoConsumo "
					sql = sql & " FROM ((((PH_Consumo INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN PH_GAreaEstado ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) INNER JOIN PH_TipoConsumo ON PH_Consumo.id_TipoConsumo = PH_TipoConsumo.Id_TipoConsumo) INNER JOIN ss_Estado ON PH_GAreaEstado.Id_Estado = ss_Estado.Id_Estado "
					sql = sql & " WHERE "
					sql = sql & " PH_Consumo.Id_Semana = " & idSemana
					sql = sql & " AND PH_Consumo.Status_registro='G' "
					sql = sql & " AND PH_Consumo.Id_Hogar > 1 "
					sql = sql & " GROUP BY "
					sql = sql & " PH_Consumo.Fecha_Creacion, "
					sql = sql & " PH_GAreaEstado.Id_Area, "
					sql = sql & " PH_GArea.Area, "
					sql = sql & " PH_PanelHogar.Id_Estado, "
					sql = sql & " ss_Estado.Estado, "
					sql = sql & " PH_Consumo.id_TipoConsumo, "
					sql = sql & " PH_TipoConsumo.TipoConsumo "
					sql = sql & " HAVING "
					sql = sql & " PH_GAreaEstado.Id_Area = " & idArea
					if idEstado <> 0 then 
						sql = sql & " AND PH_PanelHogar.Id_Estado = " & idEstado
					end if
					sql = sql & " ORDER BY "
					sql = sql & " PH_Consumo.Fecha_Creacion DESC , "
					sql = sql & " PH_GArea.Area, "
					sql = sql & " ss_Estado.Estado "
					'response.write "<br>232 sql:= " & sql
					'response.end
					rsx1.Open sql ,conexion
					if rsx1.eof then
						rsx1.close
					else 
						gDatosSol1 = rsx1.GetRows
						rsx1.close
					end if
					for iReg = 0 to ubound(gDatosSol1,2)
						Response.flush
						response.write "<tr>"
							idTipoConsumo = gDatosSol1(5,iReg)
							sFecha = gDatosSol1(0,iReg)
							Dia = mid(sFecha,9,2)
							Mes = mid(sFecha,6,2)							
							Ano = mid(sFecha,1,4)
							nFecha = Dia & "/" & Mes & "/" & Ano
							response.write "<td>" & nFecha & "</td>"
							response.write "<td>" & gDatosSol1(2,iReg) & "</td>"
							response.write "<td>" & gDatosSol1(4,iReg) & "</td>"
							response.write "<td>" & gDatosSol1(6,iReg) & "</td>"
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " PH_Consumo.Id_Hogar "
							sql = sql & " FROM ((((PH_Consumo INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN PH_GAreaEstado ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) INNER JOIN PH_TipoConsumo ON PH_Consumo.id_TipoConsumo = PH_TipoConsumo.Id_TipoConsumo) INNER JOIN ss_Estado ON PH_GAreaEstado.Id_Estado = ss_Estado.Id_Estado "
							sql = sql & " WHERE "
							sql = sql & " PH_Consumo.Fecha_Creacion = '" & sFecha & "'"
							sql = sql & " AND PH_GAreaEstado.Id_Area = " & idArea 
							if idEstado <> 0 then
								sql = sql & " AND PH_PanelHogar.Id_Estado = " & idEstado
							end if
							sql = sql & " AND PH_Consumo.id_TipoConsumo = " & idTipoConsumo
							sql = sql & " AND PH_Consumo.Id_Semana = " & idSemana
							sql = sql & " AND PH_Consumo.Status_registro = 'G' "
							sql = sql & " AND PH_Consumo.Id_Hogar > 1 "
							sql = sql & " GROUP BY "
							sql = sql & " PH_Consumo.Id_Hogar "
							'response.write "<br>232 sql:= " & sql
							rsx2.Open sql ,conexion
							if rsx2.eof then
								rsx2.close
							else 
								gDatosSol2 = rsx2.GetRows
								rsx2.close
								Total = 0
								for iReg1 = 0 to ubound(gDatosSol2,2)
									idHogar = gDatosSol2(0,iReg1)
									if vHogar(idHogar) = "" then
										Total = Total + 1
										vHogar(idHogar) = 1
									end if
								next
								
							end if
							response.write "<td>" & Total & "</td>"
						response.write "</tr>"
						'Response.flush
					next 
					
					
					Response.write "<tr>"
				
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