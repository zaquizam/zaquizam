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
			var sBus = document.getElementById("Excel").value
			window.open('Sys_mUsuarioExcel.asp?bus='+sBus,'_blank');
		}	
		//**Fin Generar PDF

		function MostrarDetalle(sVariables1,sVariables2,sVariables3,sVariables4){
			debugger;
			TipoConsumo = sVariables1;
			Area = sVariables2;
			Estado = sVariables3;
			Semana = sVariables4;
			//alert("Llego a MostrarDetalle");
			//alert("MostrarDetalle TipoConsumo:= "+ TipoConsumo);
			//alert("MostrarDetalle Area:= "+ Area);
			//alert("MostrarDetalle Estado:= "+ Estado);
			//alert("MostrarDetalle Semana:= "+ Semana);
			var stodo = "num=" + TipoConsumo;
			stodo = stodo + "&are=" + Area;
			stodo = stodo + "&est=" + Estado;
			stodo = stodo + "&sem=" + Semana;
			stodo = "g_MostrarRegistraron.asp?" + stodo;
			//alert("MostrarDetalle Todo:= "+ stodo);
			document.getElementById("Programa").value = stodo;
			$.ajax({
					url: stodo,
					type: 'GET',
					cache: false,
					async: false,
					dataType: 'HTML',
					//data: ajax,
					beforeSend: function(){
						//$("#loader").html("<img src='images/ajax_small.gif'> Buscando Imagen!");
						$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
					}
				})
				//Si la consulta se realizo con exito/
				.done(function(data) {
					//debugger;
					console.log(data);
					$('#Reporte').html(data);
					swal("Detalle Generado","Hogares que Reportaron","success");
					//				
				})
				///Si la consulta Fallo/
				.fail(function() {
					alert("Fallo - bDxPF()");
				},'HTML');			
			
		}	

		function MostrarDetalleExcel(sVariables1,sVariables2,sVariables3,sVariables4){
			//debugger;
			TipoConsumo = sVariables1;
			Area = sVariables2;
			Estado = sVariables3;
			Semana = sVariables4;
			//alert("Llego a MostrarDetalleExcel");
			//alert("MostrarDetalleExcel TipoConsumo:= "+ TipoConsumo);
			//alert("MostrarDetalleExcel Area:= "+ Area);
			//alert("MostrarDetalleExcel Estado:= "+ Estado);
			//alert("MostrarDetalleExcel Semana:= "+ Semana);
			var stodo = "num=" + TipoConsumo;
			stodo = stodo + "&are=" + Area;
			stodo = stodo + "&est=" + Estado;
			stodo = stodo + "&sem=" + Semana;
			stodo = "g_MostrarRegistraronExcel.asp?" + stodo;
			//alert("MostrarDetalleExcel Todo:= "+ stodo);
			document.getElementById("Programa").value = stodo;
			window.open(stodo,"_blank");
		}	
		
		function MostrarNoRegistran(sVariables1,sVariables2,sVariables3,sVariables4){
			//debugger;
			//svar = variable.split(variable,",");
			//TipoConsumo = svar[0];
			//Area = svar[1];
			TipoConsumo = sVariables1;
			Area = sVariables2;
			Estado = sVariables3;
			Semana = sVariables4;
			//alert("Llego a MostrarNoRegistran");
			//alert("MostrarDetalle TipoConsumo:= "+ TipoConsumo);
			//alert("MostrarDetalle Area:= "+ Area);
			//alert("MostrarDetalle Estado:= "+ Estado);
			//alert("MostrarDetalle Semana:= "+ Semana);
			var stodo = "num=" + TipoConsumo;
			stodo = stodo + "&are=" + Area;
			stodo = stodo + "&est=" + Estado;
			stodo = stodo + "&sem=" + Semana;
			stodo = "g_MostrarFaltaron.asp?" + stodo;
			//alert("MostrarDetalle Todo:= "+ stodo);
			document.getElementById("Programa").value = stodo;
			$.ajax({
				url:stodo,
				beforeSend: function(objeto){
					$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
					
				},
				success:function(data){
					//debugger;
					$('#loader2').html('');
					console.log(data); 
					$('#Reporte').html(data);
					
					swal("Detalle Generado","Hogares que Faltan","success");
				}
			})
		}	

		function MostrarNoRegistranExcel(sVariables1,sVariables2,sVariables3,sVariables4){
			//debugger;
			//svar = variable.split(variable,",");
			//TipoConsumo = svar[0];
			//Area = svar[1];
			TipoConsumo = sVariables1;
			Area = sVariables2;
			Estado = sVariables3;
			Semana = sVariables4;
			//alert("Llego a MostrarNoRegistran");
			//alert("MostrarDetalle TipoConsumo:= "+ TipoConsumo);
			//alert("MostrarDetalle Area:= "+ Area);
			//alert("MostrarDetalle Estado:= "+ Estado);
			//alert("MostrarDetalle Semana:= "+ Semana);
			var stodo = "num=" + TipoConsumo;
			stodo = stodo + "&are=" + Area;
			stodo = stodo + "&est=" + Estado;
			stodo = stodo + "&sem=" + Semana;  
			stodo = "g_MostrarFaltaronExcel.asp?" + stodo;
			//alert("MostrarDetalle Todo:= "+ stodo);
			document.getElementById("Programa").value = stodo;
			window.open(stodo,"_blank");
		}	
		
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
	<input type="hidden" name="Programa" id="Programa" align="right" size=50>
	<div id="DivBuscarInformación">
		<div class="ex1">
			<table class="w3-table w3-striped w3-bordered w3-card-4 w3-small" style="width:1000px; margin-left:auto; margin-right:auto;margin-top:10px ">
				<thead>
					<tr class="w3-blue">
						<th>Tipo de Consumo</th>
						<th># Hogares que Reportaron</th>
						<th>Variacion Vs Semana Anterior</th>
						<th># Hogares que Faltan</th>
						<th>% de Cumplimiento</th>
						<th></th>
						<th></th>
						<th></th>
					</tr>
				</thead>
				<%
				Response.write "<tr>"
					'Tipos de Consumo 
					sql = ""
					sql = sql & " SELECT "
					sql = sql & " Id_TipoConsumo, "
					sql = sql & " TipoConsumo "
					sql = sql & " FROM "
					sql = sql & " PH_TipoConsumo "
					sql = sql & " Where "
					sql = sql & " Ind_Activo = 1 "
					sql = sql & " ORDER BY "
					sql = sql & " Id_TipoConsumo "
					'response.write "<br>36 sql:=" & sql
					'response.end
					rsx1.Open sql ,conexion
					if rsx1.eof then
						rsx1.close
					else 
						gDatosSol1 = rsx1.GetRows
						rsx1.close
					end if
					for iReg = 0 to ubound(gDatosSol1,2)
						Response.write "<tr>" 
							idTipoConsumo = gDatosSol1(0,iReg)
							TipoConsumo = gDatosSol1(1,iReg)
							Response.write "<td>" & TipoConsumo & "</td>"
							'# Hogares que Reportaron
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " PH_Consumo.Id_Hogar "
							sql = sql & " FROM ((PH_Consumo INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN PH_GAreaEstado ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area "
							sql = sql & " WHERE "
							sql = sql & " PH_Consumo.Id_Semana = " & idSemana
							sql = sql & " and PH_PanelHogar.Ind_activo = 1 "
							sql = sql & " AND PH_Consumo.id_TipoConsumo = " & idTipoConsumo
							if ed_sPar(1,0) <> "0" and ed_sPar(1,0) <> "" and ed_sPar(1,0) <> "Seleccionar" then
								sql = sql & " AND PH_GAreaEstado.Id_Area = " & idArea
							end if
							if ed_sPar(2,0) <> "" and ed_sPar(2,0) <> "Seleccionar" then
								sql = sql & " AND  PH_PanelHogar.Id_Estado = " & idEstado
							end if
							sql = sql & " GROUP BY "
							sql = sql & " PH_Consumo.Id_Hogar "
							sql = sql & " HAVING PH_Consumo.Id_Hogar > 1 "
							'response.write "<br>36 sql:=" & sql
							'response.end
							rsx2.Open sql ,conexion
							if rsx2.eof then
								rsx2.close
								iTotal1 = 0
							else
								gDatosSol2 = rsx2.GetRows
								rsx2.close
								iTotal1 = ubound(gDatosSol2,2) + 1
							end if
							Response.write "<td>" 
								Response.write iTotal1 
							Response.write "</td>"
							'Variacion Vs Semana Anterior
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " PH_Consumo.Id_Hogar "
							sql = sql & " FROM ((PH_Consumo INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN PH_GAreaEstado ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area "
							sql = sql & " WHERE "
							sql = sql & " PH_Consumo.Id_Semana = " & idSemana - 1
							sql = sql & " and PH_PanelHogar.Ind_activo = 1 "
							sql = sql & " AND PH_Consumo.id_TipoConsumo = " & idTipoConsumo
							if ed_sPar(1,0) <> "0" and ed_sPar(1,0) <> "" and ed_sPar(1,0) <> "Seleccionar" then
								sql = sql & " AND PH_GAreaEstado.Id_Area = " & idArea
							end if
							if ed_sPar(2,0) <> "" and ed_sPar(2,0) <> "Seleccionar" then
								sql = sql & " AND  PH_PanelHogar.Id_Estado = " & idEstado
							end if
							sql = sql & " GROUP BY "
							sql = sql & " PH_Consumo.Id_Hogar "
							sql = sql & " HAVING PH_Consumo.Id_Hogar > 1 "
							'response.write "<br>36 sql:=" & sql
							'response.end
							rsx2.Open sql ,conexion
							if rsx2.eof then
								rsx2.close
								iTotal2 = 0
							else
								gDatosSol2 = rsx2.GetRows
								rsx2.close
								iTotal2 = ubound(gDatosSol2,2) + 1
							end if
							Response.write "<td>" 
								Response.write iTotal2
								if iTotal2 = 0 then iTotal2 = 1
								iTotal2T = ((iTotal1 - iTotal2)/iTotal2)*100
								Response.write " Var:= " &  formatnumber(iTotal2T) & "%"
							Response.write "</td>"
							'# Hogares que Faltan
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " PH_PanelHogar.Id_PanelHogar "
							sql = sql & " FROM PH_PanelHogar INNER JOIN (PH_GAreaEstado INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado "
							sql = sql & " WHERE "
							sql = sql & " PH_PanelHogar.Ind_Activo = 1 "
							if ed_sPar(1,0) <> "0" and ed_sPar(1,0) <> "" and ed_sPar(1,0) <> "Seleccionar" then
								sql = sql & " AND PH_GAreaEstado.Id_Area = " & idArea
							end if
							if ed_sPar(2,0) <> "" and ed_sPar(2,0) <> "Seleccionar" then
								sql = sql & " AND  PH_PanelHogar.Id_Estado = " & idEstado
							end if
							sql = sql & " GROUP BY "
							sql = sql & " PH_PanelHogar.Id_PanelHogar "
							sql = sql & " HAVING PH_PanelHogar.Id_PanelHogar > 1 "
							'response.write "<br>36 sql:=" & sql
							'response.end
							rsx2.Open sql ,conexion
							if rsx2.eof then
								rsx2.close
								iTotal3 = 0
							else
								gDatosSol2 = rsx2.GetRows
								rsx2.close
								iTotal3 = ubound(gDatosSol2,2) + 1
							end if
							Response.write "<td>" 
								iTotal3T = iTotal3 - iTotal1
								Response.write iTotal3T
							Response.write "</td>"
						
							Response.write "<td>" 
								iTotal4T = (iTotal1 * 100) / iTotal3
								Response.write formatnumber(iTotal4T) & "%"
							Response.write "</td>"
							Response.write "<td>"
								sVariables1 = idTipoConsumo 
								sVariables2 = idArea
								sVariables3 = idEstado
								sVariables4 = idSemana
								%>
								<img src="images/CelularCB.jpg"  style="margin-left:0px;" title="Hogares que Reportaron" alt="Detalle" width="70px" onclick="MostrarDetalle(<%=sVariables1%>,<%=sVariables2%>,<%=sVariables3%>,<%=sVariables4%>)">
								<%
								%>
								<img src="images/Excel.png"  style="margin-left:0px;" title="Hogares que Reportaron" alt="Detalle" width="70px" onclick="MostrarDetalleExcel(<%=sVariables1%>,<%=sVariables2%>,<%=sVariables3%>,<%=sVariables4%>)">
								<%
							Response.write "</td>"
							Response.write "<td>"
								sVariables1 = idTipoConsumo 
								sVariables2 = idArea 
								sVariables3 = idEstado
								sVariables4 = idSemana
								%>
								<img src="images/CelularCBNo.jpg"  style="margin-left:0px;" title="Hogares que Faltan" alt="No Registran" width="70px" onclick="MostrarNoRegistran(<%=sVariables1%>,<%=sVariables2%>,<%=sVariables3%>,<%=sVariables4%>)">
								<%
								%>
								<img src="images/Excel.png"  style="margin-left:0px;" title="Hogares que Faltan" alt="No Registran" width="70px" onclick="MostrarNoRegistranExcel(<%=sVariables1%>,<%=sVariables2%>,<%=sVariables3%>,<%=sVariables4%>)">
								<%
							Response.write "</td>"
							Response.write "<td>"
								sVariables1 = idTipoConsumo 
								sVariables2 = idArea
								sVariables3 = idEstado
								sVariables4 = idSemana
							Response.write "</td>"
						Response.write "</tr>"
					next 
				
				Response.write "</tr>"
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