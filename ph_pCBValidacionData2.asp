<!Doctype html>
<html >
<head>
	<title>Validacion Data 2</title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
	<link rel="icon" href="favicon.ico" type="image/x-icon"> 
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
	<link href="css/bootstrap.min.css" rel="stylesheet" type="text/css" />	
	<script type="text/javascript" src="js/sweetalert.min.js"></script>	
</head>
<body topmargin="0">
<!--#include file="estiloscss.asp"-->
<!--#include file="encabezado.asp"-->
<!--#include file="nn_subN.asp"-->
<!--#include file="in_DataEN1.asp"-->
<!--#include file="ph_pCBValidacionDataModal.asp"-->
<script>
	//**Inicio Actualizar Promedio
	function ActualizarPromedio(num){
		//debugger; 
		swal({
                title: "Desea Actualizar el Precio con el Promedio ?",
                text: "",
                type: "warning",
                showCancelButton: true,
                confirmButtonClass: "btn-primary",
                confirmButtonText: "Si",
                cancelButtonText: "No",
                closeOnConfirm: false,
                showLoaderOnConfirm: true
            },
            function() {
                //
				promedio =document.getElementById("Promedio").value;
				//alert("Llego Actualizar Promedio:=" + num);
				//alert("Promedio:=" + promedio);
				//return;
				var stodo = "num=" + num;
				stodo = stodo + "&pro=" + promedio;

				$.ajax({
					url:'g_ActualizarPromedio.asp?'+stodo,
					beforeSend: function(objeto){
						$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
						
					},
					success:function(data){
						//debugger;
						$('#loader2').html('');
						console.log(data); 
						swal("Actualizar Precio con Promedio","Registro Actualizado","success");
						window.location.reload();
					}
				})
                
                
            });
	}	
	//**Fin Actualizar Promedio

	//**Inicio Actualizar Moda
	function ActualizarModa(num){
		//alert("Llego Actualizar Moda:=" + num);
		//debugger; 
		swal({
                title: "Desea Actualizar el Precio con la Moda ?",
                text: "",
                type: "warning",
                showCancelButton: true,
                confirmButtonClass: "btn-primary",
                confirmButtonText: "Si",
                cancelButtonText: "No",
                closeOnConfirm: false,
                showLoaderOnConfirm: true
            },
            function() {
                //
				moda =document.getElementById("Moda").value;
				//alert("Llego Actualizar Moda:=" + num);
				//alert("Moda:=" + moda);
				//return;
				var stodo = "num=" + num;
				stodo = stodo + "&mod=" + moda;

				$.ajax({
					url:'g_ActualizarModa.asp?'+stodo,
					beforeSend: function(objeto){
						$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
						
					},
					success:function(data){
						//debugger;
						$('#loader2').html('');
						console.log(data); 
						swal("Actualizar Precio con Moda","Registro Actualizado","success");
						window.location.reload();
					}
				})
            });
	}	
	//**Fin Actualizar Moda

	//**Inicio Actualizar Manual
	function ActualizarManual(num){
		//debugger;
		manual = document.getElementById(num).value;
		//alert(manual);
		//return;
		
		if(manual == "")
		{
			swal("Error","No ha Incluido Ningun Valor","error");
			//window.open(document.getElementById("Programa").value,"_parent");
			return;
		}
		swal({
                title: "Desea Actualizar el Precio con uno Manual ?",
                text: "",
                type: "warning",
                showCancelButton: true,
                confirmButtonClass: "btn-primary",
                confirmButtonText: "Si",
                cancelButtonText: "No",
                closeOnConfirm: false,
                showLoaderOnConfirm: true
            },
            function() {
                //
				//alert("Llego Actualizar Manual:=" + num);
				//alert("Manual:=" + manual);
				//return;
				var stodo = "num=" + num;
				stodo = stodo + "&man=" + manual;

				$.ajax({
					url:'g_ActualizarManual.asp?'+stodo,
					beforeSend: function(objeto){
						$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
						
					},
					success:function(data){
						//debugger;
						$('#loader2').html('');
						console.log(data); 
						swal("Actualizar Precio Manual","Registro Actualizado","success");
						window.location.reload();
					}
				})
            });
	}	
	//**Fin Actualizar Manual
	
	//**Inicio Actualizar Manual Cantidad
	function ActualizarManualCantidad(num){
		debugger;
		manualcantidad = document.getElementById(num).value;
		//alert(manual);
		//return;
		
		if(manualcantidad == "")
		{
			swal("Error","No ha Incluido Ningun Valor en Cantidad","error");
			//window.open(document.getElementById("Programa").value,"_parent");
			return;
		}
		swal({
                title: "Desea Actualizar Cantidad con una Manual ?",
                text: "",
                type: "warning",
                showCancelButton: true,
                confirmButtonClass: "btn-primary",
                confirmButtonText: "Si",
                cancelButtonText: "No",
                closeOnConfirm: false,
                showLoaderOnConfirm: true
            },
            function() {
                //
				//alert("Llego Actualizar Manual:=" + num);
				//alert("Manual:=" + manualcantidad);
				//return;
				var stodo = "num=" + num;
				stodo = stodo + "&man=" + manualcantidad;

				$.ajax({
					url:'g_ActualizarManualCantidad.asp?'+stodo,
					beforeSend: function(objeto){
						$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
						
					},
					success:function(data){
						//debugger;
						$('#loader2').html('');
						console.log(data); 
						swal("Actualizado Cantidad Manual","Registro Actualizado","success");
						window.location.reload();
					}
				})
            });
	}	
	//**Fin Actualizar Manual

</script>   
<%
  
'==========================================================================================
' Variables y Constantes
'==========================================================================================

    Apertura
	
	dim Min
	dim Max
	dim LimiteInferior 
	dim LimiteSuperior 
	dim Moda
	
	dim gDatosSol
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
	'if ed_sPar(1,0) <> "" and ed_sPar(1,0) <> "Seleccionar" then ed_sPar(1,0) = 0

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " IdSemana, "
	sql = sql & " Semana "
	sql = sql & " FROM ss_Semana "
	sql = sql & " Where "
	sql = sql & " IdSemana > 14 "
	'sql = sql & " and IdSemana < 47 "
	sql = sql & " Order By "
	sql = sql & " IdSemana Desc "
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(1,0)="Semana"
    ed_sCombo(1,1)=sql 
    'ed_sCombo(1,2)="Seleccionar"
	ed_sCombo(1,2)=""
 
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Categoria, "
	sql = sql & " Categoria "
	sql = sql & " FROM "
	sql = sql & " PH_DataValidacion "
	if ed_sPar(1,0) <> "" and ed_sPar(1,0) <> "Seleccionar" then
		sql = sql & " WHERE "
		sql = sql & " Id_Semana = " & ed_sPar(1,0)
	END IF
	sql = sql & " GROUP BY "
	sql = sql & " Id_Categoria, "
	sql = sql & " Categoria "
	sql = sql & " HAVING "
	sql = sql & " Id_Categoria Is Not Null "
	sql = sql & " and Categoria Is Not Null "
	sql = sql & " Order by Categoria "
	'response.write "<br>372 Combo2:=" & sql
    ed_sCombo(2,0)="Categoria"
    ed_sCombo(2,1)=sql 
    ed_sCombo(2,2)="Seleccionar"

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " PH_CB_Producto.Id_Producto, "
	sql = sql & " PH_DataValidacion.Producto "
	sql = sql & " FROM PH_DataValidacion INNER JOIN PH_CB_Producto ON PH_DataValidacion.CodigoBarra = PH_CB_Producto.CodigoBarra "
	if ed_sPar(1,0) <> "" and ed_sPar(1,0) <> "Seleccionar" and ed_sPar(2,0) <> "" and ed_sPar(2,0) <> "Seleccionar" then
		sql = sql & " WHERE "
		'sql = sql & " PH_DataValidacion.Id_Categoria = " & ed_sPar(2,0)
		sql = sql & " PH_CB_Producto.Id_Categoria = " & ed_sPar(2,0)
		sql = sql & " and PH_DataValidacion.Id_Semana = " & ed_sPar(1,0)
	else
		ed_iCombo = 2
	END IF
	sql = sql & " GROUP BY "
	sql = sql & " PH_CB_Producto.Id_Producto, "
	sql = sql & " PH_DataValidacion.Producto "
	sql = sql & " HAVING (((PH_DataValidacion.Producto) Is Not Null)) "
	'sql = sql & " and PH_CB_Producto.Id_Producto < 42826 "
	sql = sql & " ORDER BY "
	sql = sql & " PH_DataValidacion.Producto "
	'response.write "<br>372 Combo3:=" & sql
    'response.end
	ed_sCombo(3,0)="Producto"
    ed_sCombo(3,1)=sql 
	ed_sCombo(3,2)="Seleccionar"
	
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
	'if ed_sPar(1,0) = "" or ed_sPar(1,0) = "Seleccionar" then ed_sPar(1,0) = 17
    Combos
	'response.end
	
%>		
	<br>
	<div style="width:98%">
	</div></center>
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
	'response.end
	idSemana = ed_sPar(1,0)
	idCategoria = ed_sPar(2,0)
	idProducto = ed_sPar(3,0)
	'response.write "<br> Combo1:=Semana"  & "==>" & idSemana
	'response.write "<br> Combo2:=Categoria" & "==>" & idCategoria
	'response.write "<br> Combo3:=Producto" & "==>" & idProducto
	'hidden 
	if idSemana <> "Seleccionar" and  idCategoria <> "Seleccionar" and  idProducto <> "Seleccionar" then

		sql = ""
		sql = sql & " SELECT "
		sql = sql & " Min_Por_Val, "
		sql = sql & " Max_Por_Val "
		sql = sql & " FROM "
		sql = sql & " PH_CB_Categoria "
		sql = sql & " WHERE "
		sql = sql & " PH_CB_Categoria.id_Categoria = " & idCategoria
		'response.write "<br>36 sql:=" & sql
		'response.end
		rsx2.Open sql ,conexion
		if rsx2.eof then
			rsx2.close
		else
			gDatosSol2 = rsx2.GetRows
			rsx2.close
			Min = gDatosSol2(0,0)
			Max = gDatosSol2(1,0)
		end if

		sql = ""
		sql = sql & " SELECT "
		sql = sql & " PH_Consumo_Detalle_Productos.Precio_producto*Tasa_de_Cambio AS Expr1, "
		sql = sql & " Count(PH_Consumo.Id_Consumo) AS CuentaDeId_Consumo "
		sql = sql & " FROM ((PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON (PH_Consumo.Id_Hogar = PH_Consumo_Detalle_Productos.Id_Hogar) AND (PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo)) INNER JOIN PH_CB_Producto ON PH_Consumo_Detalle_Productos.Numero_codigo_barras = PH_CB_Producto.CodigoBarra) INNER JOIN PH_Moneda ON PH_Consumo_Detalle_Productos.id_Moneda = PH_Moneda.Id_Moneda "
		sql = sql & " WHERE "
		sql = sql & " PH_Consumo.Id_Semana =  " & idSemana
		sql = sql & " AND PH_CB_Producto.Id_Categoria = " & idCategoria 
		sql = sql & " AND PH_CB_Producto.Id_Producto = " & idProducto
		sql = sql & " GROUP BY "
		sql = sql & " PH_Consumo_Detalle_Productos.Precio_producto*Tasa_de_Cambio "
		sql = sql & " ORDER BY "
		sql = sql & " Count(PH_Consumo.Id_Consumo) DESC " 
		rsx2.Open sql ,conexion
		if rsx2.eof then
			rsx2.close
		else
			gDatosSol2 = rsx2.GetRows
			rsx2.close
			Moda = gDatosSol2(0,0)
		end if

		sql = ""
		sql = sql & " SELECT "
		sql = sql & " PH_Consumo_Detalle_Productos.Id_Consumo_Detalle_Productos, "
		sql = sql & " PH_Consumo_Detalle_Productos.Id_Hogar, "
		sql = sql & " PH_CB_Producto.CodigoBarra, "
		sql = sql & " PH_CB_Producto.Producto, "
		sql = sql & " PH_Consumo_Detalle_Productos.Cantidad, "
		sql = sql & " PH_Consumo_Detalle_Productos.Precio_producto*Tasa_de_Cambio, "
		sql = sql & " PH_GArea.Area, "
		sql = sql & " ss_Estado.Estado, "
		sql = sql & " PH_Consumo.Fecha_Creacion "
		'sql = sql & " FROM ((PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON (PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) AND (PH_Consumo.Id_Hogar = PH_Consumo_Detalle_Productos.Id_Hogar)) INNER JOIN PH_CB_Producto ON PH_Consumo_Detalle_Productos.Numero_codigo_barras = PH_CB_Producto.CodigoBarra) INNER JOIN PH_Moneda ON PH_Consumo_Detalle_Productos.id_Moneda = PH_Moneda.Id_Moneda "
		sql = sql & " FROM (((((PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON (PH_Consumo.Id_Hogar = PH_Consumo_Detalle_Productos.Id_Hogar) AND (PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo)) INNER JOIN PH_CB_Producto ON PH_Consumo_Detalle_Productos.Numero_codigo_barras = PH_CB_Producto.CodigoBarra) INNER JOIN PH_Moneda ON PH_Consumo_Detalle_Productos.id_Moneda = PH_Moneda.Id_Moneda) INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN (PH_GAreaEstado INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado "
		sql = sql & " WHERE "
		sql = sql & " PH_Consumo.Id_Semana = " & idSemana
		sql = sql & " AND PH_CB_Producto.Id_Categoria = " & idCategoria
		sql = sql & " AND PH_CB_Producto.Id_Producto = " & idProducto
		sql = sql & " ORDER BY "
		sql = sql & " PH_Consumo_Detalle_Productos.Precio_producto*Tasa_de_Cambio "
		'response.write "<br>36 sql:=" & sql
		'response.end
		rsx1.Open sql ,conexion
		iExiste = 0
		if rsx1.eof then
			rsx1.close
		else
			gDatosSol = rsx1.GetRows
			iExiste = 1
			rsx1.close
		end if
		iTotalPrecio = 0
		iCantidad = 0
		if iExiste = 1 then
			for iReg = 0 to ubound(gDatosSol,2)
				for iRegCol = 0 to 6
					if iRegCol <> 5 then
					else
						iTotalPrecio = iTotalPrecio + cdbl(gDatosSol(iRegCol,iReg))
						iCantidad = iCantidad + 1
					end if
				next
			next
			iCalculo = cdbl(iTotalPrecio)/cint(iCantidad)
			ix = (cdbl(iCalculo) * cdbl(Min)) / 100
			LimiteInferior = iCalculo - ix
			
			ix = (cdbl(iCalculo) * cdbl(Max)) / 100
			LimiteSuperior = iCalculo + ix
			
		end if
		
		sPro=Request.ServerVariables("HTTP_REFERER")
		'response.write "pro:=" & sPro
		'hidden
		'response.write "pro:=" & sPar
		%>
		<input type="hidden" name="Programa"  id="Programa"  value="<%=sPar%>" size=200>
		<input type="hidden" name="id_Semana" id="id_Semana" value="<%=idSemana%>" size=20>
		

		<div id="DivBuscarInformación">
			<div class="ex1">
				<table class="w3-table w3-striped w3-bordered w3-card-4 w3-small" style=" margin-left:auto; margin-right:auto;margin-top:10px ">
					<thead>
						<tr class="w3-blue">
							<th>Id</th>
							<th>IdHogar</th>
							<th>Codigo Barra</th>
							<th>Producto</th>
							<th>Cantidad</th>
							<th>PrecioBs</th>
							<th>Promedio</th>
							<th>Moda</th>
							<th>Manual</th>
							<th>Valor Manual</th>
							<th>Area</th>
							<th>Estado</th>
							<th>Fec. Creacion</th>
							<th>Cantidad</th>
							<th>Valor Manual</th>
							<th>Moneda</th>
						</tr>
					</thead>
					<%
					if iExiste = 1 then
						for iReg = 0 to ubound(gDatosSol,2)
							Mostrar = 0
							Response.write "<tr>"
								for iRegCol = 0 to 5
									iValor = cdbl(gDatosSol(5,iReg))
									if cdbl(iValor) < cdbl(LimiteInferior) or cdbl(iValor) > cdbl(LimiteSuperior) then 
										Response.write "<td style='color:red; font-weight:bold'>" 
										Mostrar = 1
									else
										Response.write "<td>" 
									end if
									if iRegCol <> 5  then
										Response.write  gDatosSol(iRegCol,iReg)
									else
										Response.write  formatnumber(iValor,2)
									end if
									Response.write "</td>" 

								next
								sx=gDatosSol(0,iReg)
								if Mostrar = 1 then 
									Response.write "<td>" 
										%>
										<img src="images/Boton01.png"  style="margin-left:0px;" alt="Agregar" width="32px"' onclick="ActualizarPromedio(<%=sx%>)"/>
										<%
									Response.write "</td>" 
									Response.write "<td>" 
										%>
										<img src="images/Boton02.png"  style="margin-left:0px;" alt="Agregar" width="32px"' onclick="ActualizarModa(<%=sx%>)"/>
										<%
									Response.write "</td>" 
								else
									Response.write "<td>" 
									Response.write "</td>" 
									Response.write "<td>" 
									Response.write "</td>" 
								end if
								Response.write "<td>" 
									%>
									<img src="images/Boton03.png"  style="margin-left:0px;" alt="Agregar" width="32px"' onclick="ActualizarManual(<%=sx%>)"/>
									<%
								Response.write "</td>"
								Response.write "<td>" 
								sx1 = "Manual" & sx 
									%>
									<input type="text" name="<%=sx1%>" id="<%=sx%>" align="right" maxlength=15 size=15 onchange="ActualizarManual(<%=sx%>)"/>
									<%
								Response.write "</td>"
								Response.write "<td>" 
									Response.write gDatosSol(6,iReg)
								Response.write "</td>"
								Response.write "<td>" 
									Response.write gDatosSol(7,iReg)
								Response.write "</td>"
								Response.write "<td>" 
									Response.write gDatosSol(8,iReg)
								Response.write "</td>"
								sx = sx * -1
								Response.write "<td>" 
									%>
									<img src="images/Boton03.png"  style="margin-left:0px;" alt="Agregar" width="32px"' onclick="ActualizarManualCantidad(<%=sx%>)"/>
									<%
								Response.write "</td>"
								Response.write "<td>" 
								
								'Response.write (
								sx1 = "Manual" & sx 
									%>
									<input type="text" name="<%=sx1%>" id="<%=sx%>" align="right" maxlength=15 size=15 onchange="ActualizarManualCantidad(<%=sx%>)"/>
									<%
								Response.write "</td>"
								
								Response.write "<td>" 
								
								'Moneda
								'Response.write (
								sx1 = "Manual" & sx 
									%>									
									<img src="images/Boton04.png"  style="margin-left:0px;" alt="Agregar" width="32px"' onclick="ActualizarMoneda(<%=sx%>)"/>
									<%
								Response.write "</td>"							
							Response.write "</tr>"
						next
					end if
					%>
				</table>
				<table class="w3-table w3-striped w3-bordered w3-card-4 w3-small" style="width:1200px; margin-left:auto; margin-right:auto;margin-top:10px ">
					<thead>
						<tr class="w3-blue">
							<th>%MinVal</th>
							<th>MinVal</th>
							<th>%MaxVal</th>
							<th>MaxVal</th>
							<th>Promedio</th>
							<th>Moda</th>
							<th>#</th>
						</tr>
					</thead>
					<%
					Response.write "<tr>"

						Response.write "<td>"
							'Response.write formatnumber(Min,2)
							%>
								<input type="text" name="PorcentajeMinimo" id="PorcentajeMinimo" align="right" disabled maxlength=5 size=5 value=<%=formatnumber(Min,2)%>>
							<%
						Response.write "</td>"

						Response.write "<td>"
							'Response.write formatnumber(LimiteInferior,2)
							%>
							<input type="text" name="LimiteInferior" id="LimiteInferior" align="right" disabled maxlength=15 size=15 value=<%=formatnumber(LimiteInferior,2)%>>
							<%
						Response.write "</td>"

						Response.write "<td>"
							'Response.write formatnumber(Max,2)
							%>
							<input type="text" name="PorcentajeMaximo" id="PorcentajeMaximo" align="right" disabled maxlength=5 size=5 value=<%=formatnumber(Max,2)%>>
							<%
						Response.write "</td>"

						Response.write "<td>"
							'Response.write formatnumber(LimiteSuperior,2)
							%>
							<input type="text" name="LimiteSuperior" id="LimiteSuperior" align="right" disabled maxlength=15 size=15 value=<%=formatnumber(LimiteSuperior,2)%>>
							<%
						Response.write "</td>"

						Response.write "<td>"
							'Response.write formatnumber(iCalculo,2)
							%>
							<input type="text" name="PromedioVer" id="PromedioVer" align="right" disabled maxlength=15 size=15 value=<%=formatnumber(iCalculo,2)%>>
							<input type="hidden" name="Promedio" id="Promedio" align="right" disabled maxlength=15 size=15 value=<%=iCalculo%>>
							<%
						Response.write "</td>"

						Response.write "<td>"
							'Response.write formatnumber(Moda,2)
							%>
							<input type="text" name="ModaVer" id="ModaVer" align="right" disabled maxlength=15 size=15 value=<%=formatnumber(Moda,2)%>>
							<input type="hidden" name="Moda" id="Moda" align="right" disabled maxlength=15 size=15 value=<%=Moda%>>
							<%
						Response.write "</td>"

						Response.write "<td>"
							'Response.write formatnumber(iCantidad,0)
							%>
							<input type="text" name="Cantidad" id="Cantidad" align="right" disabled maxlength=5 size=5 value=<%=formatnumber(iCantidad,0)%>>
							<%
						Response.write "</td>"
					Response.write "</tr>"
					%>
				</table>
			</div>
		</div>
		</br>
		</br>
		</br>
		</br>
		</br>
		<%
	end if
	
	conexion.close
	%>

<style>
@keyframes showSweetAlert {
  0% {
    transform: scale(0.7);
  }
  45% {
    transform: scale(1.05);
  }
  80% {
    transform: scale(0.95);
  }
  100% {
    transform: scale(1);
  }
}
@keyframes hideSweetAlert {
  0% {
    transform: scale(1);
  }
  100% {
    transform: scale(0.5);
  }
}
@keyframes slideFromTop {
  0% {
    top: 0%;
  }
  100% {
    top: 50%;
  }
}
@keyframes slideToTop {
  0% {
    top: 50%;
  }
  100% {
    top: 0%;
  }
}
@keyframes slideFromBottom {
  0% {
    top: 70%;
  }
  100% {
    top: 50%;
  }
}
@keyframes slideToBottom {
  0% {
    top: 50%;
  }
  100% {
    top: 70%;
  }
}
.showSweetAlert {
  animation: showSweetAlert 0.3s;
}
.showSweetAlert[data-animation=none] {
  animation: none;
}
.showSweetAlert[data-animation=slide-from-top] {
  animation: slideFromTop 0.3s;
}
.showSweetAlert[data-animation=slide-from-bottom] {
  animation: slideFromBottom 0.3s;
}
.hideSweetAlert {
  animation: hideSweetAlert 0.3s;
}
.hideSweetAlert[data-animation=none] {
  animation: none;
}
.hideSweetAlert[data-animation=slide-from-top] {
  animation: slideToTop 0.3s;
}
.hideSweetAlert[data-animation=slide-from-bottom] {
  animation: slideToBottom 0.3s;
}
@keyframes animateSuccessTip {
  0% {
    width: 0;
    left: 1px;
    top: 19px;
  }
  54% {
    width: 0;
    left: 1px;
    top: 19px;
  }
  70% {
    width: 50px;
    left: -8px;
    top: 37px;
  }
  84% {
    width: 17px;
    left: 21px;
    top: 48px;
  }
  100% {
    width: 25px;
    left: 14px;
    top: 45px;
  }
}
@keyframes animateSuccessLong {
  0% {
    width: 0;
    right: 46px;
    top: 54px;
  }
  65% {
    width: 0;
    right: 46px;
    top: 54px;
  }
  84% {
    width: 55px;
    right: 0px;
    top: 35px;
  }
  100% {
    width: 47px;
    right: 8px;
    top: 38px;
  }
}
@keyframes rotatePlaceholder {
  0% {
    transform: rotate(-45deg);
  }
  5% {
    transform: rotate(-45deg);
  }
  12% {
    transform: rotate(-405deg);
  }
  100% {
    transform: rotate(-405deg);
  }
}
.animateSuccessTip {
  animation: animateSuccessTip 0.75s;
}
.animateSuccessLong {
  animation: animateSuccessLong 0.75s;
}
.sa-icon.sa-success.animate::after {
  animation: rotatePlaceholder 4.25s ease-in;
}
@keyframes animateErrorIcon {
  0% {
    transform: rotateX(100deg);
    opacity: 0;
  }
  100% {
    transform: rotateX(0deg);
    opacity: 1;
  }
}
.animateErrorIcon {
  animation: animateErrorIcon 0.5s;
}
@keyframes animateXMark {
  0% {
    transform: scale(0.4);
    margin-top: 26px;
    opacity: 0;
  }
  50% {
    transform: scale(0.4);
    margin-top: 26px;
    opacity: 0;
  }
  80% {
    transform: scale(1.15);
    margin-top: -6px;
  }
  100% {
    transform: scale(1);
    margin-top: 0;
    opacity: 1;
  }
}
.animateXMark {
  animation: animateXMark 0.5s;
}
@keyframes pulseWarning {
  0% {
    border-color: #F8D486;
  }
  100% {
    border-color: #F8BB86;
  }
}
.pulseWarning {
  animation: pulseWarning 0.75s infinite alternate;
}
@keyframes pulseWarningIns {
  0% {
    background-color: #F8D486;
  }
  100% {
    background-color: #F8BB86;
  }
}
.pulseWarningIns {
  animation: pulseWarningIns 0.75s infinite alternate;
}
@keyframes rotate-loading {
  0% {
    transform: rotate(0deg);
  }
  100% {
    transform: rotate(360deg);
  }
}
body.stop-scrolling {
  height: 100%;
  overflow: hidden;
}
.sweet-overlay {
  background-color: rgba(0, 0, 0, 0.4);
  position: fixed;
  left: 0;
  right: 0;
  top: 0;
  bottom: 0;
  display: none;
  z-index: 1040;
}
.sweet-alert {
  background-color: #ffffff;
  width: 478px;
  padding: 17px;
  border-radius: 5px;
  text-align: center;
  position: fixed;
  left: 50%;
  top: 50%;
  margin-left: -256px;
  margin-top: -200px;
  overflow: hidden;
  display: none;
  z-index: 2000;
}
@media all and (max-width: 767px) {
  .sweet-alert {
    width: auto;
    margin-left: 0;
    margin-right: 0;
    left: 15px;
    right: 15px;
  }
}
.sweet-alert .form-group {
  display: none;
}
.sweet-alert .form-group .sa-input-error {
  display: none;
}
.sweet-alert.show-input .form-group {
  display: block;
}
.sweet-alert .sa-confirm-button-container {
  display: inline-block;
  position: relative;
}
.sweet-alert .la-ball-fall {
  position: absolute;
  left: 50%;
  top: 50%;
  margin-left: -27px;
  margin-top: -9px;
  opacity: 0;
  visibility: hidden;
}
.sweet-alert button[disabled] {
  opacity: .6;
  cursor: default;
}
.sweet-alert button.confirm[disabled] {
  color: transparent;
}
.sweet-alert button.confirm[disabled] ~ .la-ball-fall {
  opacity: 1;
  visibility: visible;
  transition-delay: 0s;
}
.sweet-alert .sa-icon {
  width: 80px;
  height: 80px;
  border: 4px solid gray;
  border-radius: 50%;
  margin: 20px auto;
  position: relative;
  box-sizing: content-box;
}
.sweet-alert .sa-icon.sa-error {
  border-color: #d43f3a;
}
.sweet-alert .sa-icon.sa-error .sa-x-mark {
  position: relative;
  display: block;
}
.sweet-alert .sa-icon.sa-error .sa-line {
  position: absolute;
  height: 5px;
  width: 47px;
  background-color: #d9534f;
  display: block;
  top: 37px;
  border-radius: 2px;
}
.sweet-alert .sa-icon.sa-error .sa-line.sa-left {
  transform: rotate(45deg);
  left: 17px;
}
.sweet-alert .sa-icon.sa-error .sa-line.sa-right {
  transform: rotate(-45deg);
  right: 16px;
}
.sweet-alert .sa-icon.sa-warning {
  border-color: #eea236;
}
.sweet-alert .sa-icon.sa-warning .sa-body {
  position: absolute;
  width: 5px;
  height: 47px;
  left: 50%;
  top: 10px;
  border-radius: 2px;
  margin-left: -2px;
  background-color: #f0ad4e;
}
.sweet-alert .sa-icon.sa-warning .sa-dot {
  position: absolute;
  width: 7px;
  height: 7px;
  border-radius: 50%;
  margin-left: -3px;
  left: 50%;
  bottom: 10px;
  background-color: #f0ad4e;
}
.sweet-alert .sa-icon.sa-info {
  border-color: #46b8da;
}
.sweet-alert .sa-icon.sa-info::before {
  content: "";
  position: absolute;
  width: 5px;
  height: 29px;
  left: 50%;
  bottom: 17px;
  border-radius: 2px;
  margin-left: -2px;
  background-color: #5bc0de;
}
.sweet-alert .sa-icon.sa-info::after {
  content: "";
  position: absolute;
  width: 7px;
  height: 7px;
  border-radius: 50%;
  margin-left: -3px;
  top: 19px;
  background-color: #5bc0de;
}
.sweet-alert .sa-icon.sa-success {
  border-color: #4cae4c;
}
.sweet-alert .sa-icon.sa-success::before,
.sweet-alert .sa-icon.sa-success::after {
  content: '';
  border-radius: 50%;
  position: absolute;
  width: 60px;
  height: 120px;
  background: #ffffff;
  transform: rotate(45deg);
}
.sweet-alert .sa-icon.sa-success::before {
  border-radius: 120px 0 0 120px;
  top: -7px;
  left: -33px;
  transform: rotate(-45deg);
  transform-origin: 60px 60px;
}
.sweet-alert .sa-icon.sa-success::after {
  border-radius: 0 120px 120px 0;
  top: -11px;
  left: 30px;
  transform: rotate(-45deg);
  transform-origin: 0px 60px;
}
.sweet-alert .sa-icon.sa-success .sa-placeholder {
  width: 80px;
  height: 80px;
  border: 4px solid rgba(92, 184, 92, 0.2);
  border-radius: 50%;
  box-sizing: content-box;
  position: absolute;
  left: -4px;
  top: -4px;
  z-index: 2;
}
.sweet-alert .sa-icon.sa-success .sa-fix {
  width: 5px;
  height: 90px;
  background-color: #ffffff;
  position: absolute;
  left: 28px;
  top: 8px;
  z-index: 1;
  transform: rotate(-45deg);
}
.sweet-alert .sa-icon.sa-success .sa-line {
  height: 5px;
  background-color: #5cb85c;
  display: block;
  border-radius: 2px;
  position: absolute;
  z-index: 2;
}
.sweet-alert .sa-icon.sa-success .sa-line.sa-tip {
  width: 25px;
  left: 14px;
  top: 46px;
  transform: rotate(45deg);
}
.sweet-alert .sa-icon.sa-success .sa-line.sa-long {
  width: 47px;
  right: 8px;
  top: 38px;
  transform: rotate(-45deg);
}
.sweet-alert .sa-icon.sa-custom {
  background-size: contain;
  border-radius: 0;
  border: none;
  background-position: center center;
  background-repeat: no-repeat;
}
.sweet-alert .btn-default:focus {
  border-color: #cccccc;
  outline: 0;
  -webkit-box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(204, 204, 204, 0.6);
  box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(204, 204, 204, 0.6);
}
.sweet-alert .btn-success:focus {
  border-color: #4cae4c;
  outline: 0;
  -webkit-box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(76, 174, 76, 0.6);
  box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(76, 174, 76, 0.6);
}
.sweet-alert .btn-info:focus {
  border-color: #46b8da;
  outline: 0;
  -webkit-box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(70, 184, 218, 0.6);
  box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(70, 184, 218, 0.6);
}
.sweet-alert .btn-danger:focus {
  border-color: #d43f3a;
  outline: 0;
  -webkit-box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(212, 63, 58, 0.6);
  box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(212, 63, 58, 0.6);
}
.sweet-alert .btn-warning:focus {
  border-color: #eea236;
  outline: 0;
  -webkit-box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(238, 162, 54, 0.6);
  box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(238, 162, 54, 0.6);
}
.sweet-alert button::-moz-focus-inner {
  border: 0;
}
/*!
 * Load Awesome v1.1.0 (http://github.danielcardoso.net/load-awesome/)
 * Copyright 2015 Daniel Cardoso <@DanielCardoso>
 * Licensed under MIT
 */
.la-ball-fall,
.la-ball-fall > div {
  position: relative;
  -webkit-box-sizing: border-box;
  -moz-box-sizing: border-box;
  box-sizing: border-box;
}
.la-ball-fall {
  display: block;
  font-size: 0;
  color: #fff;
}
.la-ball-fall.la-dark {
  color: #333;
}
.la-ball-fall > div {
  display: inline-block;
  float: none;
  background-color: currentColor;
  border: 0 solid currentColor;
}
.la-ball-fall {
  width: 54px;
  height: 18px;
}
.la-ball-fall > div {
  width: 10px;
  height: 10px;
  margin: 4px;
  border-radius: 100%;
  opacity: 0;
  -webkit-animation: ball-fall 1s ease-in-out infinite;
  -moz-animation: ball-fall 1s ease-in-out infinite;
  -o-animation: ball-fall 1s ease-in-out infinite;
  animation: ball-fall 1s ease-in-out infinite;
}
.la-ball-fall > div:nth-child(1) {
  -webkit-animation-delay: -200ms;
  -moz-animation-delay: -200ms;
  -o-animation-delay: -200ms;
  animation-delay: -200ms;
}
.la-ball-fall > div:nth-child(2) {
  -webkit-animation-delay: -100ms;
  -moz-animation-delay: -100ms;
  -o-animation-delay: -100ms;
  animation-delay: -100ms;
}
.la-ball-fall > div:nth-child(3) {
  -webkit-animation-delay: 0ms;
  -moz-animation-delay: 0ms;
  -o-animation-delay: 0ms;
  animation-delay: 0ms;
}
.la-ball-fall.la-sm {
  width: 26px;
  height: 8px;
}
.la-ball-fall.la-sm > div {
  width: 4px;
  height: 4px;
  margin: 2px;
}
.la-ball-fall.la-2x {
  width: 108px;
  height: 36px;
}
.la-ball-fall.la-2x > div {
  width: 20px;
  height: 20px;
  margin: 8px;
}
.la-ball-fall.la-3x {
  width: 162px;
  height: 54px;
}
.la-ball-fall.la-3x > div {
  width: 30px;
  height: 30px;
  margin: 12px;
}
/*
 * Animation
 */
@-webkit-keyframes ball-fall {
  0% {
    opacity: 0;
    -webkit-transform: translateY(-145%);
    transform: translateY(-145%);
  }
  10% {
    opacity: .5;
  }
  20% {
    opacity: 1;
    -webkit-transform: translateY(0);
    transform: translateY(0);
  }
  80% {
    opacity: 1;
    -webkit-transform: translateY(0);
    transform: translateY(0);
  }
  90% {
    opacity: .5;
  }
  100% {
    opacity: 0;
    -webkit-transform: translateY(145%);
    transform: translateY(145%);
  }
}
@-moz-keyframes ball-fall {
  0% {
    opacity: 0;
    -moz-transform: translateY(-145%);
    transform: translateY(-145%);
  }
  10% {
    opacity: .5;
  }
  20% {
    opacity: 1;
    -moz-transform: translateY(0);
    transform: translateY(0);
  }
  80% {
    opacity: 1;
    -moz-transform: translateY(0);
    transform: translateY(0);
  }
  90% {
    opacity: .5;
  }
  100% {
    opacity: 0;
    -moz-transform: translateY(145%);
    transform: translateY(145%);
  }
}
@-o-keyframes ball-fall {
  0% {
    opacity: 0;
    -o-transform: translateY(-145%);
    transform: translateY(-145%);
  }
  10% {
    opacity: .5;
  }
  20% {
    opacity: 1;
    -o-transform: translateY(0);
    transform: translateY(0);
  }
  80% {
    opacity: 1;
    -o-transform: translateY(0);
    transform: translateY(0);
  }
  90% {
    opacity: .5;
  }
  100% {
    opacity: 0;
    -o-transform: translateY(145%);
    transform: translateY(145%);
  }
}
@keyframes ball-fall {
  0% {
    opacity: 0;
    -webkit-transform: translateY(-145%);
    -moz-transform: translateY(-145%);
    -o-transform: translateY(-145%);
    transform: translateY(-145%);
  }
  10% {
    opacity: .5;
  }
  20% {
    opacity: 1;
    -webkit-transform: translateY(0);
    -moz-transform: translateY(0);
    -o-transform: translateY(0);
    transform: translateY(0);
  }
  80% {
    opacity: 1;
    -webkit-transform: translateY(0);
    -moz-transform: translateY(0);
    -o-transform: translateY(0);
    transform: translateY(0);
  }
  90% {
    opacity: .5;
  }
  100% {
    opacity: 0;
    -webkit-transform: translateY(145%);
    -moz-transform: translateY(145%);
    -o-transform: translateY(145%);
    transform: translateY(145%);
  }
}

.accordion {
  background-color: #eee;
  color: #444;
  cursor: pointer;
  padding: 18px;
  width: 100%;
  border: none;
  text-align: left;
  outline: none;
  font-size: 20px;
  transition: 0.4s;
}

.active, .accordion:hover {
  background-color: #ccc;
}

.accordion:after {
  content: '\002B';
  color: #777;
  font-weight: bold;
  float: right;
  margin-left: 5px;
}

.active:after {
  content: "\2212";
}

.panel {
  padding: 0 18px;
  background-color: white;
  max-height: 0;
  overflow: hidden;
  transition: max-height 0.2s ease-out;
}
</style>
	<!-- Editar Moneda -->
	<script src="https://kit.fontawesome.com/9d7cfbccc5.js" crossorigin="anonymous"></script>
	<script src="js/jquery-3.1.1.min.js"></script>
	<script src="js/bootstrap.min.js"></script>
	<script src="validaciondata/monedaV1.js"></script>

</body>
</html>