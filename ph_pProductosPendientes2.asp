<!DOCTYPE HTML>
<html >
<head>
	<title>Productos Pendientes</title>
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
	<link rel="icon" href="favicon.ico" type="image/x-icon"> 
	<link href="css/sweetalert.css" rel="stylesheet" type="text/css" />
	<link href="css/factura.css" rel="stylesheet" type="text/css" />	
	<link href="css/bootstrap.min.css" rel="stylesheet" type="text/css" />	
</head>
<body topmargin="0">
	<!--#include file="estiloscss.asp"-->
	<!--#include file="meta.asp"-->
	<!--#include file="encabezado.asp"-->
	<!--#include file="nn_subN.asp"-->
	<!--#include file="in_DataEN.asp"-->
	<%
		' ph_pProductosPendientes.asp -- 02mar21 - 03mar21
		
		Apertura
		
		' Parámetros del Manteniemiento
		
		LeePar
		
		if ed_iPas<>4 then 
			Encabezado
		end if    	
		'
		Session.lcid		= 1034
		Response.CodePage 	= 65001
		Response.CharSet 	= "utf-8"	
		'
		Dim rsProductosPendientes, arrProductosPendientes
		'	
		' Buscar Los Productos Pendientes por informacion completa del codigo de Barras
		'	
		QrySql = vbnullstring
		QrySql = QrySql & " SELECT TOP (100)  PERCENT"
		QrySql = QrySql & " cacevedo_atenas.PH_Consumo_Detalle_Productos.Numero_codigo_barras,"
		QrySql = QrySql & " COUNT ( cacevedo_atenas.PH_Consumo_Detalle_Productos.Id_Consumo_Detalle_Productos ) AS Total"
		QrySql = QrySql & " FROM"
		QrySql = QrySql & " cacevedo_atenas.PH_Consumo_Detalle_Productos"
		QrySql = QrySql & " LEFT OUTER JOIN cacevedo_atenas.PH_CB_Producto"
		QrySql = QrySql & " ON cacevedo_atenas.PH_Consumo_Detalle_Productos.Numero_codigo_barras = cacevedo_atenas.PH_CB_Producto.CodigoBarra"
		QrySql = QrySql & " WHERE"
		QrySql = QrySql & " ( cacevedo_atenas.PH_Consumo_Detalle_Productos.Id_Hogar > 1 )"
		QrySql = QrySql & " AND"
		QrySql = QrySql & " ( cacevedo_atenas.PH_Consumo_Detalle_Productos.Pendiente = 1 )"
		QrySql = QrySql & " AND"
		QrySql = QrySql & " ( cacevedo_atenas.PH_Consumo_Detalle_Productos.Tiene_Codigo_Barras = 1 )"
		QrySql = QrySql & " AND"
		QrySql = QrySql & " ( cacevedo_atenas.PH_Consumo_Detalle_Productos.Status_registro='G')"
		QrySql = QrySql & " AND"
		QrySql = QrySql & " ( cacevedo_atenas.PH_CB_Producto.Id_Producto IS NULL )"
		QrySql = QrySql & " GROUP BY"
		QrySql = QrySql & " cacevedo_atenas.PH_Consumo_Detalle_Productos.Numero_codigo_barras"
		QrySql = QrySql & " ORDER BY"
		QrySql = QrySql & " COUNT ( cacevedo_atenas.PH_Consumo_Detalle_Productos.Id_Consumo_Detalle_Productos ) DESC"
		'		
		Set rsProductosPendientes = Server.CreateObject("ADODB.recordset")
		rsProductosPendientes.Open QrySql, conexion
		'
		if not rsProductosPendientes.EOF then
			arrProductosPendientes = rsProductosPendientes.GetRows()  ' Convert recordset to 2D Array
		end if
		rsProductosPendientes.Close : Set rsProductosPendientes = Nothing 
		'		
	%>
	<div class="container-fluid" id="grad1">  
	
		<div class="form-group">
			<div class="col-sm-3">
				<div class="form-group">				
					<label>Maestro de Productos:</label><span id="loader"></span>	
					<select class="form-control input-sm" title="Maestro Tabla de Productos" name="cboMasterProductos" id="cboMasterProductos" />
						<option value="0" select>-- Seleccionar --</option> 					
					</select>
				</div>
			</div>
			
			<div class="col-sm-3">
				<div class="form-group">				
					<label><i class="fas fa-barcode"></i>&nbsp;Seleccione el C&oacute;digo de Barra Pendiente:</label>					
						<select class="form-control input-sm" title="Seleccionar codigo de Barra a Procesar" name="cboProductosPendientes" id="cboProductosPendientes" onchange="buscarProductosPendientes();" />
							<option value="0" selected disabled >-- Seleccione -- </option>
							<% 'Check si es una array
							if IsArray(arrProductosPendientes) then
								For i = 0 to ubound(arrProductosPendientes, 2) %>
								<option value="<%= arrProductosPendientes(0,i)%>"> <%= uCase(arrProductosPendientes(0,i)) & " - (" & uCase(arrProductosPendientes(1,i)) &")"   %> </option>								
								<% next %>
							<% else %>
								<option value="0" disabled>-- No hay Datos -- </option>
							<% end if %>
					</select>
				</div>
			</div>			
		</div>					
	</div>
	<hr>
	
	<div class="container-fluid text-center text-primary" id="cargando" style="display:none;">
		<span ><img src="images/ajax-loader7.gif"><strong>&nbsp;Espere.., Procesando!</strong></span>
	</div> 
		
	<!-- 0 -->		
		
	<%conexion.close%>

</body>
</html>

<script src="https://kit.fontawesome.com/9d7cfbccc5.js" crossorigin="anonymous"></script>
<script src="js/jquery-3.1.1.min.js"></script>
<script src="js/sweetalert.min.js"></script>
<script src="js/bootstrap.min.js"></script>
<script src="validarprodpend/autoNumeric-1.9.18.js"></script>
<script src="validarprodpend/funcionesV1.js"></script>

<script>
	
	$(document).ready(function() {
		
		$(function($) {
			$('#txtPrecio').autoNumeric('init', {
				lZero: 'deny',
				aSep: '.',
				aDec: ','
			});			
    	});
					
		$(function() {
			//buscarProductosPendientes();
		});
									
	});	
	
</script>
