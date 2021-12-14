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
	<link href="css/bootstrap.min.css" rel="stylesheet" type="text/css" />	
	<link href="css/prodpend.css" rel="stylesheet" type="text/css" />	
</head>
<body topmargin="0">
	<!--#include file="estiloscss.asp"-->
	<!--#include file="meta.asp"-->
	<!--#include file="encabezado.asp"-->
	<!--#include file="nn_subN.asp"-->
	<!--#include file="in_DataEN.asp"-->
	<%
		' ph_pProductosPendientes.asp - 02mar21 - 01jul21
		
		Apertura
		
		' Parámetros del Manteniemiento
		
		LeePar
		
		if ed_iPas<>4 then 
			Encabezado
		end if    	
		'
		'Response.write "<br>34 llego"
		'Response.end
		Session.lcid		= 1034
		Response.CodePage 	= 65001
		Response.CharSet 	= "utf-8"	
		'
		Dim rsProductosPendientes, arrProductosPendientes
		'	
		' Buscar Los Productos Pendientes por informacion completa del codigo de Barras
		'					
		'01jul21 - Medicinas
		'
		QrySql = vbnullstring
		QrySql = QrySql & " (SELECT PH_Consumo_Detalle_Productos.Numero_codigo_barras,"
		QrySql = QrySql & " Count(PH_Consumo_Detalle_Productos.Id_Consumo_Detalle_Productos) AS Total, '' as Estatus,'' as TipMed"
		QrySql = QrySql & " FROM PH_Consumo_Detalle_Productos LEFT JOIN PH_CB_Producto ON PH_Consumo_Detalle_Productos.Numero_codigo_barras = PH_CB_Producto.CodigoBarra"
		QrySql = QrySql & " WHERE PH_Consumo_Detalle_Productos.Id_Hogar>1 AND PH_Consumo_Detalle_Productos.Pendiente=0 AND PH_Consumo_Detalle_Productos.Status_registro='G' AND PH_CB_Producto.Id_Producto Is Null"
		QrySql = QrySql & " GROUP BY PH_Consumo_Detalle_Productos.Numero_codigo_barras"
		QrySql = QrySql & " HAVING PH_Consumo_Detalle_Productos.Numero_codigo_barras<>'0' And PH_Consumo_Detalle_Productos.Numero_codigo_barras<>'00000000' And PH_Consumo_Detalle_Productos.Numero_codigo_barras is Not Null)"
		QrySql = QrySql & " UNION"
		QrySql = QrySql & " ( SELECT PH_Consumo_Detalle_Productos.Numero_codigo_barras, Count(PH_Consumo_Detalle_Productos.Id_Consumo_Detalle_Productos) AS Total, 'P' as Estatus, PH_CB_Categoria.Ind_Medicina as TipMed"
		QrySql = QrySql & " FROM (PH_Consumo_Detalle_Productos LEFT JOIN PH_CB_Producto ON PH_Consumo_Detalle_Productos.Numero_codigo_barras = PH_CB_Producto.CodigoBarra) LEFT JOIN PH_CB_Categoria ON PH_CB_Producto.Id_Categoria = PH_CB_Categoria.id_Categoria"
		QrySql = QrySql & " WHERE PH_Consumo_Detalle_Productos.Id_Hogar>1 AND PH_Consumo_Detalle_Productos.Pendiente=0 AND PH_Consumo_Detalle_Productos.Status_registro='G' AND PH_CB_Producto.Ind_Pendiente=1"
		QrySql = QrySql & " GROUP BY PH_Consumo_Detalle_Productos.Numero_codigo_barras, PH_CB_Categoria.Ind_Medicina"
		QrySql = QrySql & " HAVING PH_Consumo_Detalle_Productos.Numero_codigo_barras<>'0' And PH_Consumo_Detalle_Productos.Numero_codigo_barras<>'00000000')"
		QrySql = QrySql & " ORDER BY 2 desc"
		'response.write QrySql
		'response.end		
		'
		Set rsProductosPendientes = Server.CreateObject("ADODB.recordset")
		rsProductosPendientes.Open QrySql, conexion
		'
		if not rsProductosPendientes.EOF then
			arrProductosPendientes = rsProductosPendientes.GetRows()  ' Convert recordset to 2D Array
		end if
		rsProductosPendientes.Close : Set rsProductosPendientes = Nothing 
		'Response.write "<br>70 llego"
		'Response.end

		
		'		
	%>
	
	<div class="container-fluid" id="grad1">  
		</br>
		
		<div class="form-group row">	
		
			<div class="col-sm-3">
				<div class="form-group">
					<label>Maestro de Productos:</label>
					<button title="Crear Productos" type="submit" class="btn btn-block btn-sm btn-primary" onclick="MostrarModalMaestroProductos();"><i class="fas fa-plus"></i>&nbsp;CREAR</button>							
				</div>
			</div>
			
			<div class="col-sm-3">
				<div class="form-group">
					<label>Cierre Productos:</label>
					<button title="Crear un Productos" type="submit" class="btn btn-block btn-sm btn-success" onclick="ValidarProductosPendientes();"><i class="fas fa-sign-in-alt"></i>&nbsp;CERRAR</button>							
				</div>
			</div>
			
			<div class="col-sm-3">
				<div class="form-group text-left">
					<label>Masivo de Precios:</label>
					<button title="Crear un Productos" type="submit" class="btn btn-block btn-sm btn-info"  onclick="showMostrarMasivoPrecios();"><i class="fas fa-check-double"></i>&nbsp;PROCESAR</button>							
				</div>
			</div>
<%
					'response.write "<br> Llego:= " & ubound(arrProductosPendientes, 2)
					'response.end 
%>						
			<div class="col-sm-3">
				<div class="form-group">				
					<label>Seleccione C&oacute;digo de Barra:</label>&nbsp;&nbsp;<a href="#" onClick="Reset(); return false;" title="Borrar Pantalla" class="label label-danger badge-pill">RESET</a>					
					<select class="form-control input-sm" title="Seleccionar codigo de Barra a Procesar" name="cboProductosPendientes" id="cboProductosPendientes" onchange="ProcesarCodigoBarras();"  />
						<option value="0" selected disabled >-- Seleccione -- </option>
						<%
							if IsArray(arrProductosPendientes) then
							'Check si es una array
							
							total = ubound(arrProductosPendientes, 2)
							
							if total > 44470 then total = 44470
														
							'For i = 0 to ubound(arrProductosPendientes, 2) - 827 
							For i = 0 to total 
							
							if(arrProductosPendientes(3,i)=true) then TipoMed="Med" else TipoMed=""
							
							
						%>
								<option value="<%= arrProductosPendientes(0,i)%>"> <%= uCase(arrProductosPendientes(0,i)) & " - (" & uCase(arrProductosPendientes(1,i)) &")" & " - " & uCase(arrProductosPendientes(2,i)) & " - " & TipoMed   %> </option>								
							<% next %>
						<% else %>
								<option value="0" disabled>-- No hay Datos -- </option>
						<% end if 
						%>
					</select>						
				</div>
			</div>	
						
		</div>
		
		<div class="form-group row">
		
			<div class="col-sm-4">
				<div class="form-group text-center">
					<span class="label label-info" id="totalProductos">0</span>	
				</div>
			</div>		
			
			<div class="col-sm-4">	
				<div class="form-group text-center">
					<span class="label label-danger " id="totalPendientes">0</span>				
				</div>	
			</div>	
			
			<div class="col-sm-4">	
				<div class="form-group text-center">
					<span class="label label-warning" id="totalValidados">0</span>		
				</div>	
			</div>	
					
		</div>	
					
		<div class="form-group row" id="DatosProductos" style="display:none;">		
		
			<div class="col-sm-2">
				<strong><p id="categoria" class="text-primary"></p></strong>
			</div>
			
			<div class="col-sm-10">
				<strong><p id="descripcion" class="text-primary"></p></strong>
			</div>
			
		</div>
		
	</div>
	<!-- 0 -->	
	
	<div class="container-fluid text-center text-primary" id="cargando" style="display:none;">
		<span ><img src="images/ajax-loader7.gif"><strong>&nbsp;Espere, Procesando..!</strong></span>
	</div> 
			
	<div class="container-fluid" id="tabla-DetalleProductosPendientes" style="display:none;">
		<!-- 0 -->		
	</div>
	
	<!-- /.modal -->
	<div class="modal" id="MostrarDetalleRegistro" tabindex="-1" data-backdrop="static" data-keyboard="false" role="dialog" aria-labelledby="myModalLabel" data-focus-on="input:first">

		<div class="modal-dialog modal-dialog-centered"  role="document">

			<div class="modal-content">

				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>					
					<h4 class="modal-title">Large Modal</h4>					
				</div>

				<div class="modal-body">
																					
					<div class="form-group">
						<label>Codigo de Barras:</label>
						<input type="text" class="form-control input-sm text-right" id="txtCodigoBar" placeholder="...."  readonly />
					</div>
																									
					<div class="form-group">
						 <label>Cantidad:</label>
						 <input type="text" class="form-control input-sm text-right" id="txtCantidad"  placeholder="...."  readonly />
					</div>
					
					<div class="form-group">
						 <label>Precio Unitario:</label>
						 <input type="text" class="form-control input-sm text-right" id="txtPrecio" placeholder="...." readonly />
					</div>											
					
					<div class="form-group">
						 <label>Tasa de Cambio:</label>
						 <input type="text" class="form-control input-sm text-right" id="txtTasa" placeholder="...." readonly />							 
					</div>											
					
					<div class="form-group">
						<label class="text-primary">Moneda:</label>
						<input type="text" class="form-control input-sm text-right" id="txtMoneda" placeholder="...." readonly />							 
					</div>			
										
					<div class="form-group">
						 <label class="text-primary">Total Compra:</label>
						 <input type="text" class="form-control input-sm text-right" id="txtTotal" placeholder="...." readonly />
					</div>		
					
					<div class="form-group">
						 <label class="text-primary">Fecha:</label>
						 <input type="text" class="form-control input-sm text-right" id="txtFecha" placeholder="...." readonly />
					</div>		
															
				</div>
				
				<div class="modal-footer">
					<button type="button" class="btn btn-danger" data-dismiss="modal" title="Salir"><i class='fas fa-sign-out-alt'></i> Salir</button>
				</div>
			</div>
			<!-- /.modal-content -->
		</div>
		<!-- /.modal-dialog -->
	</div>
    <!-- /.modal -->
	
	<!-- /.modal AgregarProducto -->		
	<div class="modal" id="CrearProductos" tabindex="-1" data-backdrop="static" data-keyboard="false" role="dialog" aria-labelledby="myModalLabel" data-focus-on="input:first">

		<div class="modal-dialog modal-lg modal-dialog-centered"  role="document">

			<div class="modal-content">

				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>					
					<h4 class="modal-title">Large Modal</h4>					
				</div>

				<div class="modal-body">
						
					<div class="form-group row" id="loader" style="display:none;" >
						<span class="text-danger text-center"><img src="images/ajax_small.gif">&nbsp;Grabando, Espere....!</span>
					</div>
				
					<div class="form-group row">
						<div class="col-md-6">						
							<label class="text-danger">C&oacute;digo de Barras:</label>
							<input type="text" class="form-control input-sm text-right" id="codigoBarras" name="codigoBarras" placeholder="...." Readonly />										
							<div class="error" id="codigobarErr"></div>	
						</div>						
					</div>				
				
					<div class="form-group row">
						<div class="col-md-6">
							<label class="text-primary">Seleccione Categoria:</label>
							<select class="form-control input-sm" name="cboCategoria" id="cboCategoria" onchange="llenarComboFabricantes();"  required>
								<option value="" selected disabled >-- Seleccionar -- </option>							
							</select>
							<div class="error" id="categoriaErr"></div>
						</div>
						<div class="col-md-6">
							<label class="text-primary">Seleccione Fabricante:</label>
							<select class="form-control input-sm" name="cboFabricante" id="cboFabricante" onchange="llenarComboMarca();" required>
								<option value="" selected disabled >-- Seleccionar -- </option>							
							</select>
							<div class="error" id="fabricanteErr"></div>
						</div>
					</div>				
					<div class="form-group row">
						<div class="col-md-6">
							<label class="text-primary">Seleccione Marca:</label>
							<select class="form-control input-sm" name="cboMarcas" id="cboMarcas"  required>
								<option value="" selected disabled >-- Seleccionar -- </option>							
							</select>
							<div class="error" id="marcaErr"></div>
						</div>
						<div class="col-md-6">						
							<label class="text-primary">Seleccione Segmento:</label>
							<select class="form-control input-sm" name="cboSegmento" id="cboSegmento"  required>
								<option value="" selected disabled >-- Seleccionar -- </option>							
							</select>
							<div class="error" id="segmentoErr"></div>
						</div>						
					</div>
					
					<div class="form-group row">
						<div class="col-md-6">
							<label class="text-primary">Seleccione Tamaño:</label>
							<select class="form-control input-sm" name="cboTamano" id="cboTamano"  required>
								<option value="" selected disabled >-- Seleccionar -- </option>							
							</select>							
							<div class="error" id="tamanoErr"></div>
						</div>
						<div class="col-md-6">						
							<label class="text-primary">Seleccione Rango:</label>
							<select class="form-control input-sm" name="cboRango" id="cboRango" required>
								<option value="" selected disabled >-- Seleccionar -- </option>							
							</select>
							<div class="error" id="rangoErr"></div>
						</div>						
					</div>
					
					<div class="form-group row">
						<div class="col-md-6">
							<label class="text-primary">Seleccione Unidad / Medida:</label>
							<select class="form-control input-sm" name="cboUnidadMedida" id="cboUnidadMedida"  required>
								<option value="" selected disabled >-- Seleccionar -- </option>							
							</select>					
							<div class="error" id="unidadErr"></div>
						</div>
						<div class="col-md-6">						
							<label class="text-primary">Fecha de Alta:</label>
							<input type="text" class="form-control input-sm text-right" id="fechaCreacion" name="fechaCreacion" placeholder="Fecha de Creacion" Readonly />										
						</div>						
					</div>
					
					<div class="form-group row">
					
						<div class="col-md-12">
							<label class="text-primary">Descripci&oacute;n Producto:</label>
							<!--<textarea id="descripcionProducto" class="form-control input-sm" rows="4" placeholder="...." style="resize: none;"></textarea>-->
							<input type="text" class="form-control input-sm" id="descripcionProducto" placeholder="...." minlength="5" maxlength="100" required />
							<div class="error" id="productoErr"></div>
						</div>
						
					</div>
					
					<div class="form-group row">
						<div class="col-md-12">
							<label class="text-primary">Fragmentaci&oacute;n Producto:</label>
							<!--<textarea id="fragmentacion" class="form-control input-sm" rows="4" placeholder="...." style="resize: none;"></textarea>-->
							<input type="text" class="form-control input-sm" id="fragmentacion" placeholder="...." minlength="5" maxlength="50" />
							<div class="error" id="fragmentoErr"></div>
						</div>						
						
					</div>					
																															
				</div>
				
				<div class="modal-footer">
					<button type="button" class="btn btn-danger" data-dismiss="modal" title="Salir"><i class='fas fa-sign-out-alt'></i> Salir</button>
					<button type="button" class="btn btn-primary" title="Grabar" onclick="CrearProductos();" id="btn-crearProd"><i class='fas fa-save'></i> Grabar</button>
				</div>
				
			</div>
			<!-- /.modal-content -->
		</div>
		<!-- /.modal-dialog -->
	</div>
    <!-- /.modal -->
	
	<!-- /.modal Masivo precios -->
	<div class="modal" id="MasivoPrecios" tabindex="-1" data-backdrop="static" data-keyboard="false" role="dialog" aria-labelledby="myModalLabel" data-focus-on="input:first">

		<div class="modal-dialog modal-dialog-centered"  role="document">

			<div class="modal-content">

				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>					
					<h4 class="modal-title">Large Modal</h4>					
				</div>

				<div class="modal-body">
		
					<div class="form-group">
						<label>Codigo de Barras:</label>
						<input type="text" class="form-control input-sm text-center text-danger" id="txtBarcode" placeholder="...." readonly />						
					</div>
										
					<div class="form-group">
						 <label>Indique el Precio:</label>
						 <input type="text" class="form-control input-sm text-right" id="txtPrecioMasivo" placeholder="...." onblur="ActualizarCalculoTotales();" required />
						 <div class="error" id="precioErr"></div>
					</div>											
																
				</div>
				
				<div class="modal-footer">
					<button type="button" class="btn btn-danger" data-dismiss="modal" title="Salir"><i class='fas fa-sign-out-alt'></i> Salir</button>
					<button type="button" class="btn btn-primary" title="Grabar" onclick="actualizarPrecioMasivo();" id="btn-salvarMasivo"><i class='fas fa-save'></i> Grabar</button>
				</div>
			</div>
			<!-- /.modal-content -->
		</div>
		<!-- /.modal-dialog -->
	</div>
    <!-- /.modal -->
							
	<%
	conexion.close
	%>

</body>
</html>

<script src="https://kit.fontawesome.com/9d7cfbccc5.js" crossorigin="anonymous"></script>
<script src="js/jquery-3.1.1.min.js"></script>
<script src="js/sweetalert.min.js"></script>
<script src="js/bootstrap.min.js"></script>
<script src="validarprodpend/funcionesV7.js"></script>
<script src="validarprodpend/crudV12.js"></script>
<script src="validarprodpend/llenarcombos.js"></script>

<script>
	
	$(document).ready(function() {
				
		$(function() {
			$('#cboProductosPendientes').focus();	
		});
		
		$(function() {
			llenarComboCategoria();	
		});
								
	});	
	
</script>
