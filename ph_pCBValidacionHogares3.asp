<!DOCTYPE HTML>
<html >
<head>
	<title>Validacion Hogares</title>
	<meta charset="UTF-8">
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
' 29dic20 - 05ene21
'==========================================================================================
' Variables y Constantes
'==========================================================================================
    Apertura
'==========================================================================================
' Parámetros del Manteniemiento
'==========================================================================================
    LeePar
	
    if ed_iPas<>4 then 
        Encabezado
    end if    	
    '	
%>

	<div class="container-fluid" id="grad1">  
	
		<div class="form-group">
	
			<div class="col-sm-3">
				<div class="form-group">				
					<label>Seleccione Semana:</label><span id="loader"></span>	
					<select class="form-control input-sm" title="Seleccionar Semana" name="cboSemana" id="cboSemana" onchange="buscarArea();"  />
						<option value="0" select>-- Seleccionar --</option> 					
					</select>
				</div>
			</div>
			
			<div class="col-sm-3">
				<div class="form-group">				
					<label>Seleccione Area:</label>
					<select class="form-control input-sm" title="Seleccionar Area" name="cboArea" id="cboArea" onchange="buscarEstado();" />
						<option value="0" select>-- Seleccionar --</option> 					
					</select>
				</div>
			</div>
			
			<div class="col-sm-2">
				<div class="form-group">				
					<label>Seleccione Estado:</label>
					<select class="form-control input-sm" title="Seleccionar Estado" name="cboEstado" id="cboEstado" onchange="buscarHogar();" />
						<option value="0" select>-- Seleccionar --</option> 					
					</select>
				</div>
			</div>
						
			<div class="col-sm-2">
				<div class="form-group">				
					<label>Seleccione Hogar:</label>
					<select class="form-control input-sm" title="Seleccionar Hogar" name="cboHogar" id="cboHogar" onchange="buscarTipoConsumo();" />
						<option value="0" select>-- Seleccionar --</option> 					
					</select>
				</div>
			</div>
			
			<div class="col-sm-2">
				<div class="form-group">				
					<label>Alta Hogar:</label>
					<input type="text" class="form-control input-sm text-center text-danger" id="txtAltaHogar" placeholder="...." readonly />
				</div>
			</div>
			
		</div>
		
		<div class="form-group">
				
			<div class="col-sm-3">
				<div class="form-group">				
					<label>Seleccione Tipo Consumo:</label>
					<select class="form-control input-sm" title="Seleccionar Area" name="cboTipoConsumo" id="cboTipoConsumo" onchange="buscarTotalDiaSemana();" />
						<option value="0" select>-- Seleccionar --</option> 					
					</select>
				</div>
			</div>		
			
			<div class="col-sm-3">
				<div class="form-group">				
					<label>Seleccione Dia Semana:</label>
					<select class="form-control input-sm" title="Seleccionar Dia" name="cboTotalxDiaSemana" id="cboTotalxDiaSemana" onchange="buscarDetalleDiaSemana();" />
						<option value="0" select>-- Seleccionar --</option> 					
					</select>
				</div>
			</div>
	
			<div class="col-sm-3">
				<div class="form-group">				
					<label>Seleccione Consumo:</label>
					<select class="form-control input-sm" title="Seleccionar Consumo a Validar" name="cboDetallexDiaSemana" id="cboDetallexDiaSemana" onchange="buscarDetalleConsumoxDia();" />
						<option value="0" select>-- Seleccionar --</option> 					
					</select>
				</div>
			</div>
			
			<div class="col-sm-3">				
				<div class="form-group">
					<div class="col-sm-6">
						<label for="usr">Reset</label>
						<button id="borrar"   type="submit" class="btn btn-block btn-sm btn-info" onclick="Reset();"><i class="fas fa-recycle fa-2x"></i></button>
					</div>
					<div class="col-sm-6">
						<label for="usr">Procesar</label>
						<button id="procesar" type="submit" class="btn btn-block btn-sm btn-success" onclick="buscarImagenFactura();"><i class="fas fa-check fa-2x"></i></button>
					</div>
				</div>				
			</div>
																		
		</div>
		
	</div>        
	<hr>	
	<div class="container-fluid" id="DetalleFactura" style="display:none;">
	
		<div class="form-group-row text-center alert alert-success" role="alert" id="hogarValidado" style="display:none;" >
			<span class="bg-success"><h5><strong>CONSUMO VALIDADO&nbsp;<i class="fas fa-check"></i></strong></h5></span>			
		</div>
		<div class="form-group-row text-center alert alert-danger" role="alert" id="hogarEliminado" style="display:none;" >
			<span class="bg-danger"><h5><strong>CONSUMO ELIMINADO&nbsp;<i class="fas fa-times"></i></strong></h5></span>			
		</div>
		
		<div class="form-group-row">
		
			<div class="col-sm-3">
			
				<h4 class="text-danger"><strong>Detalle Factura:</strong></h4>
				<input type="hidden" id="tieneFactura" name="tieneFactura" value="0" />									
								
				<fieldset> 					
					<div class="my-funky-img-box">
					   <!--/.col-bdr.left-border-col-->
					   <div class="col-bdr left-border-col"></div>				   
					   <!--/.background-col-->
					   <div class="background-col">
							<img class="img-responsive" id="imgfactura" name="imgfactura" src="images/loader/cargador1.gif" style="width:480px; height:640px;" />
					   </div>
					   <!--/.col-bdr.left-border-col-->
					   <div class="col-bdr right-border-col"></div>				   
					</div>
				</fieldset> 
						
				<!---->
				<br>
				<div class="form-group">
					
					<div class="col-sm-6">
						<label class="text-danger">Canal:</label>
						<select class="form-control" name="cboCanal" id="cboCanal" onChange="buscarCadena(this.value);" required>
							<option value="" selected disabled >-- Seleccione -- </option>							
						</select>
						<div class="error" id="canalErr"></div>	
					</div>
														
					<div class="col-sm-6">
						<label class="text-danger">Cadena:</label>
						<select class="form-control" name="cboCadena" id="cboCadena" required>
							<!-- Mostrar Paneles -->
							<option value="" selected disabled >-- Seleccione -- </option>							
						</select>
					</div>
				</div>
				
				<div class="form-group">
					<div class="col-sm-6">
						<div class="text-danger">
							<!--<strong><p id="tienefactura"></p></strong>-->
							<br>
							<h5><strong><span class="text-danger" id="tienefactura"></span></strong></h5>
						</div>	
					</div>		
					<!-- Monto Total Factura-->
					<div class="col-sm-6" id="verFactura" style="display:none;">
						<span class="text-danger"><h6><strong>Monto Total Factura:</strong></h6></span>					
						<input type="text" class="form-control input-sm text-right" id="totalFactura" name="totalFactura" placeholder="Monto total de la Factura" onblur="formatMonto(this.value)" />					
						<div class="error" id="totalfacturaErr"></div>	
					</div>							
					
				</div>
				
				<div class="col-sm-12">
					<h5><strong><span class="text-danger"></span></strong></h5>				
					<button type="button" title="Grabar" class="btn btn-block btn-primary btn-sm" id="submit" onclick="grabarCambiosFactura();"><i class='fas fa-save'></i> Grabar Cambios</button>
				</div>												
										
			</div>
			
			<div class="col-sm-9">					
			
				<h4 class="text-danger"><strong>Detalle Productos:</strong></h4>
								
				<div class="table-responsive" id="tabla-resultados">
					<!-- // ** // -->
					<!-- Matriz de Datos Resultados -->
					<!-- // ** // -->										
				</div>		
					
			</div>
							
		</div>
			
	</div>
		
	</div>
	
	<hr>

	<div class="modal" id="Detalleconsumo" tabindex="-1" data-backdrop="static" data-keyboard="false" role="dialog" aria-labelledby="myModalLabel" data-focus-on="input:first">

		<div class="modal-dialog modal-dialog-centered"  role="document">

			<div class="modal-content">

				<div class="modal-header">
					<h4 class="modal-title">Large Modal</h4>					
				</div>

				<div class="modal-body">
				
						<div class="form-group">
							<label>ID:</label>
							<input type="text" class="form-control input-sm text-right" id="txtId" placeholder="...." readonly />
						</div>
											
						<div class="form-group">
							<label>Codigo de Barras:</label>
							<input type="text" class="form-control input-sm text-right" id="txtCodigoBar" placeholder="...." minlength="7" maxlength="16" onkeypress="return onlyNumberKey(event);" required />
							<div class="error" id="codigobarErr"></div>							 
						</div>
						<div class="form-group">
							 <label>Cantidad:</label>
							 <input type="text" class="form-control input-sm text-right" id="txtCantidad" placeholder="...." required />
							 <div class="error" id="cantidadErr"></div>							 
						</div>
						
						<div class="form-group">
							 <label>Precio Unitario:</label>
							 <input type="text" class="form-control input-sm text-right" id="txtPrecio" placeholder="...." onblur="formatMoney(this.value)" required />
							 <div class="error" id="precioErr"></div>
						</div>											
						
						<div class="form-group">
							 <label>Tasa de Cambio:</label>
							 <input type="text" class="form-control input-sm text-right" id="txtTasa" placeholder="...." readonly />							 
						</div>											
						
						<div class="form-group">
							 <label>Moneda de Pago:</label>
							 <input type="text" class="form-control input-sm text-right" id="txtMoneda" placeholder="...." readonly />							 
						</div>											
															
				</div>
				
				<div class="modal-footer">
					<button type="button" class="btn btn-danger" data-dismiss="modal" title="Salir"><i class='fas fa-sign-out-alt'></i> Salir</button>
					<button type="button" class="btn btn-primary" title="Grabar" onclick="salvarCambioProductos();" id="btn-salvar"><i class='fas fa-save'></i> Grabar</button>
				</div>
			</div>
			<!-- /.modal-content -->
		</div>
		<!-- /.modal-dialog -->
	</div>
    <!-- /.modal -->
		
	<div class="modal" id="AgregarProducto" tabindex="-1" data-backdrop="static" data-keyboard="false" role="dialog" aria-labelledby="myModalLabel" data-focus-on="input:first">

		<div class="modal-dialog modal-dialog-centered"  role="document">

			<div class="modal-content">

				<div class="modal-header">
					<h4 class="modal-title">Large Modal</h4>					
				</div>

				<div class="modal-body">
				
						<div class="form-group">
							<label>ID:</label>
							<input type="text" class="form-control input-sm text-right" id="txtIdConsumo" placeholder="...." readonly />
						</div>
						
						<div class="form-group">
							<label class="text-primary">Seleccione Categoria:</label>
							<select class="form-control" name="cboCategoria" id="cboCategoria" required>
								<option value="" selected disabled >-- Seleccione -- </option>							
							</select>
						</div>
						
						<div class="form-group">							
							<div class="col-sm-10">          
								<label class="text-primary">Describa el producto a buscar:</label>
								<input type="text" class="form-control input-sm" id="txtBuscarDescripcion" placeholder="Introduzca todo o parte del producto a buscar...." required />
							</div>
							<label class="control-label col-sm-2"><button type="button" title="Buscar" class="btn btn-primary btn-sm" onclick="buscarProducto();"><i class='fas fa-search'></i></button></label>
						</div>
												
						<div class="form-group">
							<label class="text-Primary">Seleccione Producto:</label>
							<select class="form-control" name="cboProducto" id="cboProducto"  onchange="mostrarBarcode();" required>
								<option value="" selected disabled >-- Seleccione -- </option>							
							</select>
						</div>
						
						<div class="form-group">
							 <label class="text-primary">C&oacute;digo de Barras:</label>
							 <input type="text" class="form-control input-sm text-right" id="txtCodigoBarras" name="txtCodigoBarras" placeholder="...." readonly />
						</div>
						
						<div class="form-group">
							 <label class="text-primary">Cantidad:</label>
							 <input type="text" class="form-control input-sm text-right" id="txtCantidadProductos" placeholder="...." required />
							 <div class="error" id="cantidadErr"></div>							 
						</div>
						
						<div class="form-group">
							 <label class="text-primary">Precio Unitario:</label>
							 <input type="text" class="form-control input-sm text-right" id="txtPrecioUnitario" placeholder="...." onblur="formatMoney(this.value)" required />
							 <div class="error" id="precioErr"></div>
						</div>											
						
						<div class="form-group">
							 <label class="text-primary">Tasa de Cambio:</label>
							 <input type="text" class="form-control input-sm text-right" id="txtTasaCambio" placeholder="...." readonly />							 
						</div>											
						
						<div class="form-group">
							<label class="text-primary">Seleccione Moneda de Pago:</label>
							<select class="form-control" name="cboMonedaPago" id="cboMonedaPago" required>
								<option value="" selected disabled >-- Seleccione -- </option>							
							</select>
						</div>										
															
				</div>
				
				<div class="modal-footer">
					<button type="button" class="btn btn-danger" data-dismiss="modal" title="Salir"><i class='fas fa-sign-out-alt'></i> Salir</button>
					<button type="button" class="btn btn-primary" title="Grabar" onclick="salvarAgregarProductos();" id="btn-salvar"><i class='fas fa-save'></i> Grabar</button>
				</div>
			</div>
			<!-- /.modal-content -->
		</div>
		<!-- /.modal-dialog -->
	</div>
    <!-- /.modal -->
	
	<%conexion.close%>
	
</body>
</html>
<script src="https://kit.fontawesome.com/9d7cfbccc5.js" crossorigin="anonymous"></script>
<script src="js/jquery-3.1.1.min.js"></script>
<script src="js/sweetalert.min.js"></script>
<script src="js/bootstrap.min.js"></script>
<script src="validacion/utilitarios.js"></script>
<script src="validacion/funcionesV5.js"></script>
<script src="validacion/crud.js"></script>

<script>
	//jQuery.noConflict();
	$(document).ready(function() {
		$(function() {
			buscarSemanas();
		});
		$(function() {
			buscarCanal();
		});
		$(function() {
			buscarCadena(0);
		});		
		$(function() {
			buscarCategoria();
		});	
	});		
</script>
