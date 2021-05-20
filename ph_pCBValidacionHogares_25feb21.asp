<!Doctype html>
<!-- ph_pCBValidacionHogares - 28dic20 - 23feb21 -->
<html >
<head>
	<title>Validaci&oacute;n Hogares</title>
	<meta charset="UTF-8">
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
	<link href="favicon.ico" rel="icon" type="image/x-icon">
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
		' 29dic20 - 29ene21
		Apertura
		' Parámetros del Manteniemiento
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
												
			<div class="col-sm-2">
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
			
			<div class="col-sm-3">
				<div class="form-group">				
					<label>Seleccione Tipo Consumo:</label>
					<select class="form-control input-sm" title="Seleccionar Area" name="cboTipoConsumo" id="cboTipoConsumo" onchange="buscarTotalDiaSemana();" />
						<option value="0" select>-- Seleccionar --</option> 					
					</select>
				</div>
			</div>		
						
		</div>
		
		<!-- 0 -->
		
		<div class="form-group">
																		
			<div class="col-sm-3">
				<div class="form-group">				
					<label>Seleccione Dia Semana:</label>
					<select class="form-control input-sm" title="Seleccionar Dia" name="cboTotalxDiaSemana" id="cboTotalxDiaSemana" onchange="buscarDetalleDiaSemana();" />
						<option value="0" select>-- Seleccionar --</option> 					
					</select>
				</div>
			</div>
	
			<div class="col-sm-2">
				<div class="form-group">				
					<label>Seleccione Fecha:</label>
					<select class="form-control input-sm" title="Seleccionar Consumo a Validar" name="cboDetallexDiaSemana" id="cboDetallexDiaSemana" onchange="buscarDetalleConsumoxDia();" />
						<option value="0" select>-- Seleccionar --</option> 					
					</select>
				</div>
			</div>
			
			<div class="col-sm-4">
				<div class="form-group">				
					<label>Seleccione Consumos Investigados:</label>
					<select class="form-control input-sm" title="Seleccionar Consumo Investigado" name="cboConsumoInvestigado" id="cboConsumoInvestigado" onchange="buscarDetalleConsumoResueltoxDia();" />
						<option value="0" select>-- Seleccionar --</option> 					
					</select>					
				</div>				
			</div>
			
			<div class="col-sm-3">				
			
				<div class="form-group">
				
					<div class="col-sm-4">
						<label for="usr">Reset</label>
						<button id="borrar"  title="Borrar Pantalla" type="submit" class="btn btn-block btn-xs btn-info" onclick="Reset();"><i class="fas fa-recycle fa-2x"></i></button>
					</div>
										
					<div class="col-sm-4">
						<label for="usr">Investigar</label>
						<button id="investigar" title="Enviar Consumo a Investigar" type="submit" class="btn btn-block btn-xs btn-primary" onclick="showMostrarInvestigarHogar();"><i class="fas fa-info-circle fa-2x"></i></button>
					</div>
					
					<div class="col-sm-4">
						<label for="usr">Procesar</label>
						<button id="procesar" title="Procesar Selecciones" type="submit" class="btn btn-block btn-xs btn-success" onclick="buscarImagenFactura();"><i class="fas fa-check fa-2x"></i></button>
					</div>
					
				</div>				
				
			</div>
																		
		</div>
			
		<!-- TABLA: RESUMEN -->
		
		<div class="form-group"> 
			<table class="table table-responsive">			
				<thead>
					<tr>
						<th>Hogares:&nbsp;<span class="label label-default" id="totalHogares">0</span></th>
						<th>Consumos:&nbsp;<span class="label label-info" id="totalConsumos">0</span></th>
						<th>Validados:&nbsp;<span class="label label-success" id="totalValidados">0</span></th>
						<th>Pendientes:&nbsp;<span class="label label-danger" id="totalPendientes">0</span></th>
						<th>Investigados:&nbsp;<span class="label label-primary" id="totalInvestigados">0</span></th>
						<th>Resueltos:&nbsp;<span class="label label-success" id="totalResueltos">0</span></th>
						<th>Alta:&nbsp;<span class="label label-warning" id="altaHogar"></span></th>
						<th>Responsable:&nbsp;<span class="label label-warning" id="responsableHogar"></span></th>
						<th>Celular:&nbsp;<span class="label label-warning" id="celularHogar"></span></th>
					</tr>
				</thead>			
			</table>		 
		</div>
            								
	</div>        
	<hr>
	
	<div class="container-fluid text-center text-primary" id="cargando" style="display:none;">
		<span ><img src="images/ajax-loader7.gif"><strong>&nbsp;Espere, Procesando...!</strong></span>
	</div>
		
	<div class="container-fluid" id="DetalleFactura" style="display:none;">
	
		<div class="form-group-row text-center alert alert-success" role="alert" id="hogarValidado" style="display:none;" >
			<span class="bg-success"><h5><strong>CONSUMO VALIDADO&nbsp;<i class="fas fa-check-double"></i></strong></h5></span>			
		</div>
		
		<div class="form-group-row text-center alert alert-danger" role="alert"  id="hogarEliminado" style="display:none;" >
			<span class="bg-danger"><h5><strong>CONSUMO ELIMINADO&nbsp;<i class="fas fa-times"></i></strong></h5></span>			
		</div>
		
		<div class="form-group-row text-center alert alert-warning" role="alert"  id="hogarInvestigado" style="display:none;" >
			<span class="bg-warning"><h5><strong>..:&nbsp;CONSUMO EN PROCESO DE INVESTIGACI&Oacute;N&nbsp;:..</strong></h5></span>			
			<span class="bg-warning" id="motivo" style="color:red;"></span>
		</div>
		
		<div class="form-group-row alert alert-info" role="alert" id="hogarResuelto" style="display:none;" >
			<h5>
				<strong>
					<p class="text-center"><i class="fas fa-check-double"></i>&nbsp;INVESTIGACI&Oacute;N DE CONSUMO RESUELTO&nbsp;<i class="fas fa-check-double"></i></p>			
					<p id="motivoInv"></p>
					<p id="comentarioInv"></p>
					<p id="motivoRsp" class="text-danger"></p>
				</strong>
			</h5>
		</div>
			
		<div class="form-group-row">
		
			<div class="col-sm-3">
			
				<h4 class="text-danger"><strong><i class="fas fa-file-image"></i>&nbsp;Detalle Factura:</strong></h4>
				<input type="hidden" id="tieneFactura" name="tieneFactura" value="0" />									
								
				<fieldset> 					
					
					<div class="my-funky-img-box">
					   <!--/.col-bdr.left-border-col-->
					   <div class="col-bdr left-border-col"></div>				   
					   <!--/.background-col-->					   
					   <div class="background-col">
							<img class="img-thumbnail" id="imgfactura" name="imgfactura" src="images/loader/cargador1.gif" style="width:480px; height:640px;" />
					   </div>
					   <!--/.col-bdr.left-border-col-->
					   <div class="col-bdr right-border-col"></div>				   
					</div>
					
				</fieldset> 
						
				<!---->
				<br>				
				<div class="form-group">
					<label class="control-label col-sm-3">Canal:</label>
					<div class="col-sm-9">
						<select class="form-control input-sm" name="cboCanal" id="cboCanal" onChange="buscarCadena(this.value);" required>
							<option value="" selected disabled >-- Seleccione -- </option>							
						</select>					  
					</div>
					<div class="error" id="canalErr"></div>
				</div>
				<br>		
				
				<div class="form-group">
					<label class="control-label col-sm-3">Cadena:</label>
					<div class="col-sm-9">
						<select class="form-control input-sm" name="cboCadena" id="cboCadena" required>							
							<option value="" selected disabled >-- Seleccione -- </option>							
						</select>			
					</div>
					<div class="error" id="cadenaErr"></div>	
				</div>
								
				<div class="form-group">
					<strong><p class="text-danger" id="tienefactura"></p></strong>
				</div>
				
				<div class="form-group">
					<label class="control-label col-sm-9">Total Productos Comprados:</label>
					<div class="col-sm-3">
						<input type="number" class="form-control input-sm text-right" id="totalProductos" name="totalProductos" placeholder="Cantidad Productos Comprados" required/>					
					</div>
					<div class="text-right" id="totalProductosErr"></div>	
				</div>
			    	
				<div class="form-group">
					<label class="control-label">Tipo Moneda:</label>
					<div class="">
						<select class="form-control input-sm" name="MonedaPagoFactura" id="MonedaPagoFactura" required>							
							<option value="0" selected disabled >-- Seleccione -- </option>							
						</select>			
					</div>
					<div class="error" id="monedapagofacturaErr"></div>	
				</div>
				
				<div class="form-group">
					<label class="control-label"><strong>Monto Total Factura:</strong></label>
					<div>
						<input type="text" class="form-control input-sm text-right" id="totalFactura" name="totalFactura" placeholder="Monto total de la Factura" onblur="formatMonto(this.value)" required/>					
					</div>
					<div class="text-right" id="totalfacturaErr"></div>	
				</div>								
													
				<div class="form-group">
					<button type="button" title="Grabar" class="btn btn-block btn-primary btn-sm" id="submit" onclick="grabarCambiosFactura();"><i class='fas fa-save'></i> Grabar Cambios</button>
				</div>												
										
			</div>
			
			<!-- TABLA DE PRODUCTOS REGISTRADOS -->
			<div class="col-sm-9">					
			
				<h4 class="text-danger"><strong>Detalle Productos: </strong></h4>
						
				<div class="form-group"  id="tabla-resultados">
					<!-- // ** // -->
					<!-- Matriz de Datos Resultados -->
					<!-- // ** // -->										
				</div>	
				
				<!-- PROMEDIOS SEMANALES X TIPO DE PRODUCTO -->
				<div class="form-group">					
			
					<h4><i class="fas fa-tachometer-alt"></i><strong>&nbsp;Indicador por Semana:</strong></h4>
								
					<div class="table-responsive" id="tabla-resumen">
						<!-- // ** // -->
						<!-- Matriz de Datos Resultados -->
						<!-- // ** // -->										
					</div>		
					
				</div>
				<!-- ./PROMEDIOS SEMANALES X TIPO DE PRODUCTO -->	
				
				
			</div>
			<!-- ./ TABLA DE PRODUCTOS REGISTRADOS -->
										
		</div>
				
	</div>			
	<hr>
	<!-- /.modal Editarconsumo Medicina-Mercado-->
	<div class="modal" id="EditarConsumo" tabindex="-1" data-backdrop="static" data-keyboard="false" role="dialog" aria-labelledby="myModalLabel" data-focus-on="input:first">

		<div class="modal-dialog modal-dialog-centered"  role="document">

			<div class="modal-content">

				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>					
					<h4 class="modal-title">Large Modal</h4>					
				</div>

				<div class="modal-body">
					
					<div class="form-group">
						<span class="text-danger text-center" id="waiting2" style="display:none;"><img src="images/ajax_small.gif">&nbsp;Buscando, Espere....!</span>
						<input type="hidden" class="form-control input-sm text-right" id="txtId" placeholder="...." readonly />
					</div>
											
					<div class="form-group">
						<label>Codigo de Barras:</label>
						<input type="text" class="form-control input-sm text-right" id="txtCodigoBar" placeholder="...." minlength="7" maxlength="16" onkeypress="return onlyNumberKey(event);" required />
						<div class="error" id="codigobarErr"></div>							 
					</div>
					
					<div class="form-group">
						<label class="checkbox-inline"><input type="checkbox" id="chkSinBarras" value="">Producto sin C&oacute;digo de barras</label>
					</div>
										
					<div class="form-group">
						<label class="text-primary">Moneda de Pago:</label>
						<select class="form-control input-sm" name="cboTMonedaPago" id="cboTMonedaPago" onchange="buscarTipoTasadeCambio();"  required>
							<option value="" selected disabled >-- Seleccione -- </option>							
						</select>
						<div class="error" id="txttmonedapagoErr"></div>	
					</div>			
										
					<div class="form-group">
						 <label>Tasa de Cambio:</label>
						 <input type="text" class="form-control input-sm text-right" id="txtTasa" placeholder="...." readonly />							 
					</div>											
					
					<div class="form-group">
						 <label>Cantidad:</label>
						 <input type="text" class="form-control input-sm text-right" id="txtCantidad" step="0.1" placeholder="...." onblur="ActualizarCalculoTotales();" required />
						 <div class="error" id="cantidadErr"></div>							 
					</div>
					
					<div class="form-group">
						 <label>Precio Unitario:</label>
						 <input type="text" class="form-control input-sm text-right" id="txtPrecio" placeholder="...." onblur="ActualizarCalculoTotales();" required />
						 <div class="error" id="precioErr"></div>
					</div>											
					
					<div class="form-group">
						 <label class="text-primary">Total Compra:</label>
						 <input type="text" class="form-control input-sm text-right" id="txtTotal" placeholder="...." readonly />
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
	
	<!-- /.modal AgregarProducto -->		
	<div class="modal" id="AgregarProducto" tabindex="-1" data-backdrop="static" data-keyboard="false" role="dialog" aria-labelledby="myModalLabel" data-focus-on="input:first">

		<div class="modal-dialog modal-dialog-centered"  role="document">

			<div class="modal-content">

				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>					
					<h4 class="modal-title">Large Modal</h4>					
				</div>

				<div class="modal-body">
				
						<div class="form-group">
							<span class="text-danger text-center" id="waiting" style="display:none;"><img src="images/ajax_small.gif">&nbsp;Buscando, Espere....!</span>
							<input type="hidden" class="form-control input-sm text-right" id="txtIdConsumo" placeholder="...." readonly />
						</div>
						
						<div class="form-group">
							<label class="text-primary">Seleccione Categoria:</label>
							<select class="form-control input-sm" name="cboCategoria" id="cboCategoria" required>
								<option value="" selected disabled >-- Seleccione -- </option>							
							</select>
						</div>
						
						<div class="form-group">							
							<div class="col-sm-10">          
								<label class="text-primary">Describa el producto a buscar:</label>
								<input type="text" class="form-control input-sm" id="txtBuscarDescripcion" placeholder="Introduzca todo o parte del producto a buscar...." required />
							</div>
							<div class="col-sm-2">
								<label class="text-pri0mary">&nbsp;</label>
								<button type="button" title="Buscar Producto" class="btn btn-primary btn-sm" onclick="buscarProducto();"><span class="glyphicon glyphicon-search"></span></button>
							</div>
						</div>
												
						<div class="form-group">						 	
							<label class="text-Primary">Seleccione Producto:</label>
							<select class="form-control" name="cboProducto" id="cboProducto" onchange="mostrarBarcode();" required>
								<option value="" selected disabled >-- Seleccione -- </option>							
							</select>
						</div>
						
						<div class="form-group">
							 <label class="text-primary">C&oacute;digo de Barras:</label>
							 <input type="text" class="form-control input-sm text-right" id="txtCodigoBarras" name="txtCodigoBarras" placeholder="...." readonly />
							 <div class="error" id="txtcodigobarErr"></div>	
						</div>
						
						<div class="form-group">
							<label class="text-primary">Seleccione Tipo de Moneda:</label>
							<select class="form-control input-sm" name="cboMonedaPago" id="cboMonedaPago" onchange="buscarTasadeCambio();"  required>
								<option value="" selected disabled >-- Seleccione -- </option>							
							</select>
							<div class="error" id="txtmonedapagoErr"></div>	
						</div>										
						
						<div class="form-group">
							 <label class="text-primary">Tasa de Cambio:</label>
							 <input type="text" class="form-control input-sm text-right" id="txtTasaCambio" placeholder="...." readonly />							 
						</div>											
						
						<div class="form-group">
							 <label class="text-primary">Cantidad:</label>
							 <input type="text" class="form-control input-sm text-right" id="txtCantidadProductos" placeholder="...." onblur="calcularTotales();" required />
							 <div class="error" id="txtcantidadErr"></div>							 
						</div>
						
						<div class="form-group">
							 <label class="text-primary">Precio Unitario:</label>
							 <input type="text" class="form-control input-sm text-right" id="txtPrecioUnitario" placeholder="...."  onblur="calcularTotales();" required />
							 <div class="error" id="txtprecioErr"></div>
						</div>											
						
						<div class="form-group">
							 <label class="text-primary">Total Compra:</label>
							 <input type="text" class="form-control input-sm text-right" id="txtTotalCompra" placeholder="...." readonly />
						</div>											
																					
				</div>
				
				<div class="modal-footer">
					<button type="button" class="btn btn-danger" data-dismiss="modal" title="Salir"><i class='fas fa-sign-out-alt'></i> Salir</button>
					<button type="button" class="btn btn-primary" title="Grabar" onclick="salvarAgregarProductos();" id="btn-salvarProd"><i class='fas fa-save'></i> Grabar</button>
				</div>
				
			</div>
			<!-- /.modal-content -->
		</div>
		<!-- /.modal-dialog -->
	</div>
    <!-- /.modal -->
	
	<!-- /.modal investigarConsumo -->
	<div class="modal" id="investigarConsumo" tabindex="-1" data-backdrop="static" data-keyboard="false" role="dialog" aria-labelledby="myModalLabel" data-focus-on="input:first">

		<div class="modal-dialog modal-dialog-centered"  role="document">

			<div class="modal-content">
				
				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>					
					<h4 class="modal-title">Large Modal</h4>					
				</div>

				<div class="modal-body">
																					
						<div class="form-group">
							<input type="hidden" class="form-control input-sm text-right" id="txtIdConsumoInvestigar" placeholder="...." readonly />
							<label class="text-primary">Seleccione Motivo de la Investigacion:</label>
							<select class="form-control input-sm" name="cboInvestigar" id="cboInvestigar" onchange="enviarInvestigacion();"  required>
								<option value="" selected disabled >-- Seleccione -- </option>							
							</select>
							<div class="error" id="txtinvestigacionErr"></div>	
						</div>
						
						<div class="form-group"> 
        					<label for="comentario">Deje aqu&iacute; comentario adicional:</label>
        					<textarea class="form-control" rows="5" id="txtComentarios" style="resize: none;"></textarea>
      					</div> 
															
				</div>
				
				<div class="modal-footer">
					<button type="button" class="btn btn-danger" data-dismiss="modal" title="Salir"><i class='fas fa-sign-out-alt'></i> Salir</button>
					<button type="button" class="btn btn-primary" title="Grabar" onclick="enviarConsumoInvestigar();" id="btn-investigar"><i class='fas fa-paper-plane'></i> Enviar</button>
				</div>
			</div>
			<!-- /.modal-content -->
		</div>
		<!-- /.modal-dialog -->
	</div>
    <!-- /.modal -->
		
	<!-- /.modal Editar consumo Comida-Juguetes-Electro-Vehiculo-hogar-->
	<div class="modal" id="EditarConsumo2" tabindex="-1" data-backdrop="static" data-keyboard="false" role="dialog" aria-labelledby="myModalLabel" data-focus-on="input:first">

		<div class="modal-dialog modal-dialog-centered"  role="document">

			<div class="modal-content">

				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>					
					<h4 class="modal-title">Large Modal</h4>					
				</div>

				<div class="modal-body">
					
					<div class="form-group">
						<span class="text-danger text-center" id="waiting3" style="display:none;"><img src="images/ajax_small.gif">&nbsp;Buscando, Espere....!</span>
						<input type="hidden" class="form-control input-sm text-right" id="txtId2" placeholder="...." readonly />
					</div>
											
					<div class="form-group">
						<label class="text-primary">Tipo de Comida:</label>
						<select class="form-control input-sm" name="cboTipoComida" id="cboTipoComida" required />
							<option value="" selected disabled >-- Seleccione -- </option>							
						</select>
						<div class="error" id="txttipocomidaErr"></div>	
					</div>			
					
					<div class="form-group">
						 <label>Nombre del Local:</label>
						 <input type="text" class="form-control input-sm text-left" id="txtNombreLocal" placeholder="...." required />							 
						 <div class="error" id="txtnombrelocalErr"></div>	
					</div>											
																										
					<div class="form-group">
						 <label class="text-primary">Total Compra:</label>
						 <input type="text" class="form-control input-sm text-right" id="txtTotalCompra2" placeholder="...." required />
						 <div class="error" id="txttotalcompra2"></div>
					</div>		
					
					<div class="form-group">
						<label class="text-primary">Moneda de Pago:</label>
						<select class="form-control input-sm" name="cboMonedaPagoNoMercado" id="cboMonedaPagoNoMercado" required />
							<option value="" selected disabled >-- Seleccione -- </option>							
						</select>
						<div class="error" id="txtmonedapagonomercadoErr"></div>	
					</div>		
															
				</div>
				
				<div class="modal-footer">
					<button type="button" class="btn btn-danger" data-dismiss="modal" title="Salir"><i class='fas fa-sign-out-alt'></i> Salir</button>
					<button type="button" class="btn btn-primary" title="Grabar" onclick="salvarCambioProductosNoMercado();" id="btn-salvar"><i class='fas fa-save'></i> Grabar</button>
				</div>
			</div>
			<!-- /.modal-content -->
		</div>
		<!-- /.modal-dialog -->
	</div>
    <!-- /.modal -->
	
	<!-- /.modal Editar cambio de moneda masivo -->
	<div class="modal" id="CambioMoneda" tabindex="-1" data-backdrop="static" data-keyboard="false" role="dialog" aria-labelledby="myModalLabel" data-focus-on="input:first">

		<div class="modal-dialog modal-dialog-centered"  role="document">

			<div class="modal-content">

				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>					
					<h4 class="modal-title">Large Modal</h4>					
				</div>

				<div class="modal-body">
				
					<div class="form-group">
						<input type="hidden" class="form-control input-sm text-right" id="txtIdCambio" placeholder="...." readonly />
						<label class="text-primary">Moneda de Pago:</label>
						<select class="form-control input-sm" name="cboCambioMonedaPago" id="cboCambioMonedaPago" required />
							<option value="" selected disabled >-- Seleccione -- </option>							
						</select>
						<div class="error" id="txtcambiomonedapagoErr"></div>	
					</div>		
															
				</div>
				
				<div class="modal-footer">
					<button type="button" class="btn btn-danger" data-dismiss="modal" title="Salir"><i class='fas fa-sign-out-alt'></i> Salir</button>
					<button type="button" class="btn btn-primary" title="Grabar" onclick="salvarCambioMoneda();" id="btn-salvar"><i class='fas fa-save'></i> Grabar</button>
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
<script src="validacion/autoNumeric-1.9.18.js"></script>
<script src="validacion/utilitariosV2.js"></script>
<script src="validacion/crudV36.js"></script>
<script src="validacion/funcionesV36.js"></script>
<script src="validacion/funResueltoV35.js"></script>

<script>
	
	$(document).ready(function() {
		
		$(function($) {
			$('#txtPrecio').autoNumeric('init', {
				lZero: 'deny',
				aSep: '.',
				aDec: ','
			});
			$('#txtPrecioUnitario').autoNumeric('init', {
				lZero: 'deny',
				aSep: '.',
				aDec: ','
			});
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
		$(function() {
			buscarMonedaPagoFactura();
		});						
		$(function() {
			buscarTipoInvestigacion();
		});				
			
		$(function() {
			buscarSemanas();
		});
		
		sessionStorage.setItem('validado',false );
		sessionStorage.setItem('investigado', false );
		sessionStorage.setItem('resuelto', false );							
		sessionStorage.setItem("Convalidado", false );						
						
	});	
	
</script>

