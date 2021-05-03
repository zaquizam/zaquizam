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
'29dic20
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
			
			<div class="col-sm-3">
				<div class="form-group">				
					<label>Seleccione Estado:</label>
					<select class="form-control input-sm" title="Seleccionar Estado" name="cboEstado" id="cboEstado" onchange="buscarHogar();" />
						<option value="0" select>-- Seleccionar --</option> 					
					</select>
				</div>
			</div>
						
			<div class="col-sm-3">
				<div class="form-group">				
					<label>Seleccione Hogar:</label>
					<select class="form-control input-sm" title="Seleccionar Hogar" name="cboHogar" id="cboHogar" onchange="buscarTipoConsumo();" />
						<option value="0" select>-- Seleccionar --</option> 					
					</select>
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
		
	<div class="container-fluid" id="DetalleFactura">
			
		<div class="form-group-row">
		
			<div class="col-sm-3">
			
				<h4 class="text-danger"><strong>Factura:</strong></h4>
				<input type="hidden" id="tieneFactura" name="tieneFactura" value="0" />									
				
				<!--
				<div class="thumbnail">				
					<img class="img-responsive" id="imgfactura" name="imgfactura" src="images/loader/cargador4.gif" style="width:600px; height:720px;" />						
				</div>
				-->
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
						
				<div class="form-group">				
					<div class="text-primary"><strong><p id="canal"></p></strong></div>
					<div class="text-primary"><strong><p id="cadena"></p></strong></div>
					<div class="text-primary"><strong><p id="tienefactura"></p></strong></div>				
					<div class="text-primary"><strong><p id="montoFactura"></p></strong></div>
				</div>		
				
				<!-- Monto Total Factura-->
				<div class="form-group" id="verFactura" style="display:none;">
					<span class="text-danger"><h5><strong>Introduzca Monto Total Factura:</strong></h5></span>					
					<input type="text" class="form-control input-sm text-right" id="totalFactura" name="totalFactura" style="width:300px;" placeholder="Monto total de la Factura" onblur="formatMonto(this.value)" />					
					<div class="error" id="totalfacturaErr"></div>	
				</div>		
										
			</div>
			
			<div class="col-sm-9">					
			
				<h4 class="text-danger"><strong>Detalle Productos:</strong></h4>
								
				<div class="table-responsive" id="tabla-resultados">
					<!-- // ** // -->
					<!-- Matriz de Datos Resultados -->
					<!-- // ** // -->
					...						
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
					<!--
					<button type="button" class="close" data-dismiss="modal" aria-label="Close">
					<span aria-hidden="true">&times;</span>
					</button>
					-->
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
					<button type="button" class="btn btn-primary" title="Grabar" onclick="salvarCambio();" id="btn-save"><i class='fas fa-save'></i> Grabar</button>
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
<script src="validacion/funciones.js"></script>

<script>
	//jQuery.noConflict();
	$(document).ready(function() {
		$(function() {
			buscarSemanas();
		});
	});		
</script>
