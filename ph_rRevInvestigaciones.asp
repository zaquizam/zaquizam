<!Doctype html>
<!-- ph_rRevInvestigaciones.asp // 13ene21 - 26abr21 -->
<html >
<head>
	<title>Revisar Investigaciones</title>
	<meta charset="UTF-8">
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
	<link href="favicon.ico" rel="icon"  type="image/x-icon">
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
' 29dic20 - 02feb21
'========================================================================================== 
'Variables y Constantes 
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
				
			<div class="col-sm-4">
				<div class="form-group">				
					<label>Seleccione Hogar:</label><span id="loader"></span>
					<select class="form-control input-sm" title="Seleccionar Hogar" name="cboHogar" id="cboHogar" />
						<option value="0" select>-- Seleccionar --</option> 					
					</select>
				</div>
			</div>
			
			<div class="col-sm-2">
				<div class="form-group">				
					<label>Seleccione Fecha:</label>	
					<select class="form-control input-sm" title="Seleccionar Fecha Consumo" name="cboIdConsumo" id="cboIdConsumo" />
						<option value="0" select>-- Seleccionar --</option> 					
					</select>
				</div>
			</div>
	
			<div class="col-sm-2">
				<div class="form-group">				
					<label>Semana:</label>	
					<input type="text" class="form-control input-sm" title="Semana" name="txtSemana" id="txtSemana" readonly />
				</div>
			</div>
												
			<div class="col-sm-2">
				<div class="form-group">				
					<label>Area:</label>
					<input type="text" class="form-control input-sm" title="Area" name="txtArea" id="txtArea" readonly />
				</div>
			</div>
			
			<div class="col-sm-2">
				<div class="form-group">				
					<label>Estado:</label>
					<input type="text" class="form-control input-sm" title="Estado" name="txtEstado" id="txtEstado" readonly />
				</div>
			</div>
					
						
		</div>
		
		<!---->
		
		<div class="form-group">
		
			<div class="col-sm-2">
				<div class="form-group">				
					<label>Tipo Consumo:</label>
					<input type="text" class="form-control input-sm" title="Tipo Consumo" name="txtTipoConsumo" id="txtTipoConsumo" readonly />					
				</div>
			</div>		
			
			<div class="col-sm-4">
				<div class="form-group">				
					<label>Motivo a Investigar:</label>
					<input type="text" class="form-control input-sm" title="Motivo de Investigacion" name="txtMotivoInvestigar" id="txtMotivoInvestigar" readonly />
				</div>
			</div>
			
			<div class="col-sm-4">
				<div class="form-group">				
					<label>Comentarios Adicionales a Investigar:</label>
					<textarea class="form-control" rows="4" id="txtComentarioAdicional" name="txtComentarioAdicional" style="resize: none;" readonly /></textarea>
				</div>
			</div>
						
			<div class="col-sm-2">				
			
				<div class="form-group">
				
					<div class="col-sm-6">
						<label for="usr">Reset</label>
						<button id="borrar"  title="Borrar Pantalla" type="submit" class="btn btn-block btn-xs btn-info" onclick="Reset();"><i class="fas fa-recycle fa-2x"></i></button>
					</div>
										
					<div class="col-sm-6">
						<label for="usr">Responder</label>
						<button id="investigar" title="Enviar Respuesta del Panelista" type="submit" class="btn btn-block btn-xs btn-success" onclick="resultadoInvestigacionHogar();"><i class="fas fa-check fa-2x"></i></button>
					</div>					
					<!--
					<div class="col-sm-3">
						<label for="usr">Anterior</label>
						<button id="prev"  title="Fecha Anterior" type="submit" class="btn btn-block btn-xs btn-info"><i class="fas fa-backward fa-2x"></i></button>
					</div>
					
					<div class="col-sm-3">
						<label for="usr">Siguiente</label>
						<button id="next"  title="Fecha Siguiente" type="submit" class="btn btn-block btn-xs btn-success"><i class="fas fa-forward fa-2x"></i></button>
					</div>
					-->
					
				</div>				
				
			</div>
																		
		</div>
			
		<!-- TABLE: RESUMEN -->
        <div class="form-group"> 
		
	  		<table class="table no-margin">			
				<thead>
					<tr>
						<th ><i class="fas fa-home"></i>&nbsp;Hogares por Investigar:&nbsp;<span class="label label-primary" id="totalHogares">0</span></th>
						<!--<th ><i class="fas fa-calculator"></i>&nbsp;Consumos:&nbsp;<span class="label label-info" id="totalConsumos">0</span></th>-->
						<!--<th ><i class="fas fa-check"></i>&nbsp;Validados:&nbsp;<span class="label label-success" id="totalValidados">0</span></th>-->
						<th ><i class="fas fa-eye-slash"></i>&nbsp;Consumos Pendientes por Hogar:&nbsp;<span class="label label-danger" id="totalPendientes">0</span></th>
						<th ><i class="fas fa-calendar-check"></i>&nbsp;Alta Hogar:&nbsp;<span class="label label-info" id="altaHogar"><i class="fas fa-calendar-day"></i></span></th>
						<th ><i class="fas fa-user-check"></i>&nbsp;Responsable Hogar:&nbsp;<span class="label label-warning" id="responsableHogar"><i class="fas fa-user-check"></i></span></th>
						<th ><i class="fas fa-phone-square-alt"></i>&nbsp;Celular Hogar:&nbsp;<span class="label label-warning" id="celularHogar"><i class="fas fa-phone-square-alt"></i></span></th>
					</tr>
				</thead>			
			</table>		 
			
		</div>
             								
	</div>        
	<hr>
	
	<div class="container-fluid" id="DetalleFactura" style="display:none;">
	
		<div class="form-group-row text-center alert alert-success" role="alert" id="hogarValidado" style="display:none;" >
			<span class="bg-success"><h5><strong>CONSUMO VALIDADO&nbsp;<i class="fas fa-check"></i></strong></h5></span>			
		</div>
		<!--
		<div class="form-group-row text-center alert alert-danger" role="alert"  id="hogarEliminado" style="display:none;" >
			<span class="bg-danger"><h5><strong>CONSUMO ELIMINADO&nbsp;<i class="fas fa-times"></i></strong></h5></span>			
		</div>
		
		<div class="form-group-row text-center alert alert-danger" role="alert"  id="hogarInvestigado" style="display:none;" >
			<span class="bg-warning"><h5><strong><i class="fas fa-eye"></i>&nbsp;CONSUMO EN PROCESO DE INVESTIGACION&nbsp;<i class="fas fa-eye"></i></strong></h5></span>			
		</div>
		-->
		
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
													
				<!--
				<div class="form-group">
					<button type="button" title="Grabar" class="btn btn-block btn-primary btn-sm" id="submit" onclick="grabarCambiosFactura();"><i class='fas fa-save'></i> Grabar Cambios</button>
				</div>
				-->

			</div>
			
			<!-- TABLA DE PRODUCTOS REGISTRADOS -->
			<div class="col-sm-9">					
			
				<h4 class="text-danger"><strong>Detalle Productos: </strong></h4>
						
				<div class="table-responsive" id="tabla-resultados">
					<!-- // ** // -->
					<!-- Matriz de Datos Resultados -->
					<!-- // ** // -->										
				</div>	
				
				<!-- PROMEDIOS SEMANALES X TIPO DE PRODUCTO -->
				<!--
				<div class="form-group">					
			
					<h4 class="text-danger"><strong>Resumen Semanal:</strong></h4>
								
					<div class="table-responsive" id="tabla-resumen">
						
					</div>		
					
				</div>
				-->
				<!-- ./PROMEDIOS SEMANALES X TIPO DE PRODUCTO -->	
				
				
			</div>
<!-- ./ TABLA DE PRODUCTOS REGISTRADOS -->

		</div>
				
</div>			
	<hr>				
	<!-- /.modal Respuesta de investigar -->
	<div class="modal" id="responderInvestigacion" tabindex="-1" data-backdrop="static" data-keyboard="false" role="dialog" aria-labelledby="myModalLabel" data-focus-on="input:first">

		<div class="modal-dialog modal-dialog-centered"  role="document">

			<div class="modal-content">
				
				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>					
					<h4 class="modal-title">Large Modal</h4>					
				</div>

				<div class="modal-body">
						
						<div class="form-group">				
							<label>Motivo a Investigar:</label>
							<input type="text" class="form-control input-sm" title="Motivo Investigacion" name="txtMotivoInvestigacion" id="txtMotivoInvestigacion" readonly />
						</div>
						
						<div class="form-group"> 
        					<label for="comentario">Comentario Adicional a Investigar:</label>
        					<textarea class="form-control" rows="10" id="txtPregunta" style="resize: none;" readonly /></textarea>
      					</div> 			
																		
						<div class="form-group"> 
        					<label for="comentario">Deje aqu&iacute; la respuesta del panelista:</label>
        					<textarea class="form-control" rows="10" id="txtRespuesta" style="resize: none;"></textarea>
      					</div> 						
															
				</div>
				
				<div class="modal-footer">
					<button type="button" class="btn btn-danger" data-dismiss="modal" title="Salir"><i class='fas fa-sign-out-alt'></i> Salir</button>
					<button type="button" class="btn btn-primary" title="Grabar" onclick="enviarRespuestaInvestigacion();" id="btn-investigar"><i class='fas fa-paper-plane'></i> Enviar</button>
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
<script src="investigar/utilitarios.js"></script>
<script src="investigar/funcionesV8.js"></script>

<script>
							
	$(function() {
		debugger;
		llenarCmbHogaresInvestigados();
	});

</script>