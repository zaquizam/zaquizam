<!-- /.modal Editar Moneda -->
	<div class="modal" id="EditarMoneda" tabindex="-1" data-backdrop="static" data-keyboard="false" role="dialog" aria-labelledby="myModalLabel" data-focus-on="input:first">

		<div class="modal-dialog modal-dialog-centered"  role="document">

			<div class="modal-content">

				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>					
					<h4 class="modal-title">Large Modal</h4>					
				</div>

				<div class="modal-body">
					
					<div class="form-group">
						<input type="text" class="form-control input-sm text-right" id="txtIdConsumo" placeholder="...." readonly />
					</div>
											
					<div class="form-group">
						<label>Codigo de Barras:</label>
						<input type="text" class="form-control input-sm text-right" id="txtCodigoBar" placeholder="...." readonly />
					</div>
					
					<div class="form-group">
						 <label>Moneda Actual:</label>
						 <input type="text" class="form-control input-sm text-right text-danger" id="txtMoneda" placeholder="...." readonly />							 
					</div>											
																			
					<div class="form-group">
						<label class="text-primary">Seleccione Moneda de Pago:</label>
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
						 <input type="text" class="form-control input-sm text-right" id="txtCantidad"  placeholder="...." readonly />
					</div>
					
					<div class="form-group">
						 <label>Precio Unitario:</label>
						 <input type="text" class="form-control input-sm text-right" id="txtPrecio" placeholder="...." readonly />
					</div>											
					
					<div class="form-group">
						 <label class="text-primary">Total Compra:</label>
						 <input type="text" class="form-control input-sm text-right" id="txtTotal" placeholder="...." readonly />
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