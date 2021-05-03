//
// CrudV38.JS // 05ENE21 - 14abr21 *
//
function obtener_DetallexProducto(id) {
	//
	debugger;
	//	
	var idConsumoDetalle	=	id;	
	var idConsumo =	$("#cboDetallexDiaSemana").val();
	//
	if (idConsumo==="0" || idConsumo==null || idConsumo==undefined){
		// Ajustar y Validar consumos enviados a Investigar
		/* Comida - Vehiculos - Hogar - Electro */
		//
		buscarTipoComida();
		buscarMonedaPagoFacturaNoMercado();
		var idConsumo =	$("#cboConsumoInvestigado").val();		
		//			
		$.ajax({
			url: "g_ValBuscarDetallesxProductosxUnicoNoreposicionMercado.asp",
			type: "GET",
			cache: false,
			async: false,
			dataType: "json",
			data: {id_Consumo : idConsumo},
			beforeSend: function(){
				$("#loader").html("<img src='images/ajax_small.gif'> Buscando Detalles...!");
			}
		})
		.done(function(data) {
			debugger;
			console.log(data);
			$("#loader").html("");		
			//
			var idtipocomida= parseInt(data[0].idtipocomida);
			var total		= Number(data[0].totalcompra).toLocaleString("es-ES", {minimumFractionDigits: 2}); //data[0].precio;
			var idmoneda	= parseInt(data[0].idmoneda);
			//	
			$("#cboTipoComida").val(idtipocomida);
			$("#cboTipoComida").trigger("change"); 
			$("#cboMonedaPagoNoMercado").val(idmoneda);
			$("#cboMonedaPagoNoMercado").trigger("change"); 
			$("#txtNombreLocal").val(data[0].nombrelocal);			
			$("#txtTotalCompra2").val(total);					
			$("#txtMonedaPago").val(data[0].moneda);								
			$("#txtId2").val(idConsumo);
			//			
			ActualizarCalculoTotales();
			//
			$("#EditarConsumo2").modal("show");
			$(".modal-title").html("<i class='fas fa-edit'></i> Actualizar Consumo");		
			//				
		})
		/*Si la consulta Fallo*/
		.fail(function() {
			alert("Fallo - oDxP(1)");
		},'HTML');	
		//		
	} else {	
		//
		// Ajustar y Validar Consumos 
		//
		var idTipConsumo	=	$("#cboTipoConsumo").val();	
		//	
		if (idTipConsumo === "1" || idTipConsumo === "8") {
			//	Medicinas y Mercado de Reposicion
			buscarTipoMonedaPago();
			//		
			$.ajax({
				url: "g_ValBuscarDetallesxProductosxUnico.asp",
				type: 'GET',
				cache: false,
				async: false,
				dataType: 'json',
				data: {id_ConsumoDetalle : idConsumoDetalle},
				beforeSend: function(){
					$("#loader").html("<img src='images/ajax_small.gif'> Buscando Detalles...!");
				}
			})
			.done(function(data) {
				//debugger;
				console.log(data);
				$("#loader").html("");		
				//
				var precio 	 = Number(data[0].precio).toLocaleString("es-ES", {minimumFractionDigits: 2}); //data[0].precio;
				var cantidad = data[0].cantidad;
				var barcode  = data[0].barcode;
				var tasa	 = Number(data[0].tasa).toLocaleString("es-ES", {minimumFractionDigits: 2}); //data[0].precio;
				var moneda	 = data[0].moneda;
				var idmoneda = parseInt(data[0].idmoneda);
				//	
				$("#txtPrecio").val(precio);
				$("#txtCantidad").val(cantidad);
				$("#txtCodigoBar").val(barcode);		
				$("#txtTasa").val(tasa);
				$("#txtId").val(idConsumoDetalle);		
				$("#cboTMonedaPago").val(idmoneda);	
				$("#chkSinBarras").prop("checked", false);
				$("#txtCodigoBar").prop("disabled", false);
				ActualizarCalculoTotales();
				//
				$("#EditarConsumo").modal("show");
				$(".modal-title").html("<i class='fas fa-edit'></i> Actualizar Consumo");		
				//				
			})
			/*Si la consulta Fallo*/
			.fail(function() {
				alert("Fallo - oDxP(2)");
			},'HTML');			
		
		} else {
			//
			// Ajustar y Validar consumos enviados a Investigar
			/* Comida - Vehiculos - Hogar - Electro */
			//
			buscarTipoComida();
			buscarMonedaPagoFacturaNoMercado();
			var idConsumo =	$("#cboConsumoInvestigado").val();			
			//
			if (idConsumo==="0" || idConsumo==null || idConsumo==undefined){
				var idConsumo =	$("#cboDetallexDiaSemana").val();
			}
			//					
			$.ajax({
				url: "g_ValBuscarDetallesxProductosxUnicoNoreposicionMercado.asp",
				type: "GET",
				cache: false,
				async: false,
				dataType: "json",
				data: {id_Consumo : idConsumo},
				beforeSend: function(){
					$("#loader").html("<img src='images/ajax_small.gif'> Buscando Detalles...!");
				}
			})
			.done(function(data) {
				debugger;
				console.log(data);
				$("#loader").html("");		
				//
				var idtipocomida= parseInt(data[0].idtipocomida);
				var total		= Number(data[0].totalcompra).toLocaleString("es-ES", {minimumFractionDigits: 2}); //data[0].precio;
				var idmoneda	= parseInt(data[0].idmoneda);
				//	
				$("#cboTipoComida").val(idtipocomida);
				$("#cboTipoComida").trigger("change"); 
				$("#cboMonedaPagoNoMercado").val(idmoneda);
				$("#cboMonedaPagoNoMercado").trigger("change"); 
				$("#txtNombreLocal").val(data[0].nombrelocal);			
				$("#txtTotalCompra2").val(total);					
				$("#txtMonedaPago").val(data[0].moneda);								
				$("#txtId2").val(idConsumo);
				//			
				ActualizarCalculoTotales();
				//
				$("#EditarConsumo2").modal("show");
				$(".modal-title").html("<i class='fas fa-edit'></i> Actualizar Consumo");		
				//				
			})
			/*Si la consulta Fallo*/
			.fail(function() {
				alert("Fallo - oDxP(3)");
			},'HTML');	
	
		}
	
	}
}
//
function preformatFloat(float){
   if(!float){
      return '';
   };

   //Index of first comma
   const posC = float.indexOf(',');

   if(posC === -1){
      //No commas found, treat as float
      return float;
   };

   //Index of first full stop
   const posFS = float.indexOf('.');

   if(posFS === -1){
      //Uses commas and not full stops - swap them (e.g. 1,23 --> 1.23)
      return float.replace(/\,/g, '.');
   };

   //Uses both commas and full stops - ensure correct order and remove 1000s separators
   return ((posC < posFS) ? (float.replace(/\,/g,'')) : (float.replace(/\./g,'').replace(',', '.')));
};
//
function salvarCambioProductos() {
	//
	debugger
	//
	if(validarAjustesProductos()){
	  //
	  swal(
		{
		  title: "Estan Correctos Todos",
		  text: ".. los ajustes realizados ?",
		  type: "warning",
		  showCancelButton: true,
		  confirmButtonClass: "btn-danger",
		  confirmButtonText: "Si",
		  cancelButtonText: "No",
		  closeOnConfirm: true,
		  showLoaderOnConfirm: true,
		},
		function () {
			//
			debugger;
			//				
			var Precio = parseFloat(preformatFloat($("#txtPrecio").val()));
			var TasaCambio = parseFloat(preformatFloat($("#txtTasa").val()));
			var TotalCompra = parseFloat(preformatFloat($("#txtTotal").val()));
			//
			let ajax = {
				idConsumoDetalle	:	$("#txtId").val(),				
				cantidad			:	$("#txtCantidad").val(),
				barcode				:	$("#txtCodigoBar").val(),
				idConsumo			:	$("#cboDetallexDiaSemana").val(),
				idMoneda			:	$("#cboTMonedaPago").val(),
				moneda				:	$("#cboTMonedaPago option:selected" ).text().trim(),
				precio 				:	Precio,
				tasa				:	TasaCambio,
				total				:	TotalCompra,
			};				
			//
			$.ajax({		
				url: "g_ValUpdateDetallesxProductosxUnico.asp",
				type: 'GET',
				cache: false,
				async: false,
				data: ajax,
				//dataType: "json",
				beforeSend: function(objeto){
					$("#loader").html("<img src='images/ajax_small.gif'> Espere, Grabando Ajustes..!");
				},
				success: function (data) {
					debugger;
					console.log(data);				
					$("#EditarConsumo").modal("hide");					
					$("#loader").html("");
					swal("Aviso..!", "Producto Actualizado...!", "success");
					buscarDetallexProductoFactura();
					//
					CalcularTotalesConsumos();
				},
			});
			//
		}
	  );

	}else{
	  //swal("Aviso..!", "Hay Errores revise los mensajes...!", "error");
	}				
}
//
function grabarCambiosFactura() {
	//
	 debugger;
	//
	if(validarAjustesFactura()){
	  //
	  swal(
		{
		  title: "Estan Correctos Todos",
		  text: ".. los ajustes realizados ?",
		  type: "warning",
		  showCancelButton: true,
		  confirmButtonClass: "btn-danger",
		  confirmButtonText: "Si",
		  cancelButtonText: "No",
		  closeOnConfirm: true,
		  showLoaderOnConfirm: true,
		},
		function () {			
			//
			let ajax = {
				idConsumo		:	$("#cboDetallexDiaSemana").val(),
				canal 			:	$("#cboCanal").val(),
				cadena			:	$("#cboCadena").val(),
				idMoneda		:	$("#MonedaPagoFactura").val(),
				totalFact		:	parseFloat(preformatFloat($("#totalFactura").val())),										
				totalProd		:	parseFloat(preformatFloat( $("#totalProductos").val())),
				
			};				
			//			
			$.ajax({		
				url: "g_ValUpdateDetallesxFacturaxUnico.asp",
				type: 'GET',
				cache: false,
				async: false,
				data: ajax,
				beforeSend: function(objeto){
					$("#loader").html("<img src='images/ajax_small.gif'> Espere, Grabando Ajustes..!");
				},
				success: function (data) {
					debugger;
					console.log(data);									
					$("#loader").html("");
					if(data==="True"){
						swal("Aviso..!", "Factura Actualizada...!", "success");						
					}
				},
			});			
		}
	  );

	}else{
	  swal("Aviso..!", "Hay Errores revise los mensajes...!", "error");
	}				
}
//
function validarAjustesFactura() {
	//	
	var Error = 0;
	//
	$("#btn-salvar").attr("disabled", true);
	//
	// Tiene Factura	
	let tieneFactura = $("#tieneFactura").val().trim();  
	// Total Factura		
	let totalFactura = $("#totalFactura").val().trim();  
	if (totalFactura == null || totalFactura == "" || totalFactura.length == 0 || totalFactura == undefined ) {
		$("#totalfacturaErr").html("<span style='color:red;'>Monto Factura esta vacio..!</span>");
		Error++;
	} else {
		//
		totalFactura = totalFactura.replace(/[.]/g, "");
		totalFactura = totalFactura.replace(/[,]/g, ".");
		//
		let regex = /^[0-9.,]+$/;
		if (regex.test(totalFactura) === false) {
		  $("#totalfacturaErr").html("<span style='color:red;'>Introduzca un Monto Factura valido!</span>");
		  swal("Aviso..!", "Introduzca un Monto Factura valido!", "error");
		  Error++;
		} else {
		  if (parseFloat(totalFactura) <= 0) {
			$("#totalfacturaErr").html("<span style='color:red;'>Monto Factura debe ser mayor a cero!</span>");
			swal("Aviso..!", "Monto Factura debe ser mayor a cero!", "error");
			Error++;
		  } else {
			$("#totalfacturaErr").html("");
		  }
		}
	}	
	// Canal
	let cmbCanal = document.getElementById("cboCanal").selectedIndex;
	if (cmbCanal == null || cmbCanal == 0 || cmbCanal < 0) {
	  $("#canalErr").html("<span style='color:red;'>Seleccione Canal!</span>");
	  Error++;
	} else {
	  $("#canalErr").html("");
	}
	// Cadena
	let cmbCadena = document.getElementById("cboCadena").selectedIndex;
	if (cmbCadena == null || cmbCadena == undefined) {
	  $("#cadenaErr").html("<span style='color:red;'>Seleccione Cadena!</span>");
	  Error++;
	} else {
	  $("#cadenaErr").html("");
	}
	//
	// Moneda Pago Factura
	let cmbMonedaPagoFactura = document.getElementById("MonedaPagoFactura").selectedIndex;
	if (cmbMonedaPagoFactura == null || cmbMonedaPagoFactura == undefined || cmbMonedaPagoFactura == 0) {
	  $("#monedapagofacturaErr").html("<span style='color:red;'>Seleccione Moneda Pago!</span>");
	  Error++;
	} else {
	  $("#monedapagofacturaErr").html("");
	}
	//
	let totalProductos = $("#totalProductos").val().trim();  
	if (totalProductos == null || totalProductos== "" || totalProductos.length == 0 || totalProductos == undefined ) {
		$("#totalProductosErr").html("<span style='color:red;'>Total Productos esta vacio..!</span>");
		Error++;
	} else {
		//
		totalProductos = totalProductos.replace(/[.]/g, "");
		totalProductos = totalProductos.replace(/[,]/g, ".");
		//
		let regex = /^[0-9.,]+$/;
		if (regex.test(totalProductos) === false) {
		  $("#totalProductosErr").html("<span style='color:red;'>Introduzca un Total valido!</span>");
		  swal("Aviso..!", "Introduzca un Monto Factura valido!", "error");
		  Error++;
		} else {
		  if (parseFloat(totalProductos) <= 0) {
			$("#totalProductosErr").html("<span style='color:red;'>Total debe ser mayor a cero!</span>");
			swal("Aviso..!", "Monto Factura debe ser mayor a cero!", "error");
			Error++;
		  } else {
			$("#totalProductosErr").html("");
		  }
		}
	}	
	//
	if (Error == 0) {
		$("#btn-salvar").attr("disabled", false);
		return true;
	} else {
		$("#btn-salvar").attr("disabled", false);
		return false;
	}
	//
}
//
function validar_Directo(id) {
	//
	debugger;
	//
	var idConsumo =	$("#cboDetallexDiaSemana").val();
	//
	if (idConsumo==="0" || idConsumo==null || idConsumo==undefined){
		var idConsumo =	$("#cboConsumoInvestigado").val();			
	}	
	//
	swal({
	  title: "¿ Esta Seguro de Validar ?",
	  text: ".. ",
	  type: "warning",
	  showCancelButton: true,
	  confirmButtonClass: "btn-danger",
	  confirmButtonText: "Si",
	  cancelButtonText: "No",
	  closeOnConfirm: true,
	  showLoaderOnConfirm: true,
	},
	function () {
		//				
		let ajax = {				
			idConsumoDetalle	:	id,
			idConsumo			:	idConsumo,
		};				
		//
		$.ajax({		
			url: "g_ValUpdateDetallesxProductosxUnicoDirecto.asp",
			type: 'GET',
			cache: false,
			async: false,
			data: ajax,
			beforeSend: function(objeto){
				$("#loader").html("<img src='images/ajax_small.gif'> Espere, Validando..!");
			},
			success: function (data) {
				swal("Aviso..!", "Producto Validado...!", "success");
				debugger;
				console.log(data);				
				$("#loader").html("");
				if(data==="0"){
					$("#hogarValidado").css("display", "block");
					sessionStorage.setItem("Convalidado", 1 );
				}else{
					sessionStorage.setItem("Convalidado", 0 );
					$("#hogarValidado").css("display", "none");
				}
				//
				if($("#cboDetallexDiaSemana").val()==="0"){
					buscarDetallexProductoFacturaResuelto();
				}else{
					buscarDetallexProductoFactura();						
				}
				CalcularTotalesConsumos();
			},
		});
		//
	}
  );
	
}
//
function validarMasivo(){
	//
	debugger;
	//	
	var idConsumo =	$("#cboDetallexDiaSemana").val();
	//
	if (idConsumo==="0" || idConsumo==null || idConsumo==undefined){
		var idConsumo =	$("#cboConsumoInvestigado").val();			
	}	
	//
	swal({
	  title: "¿ Seguro de Validarlo Todo ?",
	  text: ".. ",
	  type: "warning",
	  showCancelButton: true,
	  confirmButtonClass: "btn-danger",
	  confirmButtonText: "Si",
	  cancelButtonText: "No",
	  closeOnConfirm: true,
	  showLoaderOnConfirm: true,
	},
	function () {
		//				
		let ajax = {							
			idConsumo : idConsumo,
			//idResuelto: sessionStorage.getItem('resuelto'),
		};				
		//
		$.ajax({
			url: "g_ValUpdateDetallesxProductosMasivo.asp",
			type: 'GET',
			cache: false,
			async: false,
			data: ajax,
			beforeSend: function(objeto){
				$("#loader").html("<img src='images/ajax_small.gif'> Espere, Validando..!");
			},
		})
		/*Si la consulta se realizo con exito*/
		.done(function(data) {
			swal("Aviso..!", "Producto Validado...!", "success");
			debugger;
			console.log(data);				
			$("#loader").html("");
			if(data==="0"){
				$("#hogarValidado").css("display", "block");
				sessionStorage.setItem("Convalidado", 1 );
			}else{
				$("#hogarValidado").css("display", "none");
				sessionStorage.setItem("Convalidado", 0 );
			}
			if($("#cboDetallexDiaSemana").val()==="0"){
				buscarDetallexProductoFacturaResuelto();
			}else{
				buscarDetallexProductoFactura();
			}
			CalcularTotalesConsumos();
		})
		/*Si la consulta Fallo*/
		.fail(function() {
			$("#loader").html("");	
			swal("Algo salio mal.!","Intentelo de nuevo..! vM()", "error");
		},'html')
		//				
	}
  );		
}
//
function eliminar_Status_Producto(id) {
	//
	debugger;
	//
	var idConsumo =	$("#cboDetallexDiaSemana").val();
	//
	if (idConsumo==="0" || idConsumo==null || idConsumo==undefined){
		var idConsumo =	$("#cboConsumoInvestigado").val();			
	}	
	//
	swal({
	  title: "¿ Esta Seguro de Eliminar ?",
	  text: "El status del Producto.. ",
	  type: "warning",
	  showCancelButton: true,
	  confirmButtonClass: "btn-danger",
	  confirmButtonText: "Si",
	  cancelButtonText: "No",
	  closeOnConfirm: true,
	  showLoaderOnConfirm: true,
	},
	function () {
		//				
		let ajax = {				
			idConsumoDetalle	:	id,
			idConsumo			:	idConsumo,
		};				
		//
		$.ajax({		
			url: "g_ValUpdateEliminarStatusProductos.asp",
			type: 'GET',
			cache: false,
			async: false,
			data: ajax,
			beforeSend: function(objeto){
				$("#loader").html("<img src='images/ajax_small.gif'> Espere, Validando..!");
			},
			success: function (data) {
				swal("Aviso..!", "Status Producto Eliminado.!", "success");
				debugger;
				console.log(data);				
				$("#loader").html("");
				if(data==="0"){
					$("#hogarValidado").css("display", "block");
					sessionStorage.setItem("Convalidado", 1 );
				}else{
					$("#hogarValidado").css("display", "none");
					sessionStorage.setItem("Convalidado", 0 );
				}
				buscarDetallexProductoFactura();				
				CalcularTotalesConsumos();
			},
		});
		//
	}
  );
	
}
//
function validarAjustesProductos() {
	//	
	var Error = 0;
	//	
	// Nro del Codigo de Barras
	//	
	//
	let barcode = $("#txtCodigoBar").val().trim();
	if (barcode == null || barcode == "" || barcode.length == 0 || barcode == undefined ) {
		$("#codigobarErr").html("<span style='color:red;'>Codigo de barras esta vacio o en blanco..!</span>");
		Error++;
	}else {
		let regex = /^[0-9]+$/;
		if (regex.test(barcode) === false) {
			$("#codigobarErr").html("<span style='color:red;'>Introduzca una Codigo de barras valido (Solo numeros)!</span>");
			Error++;
		} else {
			var chk = document.getElementById("chkSinBarras").checked;
			if (parseFloat(barcode) <= 0 && chk===false) {
				$("#codigobarErr").html("<span style='color:red;'>Codigo de barras invalido..!</span>");
				Error++;
			} else {
				$("#codigobarErr").html("");
			}
			if (barcode.length < 7 || barcode.length > 16) {
				$("#codigobarErr").html("<span style='color:red;'>Codigo de barras errado, Min 7 y Max 16 Caracteres..!</span>");
				Error++;
			} else {
				$("#codigobarErr").html("");
			}
		}
	}	
	//
	// Cantidad
	//
	let cantidad = $("#txtCantidad").val().trim();
	if (cantidad == null || cantidad == "" || cantidad.length == 0 || cantidad == undefined ) {
		$("#cantidadErr").html("<span style='color:red;'>Cantidad esta vacio o en blanco..!</span>");
		Error++;
	} else {
		cantidad = cantidad.replace(/[.]/g, "");
		cantidad = cantidad.replace(/[,]/g, ".");
		let regex = /^[0-9.,]+$/;
		if (regex.test(cantidad) === false) {
		  $("#cantidadErr").html("<span style='color:red;'>Introduzca una Cantidad valida!</span>");
		  Error++;
		} else {
		  if (parseFloat(cantidad) <= 0) {
			  $("#cantidadErr").html("<span style='color:red;'>Cantidad debe ser mayor a cero!</span>");
			Error++;
		  } else {
			$("#cantidadErr").html("");
		  }
		}
	}
	//
	// Precio
	let precio = $("#txtPrecio").val().trim();  
	if (precio == null || precio == "" || precio.length == 0 || precio == undefined ) {
		$("#precioErr").html("<span style='color:red;'>Precio esta vacio o en blanco..!</span>");
		Error++;
	} else {
		//		
		precio = precio.replace(/[.]/g, "");
		precio = precio.replace(/[,]/g, ".");
		let regex = /^[0-9.,]+$/;
		if (regex.test(precio) === false) {
		  $("#precioErr").html("<span style='color:red;'>Introduzca una Precio valido!</span>");
		  Error++;
		} else {
		  if (parseFloat(precio) <= 0) {
			$("#precioErr").html("<span style='color:red;'>Precio debe ser mayor a cero!</span>");
			Error++;
		  } else {
			$("#precioErr").html("");
		  }
		}
	}	
	//
	let tieneFactura = $("#tieneFactura").val().trim();  
	//if (tieneFactura === "1") {
		
		let totalFactura = $("#totalFactura").val().trim();  
		if (totalFactura == null || totalFactura == "" || totalFactura.length == 0 || totalFactura == undefined ) {
			$("#totalfacturaErr").html("<span style='color:red;'>Monto Factura esta vacio..!</span>");
			swal("Aviso..!", "Monto Factura esta vacio..!", "error");
			Error++;
		} else {
			//
			totalFactura = totalFactura.replace(/[.]/g, "");
			totalFactura = totalFactura.replace(/[,]/g, ".");
			//
			let regex = /^[0-9.,]+$/;
			if (regex.test(totalFactura) === false) {
			  $("#totalfacturaErr").html("<span style='color:red;'>Introduzca un Monto Factura valido!</span>");
			  swal("Aviso..!", "Introduzca un Monto Factura valido!", "error");
			  Error++;
			} else {
			  if (parseFloat(totalFactura) <= 0) {
				$("#totalfacturaErr").html("<span style='color:red;'>Monto Factura debe ser mayor a cero!</span>");
				swal("Aviso..!", "Monto Factura debe ser mayor a cero!", "error");
				Error++;
			  } else {
				$("#totalfacturaErr").html("");
			  }
			}
		}	
	//}
	// Canal
	let cmbCanal = document.getElementById("cboCanal").selectedIndex;
	if (cmbCanal == null || cmbCanal == 0 || cmbCanal < 0) {
	  $("#canalErr").html("<span style='color:red;'>Seleccione Canal!</span>");
	  Error++;
	} else {
	  $("#canalErr").html("");
	}
	// Cadena
	let cmbCadena = document.getElementById("cboCadena").selectedIndex;
	if (cmbCadena == null || cmbCadena == undefined ) {
	  $("#cadenaErr").html("<span style='color:red;'>Seleccione Cadena!</span>");
	  Error++;
	} else {
	  $("#cadenaErr").html("");
	}
	//
	if (Error == 0) {
		$("#btn- save").attr("disabled", false);
		return true;
	} else {
		$("#btn-save").attr("disabled", false);
		return false;
	}
}
//
function eliminarProducto(){
	//
	debugger;
	//
	var idConsumo =	$("#cboDetallexDiaSemana").val();	
	//
	swal({
	  title: "¿ Seguro de Eliminarlo Todo ?",
	  text: " Esta accion no se puede reversar..! ",
	  type: "warning",
	  showCancelButton: true,
	  confirmButtonClass: "btn-danger",
	  confirmButtonText: "Si",
	  cancelButtonText: "No",
	  closeOnConfirm: true,
	  showLoaderOnConfirm: true,
	},
	function () {
		//				
		let ajax = {				
			idConsumo :	idConsumo,
		};				
		//
		$.ajax({		
			url: "g_ValEliminarTodoelConsumo.asp",
			type: 'GET',
			cache: false,
			async: false,
			data: ajax,
			beforeSend: function(objeto){
				$("#loader").html("<img src='images/ajax_small.gif'> Espere, Eliminando..!");
			},
			success: function (data) {
				//debugger;
				console.log(data);				
				swal("Aviso..!", "Consumo Eliminado...!", "success");
				$("#loader").html("");
				if(data==="0"){
					$("#hogarEliminado").css("display", "block");						
				}else{
					$("#hogarEliminado").css("display", "none");
				}
				buscarDetallexProductoFactura();
			},
		});
		//
	}
  );
}
//
function eliminar_Detalle_Producto(id){
	//
	// debugger;
	//
	swal({
	  title: "¿ Seguro de Eliminar Item ?",
	  text: " Esta accion no se puede reversar..! ",
	  type: "warning",
	  showCancelButton: true,
	  confirmButtonClass: "btn-danger",
	  confirmButtonText: "Si",
	  cancelButtonText: "No",
	  closeOnConfirm: true,
	  showLoaderOnConfirm: true,
	},
	function () {
		//				
		let ajax = {				
			idConsumo :	id,
		};				
		//
		$.ajax({		
			url: "g_ValEliminarDetalledelConsumo.asp",
			type: 'GET',
			cache: false,
			async: false,
			data: ajax,
			beforeSend: function(objeto){
				$("#loader").html("<img src='images/ajax_small.gif'> Espere, Eliminando Detalle..!");
			},
			success: function (data) {
				//debugger;
				console.log(data);				
				swal("Aviso..!", "Detalle del Consumo Eliminado...!", "success");
				$("#loader").html("");				
				buscarDetallexProductoFactura();
			},
		});
		//
	}
  );
}
//
function marcar_Producto_Pendiente(id){
	//
	// debugger;
	//
	swal({
	  title: "¿ Seguro de Marcalo Pendiente ?",
	  text: " ... ",
	  type: "warning",
	  showCancelButton: true,
	  confirmButtonClass: "btn-danger",
	  confirmButtonText: "Si",
	  cancelButtonText: "No",
	  closeOnConfirm: true,
	  showLoaderOnConfirm: true,
	},
	function () {
		//				
		let ajax = {				
			idConsumo :	id,
		};				
		//
		$.ajax({		
			url: "g_ValMarcarProductoPendiente.asp",
			type: 'GET',
			cache: false,
			async: false,
			data: ajax,
			beforeSend: function(objeto){
				$("#loader").html("<img src='images/ajax_small.gif'> Espere, Marcando Detalle..!");
			},
			success: function (data) {
				swal("Aviso..!", "Consumo Marcado Pendiente...!", "success");
				// debugger;
				// console.log(data);				
				// $("#loader").html("");	
				//
				debugger;
				console.log(data);				
				$("#loader").html("");
				if(data==="0"){
					$("#hogarValidado").css("display", "block");
					sessionStorage.setItem("Convalidado", 1 );
				}else{
					sessionStorage.setItem("Convalidado", 0 );
					$("#hogarValidado").css("display", "none");
				}
				//
				if($("#cboDetallexDiaSemana").val()==="0"){
					buscarDetallexProductoFacturaResuelto();
				}else{
					buscarDetallexProductoFactura();						
				}
				CalcularTotalesConsumos();
				//				
				//buscarDetallexProductoFactura();
				//CalcularTotalesConsumos();
			},
		});
		//
	}
  );
}
//
function agregarProducto() {
	//
	//debugger;
	//
	Reset_FormAddProductos();
	//
	var idConsumo =	$("#cboDetallexDiaSemana").val();		
	$("#loader").html("");		
	//	
	$("#txtIdConsumo").val(idConsumo);		
	//	
	$("#AgregarProducto").modal("show");
	$('#otrosProductos').prop('checked', false); // Unchecks it
	$("#MasterProductos").css("display", "block");
	$(".modal-title").html("<i class='fas fa-plus'></i> Agregar Producto");		
	//
	buscarMonedaPago();
	buscarOtrasCategorias();
	buscarMarcasCubitos();	
	//			
}
//
function ActualizarCalculoTotales() {
	// Edit Productos
	// debugger;
	$("#txtTotal").val("");
	//
	var Cantidad = parseFloat(preformatFloat($("#txtCantidad").val()));
	var Precio = parseFloat(preformatFloat($("#txtPrecio").val()));
	var TasaCambio = parseFloat(preformatFloat($("#txtTasa").val()));
	//
	valid = !isNaN(Cantidad) && !isNaN(Precio);
    if (!valid) {
    	return false;
    }	
	//			
	var totalCompra = 0;
	totalCompra =(TasaCambio * Precio) * Cantidad;
	//	
	$("#txtTotal").val(number_format_js(totalCompra,2,',','.'));
	//
}
//
function calcularTotales() {
	// add Productos
	debugger;
	$("#txtTotalCompra").val("");
	//
	var Cantidad = parseFloat(preformatFloat($("#txtCantidadProductos").val()));
	var Precio = parseFloat(preformatFloat($("#txtPrecioUnitario").val()));
	var TasaCambio = parseFloat(preformatFloat($("#txtTasaCambio").val()));	
	//
	valid = !isNaN(Cantidad) && !isNaN(Precio);
    if (!valid) {
    	return false;
    }	
	//			
	var totalCompra = 0;
	var value = $("#cboCategoriaOtros").val();
	if (parseInt(value) == 9) {
		//queso
		totalCompra =(( parseFloat(TasaCambio) * parseFloat(Precio) ));
	}else{
		totalCompra =((parseFloat(TasaCambio) * parseFloat(Precio)) * parseFloat(Cantidad) );
	}
	//	
	$("#txtTotalCompra").val(number_format_js(totalCompra,2,',','.'));
	//
}
// 
function salvarAgregarProductos() {
	// add Productos
	// debugger;
	//		
	if(validarNuevosProductos()){		
	  	//
	  swal(
		{
		  title: "Estan Correctos todos",
		  text: ".. los Datos ?",
		  type: "warning",
		  showCancelButton: true,
		  confirmButtonClass: "btn-danger",
		  confirmButtonText: "Si",
		  cancelButtonText: "No",
		  closeOnConfirm: true,
		  showLoaderOnConfirm: true,
		},
		function () {
			//
			// var sMoneda =$("#cboMonedaPago").val();
			// var fields = sMoneda.split('-');
			// var buscar = fields[0];
			//
			debugger;
			var id = $("#cboMonedaPago").val();	
			var idConsumoDetalle	=	id;	
			//
			var idConsumo	=	$("#cboDetallexDiaSemana").val();
			var idHogar	 	= 	$("#cboHogar").val(); 
			//
			if(idConsumo==="0" || idConsumo==null || idConsumo==undefined ){				
				idConsumo	= $("#cboConsumoInvestigado").val();								
			}			
			if(idHogar==="0" || idHogar==null || idHogar==undefined){
				var id_Hogar	=	$("#cboConsumoInvestigado").val();				
				var fields		=	id_Hogar.split('-');
				idHogar			= 	fields[0];
			}			
			//	
			// idConsumo		:	idConsumo, $("#cboDetallexDiaSemana").val(),						
			// ("#cboHogar").val(), 
			let ajax = {
				
				idConsumo		:	idConsumo, 			
				idHogar	 		:	idHogar,
				idSemana 		:	$("#cboSemana").val(),
				tieneFactura 	:	$("#tieneFactura").val().trim(),				
				pRecio 			:	parseFloat(preformatFloat($("#txtPrecioUnitario").val())),
				cAntidad		:	parseFloat(preformatFloat($("#txtCantidadProductos").val())),
				bArcode			:	$("#txtCodigoBarras").val(),
				moneda			:	$("#cboMonedaPago option:selected" ).text().trim(),
				idMoneda		: 	id,
				tasaCambio		:	parseFloat(preformatFloat($("#txtTasaCambio").val())),
				categoria		:	$("#cboCategoriaOtros").val(),
				marcaCubito		:	$("#cboMarcaCubitos").val(),
				unidades		: 	$("#unidades").val(),
				
			};				
			//
			$.ajax({		
				url: "g_ValAgregarProductosalConsumo.asp",
				type: 'POST',
				cache: false,
				async: false,
				data: ajax,
				//dataType: "json",
				beforeSend: function(objeto){
					$("#loader").html("<img src='images/ajax_small.gif'> Grabando Nuevo..!");
				},
				success: function (data) {
					debugger;
					console.log(data);				
					//$("#AgregarProducto").modal("hide");					
					Reset_FormAddProductos();
					$("#loader").html("");
					swal("Aviso..!", "Producto Agregado...!", "success");
					buscarDetallexProductoFactura();
				},
			});
			//
		}
	  );

	}else{
	  //swal("Aviso..!", "Hay Errores revise los mensajes...!", "error");
	}				
}
//
function validarNuevosProductos() {		
	//
	$("#btn-salvarProd").attr("disabled", true);
	//	
	 debugger;
	//
	var Error = 0;
	//	
	if($("#otrosProductos").is(':checked')) {  
		//	
		$("#txtCodigoBarras").val("0");
		$("#cboCategoria").prop("selectedIndex", 0);		
		$("#cboProducto").prop("selectedIndex", 0);
		//
	} else {
		//	
		// Nro del Codigo de Barras
		//	
		let barcode = $("#txtCodigoBarras").val().trim();
		if (barcode == null || barcode == "" || barcode.length == 0 || barcode == undefined ) {
			$("#txtcodigobarErr").html("<span style='color:red;'>Codigo de barras esta vacio o en blanco..!</span>");
			Error++;
		}else {
			let regex = /^[0-9]+$/;
			if (regex.test(barcode) === false) {
				$("#txtcodigobarErr").html("<span style='color:red;'>Introduzca una Codigo de barras valido (Solo numeros)!</span>");
				Error++;
			} else {
				if (parseFloat(barcode) <= 0) {
					$("#txtcodigobarErr").html("<span style='color:red;'>Codigo de barras invalido..!</span>");
					Error++;
				} else {
					$("#txtcodigobarErr").html("");
				}
				if (barcode.length < 7 || barcode.length > 16) {
					$("#txtcodigobarErr").html("<span style='color:red;'>Codigo de barras errado, Min 7 y Max 16 Caracteres..!</span>");
					Error++;
				} else {
					$("#txtcodigobarErr").html("");
				}
			}
		}
	}
	//
	// Cantidad
	//
	let cantidad = $("#txtCantidadProductos").val().trim();
	if (cantidad == null || cantidad == "" || cantidad.length == 0 || cantidad == undefined ) {
		$("#txtcantidadErr").html("<span style='color:red;'>Cantidad esta vacio o en blanco..!</span>");
		Error++;
	} else {
		cantidad = cantidad.replace(/[.]/g, "");
		cantidad = cantidad.replace(/[,]/g, ".");
		let regex = /^[0-9.,]+$/;
		if (regex.test(cantidad) === false) {
		  $("#txtcantidadErr").html("<span style='color:red;'>Introduzca una Cantidad valida!</span>");
		  Error++;
		} else {
		  if (parseFloat(cantidad) <= 0) {
			  $("#txtcantidadErr").html("<span style='color:red;'>Cantidad debe ser mayor a cero!</span>");
			Error++;
		  } else {
			$("#txtcantidadErr").html("");
		  }
		}
	}
	//
	// Precio
	//
	let precio = $("#txtPrecioUnitario").val().trim();  
	if (precio == null || precio == "" || precio.length == 0 || precio == undefined ) {
		$("#txtprecioErr").html("<span style='color:red;'>Precio esta vacio o en blanco..!</span>");
		Error++;
	} else {
		//		
		precio = precio.replace(/[.]/g, "");
		precio = precio.replace(/[,]/g, ".");
		let regex = /^[0-9.,]+$/;
		if (regex.test(precio) === false) {
		  $("#txtprecioErr").html("<span style='color:red;'>Introduzca una Precio valido!</span>");
		  Error++;
		} else {
		  if (parseFloat(precio) <= 0) {
			$("#txtprecioErr").html("<span style='color:red;'>Precio debe ser mayor a cero!</span>");
			Error++;
		  } else {
			$("#txtprecioErr").html("");
		  }
		}
	}		
	//
	// Tipo moneda
	let cmbMonedaPago = document.getElementById("cboMonedaPago").selectedIndex;
	if (cmbMonedaPago == null || cmbMonedaPago == 0 || cmbMonedaPago < 0) {
	  $("#canalErr").html("<span style='color:red;'>Seleccione Tipo Moneda de Pago!</span>");
	  Error++;
	} else {
	  $("#txtmonedapagoErr").html("");
	}	
	//	
	if (Error == 0) {
		$("#btn-salvarProd").attr("disabled", false);
		return true;
	} else {
		$("#btn-salvarProd").attr("disabled", false);
		return false;
	}
}
//
function showMostrarInvestigarHogar() {
	//
	//debugger;
	var idConsumo =	$("#cboDetallexDiaSemana").val();
	$("#txtComentarios").val("")
	//	
	if (idConsumo == null || idConsumo == "" || idConsumo.Length == 0 || idConsumo== undefined || idConsumo== "0" ) {
		swal("Aviso..!", "Debe Seleccionar un Consumo...!", "error");		
	}else{
		$("#txtIdConsumoInvestigar").val(idConsumo);		
		$("#investigarConsumo").modal("show");
		$(".modal-title").html("<i class='fas fa-eye'></i> Investigar Consumo");	
	}
}
//
function enviarConsumoInvestigar() {
	//
	debugger;
	var idConsumo 	=	$("#txtIdConsumoInvestigar").val();	
	var idItemsInv	=	$("#cboInvestigar").val();
	var idHogar		=	$("#cboHogar").val();
	var observa  	=	$("#txtComentarios").val()
	//
	if (idItemsInv == null || idItemsInv == "" || idItemsInv.Length == 0 || idItemsInv== undefined || idItemsInv== "0" ) {
		swal("Aviso..!", "Debe Seleccionar un Motivo...!", "error");
		return false;	
	}
	//
	swal({
		title: "¿ Seguro de Investigarlo ?",
	  text: " Esta accion no se puede reversar..! ",
	  type: "warning",
	  showCancelButton: true,
	  confirmButtonClass: "btn-danger",
	  confirmButtonText: "Si",
	  cancelButtonText: "No",
	  closeOnConfirm: true,
	  showLoaderOnConfirm: true,
	},
	function () {
		//				
		let ajax = {				
			id_Consumo  : idConsumo,
			id_ItemsInv : idItemsInv,
			id_Hogar	: idHogar,
			observacion	: observa,			
		};				
		//
		$.ajax({
			url: "g_ValEnviarInvestigarConsumo.asp",
			type: 'GET',
			cache: false,
			async: false,
			data: ajax,
			beforeSend: function(objeto){
				$("#loader").html("<img src='images/ajax_small.gif'> Enviado, Investigacion!");
			},
		})
		/*Si la consulta se realizo con exito*/
		.done(function(data) {
			debugger;
				console.log(data);				
				$("#loader").html("");								
				if(data==="True"){					
					$("#cboInvestigar").prop("selectedIndex", 0);
					$("#txtIdConsumoInvestigar").val("");
					$("#investigarConsumo").modal("hide");					
					swal("Aviso..!", "Consumo enviado a Investigar.!", "success");
					$("#hogarInvestigado").css("display", "block");					
					CalcularTotalesConsumos();
				}else{
					swal("Aviso..!", "Algo Salio Mal.., Intente de nuevo!","error");
				}							
		})
		/*Si la consulta Fallo*/
		.fail(function() {
			$("#loader").html("");	
			swal("Algo salio mal.!","Intentelo de nuevo..! eCI()", "error");
		},'html');
		//		
	}
  );
}
//
function deshacerMasivo(){
	//
	//debugger;
	//
	var idConsumo =	$("#cboDetallexDiaSemana").val();	
	//
	swal({
	  title: "¿ Seguro Deshacer Validacion ?",
	  text: ".. ",
	  type: "warning",
	  showCancelButton: true,
	  confirmButtonClass: "btn-danger",
	  confirmButtonText: "Si",
	  cancelButtonText: "No",
	  closeOnConfirm: true,
	  showLoaderOnConfirm: true,
	},
	function () {
		//				
		let ajax = {							
			idConsumo		:	idConsumo,
		};				
		//
		$.ajax({
			url: "g_ValUpdateDetallesxProductosMasivoDeshacer.asp",
			type: 'GET',
			cache: false,
			async: false,
			data: ajax,
			beforeSend: function(objeto){
				$("#loader").html("<img src='images/ajax_small.gif'> Espere, Validando..!");
			},
		})
		/*Si la consulta se realizo con exito*/
		.done(function(data) {
			swal("Aviso..!", "Producto Sin Validar...!", "success");
			debugger;
			console.log(data);				
			$("#loader").html("");
			if(data==="0"){
				$("#hogarValidado").css("display", "block");						
			}else{
				$("#hogarValidado").css("display", "none");
			}
			buscarDetallexProductoFactura();						
		})
		/*Si la consulta Fallo*/
		.fail(function() {
			$("#loader").html("");	
			swal("Algo salio mal.!","Intentelo de nuevo..! dM()", "error");
		},'html');
		//		
	}
  );		
}
//
function Reset_FormAddProductos(){
	$("#txtIdConsumo").val("");
	$("#txtBuscarDescripcion").val("");
	$("#txtCodigoBarras").val("");
	$("#txtTasaCambio").val("");
	$("#txtCantidadProductos").val("");
	$("#txtPrecioUnitario").val("");
	$("#txtTotalCompra").val("");	
	$("#cboCategoria").prop("selectedIndex", 0);
	$("#cboMonedaPago").prop("selectedIndex", 0);	
	$("#cboProducto").val("");
	//
	$("#MasterProductos").css("display", "block");
	$("#cboCategoria").removeAttr("disabled");
	$("#cboCategoria").prop("selectedIndex", 0);
	$("#txtBuscarDescripcion").removeAttr("disabled");
	$("#txtBuscarDescripcion").val("");
	$("#cboProducto").prop("selectedIndex", 0);
	$("#cboProducto").removeAttr("disabled");
	$("#txtCodigoBarras").val("");
	//		
	$("#cboCategoria").focus();
	$("#showOtrosProductos").css("display", "none");
	$("#showCubitos").css("display", "none");
	$("#showUnidades").css("display", "none");
	$("#cboCategoriaOtros").prop("selectedIndex", 0);
	$("#cboMarcaCubitos").prop("selectedIndex", 0);
	//
	$("#unidades").val("0");
	$('#otrosProductos').prop('checked', false); 
	//
	$("#nombreCantidad").html("Cantidad:");
	$("#nombrePrecio").html("Precio Unitario:"); 
	
	
}
//
$("#chkSinBarras").change(function() {
	if ($(this).is(':checked')) {
		$("#txtCodigoBar").prop("disabled", true);
     	$("#txtCodigoBar").val("00000000");
  	} else {
  		$("#txtCodigoBar").prop("disabled", false);
     	$("#txtCodigoBar").val("");
  	}
});
//
function salvarCambioProductosNoMercado() {
	//
	debugger
	//
	if(validarAjustesProductosNoMercado()){
	  //
	  swal(
		{
		  title: "Estan Correctos Todos",
		  text: ".. los ajustes realizados ?",
		  type: "warning",
		  showCancelButton: true,
		  confirmButtonClass: "btn-danger",
		  confirmButtonText: "Si",
		  cancelButtonText: "No",
		  closeOnConfirm: true,
		  showLoaderOnConfirm: true,
		},
		function () {
			//
			debugger;
			//							
			var TotalCompra = parseFloat(preformatFloat($("#txtTotalCompra2").val()));
			//
			let ajax = {
				idConsumo			:	$("#txtId2").val(),				
				idtipocomida		:	$("#cboTipoComida").val(),
				idMoneda			:	$("#cboMonedaPagoNoMercado").val(),
				total				:	TotalCompra,
				nombreLocal			:	$("#txtNombreLocal").val().trim(),
			};				
			//
			$.ajax({		
				url: "g_ValUpdateDetallesxProductosxUnicoNoMercado.asp",
				type: 'GET',
				cache: false,
				async: false,
				data: ajax,
				//dataType: "json",
				beforeSend: function(objeto){
					$("#loader").html("<img src='images/ajax_small.gif'> Espere, Grabando Ajustes..!");
				},
				success: function (data) {
					debugger;
					console.log(data);				
					if(data==="True"){
						$("#EditarConsumo2").modal("hide");					
						$("#loader").html("");
						sessionStorage.setItem("Convalidado", 1 );
						$("#hogarValidado").css("display", "block");						
						swal("Aviso..!", "Producto Actualizado...!", "success");
						//
						buscarDetallexProductoFacturaResuelto();
						buscarDetalleConsumoResueltoxDia();
						CalcularTotalesConsumos();
						//					
						
					}else{
						sessionStorage.setItem("Convalidado", 0 );
						$("#hogarValidado").css("display", "none");
					}
				},
			});
			//
		}
	  );

	}else{
	  //swal("Aviso..!", "Hay Errores revise los mensajes...!", "error");
	}				
}
//
function validarAjustesProductosNoMercado() {
	//	
	var Error = 0;
	//
	// Tipo Comida
	//
	let cmbComida = document.getElementById("cboTipoComida").selectedIndex;
	if (cmbComida == null || cmbComida == "") {
		$("#txttipocomidaErr").html("<span style='color:red;'>Seleccione Tipo Comida!</span>");
	  Error++;
	} else {
	  $("#txttipocomidaErr").html("");
	}	
	let NombreLocal = $("#txtNombreLocal").val().trim();
	if (NombreLocal == null || NombreLocal == "" || NombreLocal.length == 0 || NombreLocal== undefined ) {
		$("#txtnombrelocalErr").html("<span style='color:red;'>Cantidad esta vacio o en blanco..!</span>");
		Error++;
	} else {
		$("#txtnombrelocalErr").html("");		
	}
	//
	// Precio
	let precio = $("#txtTotalCompra2").val().trim();  
	if (precio == null || precio == "" || precio.length == 0 || precio == undefined ) {
		$("#txttotalcompra2").html("<span style='color:red;'>Total compra esta vacio..!</span>");
		Error++;
	} else {
		//		
		precio = parseFloat(preformatFloat(precio));		
		let regex = /^[0-9.,]+$/;
		if (regex.test(precio) === false) {
		  $("#txttotalcompra2").html("<span style='color:red;'>Introduzca un Total Compra valido!</span>");
		  Error++;
		} else {
		  if (parseFloat(precio) <= 0) {
			$("#txttotalcompra2").html("<span style='color:red;'>Total Compra debe ser mayor a cero!</span>");
			Error++;
		  } else {
			$("#txttotalcompra2").html("");
		  }
		}
	}		
	//
	if (Error == 0) {
		$("#btn- save").attr("disabled", false);
		return true;
	} else {
		$("#btn-save").attr("disabled", false);
		return false;
	}
}
//
function pendientesMasivo(){
	//
	debugger;	
	//
	// Validar Pendientes Masivo al menos uno activo este Marcado
	//
	var selected = [];
	var checked = 0;
	$('.data input:checked').each(function() {
		selected.push($(this).val());
		checked ++;
	});		
	if (checked == 0) {
		swal("Aviso..!", "Debe marcar al menos uno como Pendiente...!", "error"); 
		return false;
	}else{
		GetSelectedValues();
	}
	var idConsumo =	$("#cboDetallexDiaSemana").val();		
	//
	swal({
	 	title: "Esta Seguro de Enviar",
		text: "seleccionados a Pendientes ?",
	  	type: "warning",
	  	showCancelButton: true,
	  	confirmButtonClass: "btn-danger",
	  	confirmButtonText: "Si",
	  	cancelButtonText: "No",
	  	closeOnConfirm: true,
	  	showLoaderOnConfirm: true,
	},
	function () {
		//
		let ajax = {							
			checkboxes  : $("#Hiddenfield1").val(),
			idConsumo	: idConsumo,			
		};				
		//	
		$.ajax({
			url: "g_ValUpdatePendientesMasivo.asp",
			type: 'POST',
			cache: false,
			async: false,
			data: ajax,
			beforeSend: function(objeto){
				$("#loader").html("<img src='images/ajax_small.gif'> Espere, Validando..!");
			},
		})
		/*Si la consulta se realizo con exito*/
		.done(function(data) {
			$("#Hiddenfield1").val("");
			swal("Aviso..!", "Seleccion Pendientes Marcados...!", "success");
			debugger;
			console.log(data);				
			$("#loader").html("");
			if(data==="0"){
				$("#hogarValidado").css("display", "block");
				CalcularTotalesConsumos();
			}else{
				$("#hogarValidado").css("display", "none");
			}
			buscarDetallexProductoFactura();						
		})
		/*Si la consulta Fallo*/
		.fail(function() {
			$("#loader").html("");	
			$("#Hiddenfield1").val("");
			swal("Algo salio mal.!","Intentelo de nuevo..! pM()", "error");
		},'html');
		//		
	}
  );		
}
//
function GetSelectedValues() {
	//Get the checkbox values and assigned it as a comma separated string to hiddenfield
	$("#Hiddenfield1").val($("input[name=pendientes]:checked").map(function () {return this.value;}).get().join(","));
	//alert($("#Hiddenfield1").val());
}
//
function selects(){  
	var ele=document.getElementsByName('pendientes');  
	for(var i=0; i<ele.length; i++){  
		if(ele[i].type=='checkbox')  
			ele[i].checked=true;  
	}  
}
//
function deSelect(){  
	var ele=document.getElementsByName('pendientes');  
	for(var i=0; i<ele.length; i++){  
		if(ele[i].type=='checkbox')  
			ele[i].checked=false;  
		  
	}  
}    
//
function monedaMasivo() {
	//
	//debugger;
	//
	var idConsumo =	$("#cboDetallexDiaSemana").val();		
	$("#loader").html("");		
	//	
	$("#txtIdCambio").val(idConsumo);		
	//
	$("#CambioMoneda").modal("show");
	$(".modal-title").html("<i class='fas fa-edit'></i> Actualizar Moneda de Pago");	
	//
	buscarCambioTipoMonedaPago();
	//			
}
//
function salvarCambioMoneda() {
	//
	debugger
	//
	let cmbMoneda = document.getElementById("cboCambioMonedaPago").selectedIndex;
	if (cmbMoneda == null || cmbMoneda == "" || cmbMoneda == 0) {
		$("#txtcambiomonedapagoErr").html("<span style='color:red;'>Seleccione un Tipo de Moneda..!</span>");
		return false;  
	} else {
	  $("#txtcambiomonedapagoErr").html("");
	}		
	  //
	  swal(
		{
		  title: "Seguro de Cambiar",
		  text: "el Tipo de Moneda ?",
		  type: "warning",
		  showCancelButton: true,
		  confirmButtonClass: "btn-danger",
		  confirmButtonText: "Si",
		  cancelButtonText: "No",
		  closeOnConfirm: true,
		  showLoaderOnConfirm: true,
		},
		function () {
			//
			let ajax = {
				idConsumo	:	$("#txtIdCambio").val(),				
				idMoneda	:	$("#cboCambioMonedaPago").val(),
				Moneda		:	$("#cboCambioMonedaPago option:selected" ).text().trim(),
				idSemana 	:	$("#cboSemana").val(),
			};				
			//
			$.ajax({		
				url: "g_ValUpdateCambioMonedaMasivo.asp",
				type: 'GET',
				cache: false,
				async: false,
				data: ajax,
				beforeSend: function(objeto){
					$("#loader").html("<img src='images/ajax_small.gif'> Espere, Grabando Ajustes..!");
				},
				success: function (data) {
					debugger;
					console.log(data);				
					if(data==="True"){
						$("#CambioMoneda").modal("hide");					
						$("#loader").html("");
						swal("Aviso..!", "Moneda Actualizada...!", "success");
						//						
						buscarDetallexProductoFactura();
						//						
					}
				},
			});
			//
		}
	  );			
}