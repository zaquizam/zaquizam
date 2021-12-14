//
// MonedaV1.JS // 07abr21 - 11oct21
//
//
function Reset_Form(){
	//
	$("#txtIdConsumo").val("");
	$("#txtCodigoBarras").val("");
	$("#txtTasaCambio").val("");
	$("#txtCantidad").val("");
	$("#txtPrecio").val("");
	$("#txtTotal").val("");	
	$("#txtMoneda").val("");	
	$("#cboTMonedaPago").prop("selectedIndex", 0);	
	//	
}
//
function ActualizarMoneda(idConsumoDetalle) {
	//
	// debugger;
	//
	Reset_Form();	
	//	
	$("#txtIdConsumo").val(Math.abs(idConsumoDetalle));		
	//	
	$("#EditarMoneda").modal("show");
	$(".modal-title").html("<i class='fas fa-edit'></i> Actualizar Moneda");		
	//
	buscarMonedaPago();
	buscarDetalleProducto(Math.abs(idConsumoDetalle));
	$("#cboTMonedaPago").prop("selectedIndex", 0);		
	//
}
//
function buscarMonedaPago() {		
	//
	$.ajax({
		url: "g_ValBuscarLlenarCmbTMonedaPago.asp",
		type: "POST",
		cache: false,
		async: false,		
		dataType: "json",
		success: function (data) {
			let select = $("#cboTMonedaPago");
			select.find("option").remove();
			select.append("<option value='0' selected disabled> -- Seleccione -- </option>");
			$.each(data.data, function (key, value) {
				select.append("<option value=" + value.id + ">" + value.nombre + "</option>");
			});						
		},
  });
}
// 
function buscarTipoTasadeCambio() {
	//
	// Buscar la tasa de cambio de la semana del consumo Modificar Moneda
	//
	var idSemana	=	$("#id_Semana").val(); 	
	var idMoneda  	=	$("#cboTMonedaPago").val();		
	//		
	if(idMoneda == "2"){
		tasa = 1
		$("#txtTasa").val(Number(tasa).toLocaleString("es-ES", {minimumFractionDigits: 5}));
		ActualizarCalculoTotales();
		return false;	
	}
	//
	let ajax = {
		idsemana : idSemana,
		idmoneda : idMoneda,
	};	
	//***
	
	$.ajax({	
		url: "g_ValBuscarLlenarCmbTipoTasaCambioValData.asp",
		type: "GET",
		cache: false,
		async: false,
		data: ajax,					
	})
	/*Si la consulta se realizo con exito*/
	.done(function(data) {
		//debugger;
		console.log(data);
		tasa = Number(data).toLocaleString("es-ES", {minimumFractionDigits: 5});	
		$("#txtTasa").val(tasa);						
		ActualizarCalculoTotales();
		$("#txttmonedapagoErr").html("");
		//				
	})
	/*Si la consulta Fallo*/
	.fail(function() {
		alert("Alerta Fallo - BTTDC()");
	},'HTML');
	//	
}
//
function buscarDetalleProducto(idConsumoDetalle) {
	//	
	// buscar la respuesta de Investigaciones			
	//	
	let ajax = {
		idConsumo: idConsumoDetalle,		
	};
	//
	$.ajax({
		url: "g_ValBuscarDetalleConsumoValDatos.asp",
		type: 'GET',
		cache: false,
		async: false,		
		 data: ajax,				
	})
	/*Si la consulta se realizo con exito*/
	.done(function(data) {
		//debugger;
		console.log(data);
		var cantidad	=	data[0].cantidad;
		var precio  	=	data[0].precio;
		var barcode     =	data[0].barcode;
		var tasa        =	data[0].tasa;
		var total       =	data[0].total;
		var moneda		=	data[0].moneda;
		//												
		$("#txtCantidad").val(cantidad);
		$("#txtTasa").val(Number(tasa).toLocaleString("es-ES", {minimumFractionDigits: 5}));
		$("#txtPrecio").val(Number(precio).toLocaleString("es-ES", {minimumFractionDigits: 2}));
		$("#txtTotal").val(Number(total).toLocaleString("es-ES", {minimumFractionDigits: 2}));
		$("#txtCodigoBar").val(barcode);
		$("#txtMoneda").val(moneda);
		//
	})
	/*Si la consulta Fallo*/
	.fail(function() {
		alert("Fallo - bDetProd("+idConsumo+")");
	},'json');
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
function salvarCambioMoneda() {
	//
	//debugger	
	//
	//Categoria
	let comboValor = document.getElementById("cboTMonedaPago").selectedIndex;
	if (comboValor == null || comboValor == 0 || comboValor < 0) {
	  $("#txttmonedapagoErr").html("<span style='color:red;'>Seleccione Moneda de Pago..!</span>");
	return false;	
	} else {
	  $("#txttmonedapagoErr").html("");
	}	
	//
	swal(
		{
		  title: "Esta Correcto el",
		  text: ".. cambio realizado ?",
		  type: "warning",
		  showCancelButton: true,
		  confirmButtonClass: "btn-danger",
		  confirmButtonText: "Si",
		  cancelButtonText: "No",
		  closeOnConfirm: false,
		  showLoaderOnConfirm: false,
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
				idConsumoDetalle	:	$("#txtIdConsumo").val(),				
				cantidad			:	$("#txtCantidad").val(),
				barcode				:	$("#txtCodigoBar").val(),				
				idMoneda			:	$("#cboTMonedaPago").val(),
				moneda				:	$("#cboTMonedaPago option:selected" ).text().trim(),
				precio 				:	Precio,
				tasa				:	TasaCambio,
				total				:	TotalCompra,
			};				
			//
			$.ajax({
				url: "g_ValUpdateDetallesxProductosxUnicoValDatos.asp",
				type: 'GET',
				cache: false,
				async: false,		
				 data: ajax,				
			})
			/*Si la consulta se realizo con exito*/
			.done(function(data) {
				debugger;
				console.log(data);				
				$("#EditarMoneda").modal("hide");					
				swal("Aviso..!", "Moneda Actualizada...!", "success");
				window.location.reload();
				//
			})
			/*Si la consulta Fallo*/
			.fail(function() {
				alert("Fallo - sCambMon()");
			},'html');		
		}
	);
		
}
//
//
function ActualizarCalculoTotales() {
	//
	// Edit Moneda
	// debugger;
	//
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
	$("#txtTotal").val(Number(totalCompra).toLocaleString("es-ES", {minimumFractionDigits: 2}));
	//	
}
//