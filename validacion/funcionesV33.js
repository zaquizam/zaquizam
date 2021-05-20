//
// funcionesV27.js // 29dic20 - 21ene21 
//
function Reset() {
	//		
	//$('#cboSemana').find('option:not(:first)').remove();	
	$('#cboArea').find('option:not(:first)').remove();
	$('#cboEstado').find('option:not(:first)').remove();
	$('#cboHogar').find('option:not(:first)').remove();
	$('#cboTipoConsumo').find('option:not(:first)').remove();
	$('#cboTotalxDiaSemana').find('option:not(:first)').remove();
	$('#cboDetallexDiaSemana').find('option:not(:first)').remove();	
	//
	$("#totalConsumos").html("0");
	$("#totalValidados").html("0");		
	$("#totalPendientes").html("0");		
	$("#altaHogar").html("");
	$("#totalHogares").html("0");
	$("#responsableHogar").html("");
	$("#celularHogar").html("");
	$("#totalInvestigados").html("0");	
	$("#totalResueltos").html("0");
	sessionStorage.setItem("Convalidado", false );	
	Reset_Detalles();
	$("#cargando").css("display", "none");
	$('#cboSemana').val(0);
	$("#cboSemana").focus();			
}
//
function Reset_Detalles() {
	//				
	$("#tabla-resultados").html("");	
	$("#montoFactura").val("");	
	$("#imgfactura").html("<img src='images/loader/cargador4.gif'");	
	$("#montoFactura").css("display", "none");
	$("#DetalleFactura").css("display", "none");
	$("#hogarValidado").css("display", "none");	
	$("#hogarEliminado").css("display", "none");
	$("#hogarInvestigado").css("display", "none");
	$("#hogarResuelto").css("display", "none");	
	$("#tienefactura").html("");
	$("#MonedaPagoFactura").val(0);	
	//
}
//
function Reset_Resuelto() {
	//		
	$('#cboArea').find('option:not(:first)').remove();
	$('#cboEstado').find('option:not(:first)').remove();
	$('#cboHogar').find('option:not(:first)').remove();
	$('#cboTipoConsumo').find('option:not(:first)').remove();
	$('#cboTotalxDiaSemana').find('option:not(:first)').remove();
	$('#cboDetallexDiaSemana').find('option:not(:first)').remove();		
	$("#altaHogar").html("");	
	$("#responsableHogar").html("");
	$("#celularHogar").html("");	
	Reset_Detalles();
	//
}
//
function Reset_Detalles() {
	//				
	$("#tabla-resultados").html("");	
	$("#montoFactura").val("");	
	$("#imgfactura").html("<img src='images/loader/cargador4.gif'");	
	$("#montoFactura").css("display", "none");
	$("#DetalleFactura").css("display", "none");
	$("#hogarValidado").css("display", "none");	
	$("#hogarEliminado").css("display", "none");
	$("#hogarInvestigado").css("display", "none");
	$("#hogarResuelto").css("display", "none");	
	$("#tienefactura").html("");
	$("#MonedaPagoFactura").val(0);	
	//
}
//
function buscarSemanas() {
	//
	$('#cboArea').find('option:not(:first)').remove();
	$('#cboEstado').find('option:not(:first)').remove();
	$('#cboHogar').find('option:not(:first)').remove();
	$('#cboTipoConsumo').find('option:not(:first)').remove();
	$('#cboTotalxDiaSemana').find('option:not(:first)').remove();
	$('#cboDetallexDiaSemana').find('option:not(:first)').remove();
	Reset();
	Reset_Detalles();	
	//
	$.ajax({		
		url: "g_ValBuscarSemanas.asp",
		cache: false,
		async: false,
		dataType: "json",
		beforeSend: function(objeto){
			//$("#cargando").css("display", "block");
			$("#cargando").css("display", "block");
		},
		success: function (data) {			
			let $select = $("#cboSemana");
			$select.find("option").remove();
			$select.append("<option value='0' selected disabled>-- Seleccione --</option>");
			$.each(data, function (i, value) {				
				$select.append("<option value=" + value.Id + ">" + value.Name + "</option>");				
			});				
			//$("#cargando").css("display", "none");
			$("#cargando").css("display", "none");
			$("#cboSemana").focus();
		},
	});		
}
//
function buscarArea() {		
	//
	$('#cboArea').find('option:not(:first)').remove();
	$('#cboEstado').find('option:not(:first)').remove();
	$('#cboHogar').find('option:not(:first)').remove();
	$('#cboTipoConsumo').find('option:not(:first)').remove();
	$('#cboTotalxDiaSemana').find('option:not(:first)').remove();
	$('#cboDetallexDiaSemana').find('option:not(:first)').remove();	
	Reset_Detalles();
	//
	var idSemana=$("#cboSemana").val();	
	if (idSemana == null || idSemana == 0) {
		swal("Aviso..!", "Seleccione una Semana", "error");
		$("#cboSemana").focus();
		return false;
	}
	//		
	$.ajax({			
		url: "g_ValBuscarAreaxSemana.asp?id_Semana=" + idSemana,
		cache: false,
		async: false,
		dataType: "json",
		beforeSend: function(objeto){
			//$("#cargando").css("display", "block");
			$("#cargando").css("display", "block");
		},
		success: function (data) {			
			let $select = $("#cboArea");
			$select.find("option").remove();
			$select.append("<option value='0' selected disabled>-- Seleccione --</option>");
			$.each(data, function (i, value) {				
				$select.append("<option value=" + value.Id + ">" + value.Name + "</option>");				
			});				
			$("#cargando").css("display", "none");
			$("#cboArea").focus();
			CalcularTotalesConsumos();
		},
	});		
}
//
cboTipoComida
function buscarEstado() {	
	//
	$('#cboEstado').find('option:not(:first)').remove();
	$('#cboHogar').find('option:not(:first)').remove();
	$('#cboTipoConsumo').find('option:not(:first)').remove();
	$('#cboTotalxDiaSemana').find('option:not(:first)').remove();
	$('#cboDetallexDiaSemana').find('option:not(:first)').remove();	
	Reset_Detalles();
	//
	var idArea=$("#cboArea").val(); 
	var idSemana=$("#cboSemana").val();	
	if (idArea == null || idArea == 0) {
		swal("Aviso..!", "Seleccione un Area", "error");
		$("#cboArea").focus();
		return false;
	}
	//		
	$.ajax({			
		url: "g_ValBuscarEstadoxArea.asp?id_Area=" + idArea + "&id_Semana=" + idSemana,
		cache: false,
		async: false,
		dataType: "json",
		beforeSend: function(objeto){
			$("#cargando").css("display", "block");
		},
		success: function (data) {			
			console.log(data);
			let $select = $("#cboEstado");
			$select.find("option").remove();
			$select.append("<option value='0' selected disabled>-- Seleccione --</option>");
			$.each(data, function (i, value) {				
				$select.append("<option value=" + value.Id + ">" + value.Name + "</option>");				
			});				
			$("#cargando").css("display", "none");
			$("#cboEstado").focus();
		},
	});		
}
//
function buscarHogar() {
	//
	$('#cboHogar').find('option:not(:first)').remove();
	$('#cboTipoConsumo').find('option:not(:first)').remove();
	$('#cboTotalxDiaSemana').find('option:not(:first)').remove();
	$('#cboDetallexDiaSemana').find('option:not(:first)').remove();
	Reset_Detalles();
	//
	var idEstado	=	$("#cboEstado").val();
	var idArea	=	$("#cboArea").val(); 
	var idSemana	=	$("#cboSemana").val();		
	//
	if (idEstado == null || idEstado == 0) {
		swal("Aviso..!", "Seleccione un Estado", "error");
		$("#cboEstado").focus();
		return false;
	}
	//		
	$.ajax({		
		url: "g_ValBuscarHogarxArea.asp?id_Estado=" + idEstado + "&id_Semana=" + idSemana + "&id_Area=" + idArea,
		cache: false,
		async: false,
		dataType: "json",
		beforeSend: function(objeto){
			$("#cargando").css("display", "block");
		},
		success: function (data) {
			//debugger;
			console.log(data);				
			let $select = $("#cboHogar");
			$select.find("option").remove();
			$select.append("<option value='0' selected disabled>-- Seleccione --</option>");
			$.each(data, function (i, value) {				
				$select.append("<option value=" + value.Id + ">" + value.Name + "</option>");				
			});				
			$("#cargando").css("display", "none");
			$("#cboHogar").focus();
		},
	});		
}
//
function buscarTipoConsumo() {
	//
	//debugger;
	$('#cboTotalxDiaSemana').find('option:not(:first)').remove();
	$('#cboDetallexDiaSemana').find('option:not(:first)').remove();	
	$("#cboDiaSemana").val("");
	Reset_Detalles();
	//
	var idHogar	= $("#cboHogar").val(); 
	var idSemana 	= $("#cboSemana").val();	

	if (idHogar == null || idHogar == 0) {
		swal("Aviso..!", "Seleccione un Hogar", "error");
		$("#cboHogar").focus();
		return false;
	}
	//		
	$.ajax({		
		url: "g_ValBuscarHogarxTipoConsumo.asp?id_Hogar=" + idHogar + "&id_Semana=" + idSemana,
		cache: false,
		async: false,
		dataType: "json",
		beforeSend: function(objeto){
			$("#cargando").css("display", "block");
		},
		success: function (data) {
			console.log(data);				
			let $select = $("#cboTipoConsumo");
			$select.find("option").remove();
			$select.append("<option value='0' selected disabled>-- Seleccione --</option>");
			$.each(data, function (i, value) {				
				$select.append("<option value=" + value.Id + ">" + value.Name + "</option>");				
			});				
			$("#cargando").css("display", "none");
			$("#cboTipoConsumo").focus();
			buscarAltaHogar();
		},
	});		
}
//
function buscarTotalDiaSemana() {
	//
	$('#cboTotalxDiaSemana').find('option:not(:first)').remove();
	$('#cboDetallexDiaSemana').find('option:not(:first)').remove();
	Reset_Detalles();
	//
	var idTipConsumo	=	$("#cboTipoConsumo").val();
	var idHogar			=	$("#cboHogar").val();
	var idSemana		=	$("#cboSemana").val(); 
	//	
	if (idTipConsumo == null || idTipConsumo == 0) {
		swal("Aviso..!", "Seleccione un Tipo Consumo", "error");
		$("#cboTipoConsumo").focus();
		return false;
	}
	//		
	$.ajax({		
		url: "g_ValBuscarTotalConsumoxDiaSemana.asp?id_TipoConsumo=" + idTipConsumo + "&id_Hogar=" + idHogar + "&id_Semana=" + idSemana,
		cache: false,
		async: false,
		dataType: "json",
		beforeSend: function(objeto){
			$("#cargando").css("display", "block");
		},
		success: function (data) {
			console.log(data);				
			let select = $("#cboTotalxDiaSemana");
			select.find("option").remove();
			select.append("<option value='0' selected disabled>-- Seleccione --</option>");
			$.each(data, function (i, value) {				
				select.append("<option value=" + value.Id + ">" + value.Name + "</option>");				
			});				
			$("#cargando").css("display", "none");
			$("#cboTotalxDiaSemana").focus();
		},
	});		
}
//
function buscarDetalleDiaSemana() {
	//
	//debugger;
	$('#cboDetallexDiaSemana').find('option:not(:first)').remove();
	Reset_Detalles();
	//
	var idFecha		=   $("#cboTotalxDiaSemana").val();
	var idTipConsumo	=	$("#cboTipoConsumo").val();
	var idHogar		=   $("#cboHogar").val();
	var idSemana		=	$("#cboSemana").val(); 
	//	
	if (idTipConsumo == null || idTipConsumo == 0) {
		swal("Aviso..!", "Seleccione un Tipo Consumo", "error");
		$("#cboTipoConsumo").focus();
		return false;
	}
	//		
	$.ajax({		
		url: "g_ValBuscarDetalleConsumoxDiaSemana.asp?id_Fecha=" + idFecha + "&id_Hogar=" + idHogar + "&id_Semana=" + idSemana + "&id_TipoConsumo=" + idTipConsumo,
		cache: false,
		async: false,
		dataType: "json",
		beforeSend: function(objeto){
			$("#cargando").css("display", "block");
		},
		success: function (data) {
			console.log(data);				
			let select = $("#cboDetallexDiaSemana");
			select.find("option").remove();
			select.append("<option value='0' selected disabled>-- Seleccione --</option>");
			$.each(data, function (i, value) {				
				select.append("<option value=" + value.Id + ">" + value.Name + "</option>");				
			});				
			$("#cargando").css("display", "none");
			$("#cboDetallexDiaSemana").focus();
		},
	});		
}
//
function buscarDetalleConsumoxDia() {
	//
	// Buscar el detalle los del consumo individual
	debugger;
	//
	$("#cboDiaSemana").val("");
	Reset_Detalles();
	buscarCadena(0);
	buscarCanal();
	buscarHogarValidado();
	buscarHogarInvestigado();
	buscarMotivoInvestigacion();
	//	
	var idConsumo		=	$("#cboDetallexDiaSemana").val();
	//	
	if ($("#cboTotalxDiaSemana").val() == null || $("#cboTotalxDiaSemana").val() == 0) {
		swal("Aviso..!", "Seleccione un Tipo Consumo", "error");
		$("#cboTipoConsumo").focus();
		return false;
	}
	//		
	$.ajax({		
		url: "g_ValBuscarDetalleConsumoxDia.asp?id_Consumo=" + idConsumo,
		cache: false,
		async: false,
		dataType: "json",
		beforeSend: function(objeto){
			$("#cargando").css("display", "block");
		},
		success: function (data) {							
			//debugger;				
			console.log(data);				
			//
			var canal  = parseInt(data[0].canal);
			var cadena = parseInt(data[0].cadena);
			var moneda = parseInt(data[0].moneda);
			var totalProductos = parseInt(data[0].totalproductos);
			//
			$("#cboCanal").val(canal).change();
			$("#cboCadena").val(cadena).change();			
			$("#MonedaPagoFactura").val(moneda).change();
			$("#totalProductos").val(totalProductos);
			var factura=data[0].tienefactura;
			//
			if(factura==="True"){
				$("#cargando").css("display", "none");				
				$("#tienefactura").html("Tiene Factura: ( SI )");
				$("#totalFactura").val("");							
				$("#tieneFactura").val("1");				
			}else{
				$("#tienefactura").html("Tiene Factura: ( NO )");					
				$("#totalFactura").val("");
				$("#tieneFactura").val("0");
			}
			if(parseFloat(data[0].totalcompra)===0){
				$("#montoFactura").html('Total Compra: Ver factura');	
			}else{
				var totalcompraGeneral = parseFloat(data[0].totalcompra);					
				totalcompra = Number(data[0].totalcompra).toLocaleString("es-ES", {minimumFractionDigits: 2});					
				$("#montoFactura").html("<span>Total Compra: " + totalcompra + "</span>");	
				$("#totalFactura").val(totalcompra);
				$("#totalFactura").attr("disabled", false);
			}								
			$("#cargando").css("display", "none");
			//
			buscarImagenFactura();
			//
		},    				
	});		
}
//
function buscarImagenFactura() {
	//	
	buscarDetallexProductoFactura();
	//
	//debugger;
	var id	=	$("#cboDetallexDiaSemana").val();
	//	
	if (id == null || id == 0) {
		swal("Aviso..!", "Faltan Datos, para Procesar...!", "error");
		$("#cboSemana").focus();
		return false;
	}
	//	
	$.ajax({
		url: "g_ValBuscarImagenFacturaxDia.asp",
		type: 'GET',
		cache: false,
		async: false,
		dataType: 'JSON',
		/*En el data se define los datos que se mandaran y como, en este ejemplo se envian los datos como tipo JSON*/
		data: {id_Consumo: id},
		/*El beforSend se ejecuta hasta que se reciba una respuesta del servidor, mientras tanto mostrara el mensaje "Cargando..."*/
		beforeSend: function(){
			//$("#loader").html("<img src='images/ajax_small.gif'> Buscando Imagen!");
			$("#cargando").css("display", "block");
		}
	})
	/*Si la consulta se realizo con exito*/
	.done(function(data) {
		//debugger;
		$("#DetalleFactura").css("display", "block");				
		console.log(data);
		if (data[0].id==="0"){
			var imagen = "images/"+ data[0].imagen;	
		}else{
			var imagen = "images/facturas/"+ data[0].imagen;	
		}
		$("#imgfactura").attr("src", imagen);		
		$("#cargando").css("display", "none");			
		//
		buscarResumenSemanal();
		//
	})
	/*Si la consulta Fallo*/
	.fail(function() {
		alert("Fallo - bIF()");
	},'json');
//	
}
//
function buscarDetallexProductoFactura() {
	//
	debugger;
	var sUrl;
	var idConsumo 	 =	$("#cboDetallexDiaSemana").val();		
	var idTipConsumo =	$("#cboTipoConsumo").val();
	if (parseInt(idTipConsumo)===0 && parseFloat(idConsumo)===0  ) {
		var idConsumo =	$("#cboConsumoInvestigado").val();
		sUrl = "g_ValBuscarDetallesxProductosxFacturaResuelto.asp";
	}else{
		sUrl = "g_ValBuscarDetallesxProductosxFactura.asp";
	}
	//
	let ajax = {
		id_Consumo   : idConsumo,
		id_TipConsumo: idTipConsumo,
	};
	//
	$.ajax({
		//url: "g_ValBuscarDetallesxProductosxFactura.asp",
		url: sUrl,
		type: 'GET',
		cache: false,
		async: false,
		dataType: 'HTML',
		/*En el data se define los datos que se mandaran y como, en este ejemplo se envian los datos como tipo JSON*/
		data: ajax,
		/*El beforSend se ejecuta hasta que se reciba una respuesta del servidor, mientras tanto mostrara el mensaje "Cargando..."*/
		beforeSend: function(){
			//$("#loader").html("<img src='images/ajax_small.gif'> Buscando Imagen!");
			$("#cargando").css("display", "block");
		}
	})
	/*Si la consulta se realizo con exito*/
	.done(function(data) {
		//debugger;
		console.log(data);
		$("#cargando").css("display", "none");
		$("#tabla-resultados").html(data);
		//				
	})
	/*Si la consulta Fallo*/
	.fail(function() {
		alert("Fallo - bDxPF()");
	},'HTML');
		
}
//
function buscarCadena(id) {
	$("#cargando").css("display", "block");
	$("#cboCadena").prop("disabled", true);
	let ajax = {
		opcion: 1,
		id: id,
	};
	$.ajax({
		url: "g_ValBuscarCadenaxConsumo.asp",
		type: "POST",
		cache: false,
		async: false,
		data: ajax,
		dataType: "json",
		success: function (data) {
			let select = $("#cboCadena");
			select.find("option").remove();
			select.append("<option value='' selected disabled> -- Seleccione -- </option>");
			$.each(data.data, function (key, value) {
				select.append("<option value=" + value.id + ">" + value.nombre + "</option>");
			});
			$("#cargando").css("display", "none");
			var len = $("#cboCadena option").length;
			if (len <= 2) {
				$("#cboCadena").prop("selectedIndex", 1);
				$("#cboCadena").prop("disabled", true);
				texto=$("#cboCadena option:selected" ).text().trim();			
			} else {
				$("#cboCadena").prop("selectedIndex", 0);
				$("#cboCadena").prop("disabled", false);
			}	  
		},
  });
}
//
function buscarCanal() {
	$("#cargando").css("display", "block");	
	let ajax = {
			id: 0,
		opcion: 2,
	};
	$.ajax({
		url: "g_ValBuscarCadenaxConsumo.asp",
		type: "POST",
		cache: false,
		async: false,
		data: ajax,
		dataType: "json",
		success: function (data) {
			//debugger;
			let select = $("#cboCanal");
			select.find("option").remove();
			select.append("<option value='0' selected disabled> -- Seleccione -- </option>");
			$.each(data.data, function (key, value) {
				select.append("<option value=" + value.id + ">" + value.nombre + "</option>");
			});
			$("#cargando").css("display", "none");			
		},
  });
}
//
function buscarHogarValidado() {
	//		
	//$("#loader").html("<img src='images/ajax_small.gif'> Buscando Hogar Validado...!");	
	$("#cargando").css("display", "block");
	var idConsumo =	$("#cboDetallexDiaSemana").val();	
	let ajax = {
		idConsumo: idConsumo,		
	};
	$.ajax({
		url: "g_ValBuscarHogarValidado.asp",
		type: "POST",
		cache: false,
		async: false,
		data: ajax,
		//dataType: "json",
		success: function (data) {			
			//debugger;
			console.log(data);				
			$("#cargando").css("display", "none");
			if(data==="True"){
				$("#hogarValidado").css("display", "block");						
			}else{
				$("#hogarValidado").css("display", "none");
			}			
		},
  });
}
//
function buscarCategoria() {
	//debugger;
	$("#cargando").css("display", "block");
	//
	$.ajax({
		url: "g_ValBuscarLlenarCmbCategoria.asp",
		type: "POST",
		cache: false,
		async: false,
		//data: ajax,
		dataType: "json",
		success: function (data) {
			let select = $("#cboCategoria");
			select.find("option").remove();
			select.append("<option value='0' selected disabled> -- Seleccione -- </option>");
			$.each(data.data, function (key, value) {
				select.append("<option value=" + value.id + ">" + value.nombre + "</option>");
			});
			$("#cargando").css("display", "none");			
		},
  });
}
//
function buscarProducto() {
	// Buscar Codigo de Barras del Producto
	// debugger;
	var txtBuscar  = $("#txtBuscarDescripcion").val().trim();
	if (txtBuscar == null || txtBuscar == "" || txtBuscar.length == 0 || txtBuscar == undefined ) {
		swal("Aviso..!", "Describa el producto a buscar..!", "error");
		return false;
	}
	$("#waiting").css("display", "block");	
	var idCategoria = $("#cboCategoria").val();
	//			
	let ajax = {
		id: idCategoria,
		find: txtBuscar,
	};
	$.ajax({
		url: "g_ValBuscarLlenarCmbProducto.asp",
		type: "GET",
		cache: false,
		async: false,
		data: ajax,
		//dataType: "json",
		success: function (data) {
			//debugger;
			console.log(data);
			let select = $("#cboProducto");
			select.find("option").remove();
			select.append("<option value='0' selected disabled> -- Seleccione -- </option>");
			$.each(data.data, function (key, value) {
				select.append("<option value=" + value.id + ">" + value.nombre + "</option>");
			});
			$("#cargando").css("display", "none");			
			$("#waiting").css("display", "none");
		},		
  });
}
//
function mostrarBarcode() {
	// Mostrar el barcode del producto	
	var barcode = $("#cboProducto").val();
	$("#txtCodigoBarras").val(barcode);
	//	
}
//
function buscarAltaHogar() {
	// Buscar la fecha de Creacion/Ingreso del hogar
	// debugger;		
	var idHogar	 = $("#cboHogar").val(); 
	$.ajax({		
		url: "g_ValBuscarAltaHogar.asp?id_Hogar=" + idHogar,
		cache: false,
		async: false,
		beforeSend: function(objeto){
			//$("#loader").html("<img src='images/ajax_small.gif'> Alta Hogar..!");
			$("#cargando").css("display", "block");
		},
		success: function (data) {
			//debugger;
			console.log(data);							
			//				
			var nombre   = data[0].nombre;
			var apellido = data[0].apellido;
			var celular  = data[0].celular;
			var fecha    = data[0].fecha;
			//
			if (fecha == null || fecha == "" || fecha.length == 0 || fecha == undefined ) {
				$("#responsableHogar").html(nombre+" "+apellido);
				$("#celularHogar").html(celular);
				$("#altaHogar").html("Sin Registro");			
			}else{
				$("#responsableHogar").html(nombre+" "+apellido);
				$("#celularHogar").html(celular);
				$("#altaHogar").html(fecha);			
			}
			$("#cargando").css("display", "none");			
		},
	});		
}
//
function buscarTasadeCambio() {
	// Add Productos
	// Buscar la tasa de cambio de la semana del consumo
	debugger;
	var idSemana	=	$("#cboSemana").val(); 	
	var idMoneda  	=	$("#cboMonedaPago").val();			
	$("#loader").html("<img src='images/ajax_small.gif'> Buscando Tasa de Cambio...!");	
	//
	if(idMoneda === '2'){
		tasa = 1
		$("#txtTasaCambio").val(Number(tasa).toLocaleString("es-ES", {minimumFractionDigits: 2}));
		return false;				
	}
	//
	$("#waiting").css("display", "block");
	let ajax = {
		id_semana: idSemana,
		id_moneda: idMoneda,
	};	
	$.ajax({
		url: "g_ValBuscarTasaCambio.asp",
		type: "POST",
		cache: false,
		async: false,
		data: ajax,
		//dataType: "json",
		success: function (data) {
			//debugger;
			console.log(data);
			tasa = Number(data).toLocaleString("es-ES", {minimumFractionDigits: 2});	
			$("#txtTasaCambio").val(tasa);						
			$("#txtCantidadProductos").focus();			
			$("#loader").html('');
			$("#waiting").css("display", "none");
		},
  });
}
//
function buscarMonedaPago() {
	//debugger;
	//$("#loader").html("<img src='images/ajax_small.gif'> Buscando Moneda de Pago!");	
	$("#cargando").css("display", "block");
	//
	$.ajax({
		url: "g_ValBuscarLlenarCmbTMonedaPago.asp",
		type: "POST",
		cache: false,
		async: false,
		//data: ajax,
		dataType: "json",
		success: function (data) {
			let select = $("#cboMonedaPago");
			select.find("option").remove();
			select.append("<option value='0' selected disabled> -- Seleccione -- </option>");
			$.each(data.data, function (key, value) {
				select.append("<option value=" + value.id + ">" + value.nombre + "</option>");
			});			
			$("#cargando").css("display", "none");			
		},
  });
}
//
function buscarTipoTasadeCambio() {
	//
	// Buscar la tasa de cambio de la semana del consumo Modificar
	// debugger;
	var idSemana	=	$("#cboSemana").val(); 	
	var idMoneda  	=	$("#cboTMonedaPago").val();		
	//		
	if(idMoneda == "2"){
		tasa = 1
		$("#txtTasa").val(Number(tasa).toLocaleString("es-ES", {minimumFractionDigits: 2}));
		ActualizarCalculoTotales();
		return false;	
	}
	//
	$("#waiting2").css("display", "block");
	let ajax = {
		id_semana: idSemana,
		idmoneda: idMoneda,
	};	
	$.ajax({
		url: "g_ValBuscarLlenarCmbTipoTasaCambio.asp",
		type: "POST",
		cache: false,
		async: false,
		data: ajax,
		//dataType: "json",
		success: function (data) {
			//debugger;
			console.log(data);
			tasa = Number(data).toLocaleString("es-ES", {minimumFractionDigits: 2});	
			$("#txtTasa").val(tasa);						
			$("#txtCantidad").focus();			
			$("#loader").html('');
			$("#waiting2").css("display", "none");
			ActualizarCalculoTotales();
		},
  });
}
//
function buscarTipoMonedaPago() {
	//debugger;
	//$("#loader").html("<img src='images/ajax_small.gif'> Buscando Moneda de Pago!");	
	$("#cargando").css("display", "block");
	//
	$.ajax({
		url: "g_ValBuscarLlenarCmbTipoMonedaPago.asp",
		type: "POST",
		cache: false,
		async: false,
		//data: ajax,
		dataType: "json",
		success: function (data) {
			let select = $("#cboTMonedaPago");
			select.find("option").remove();
			select.append("<option value='0' selected disabled> -- Seleccione -- </option>");
			$.each(data.data, function (key, value) {
				select.append("<option value=" + value.id + ">" + value.nombre + "</option>");
			});					
			$("#cargando").css("display", "none");			
		},
  });
}
//
function CalcularTotalesConsumos() {
	// Totalizar consumos x Validados y Pendientes
	// debugger;
	//
	var idSemana	=	$("#cboSemana").val(); 	
	//		
	$.ajax({		
		url: "g_ValTotalizarConsumosxValidadosyPendientes.asp?id_semana=" + idSemana,
		cache: false,
		async: false,
		dataType: "json",
		beforeSend: function(objeto){
			//$("#loader").html("<img src='images/ajax_small.gif'> Totalizando!");
			$("#cargando").css("display", "block");
		},
		success: function (data) {							
			//debugger;				
			console.log(data);				
			//
			var count = data.length;
  			console.log(count);
			//				
			if(count==2){
				var tvalidados=data[0].total;
				var tpendientes=data[1].total;
				var total = parseInt(tvalidados) + parseInt(tpendientes);
			} else if (count==1) {
				var validado = data[0].validado;
				if (validado=="True"){
					var tvalidados = data[0].total;
					var tpendientes = 0;	
				}else{
					var tvalidados = 0;
					var tpendientes = data[0].total;
				}					
				var total = parseInt(tvalidados) + parseInt(tpendientes);			
			}
			//
			valor = Number(total).toLocaleString("es-ES", {minimumFractionDigits: 0});	
			$("#totalConsumos").html(valor);		
			//
			valor = Number(tvalidados).toLocaleString("es-ES", {minimumFractionDigits: 0});	
			$("#totalValidados").html(valor);		
			//
			valor = Number(tpendientes).toLocaleString("es-ES", {minimumFractionDigits: 0});	
			$("#totalPendientes").html(valor);		
			//
			$("#altaHogar").html("<i class='fas fa-calendar-day'></i>");				
			buscarTotalHogaresReportados();
			$("#cargando").css("display", "none");
		},    				
	});		
}
//
function buscarTotalHogaresReportados() {
	// Buscar total de hogares reportan consumos x seamana
	// debugger;		
	var idSemana =	$("#cboSemana").val();
	//
	$.ajax({		
		url: "g_ValTotalizarHogaresxConsumos.asp?id_Semana=" + idSemana,
		cache: false,
		async: false,
		beforeSend: function(objeto){
			//$("#loader").html("<img src='images/ajax_small.gif'> Totalizando Hogares!");
			$("#cargando").css("display", "block");
		},
		success: function (data) {
			//debugger;
			console.log(data);									
			valor = Number(data).toLocaleString("es-ES", {minimumFractionDigits: 0});	
			$("#totalHogares").html(valor);							
			buscarTotalHogaresInvestigados();			
			$("#cargando").css("display", "none");
			
		},
	});		
}
//
function buscarTotalHogaresInvestigados() {
	// Buscar total de hogares Investigados
	// debugger;		
	var idSemana =	$("#cboSemana").val();
	//
	$.ajax({		
		url: "g_ValTotalizarHogaresInvestigados.asp?id_Semana=" + idSemana,
		cache: false,
		async: false,
		beforeSend: function(objeto){
			//$("#loader").html("<img src='images/ajax_small.gif'> Totalizando Investigados!");
			$("#cargando").css("display", "block");
		},
		success: function (data) {
			debugger;
			console.log(data);									
			valor = Number(data).toLocaleString("es-ES", {minimumFractionDigits: 0});	
			$("#totalInvestigados").html(valor);
			buscarTotalHogaresResueltos();
			$("#cargando").css("display", "none");
			
		},
	});		
}
//
function buscarMonedaPagoFactura() {		
	//	
	//$("#loader").html("<img src='images/ajax_small.gif'> Buscando Moneda!");	
	$("#cargando").css("display", "block");
	//
	$.ajax({
		url: "g_ValBuscarLlenarCmbMonedaPagoFactura.asp",
		type: "POST",
		cache: false,
		async: false,
		//data: ajax,
		dataType: "json",
		success: function (data) {
			//debugger;
			console.log(data);
			let select = $("#MonedaPagoFactura");
			select.find("option").remove();
			select.append("<option value='0' selected disabled> -- Seleccione -- </option>");
			$.each(data.data, function (key, value) {
				select.append("<option value=" + value.id + ">" + value.nombre + "</option>");
			});
			$("#cargando").css("display", "none");			
		},
  });
	
}
//
function buscarResumenSemanal(){	
	//
	debugger;	
	//
	let ajax = {
		id_semana	: $("#cboSemana").val(),
		id_Hogar	: $("#cboHogar").val(),
		id_TipCons	: $("#cboTipoConsumo").val(),
	};
	//
	$.ajax({
		url: "g_ValCalcularResumenSemanal.asp",
		type: 'GET',
		cache: false,
		async: false,
		dataType: 'HTML',
		data: ajax,
		beforeSend: function(){
			$("#cargando").css("display", "block");
		}
	})
	.done(function(data) {
		debugger;
		console.log(data);
		$("#cargando").css("display", "none");
		$("#tabla-resumen").html(data);
		//				
	})
	.fail(function() {
		alert("Fallo - bRS()");
	},'HTML');
	
}
//
function buscarTipoInvestigacion() {		
	//
	$.ajax({		
		url: "g_ValTipoInvestigacion.asp",
		cache: false,
		async: false,
		dataType: "json",
		beforeSend: function(objeto){
			//$("#loader").html("<img src='images/ajax_small.gif'> Buscando Motivos!");
			$("#cargando").css("display", "block");
		},
		success: function (data) {			
			let $select = $("#cboInvestigar");
			$select.find("option").remove();
			$select.append("<option value='0' selected disabled>-- Seleccione --</option>");
			$.each(data, function (i, value) {				
				$select.append("<option value=" + value.id + ">" + value.nombre + "</option>");				
			});				
			$("#cargando").css("display", "none");
			$("#cboSemana").focus();
		},
	});		
}
//
function buscarHogarInvestigado() {
	//
	//debugger;
	//$("#loader").html("<img src='images/ajax_small.gif'> Buscando Hogar Investigado!");	
	$("#cargando").css("display", "block");
	var idConsumo =	$("#cboDetallexDiaSemana").val();	
	let ajax = {
		idConsumo: idConsumo,		
	};
	$.ajax({
		url: "g_ValBuscarHogarInvestigado.asp",
		type: "GET",
		cache: false,
		async: false,
		data: ajax,
		//dataType: "json",
		success: function (data) {			
			//debugger;
			console.log(data);				
			$("#cargando").css("display", "none");			
			if(data==="True"){
				$("#hogarInvestigado").css("display", "block");							
			}else{
				$("#hogarInvestigado").css("display", "none");
			}			
		},
  });
}
//
function CalcularTotalConsumosPendientes() {
	// Totalizar consumos x Validados y Pendientes
	// debugger;
	//
	var idSemana	=	$("#cboSemana").val(); 	
	//		
	$.ajax({		
		url: "g_ValTotalizarConsumosPendientes.asp?id_semana=" + idSemana,
		cache: false,
		async: false,
		dataType: "json",
		beforeSend: function(objeto){
			$("#cargando").css("display", "block");
		},
		success: function (data) {							
			//debugger;				
			console.log(data);				
			var tinvestigados=parseInt(data);			
			valor = Number(tinvestigados).toLocaleString("es-ES", {minimumFractionDigits: 0});	
			$("#totalInvestigados").html(valor);		
			//buscarTotalHogaresReportados();
			$("#cargando").css("display", "none");
		},    				
	});		
}
//
function buscarMotivoInvestigacion() {
	//
	//debugger;
	$("#cargando").css("display", "block");
	var idConsumo =	$("#cboDetallexDiaSemana").val();	
	let ajax = {
		idConsumo: idConsumo,		
	};
	$.ajax({
		url: "g_ValBuscarMotivoInvestigacion.asp",
		type: "GET",
		cache: false,
		async: false,
		data: ajax,
		//dataType: "json",
		success: function (data) {			
			//debugger;
			console.log(data);				
			$("#cargando").css("display", "none");			
			if(data!=="False"){
				$("#motivo").html("<h5><strong>|&nbsp;Motivo: " + data + "&nbsp;|</strong></h5>");
			}else{
				$("#motivo").html("");
			}			
		},
  });
}
//

function buscarTipoComida() {
	//	
	$.ajax({			
		url: "g_ValBuscarTipoComida.asp",
		type: "POST",
		cache: false,
		async: false,
		dataType: "JSON",
		beforeSend: function(objeto){
			$("#cargando").css("display", "block");
		},
		success: function (data) {						
			console.log(data);
			let select = $("#cboTipoComida");
			select.find("option").remove();
			select.append("<option value='' selected disabled>-- Seleccione --</option>");
			$.each(data, function (i, value) {				
				select.append("<option value=" + value.id + ">" + value.nombre + "</option>");				
			});				
			$("#cargando").css("display", "none");			
		},
	});		
}
//
function buscarMonedaPagoFacturaNoMercado() {		
	//		
	$("#cargando").css("display", "block");
	//
	$.ajax({
		url: "g_ValBuscarLlenarCmbMonedaPagoFactura.asp",
		type: "POST",
		cache: false,
		async: false,
		//data: ajax,
		dataType: "json",
		success: function (data) {
			//debugger;
			console.log(data);
			let select = $("#cboMonedaPagoNoMercado");
			select.find("option").remove();
			select.append("<option value='0' selected disabled> -- Seleccione -- </option>");
			$.each(data.data, function (key, value) {
				select.append("<option value=" + value.id + ">" + value.nombre + "</option>");
			});
			$("#cargando").css("display", "none");			
		},
  });
	
}