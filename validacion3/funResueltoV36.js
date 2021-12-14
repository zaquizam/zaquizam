//
// funResueltoV36.js - 19ene21 - 06abr21ene21
//
function llenarCmbConsumosResueltos() {
	//	
	// debugger;
	//
	$.ajax({		
		url: "g_rRevInvLlenarCmbHogarInvestigado.asp",
		cache: false,
		async: false,
		dataType: "json",		
		beforeSend: function(objeto){
			$("#loader").html("<img src='images/ajax_small.gif'> Buscando!");
		},
		success: function (data) {
			//debugger;
			console.log(data);	
			var contador=0;
			let $select = $("#cboHogar");
			$select.find("option").remove();
			$select.append("<option value='0' selected disabled>-- Seleccione --</option>");
			$.each(data, function (i, value) {				
				contador++;
				$select.append("<option value=" + value.Id + ">" + value.Nombre + "</option>");				
			});				
			$("#loader").html("");
			$("#totalHogares").html(contador);			
			$("#cboHogar").focus();
		},
	});		
}
//
function buscarDetalleConsumoResueltoxDia() {
	//
	// Buscar el detalle los del consumo individual
	debugger;
	//
	Reset_Resuelto();
	buscarCadena(0);
	buscarCanal();
	buscarHogarResuelto();
	buscarAltaHogarResuelto();
	//buscarHogarInvestigado();
	//buscarMotivoInvestigacion();
	//	
	var idConsumo = $("#cboConsumoInvestigado").val();
	if (idConsumo==="0" || idConsumo==null || idConsumo==undefined){
		// Validacion Normal
		var idConsumo =	$("#cboDetallexDiaSemana").val();
	}else{
		//Reset_Resuelto();
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
			sessionStorage.setItem('validado', data[0].validado );
			sessionStorage.setItem('investigado', data[0].investigar );
			sessionStorage.setItem('resuelto', data[0].resuelto );						
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
			buscarImagenFacturaResuelto();
			//
		},    				
	});		
}
//
function buscarHogarResuelto() {
	//	
	// buscar la respuesta de Investigaciones			
	//debugger;
	//	
	$("#cargando").css("display", "block");
	var idConsumo =	$("#cboConsumoInvestigado").val();	
	if (idConsumo==="0" || idConsumo==null || idConsumo==undefined){
		// Validacion Normal
		var idConsumo =	$("#cboDetallexDiaSemana").val();	
	}
	//	
	let ajax = {
		idConsumo: idConsumo,		
	};
	//
	$.ajax({
		url: "g_ValBuscarConsumoHogarResuelto.asp",
		type: 'GET',
		cache: false,
		async: false,		
		 data: ajax,		
		beforeSend: function(){			
			$("#cargando").css("display", "block");
		}
	})
	/*Si la consulta se realizo con exito*/
	.done(function(data) {
		//debugger;
			console.log(data);
			var motivoInv	=	data[0].motivo;
			var comentaInv 	=	data[0].comentario;
			var comentaRsp 	=	data[0].respuesta;
			//									
			$("#cargando").css("display", "none");
			$("#hogarResuelto").css("display", "block");
			//
			if(data=="False"){
				$("#motivoInv").val("No Aplica");
				$("#comentarioInv").val("No Aplica");
				$("#motivoRsp").val("No Aplica");				
			}else{
				 $("#motivoInv").html("Motivo Investigacion: " + motivoInv);
				 $("#comentarioInv").html("Comentario Adicional: " + comentaInv);
				 $("#motivoRsp").html("Respuesta de Investigacion: " + comentaRsp);
			}
		//
	})
	/*Si la consulta Fallo*/
	.fail(function() {
		alert("Fallo - bHResuelto("+idConsumo+")");
	},'json');
}
//
function buscarImagenFacturaResuelto() {
	//	
	buscarDetallexProductoFacturaResuelto();
	//
	//debugger;
	var idConsumo	=	$("#cboConsumoInvestigado").val();
	//
	if (idConsumo==="0" || idConsumo==null || idConsumo==undefined){
		// Validacion Normal
		var idConsumo =	$("#cboDetallexDiaSemana").val();
	}
	//	
	// if (idConsumo == null || idConsumo == 0) {
		// swal("Aviso..!", "Faltan Datos, para Procesar...!", "error");
		// $("#cboSemana").focus();
		// return false;
	// }
	//	
	$.ajax({
		url: "g_ValBuscarImagenFacturaxDia.asp",
		type: 'GET',
		cache: false,
		async: false,
		dataType: 'JSON',
		data: {id_Consumo: idConsumo},
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
		buscarResumenSemanalResuelto();
		//
	})
	/*Si la consulta Fallo*/
	.fail(function() {
		alert("Fallo - bIFR()");
	},'json');
//	
}
//
function buscarDetallexProductoFacturaResuelto() {
	//
	//debugger;
	var idConsumo =	$("#cboConsumoInvestigado").val();	
	//
	if (idConsumo==="0" || idConsumo==null || idConsumo==undefined){
		// Validacion Normal
		var idConsumo =	$("#cboDetallexDiaSemana").val();
	}
	//	
	$.ajax({
		url: "g_ValBuscarDetallesxProductosxFacturaResuelto.asp",
		type: 'GET',
		cache: false,
		async: false,
		dataType: 'HTML',
		data: {id_Consumo: idConsumo},
		beforeSend: function(){
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
		alert("Fallo - bDxPFR()");
	},'HTML');
		
}
//
function buscarResumenSemanalResuelto(){	
	//
	//debugger;	
	//
	var idConsumo =	$("#cboConsumoInvestigado").val();	
	//
	if (idConsumo==="0" || idConsumo==null || idConsumo==undefined){
		// Validacion Normal
		var idConsumo =	$("#cboDetallexDiaSemana").val();
		var idHogar	  = $("#cboHogar").val(); 
		var idSemana  = $("#cboSemana").val();
	}else{	
		var idItems		=	$("#cboConsumoInvestigado option:selected" ).text().trim();
		var fields 		=	idItems.split('-');	
		var idHogar		=	fields[0];
	}
	
	let ajax = {
		id_Semana	: parseInt($("#cboSemana").val()),
		id_Hogar	: parseInt(idHogar),		
		id_Consumo 	: parseInt(idConsumo),
	};
	//
	$.ajax({
		url: "g_ValCalcularResumenSemanalResuelto.asp",
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
		//debugger;
		console.log(data);
		$("#cargando").css("display", "none");
		$("#tabla-resumen").html(data);
		//				
	})
	/*Si la consulta Fallo*/
	.fail(function() {
		alert("Fallo - bRSR()");
	},'HTML');
	
}
//
function buscarAltaHogarResuelto() {
	// Buscar la fecha de Creacion/Ingreso del hogar
	// debugger;		
	//
	var idConsumo	=	$("#cboConsumoInvestigado").val();
	//
	if (idConsumo==="0" || idConsumo==null || idConsumo==undefined){
		// Validacion normal
		var idHogar	 = $("#cboHogar").val(); 
	}else{
		var idItems		=	$("#cboConsumoInvestigado option:selected" ).text().trim();
		var fields 		=	idItems.split('-');	
		var idHogar		=	parseInt(fields[0]);
	}
	//
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
			var total    = data[0].total;
			//
			if (fecha == null || fecha == "" || fecha.length == 0 || fecha == undefined ) {
				$("#responsableHogar").html(nombre+" "+apellido);
				$("#celularHogar").html(celular);
				$("#altaHogar").html("Sin Registro");			
				$("#integrantesHogar").html(total);
			}else{
				$("#responsableHogar").html(nombre+" "+apellido);
				$("#celularHogar").html(celular);
				$("#altaHogar").html(fecha);
				$("#integrantesHogar").html(total);
			}
			$("#cargando").css("display", "none");			
		},
	});		
}
//
function buscarTotalHogaresResueltos() {
	//
	// Buscar total de hogares Resueltos
	//
	//debugger;
	var validado = sessionStorage.getItem("Convalidado");
	if ( validado === "0" ) {
		return;		
	}	
	var idSemana =	$("#cboSemana").val();	
	//
	$.ajax({
		url: "g_ValBuscarLlenarCmbConsumosResueltos.asp?id_Semana=" + idSemana,
		type: "POST",
		cache: false,
		async: false,
		//data: ajax,
		dataType: "json",
		success: function (data) {
			//debugger;
			console.log(data);
			let select = $("#cboConsumoInvestigado");
			select.find("option").remove();
			select.append("<option value='0' selected disabled> -- Seleccione -- </option>");
			$.each(data.data, function (key, value) {
				select.append("<option value=" + value.id + ">" + value.nombre + "</option>");
			});
			//
			 var length = $("#cboConsumoInvestigado > option").length;
			if (length<=2) {
				$("#cboConsumoInvestigado").prop("selectedIndex", 1);
				var value =  $("#cboConsumoInvestigado").val();
				if (value <=0){
					length = 0;	
				}else{
					$("#cboConsumoInvestigado").prop("selectedIndex", 0);
					length = length - 1
				}				
			}else{
				length = length - 1;
			}
			$("#totalResueltos").html(length);
			//
			$("#cargando").css("display", "none");			
		},
  });
}