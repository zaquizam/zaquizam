//
// funcionesV4.js // 09abr21 - 11abr21
//
function Reset() {
	$("#detallesMaestro").css("display", "none");
	$("#cboProcesarFecha").prop("selectedIndex", 0);	
	$("#cboProcesarTrimestre").prop("selectedIndex", 0);
	$("#cboProcesarFecha").focus();	
}
function clear() {
	$("#detallesMaestro").css("display", "none");	
}
//
function buscarFechas() {		
	//
	//debugger;
	$.ajax({
		url: "g_ConvBuscarFechas.asp",
		type: 'POST',
		cache: false,
		async: false,
		dataType: 'JSON',
		// data: {id_Consumo: id},
		beforeSend: function(){			
			$("#cargando").css("display", "block");
		}
	})
	/*Si la consulta se realizo con exito*/
	.done(function(data) {
		//debugger;
		let $select = $("#cboProcesarFecha");
		$select.find("option").remove();
		$select.append("<option value='0' selected disabled>-- Seleccione --</option>");
		$.each(data, function (i, value) {				
			$select.append("<option value=" + value.id + ">" + value.name + "</option>");				
		});				
		$("#cargando").css("display", "none");
		$("#cboProcesarFecha").focus();
		//
	})
	/*Si la consulta Fallo*/
	.fail(function() {
		alert("Fallo - buscarFechas()");
	},'json');		
}
//
function procesarFecha() {
	//
	$("#cargando").css("display", "block");
	$("#detallesPaso2").css("display", "none");		
	$("#detallesPaso3").css("display", "none");		
	$("#detallesPaso4").css("display", "none");		
	document.getElementById('cargando').style.display = 'block'; 	
	$("#cboProcesarTrimestre").prop("selectedIndex", 0);	
	var idMes=$("#cboProcesarFecha").val();
	clear();
	paso1(idMes); 	
	paso2(idMes); 		
	paso3(idMes); 		
	paso4(idMes); 
	document.getElementById('cargando').style.display = 'none';
	$("#detallesMaestro").css("display", "block");
	//
}
//
function procesarTrimestre() {
	//
	$("#cargando").css("display", "block");		
	$("#detallesPaso2").css("display", "none");		
	$("#detallesPaso3").css("display", "none");		
	$("#detallesPaso4").css("display", "none");		
	document.getElementById('cargando').style.display = 'block'; 
	$("#cboProcesarFecha").prop("selectedIndex", 0);	
	var idMes=$("#cboProcesarTrimestre").val();
	//debugger;
	//
	clear();
	paso1(idMes); 	
	paso2(idMes); 		
	// paso3(idMes); 		
	// paso4(idMes); 
	var id_Cliente=sessionStorage.getItem("idCliente");
	if(id_Cliente==3 || id_Cliente==1){
		paso3(idMes); 		
		paso4(idMes); 			
	}	
	document.getElementById('cargando').style.display = 'none';
	$("#detallesMaestro").css("display", "block");
	//
}
//
function paso1(id) {
	// Calculo Pregunta #1
	$("#Paso1").html("");
	//
	var idMes= id; //$("#cboProcesarFecha").val();
	//
	$.ajax({
		url: "g_ConvCalcularCombinacionesxHogaresPaso1.asp",
		type: 'GET',
		cache: false,
		async: false,
		dataType: 'html',
		data: {id_Mes: idMes},
		// beforeSend: function(){			
			// $("#cargando").css("display", "block");
		// }
	})
	/*Si la consulta se realizo con exito*/
	.done(function(data) {
		debugger
		console.log(data);
		$("#detallesPaso1").css("display", "block");
		$("#Paso1").html(data);
		//		
	})
	/*Si la consulta Fallo*/
	.fail(function() {
		alert("Fallo - paso1()");
	},'html');	
	//
}
//
function paso2(id) {
	// Calculo Pregunta #2
	$("#Paso2").html("");
	//
	var idMes= id; //$("#cboProcesarFecha").val();
	//
	$.ajax({
		url: "g_ConvCalcularCombinacionesxHogaresPaso2.asp",
		type: 'GET',
		cache: false,
		async: false,
		dataType: 'html',
		data: {id_Mes: idMes},
		// beforeSend: function(){			
			// $("#cargando").css("display", "block");
		// }
	})
	/*Si la consulta se realizo con exito*/
	.done(function(data) {
		debugger
		console.log(data);
		$("#detallesPaso2").css("display", "block");
		$("#Paso2").html(data);
		//		
	})
	/*Si la consulta Fallo*/
	.fail(function() {
		alert("Fallo - paso2()");
	},'html');	
	//
}
//
function paso3(id) {
	// Calculo Pregunta #3	
	$("#Paso3").html("");   	
	//
	var idMes=id; //$("#cboProcesarFecha").val();	
	//
	$.ajax({
		url: "g_ConvCalcularTotalHogaresxConsumoPaso3.asp",
		type: 'GET',
		cache: false,
		async: false,
		dataType: 'html',
		data: {id_Mes: idMes},
		beforeSend: function(){			
			$("#cargando").css("display", "block");
		}
	})
	/*Si la consulta se realizo con exito*/
	.done(function(data) {
		console.log(data);
		$("#detallesPaso3").css("display", "block");
		//
		if(data==="0"){
			$("#Paso3").html("....NO HAY DATOS PARA EL MES SELECCIONADO....");
		}else{
			$("#Paso3").html(data + " %");			
		}
	})
	/*Si la consulta Fallo*/
	.fail(function() {
		alert("Fallo - paso3()");
	},'html');	
	//
}
//
function paso4(id) {
	// Calculo Pregunta #4
	$("#Paso4").html("");
	//
	var idMes=id; //$("#cboProcesarFecha").val();
	//
	$.ajax({
		url: "g_ConvCalcularCombinacionesxHogaresPaso4.asp",
		type: 'GET',
		cache: false,
		async: false,
		dataType: 'html',
		data: {id_Mes: idMes},
		// beforeSend: function(){			
			// $("#cargando").css("display", "block");
		// }
	})
	/*Si la consulta se realizo con exito*/
	.done(function(data) {
		console.log(data);
		$("#detallesPaso4").css("display", "block");
		$("#Paso4").html(data);
		//		
	})
	/*Si la consulta Fallo*/
	.fail(function() {
		alert("Fallo - paso4()");
	},'html');	
	//
}
//

