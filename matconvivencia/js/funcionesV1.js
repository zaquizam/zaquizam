//
// funcionesV1.js // 24may21 - 24may21
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
function LlenarCombos() {		
	//
	//debugger;
	var id=0;
	$("#cboCategoria_A").prop("disabled", true);
	$("#cboCategoria_B").prop("disabled", true);
	let ajax = { opcion: 1,	id1: id, id2: id, id3: id, id4: id, id5: id, };		
	//	
	$.ajax({
		url: "matconvivencia/llenar_cmb_convivencias.asp",
		type: "POST", //Use "PUT" for HTTP PUT methods
		dataType: 'json',   
		data:  ajax,		
		beforeSend: function(){			
			$("#cargando").css("display", "block");
		}		
	})
	.done (function(response, textStatus, jqXHR) { 
		// console.log(response);
		// debugger;
		var len = response.data.length;
		$("#cboCategoria_A").empty();				
		$("#cboCategoria_A").append("<option selected disabled value='0'>-- Seleccione --</option>");
		for( var i = 0; i < len; i++){
			var id = response.data[i]['id'];
			var nombre = response.data[i]['nombre'];
			$("#cboCategoria_A").append("<option value='"+id+"'>"+nombre+"</option>");
		}
		$("#cboCategoria_B").empty();				
		$("#cboCategoria_A").find("option").clone().appendTo("#cboCategoria_B");
	})
	.fail (function(jqXHR, textStatus, errorThrown) { 
		//alert("Error " + errorThrown); 
		swal("Algo salio mal.!",errorThrown, "error");
	})
	.always (function(jqXHROrData, textStatus, jqXHROrErrorThrown) { 
		//alert("complete"); 
		$("#cargando").css("display", "none");
		$("#cboCategoria_A").prop("disabled", false);
		$("#cboCategoria_B").prop("disabled", false);
	});		
	//
}
//
// <!-- CATEGORIA A -->
//
$("#cboCategoria_A").on("change", function() {
    // Fill combo Fabricante
	var id=0;
	var optCat= $("#cboCategoria_A option:selected").val();
	debugger;
	//
	$("#cboFabricante_A").prop("disabled", true);
	let ajax = { opcion: 2,	id1: optCat, id2: id, id3: id, id4: id, id5: id, };
	//	
	$.ajax({
		url: "matconvivencia/llenar_cmb_convivencias.asp",
		type: "POST", 
		dataType: 'json',   
		data:  ajax,		
		beforeSend: function(){			
			$("#cargando").css("display", "block");
		}		
	})
	.done (function(response, textStatus, jqXHR) { 
		console.log(response);
		debugger;
		var len = response.data.length;
		$("#cboFabricante_A").empty();				
		$("#cboFabricante_A").append("<option selected disabled value='0'>-- Seleccione --</option>");
		for( var i = 0; i < len; i++){
			var id = response.data[i]['id'];
			var nombre = response.data[i]['nombre'];
			$("#cboFabricante_A").append("<option value='"+id+"'>"+nombre+"</option>");
		}		
	})
	.fail (function(jqXHR, textStatus, errorThrown) { 
		swal("Algo salio mal.!",errorThrown, "error");
	})
	.always (function(jqXHROrData, textStatus, jqXHROrErrorThrown) { 
		//alert("complete"); 
		$("#cargando").css("display", "none");
		$("#cboFabricante_A").prop("disabled", false);		
	});		
	//
});
//
$("#cboFabricante_A").on("change", function() {
    // Fill combo Marca
	var id=0;
	var optCat = $("#cboCategoria_A option:selected").val();
	var optFab = $("#cboFabricante_A option:selected").val();		
	//
	$("#cboMarca_A").prop("disabled", true);
	let ajax = { opcion: 3,	id1: optCat, id2: optFab, id3: id, id4: id, id5: id, };
	debugger;
	//	
	$.ajax({
		url: "matconvivencia/llenar_cmb_convivencias.asp",
		type: "POST", 
		dataType: 'json',   
		data:  ajax,		
		beforeSend: function(){			
			$("#cargando").css("display", "block");
		}		
	})
	.done (function(response, textStatus, jqXHR) { 
		console.log(response);
		debugger;
		var len = response.data.length;
		$("#cboMarca_A").empty();				
		$("#cboMarca_A").append("<option selected disabled value='0'>-- Seleccione --</option>");
		for( var i = 0; i < len; i++){
			var id = response.data[i]['id'];
			var nombre = response.data[i]['nombre'];
			$("#cboMarca_A").append("<option value='"+id+"'>"+nombre+"</option>");
		}		
	})
	.fail (function(jqXHR, textStatus, errorThrown) { 
		swal("Algo salio mal.!",errorThrown, "error"); 
	})
	.always (function(jqXHROrData, textStatus, jqXHROrErrorThrown) { 
		//alert("complete"); 
		$("#cargando").css("display", "none");
		$("#cboMarca_A").prop("disabled", false);		
	});		
	//
});
//
$("#cboMarca_A").on("change", function() {
    // Fill combo Segmento
	var id=0;
	var optCat = $("#cboCategoria_A option:selected").val();	
	//
	$("#cboSegmento_A").prop("disabled", true);
	let ajax = { opcion: 4,	id1: optCat, id2: id, id3: id, id4: id, id5: id, };
	debugger;
	//	
	$.ajax({
		url: "matconvivencia/llenar_cmb_convivencias.asp",
		type: "POST", 
		dataType: 'json',   
		data:  ajax,		
		beforeSend: function(){			
			$("#cargando").css("display", "block");
		}		
	})
	.done (function(response, textStatus, jqXHR) { 
		console.log(response);
		debugger;
		var len = response.data.length;
		$("#cboSegmento_A").empty();				
		$("#cboSegmento_A").append("<option selected disabled value='0'>-- Seleccione --</option>");
		for( var i = 0; i < len; i++){
			var id = response.data[i]['id'];
			var nombre = response.data[i]['nombre'];
			$("#cboSegmento_A").append("<option value='"+id+"'>"+nombre+"</option>");
		}		
	})
	.fail (function(jqXHR, textStatus, errorThrown) { 
		swal("Algo salio mal.!",errorThrown, "error"); 
	})
	.always (function(jqXHROrData, textStatus, jqXHROrErrorThrown) { 
		//alert("complete"); 
		$("#cargando").css("display", "none");
		$("#cboSegmento_A").prop("disabled", false);		
	});		
	//
});
//
// <!-- CATEGORIA B -->
//
$("#cboCategoria_B").on("change", function() {
    // Fill combo Fabricante
	var id=0;
	var optCat= $("#cboCategoria_B option:selected").val();
	debugger;
	//
	$("#cboFabricante_B").prop("disabled", true);
	let ajax = { opcion: 2,	id1: optCat, id2: id, id3: id, id4: id, id5: id, };
	//	
	$.ajax({
		url: "matconvivencia/llenar_cmb_convivencias.asp",
		type: "POST", 
		dataType: 'json',   
		data:  ajax,		
		beforeSend: function(){			
			$("#cargando").css("display", "block");
		}		
	})
	.done (function(response, textStatus, jqXHR) { 
		console.log(response);
		debugger;
		var len = response.data.length;
		$("#cboFabricante_B").empty();				
		$("#cboFabricante_B").append("<option selected disabled value='0'>-- Seleccione --</option>");
		for( var i = 0; i < len; i++){
			var id = response.data[i]['id'];
			var nombre = response.data[i]['nombre'];
			$("#cboFabricante_B").append("<option value='"+id+"'>"+nombre+"</option>");
		}		
	})
	.fail (function(jqXHR, textStatus, errorThrown) { 
		swal("Algo salio mal.!",errorThrown, "error");
	})
	.always (function(jqXHROrData, textStatus, jqXHROrErrorThrown) { 
		//alert("complete"); 
		$("#cargando").css("display", "none");
		$("#cboFabricante_B").prop("disabled", false);		
	});		
	//
});
//
$("#cboFabricante_B").on("change", function() {
    // Fill combo Marca
	var id=0;
	var optCat = $("#cboCategoria_B option:selected").val();
	var optFab = $("#cboFabricante_B option:selected").val();		
	//
	$("#cboMarca_B").prop("disabled", true);
	let ajax = { opcion: 3,	id1: optCat, id2: optFab, id3: id, id4: id, id5: id, };
	debugger;
	//	
	$.ajax({
		url: "matconvivencia/llenar_cmb_convivencias.asp",
		type: "POST", 
		dataType: 'json',   
		data:  ajax,		
		beforeSend: function(){			
			$("#cargando").css("display", "block");
		}		
	})
	.done (function(response, textStatus, jqXHR) { 
		console.log(response);
		debugger;
		var len = response.data.length;
		$("#cboMarca_B").empty();				
		$("#cboMarca_B").append("<option selected disabled value='0'>-- Seleccione --</option>");
		for( var i = 0; i < len; i++){
			var id = response.data[i]['id'];
			var nombre = response.data[i]['nombre'];
			$("#cboMarca_B").append("<option value='"+id+"'>"+nombre+"</option>");
		}		
	})
	.fail (function(jqXHR, textStatus, errorThrown) { 
		swal("Algo salio mal.!",errorThrown, "error"); 
	})
	.always (function(jqXHROrData, textStatus, jqXHROrErrorThrown) { 
		//alert("complete"); 
		$("#cargando").css("display", "none");
		$("#cboMarca_B").prop("disabled", false);		
	});		
	//
});
//
$("#cboMarca_B").on("change", function() {
    // Fill combo Segmento
	var id=0;
	var optCat = $("#cboCategoria_B option:selected").val();	
	//
	$("#cboSegmento_B").prop("disabled", true);
	let ajax = { opcion: 4,	id1: optCat, id2: id, id3: id, id4: id, id5: id, };
	debugger;
	//	
	$.ajax({
		url: "matconvivencia/llenar_cmb_convivencias.asp",
		type: "POST", 
		dataType: 'json',   
		data:  ajax,		
		beforeSend: function(){			
			$("#cargando").css("display", "block");
		}		
	})
	.done (function(response, textStatus, jqXHR) { 
		console.log(response);
		debugger;
		var len = response.data.length;
		$("#cboSegmento_B").empty();				
		$("#cboSegmento_B").append("<option selected disabled value='0'>-- Seleccione --</option>");
		for( var i = 0; i < len; i++){
			var id = response.data[i]['id'];
			var nombre = response.data[i]['nombre'];
			$("#cboSegmento_B").append("<option value='"+id+"'>"+nombre+"</option>");
		}		
	})
	.fail (function(jqXHR, textStatus, errorThrown) { 
		swal("Algo salio mal.!",errorThrown, "error"); 
	})
	.always (function(jqXHROrData, textStatus, jqXHROrErrorThrown) { 
		//alert("complete"); 
		$("#cargando").css("display", "none");
		$("#cboSegmento_B").prop("disabled", false);		
	});		
	//
});


