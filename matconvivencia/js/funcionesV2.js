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
	var opcion=0;
	var combo="";
	var optCat= $("#cboCategoria_A option:selected").val();	
	//
	$("#cboFabricante_A").prop("disabled", true);
	opcion = 2;	
	combo="#cboFabricante_A";
	fillAllCombos(opcion, optCat, combo);	
	//
	$("#cboMarca_A").prop("disabled", true);
	opcion = 3;	
	combo="#cboMarca_A";
	fillAllCombos(opcion, optCat, combo);	
	//
	$("#cboSegmento_A").prop("disabled", true);
	opcion = 4;	
	combo="#cboSegmento_A";
	fillAllCombos(opcion, optCat, combo);	
	//
	debugger;
	$("#cboRangTamanoA").prop("disabled", true);
	opcion = 5;	
	combo="#cboRangTamanoA";
	fillAllCombos(opcion, optCat, combo);	
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
function fillAllCombos(opc,idcat,cmb) {
	//debugger;
	let ajax = { opcion: opc, idCat: idcat, };	
	//	
	$.ajax({
		url: "matconvivencia/llenar_cmb_convivencias_todos.asp",
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
		$(cmb).empty();				
		$(cmb).append("<option selected disabled value='0'>-- Seleccione --</option>");
		for( var i = 0; i < len; i++){
			var id = response.data[i]['id'];
			var nombre = response.data[i]['nombre'];
			$(cmb).append("<option value='"+id+"'>"+nombre+"</option>");
		}
		$(cmb).prop("disabled", false);		
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
}
//
// <!-- CATEGORIA B -->
//


