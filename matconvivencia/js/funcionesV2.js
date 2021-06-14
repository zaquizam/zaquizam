//
// funcionesV1.js // 24may21 - 14jun21
//
function Reset() {
	$("#detallesMaestro").css("display", "none");
	//
	$("#cboCategoria_A").prop("selectedIndex", 0);
	$("#cboFabricante_A").find('option:not(:first)').remove();
	$("#cboFabricante_A").prop("selectedIndex", 0);
	$("#cboMarca_A").find('option:not(:first)').remove();
	$("#cboMarca_A").prop("selectedIndex", 0);
	$("#cboSegmento_A").find('option:not(:first)').remove();
	$("#cboSegmento_A").prop("selectedIndex", 0);
	$("#cboRangTamanoA").find('option:not(:first)').remove();
	$("#cboRangTamanoA").prop("selectedIndex", 0);
	//
	$("#cboCategoria_B").prop("selectedIndex", 0);
	$("#cboFabricante_B").find('option:not(:first)').remove();
	$("#cboFabricante_B").prop("selectedIndex", 0);
	$("#cboMarca_B").find('option:not(:first)').remove();
	$("#cboMarca_B").prop("selectedIndex", 0);
	$("#cboSegmento_B").find('option:not(:first)').remove();
	$("#cboSegmento_B").prop("selectedIndex", 0);
	$("#cboRangTamanoB").find('option:not(:first)').remove();
	$("#cboRangTamanoB").prop("selectedIndex", 0);
	//
	LlenarCombos();
	//
	$("#cboCategoria_A").prop("selectedIndex", 0);
	$("#cboCategoria_A").focus();
}
//
function clear() {
	$("#detallesMaestro").css("display", "none");
}
//
function LlenarCombos() {
	LlenarCategoria()
	LlenarArea();
	LlenarPeriodo();
}
//
function LlenarCategoria() {
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
		//$("#cboCategoria_A").multiselect('destroy');
		$("#cboCategoria_A").empty();
		$("#cboCategoria_A").append("<option selected disabled value='0'>-- Seleccione --</option>");
		for( var i = 0; i < len; i++){
			var id = response.data[i]['id'];
			var nombre = response.data[i]['nombre'];
			$("#cboCategoria_A").append("<option value='"+id+"'>"+nombre+"</option>");
		}
		//		
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
function LlenarArea() {
	//
	//debugger;
	var id=0;
	$("#cboArea").prop("disabled", true);

	let ajax = { opcion: 2, };
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
		$("#cboArea").multiselect('destroy');
		var len = response.data.length;
		$("#cboArea").empty();
		//$("#cboArea").append("<option selected disabled value='0'>-- Seleccione --</option>");
		for( var i = 0; i < len; i++){
			var id = response.data[i]['id'];
			var nombre = response.data[i]['nombre'];
			$("#cboArea").append("<option value='"+id+"'>"+nombre+"</option>");
		}

		$("#cboArea").multiselect({
  			nonSelectedText: '-- Seleccione --',
  			buttonWidth: '285px',
			includeSelectAllOption: true,
            //enableFiltering: true
 		});

	})
	.fail (function(jqXHR, textStatus, errorThrown) {
		//alert("Error " + errorThrown);
		swal("Algo salio mal.!",errorThrown, "error");
	})
	.always (function(jqXHROrData, textStatus, jqXHROrErrorThrown) {
		//alert("complete");
		$("#cargando").css("display", "none");
		$("#cboArea").prop("disabled", false);
		//$("#cboArea").multiselect('rebuild');
	});
	//
}
//
function LlenarPeriodo() {
	//
	//debugger;
	//
	var id=0;
	$("#cboPeriodo").prop("disabled", true);

	let ajax = { opcion: 3, };
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
		$("#cboPeriodo").multiselect('destroy');
		var len = response.data.length;
		$("#cboPeriodo").empty();
		//$("#cboPeriodo").append("<option selected disabled value='0'>-- Seleccione --</option>");
		for( var i = 0; i < len; i++){
			var id = response.data[i]['id'];
			var nombre = response.data[i]['nombre'];
			$("#cboPeriodo").append("<option value='"+id+"'>"+nombre+"</option>");
		}

		$("#cboPeriodo").multiselect({
  			nonSelectedText: '-- Seleccione --',
  			buttonWidth: '285px',
			includeSelectAllOption: true,
            //enableFiltering: true
 		});

	})
	.fail (function(jqXHR, textStatus, errorThrown) {
		//alert("Error " + errorThrown);
		swal("Algo salio mal.!",errorThrown, "error");
	})
	.always (function(jqXHROrData, textStatus, jqXHROrErrorThrown) {
		//alert("complete");
		$("#cargando").css("display", "none");
		$("#cboPeriodo").prop("disabled", false);
		//$("#cboPeriodo").multiselect('rebuild');
	});
	//
}
//
// <!-- CATEGORIA A -->
//
$("#cboCategoria_A").on("change", function() {
    // Fill combo Fabricante + Marca + Segmento + Rango Tamaño
	var opcion=0;
	var combo="";
	var optCat= $("#cboCategoria_A option:selected").val();
	//
	$("#cboFabricante_A").prop("disabled", true);
	opcion = 2;
	combo="#cboFabricante_A";
	fillAllCombos1(opcion, optCat, combo);
	//
	$("#cboMarca_A").prop("disabled", true);
	opcion = 3;
	combo="#cboMarca_A";
	fillAllCombos1(opcion, optCat, combo);
	//
	$("#cboSegmento_A").prop("disabled", true);
	opcion = 4;
	combo="#cboSegmento_A";
	fillAllCombos1(opcion, optCat, combo);
	//
	$("#cboRangTamanoA").prop("disabled", true);
	opcion = 5;
	combo="#cboRangTamanoA";
	fillAllCombos1(opcion, optCat, combo);
	//
});
//
$("#cboFabricante_A").on("change", function() {
    // Fill combo Marca + Segmento + Rango Tamaño
	//debugger;
	var opcion = 0;
	var optCat = $("#cboCategoria_A option:selected").val();
	var optFab = $("#cboFabricante_A option:selected").val();
	//
	$("#cboMarca_A").prop("disabled", true);
	opcion = 1;
	combo = "#cboMarca_A";
	fillAllCombos2(opcion, optCat, optFab, combo);
	//
	$("#cboSegmento_A").prop("disabled", true);
	opcion = 2;
	combo = "#cboSegmento_A";
	fillAllCombos2(opcion, optCat, optFab, combo);
	//
	$("#cboRangTamanoA").prop("disabled", true);
	opcion = 3;
	combo = "#cboRangTamanoA";
	fillAllCombos2(opcion, optCat, optFab, combo);
	//
});
//
$("#cboMarca_A").on("change", function() {
    // Fill combo Segmento
	//debugger;
	var opcion = 0;
	var optCat = $("#cboCategoria_A option:selected").val();
	var optFab = $("#cboFabricante_A option:selected").val();
	var optMar = $("#cboMarca_A option:selected").val();
	//
	opcion = 1;
	$("#cboSegmento_A").prop("disabled", true);
	combo = "#cboSegmento_A";
	fillAllCombos3(opcion, optCat, optFab, optMar, combo);
	//
	$("#cboRangTamanoA").prop("disabled", true);
	opcion = 2;
	combo = "#cboRangTamanoA";
	fillAllCombos3(opcion, optCat, optFab, optMar, combo);
	//
});
//
$("#cboSegmento_A").on("change", function() {
    // Fill combo Segmento
	//debugger;
	var opcion = 0;
	var optCat = $("#cboCategoria_A option:selected").val();
	var optFab = $("#cboFabricante_A option:selected").val();
	var optMar = $("#cboMarca_A option:selected").val();
	var optSeg = $("#cboSegmento_A option:selected").val();
	//
	opcion = 1;
	$("#cboRangTamanoA").prop("disabled", true);
	combo = "#cboRangTamanoA";
	fillAllCombos4(opcion, optCat, optFab, optMar, optSeg, combo);
	//
});
//
// <!-- CATEGORIA B -->
//
$("#cboCategoria_B").on("change", function() {
    // Fill combo Fabricante + Marca + Segmento + Rango Tamaño
	var opcion=0;
	var combo="";
	var optCat= $("#cboCategoria_B option:selected").val();
	//
	$("#cboFabricante_B").prop("disabled", true);
	opcion = 2;
	combo="#cboFabricante_B";
	fillAllCombos1(opcion, optCat, combo);
	//
	$("#cboMarca_B").prop("disabled", true);
	opcion = 3;
	combo="#cboMarca_B";
	fillAllCombos1(opcion, optCat, combo);
	//
	$("#cboSegmento_B").prop("disabled", true);
	opcion = 4;
	combo="#cboSegmento_B";
	fillAllCombos1(opcion, optCat, combo);
	//
	$("#cboRangTamanoB").prop("disabled", true);
	opcion = 5;
	combo="#cboRangTamanoB";
	fillAllCombos1(opcion, optCat, combo);
	//
});
//
$("#cboFabricante_B").on("change", function() {
    // Fill combo Marca + Segmento + Rango Tamaño
	//debugger;
	var opcion = 0;
	var optCat = $("#cboCategoria_B option:selected").val();
	var optFab = $("#cboFabricante_B option:selected").val();
	//
	$("#cboMarca_B").prop("disabled", true);
	opcion = 1;
	combo = "#cboMarca_B";
	fillAllCombos2(opcion, optCat, optFab, combo);
	//
	$("#cboSegmento_B").prop("disabled", true);
	opcion = 2;
	combo = "#cboSegmento_B";
	fillAllCombos2(opcion, optCat, optFab, combo);
	//
	$("#cboRangTamanoB").prop("disabled", true);
	opcion = 3;
	combo = "#cboRangTamanoB";
	fillAllCombos2(opcion, optCat, optFab, combo);
	//
});
//
$("#cboMarca_B").on("change", function() {
    // Fill combo Segmento
	//debugger;
	var opcion = 0;
	var optCat = $("#cboCategoria_B option:selected").val();
	var optFab = $("#cboFabricante_B option:selected").val();
	var optMar = $("#cboMarca_B option:selected").val();
	//
	opcion = 1;
	$("#cboSegmento_B").prop("disabled", true);
	combo = "#cboSegmento_B";
	fillAllCombos3(opcion, optCat, optFab, optMar, combo);
	//
	$("#cboRangTamanoB").prop("disabled", true);
	opcion = 2;
	combo = "#cboRangTamanoB";
	fillAllCombos3(opcion, optCat, optFab, optMar, combo);
	//
});
//
$("#cboSegmento_B").on("change", function() {
    // Fill combo Segmento
	//debugger;
	var opcion = 0;
	var optCat = $("#cboCategoria_B option:selected").val();
	var optFab = $("#cboFabricante_B option:selected").val();
	var optMar = $("#cboMarca_B option:selected").val();
	var optSeg = $("#cboSegmento_B option:selected").val();
	//
	opcion = 1;
	$("#cboRangTamanoB").prop("disabled", true);
	combo = "#cboRangTamanoB";
	fillAllCombos4(opcion, optCat, optFab, optMar, optSeg, combo);
	//
});
//
// <!-- FUNCIONES A y B -->
//
function fillAllCombos1(opc,idcat,cmb) {
	//debugger;
	let ajax = { opcion: opc, idCat: idcat, };
	//
	$.ajax({
		url: "matconvivencia/llenar_cmb_x_cat.asp",
		type: "POST",
		dataType: 'json',
		data:  ajax,
		beforeSend: function(){
			$("#cargando").css("display", "block");
		}
	})
	.done (function(response, textStatus, jqXHR) {
		//console.log(response);
		//debugger;
		var len = response.data.length;
		$(cmb).empty();
		$(cmb).append("<option selected disabled value='0'>-- Seleccione --</option>");
		for( var i = 0; i < len; i++){
			var id = response.data[i]['id'];
			var nombre = response.data[i]['nombre'];
			$(cmb).append("<option value='"+id+"'>"+nombre+"</option>");
		}
	})
	.fail (function(jqXHR, textStatus, errorThrown) {
		swal("Algo salio mal.!",errorThrown, "error");
	})
	.always (function(jqXHROrData, textStatus, jqXHROrErrorThrown) {
		//alert("complete");
		$("#cargando").css("display", "none");
		$(cmb).prop("disabled", false);
	});
	//
}
//
function fillAllCombos2(opc, idcat, idfab, cmb) {
	//debugger;
	let ajax = { opcion: opc, idCat: idcat, idFab: idfab, };
	//
	$.ajax({
		url: "matconvivencia/llenar_cmb_x_Cat-Fab.asp",
		type: "POST",
		dataType: 'json',
		data:  ajax,
		beforeSend: function(){
			$("#cargando").css("display", "block");
		}
	})
	.done (function(response, textStatus, jqXHR) {
		//console.log(response);
		//debugger;
		var len = response.data.length;
		$(cmb).empty();
		$(cmb).append("<option selected disabled value='0'>-- Seleccione --</option>");
		for( var i = 0; i < len; i++){
			var id = response.data[i]['id'];
			var nombre = response.data[i]['nombre'];
			$(cmb).append("<option value='"+id+"'>"+nombre+"</option>");
		}
	})
	.fail (function(jqXHR, textStatus, errorThrown) {
		swal("Algo salio mal.!",errorThrown, "error");
	})
	.always (function(jqXHROrData, textStatus, jqXHROrErrorThrown) {
		//alert("complete");
		$("#cargando").css("display", "none");
		$(cmb).prop("disabled", false);
	});
	//
}
//
function fillAllCombos3(opc, idcat, idfab, idmar, cmb) {
	//debugger;
	let ajax = { opcion: opc, idCat: idcat, idFab: idfab, idMar: idmar, };
	//
	$.ajax({
		url: "matconvivencia/llenar_cmb_x_Cat-Fab-Mar.asp",
		type: "POST",
		dataType: 'json',
		data:  ajax,
		beforeSend: function(){
			$("#cargando").css("display", "block");
		}
	})
	.done (function(response, textStatus, jqXHR) {
		//console.log(response);
		//debugger;
		var len = response.data.length;
		$(cmb).empty();
		$(cmb).append("<option selected disabled value='0'>-- Seleccione --</option>");
		for( var i = 0; i < len; i++){
			var id = response.data[i]['id'];
			var nombre = response.data[i]['nombre'];
			$(cmb).append("<option value='"+id+"'>"+nombre+"</option>");
		}
	})
	.fail (function(jqXHR, textStatus, errorThrown) {
		swal("Algo salio mal.!",errorThrown, "error");
	})
	.always (function(jqXHROrData, textStatus, jqXHROrErrorThrown) {
		//alert("complete");
		$("#cargando").css("display", "none");
		$(cmb).prop("disabled", false);
	});
	//
}
//
function fillAllCombos4(opc, idcat, idfab, idmar, idseg, cmb) {
	//debugger;
	let ajax = { opcion: opc, idCat: idcat, idFab: idfab, idMar: idmar, idSeg: idseg, };
	//
	$.ajax({
		url: "matconvivencia/llenar_cmb_x_Cat-Fab-Mar-Seg.asp",
		type: "POST",
		dataType: 'json',
		data:  ajax,
		beforeSend: function(){
			$("#cargando").css("display", "block");
		}
	})
	.done (function(response, textStatus, jqXHR) {
		//console.log(response);
		//debugger;
		var len = response.data.length;
		$(cmb).empty();
		$(cmb).append("<option selected disabled value='0'>-- Seleccione --</option>");
		for( var i = 0; i < len; i++){
			var id = response.data[i]['id'];
			var nombre = response.data[i]['nombre'];
			$(cmb).append("<option value='"+id+"'>"+nombre+"</option>");
		}
	})
	.fail (function(jqXHR, textStatus, errorThrown) {
		swal("Algo salio mal.!",errorThrown, "error");
	})
	.always (function(jqXHROrData, textStatus, jqXHROrErrorThrown) {
		//alert("complete");
		$("#cargando").css("display", "none");
		$(cmb).prop("disabled", false);
	});
	//
}
//

// $('#cboArea').multiselect({
	// onChange: function(option, checked, select) {
		// alert('Changed option ' + $(option).val() + '.');
	// }
// });

