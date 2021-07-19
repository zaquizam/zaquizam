//
// funcionesV2.js // 12jul21 - 
//
function Reset() {
	$("#detallesMaestro").css("display", "none");
	$("#detalleTotalHogares").css("display", "none");
	$("#tablaResultados").css("display", "none");
	//
	$("#cboFabricante").find('option:not(:first)').remove();
	$("#cboFabricante").prop("selectedIndex", 0);
	$("#cboMarca").find('option:not(:first)').remove();
	$("#cboMarca").prop("selectedIndex", 0);
	$("#cboSegmento").find('option:not(:first)').remove();
	$("#cboSegmento").prop("selectedIndex", 0);
	$("#cboTamano").find('option:not(:first)').remove();
	$("#cboTamano").prop("selectedIndex", 0);
	//
	$("#cboIndicadores").find('option:not(:first)').remove();
	$("#cboIndicadores").prop("selectedIndex", 0);
	$("#cboSemanas").find('option:not(:first)').remove();
	$("#cboSemanas").prop("selectedIndex", 0);
	$("#cboArea").find('option:not(:first)').remove();
	$("#cboArea").prop("selectedIndex", 0);
	$("#cboZona").find('option:not(:first)').remove();
	$("#cboZona").prop("selectedIndex", 0);
	$("#cboProducto").find('option:not(:first)').remove();
	$("#cboProducto").prop("selectedIndex", 0);
	$("#cboCanal").find('option:not(:first)').remove();
	$("#cboCanal").prop("selectedIndex", 0);
	//
	$("#cboCategoria").prop("selectedIndex", 0);
	$("#cboCategoria").focus();
}
//
function clear() {
	$("#detallesMaestro").css("display", "none");
}
//
// function GenerarExcel() {
	// //alert("Generar Excel");
	// num = document.getElementById("Excel").value;
	// //alert("Generar Excel:="+ num);
	// window.open("PH_Cte_RetailScanningExcel.asp?" + num,"_blank");
// }

function ValidarCliente(){
	
	debugger;
	
	let ajax = { opcion: 12, idCli: sessionStorage.getItem("idCliente"), };
	//
	$.ajax({
		url: "RetSem_llenar_cmb.asp",
		type: "GET",
		dataType: 'html',
		data:  ajax,
		beforeSend: function(){
			$("#cargando").css("display", "block");
		}
	})
	.done (function(response, textStatus, jqXHR) {
		console.log(response);
		
		if(response=="True"){
			//$("#cboCategoria").prop("disabled", false);
			return true;			
		}else{
			$("#cboCategoria").empty();
			$("#cboCategoria").append("<option selected disabled value='0'>-- Seleccione --</option>");
			$("#cboCategoria").prop("disabled", true);
			swal("Atenas Grupo Consultor","Servicio No Contratado","info");
			return false;			
		}		
		
	})
	.fail (function(jqXHR, textStatus, errorThrown) {
		//alert("Error " + errorThrown);
		swal("Algo salio mal.!","LlenarCategoria()", "error");
	})
	.always (function(jqXHROrData, textStatus, jqXHROrErrorThrown) {
		$("#cargando").css("display", "none");		
	});
		
}	

//
function LlenarCategoria() {
	//
	//debugger;		
		
	$("#cboCategoria").prop("disabled", true);
	
	let ajax = { opcion: 1,	idCli: sessionStorage.getItem("idCliente"), };
	//
	$.ajax({
		url: "RetSem_llenar_cmb.asp",
		type: "GET",
		dataType: 'json',
		data:  ajax,
		beforeSend: function(){
			$("#cargando").css("display", "block");
		}
	})
	.done (function(response, textStatus, jqXHR) {
		console.log(response);
		//debugger;
		var len = response.data.length;
		//$("#cboCategoria_A").multiselect('destroy');
		$("#cboCategoria").empty();
		$("#cboCategoria").append("<option selected disabled value='0'>-- Seleccione --</option>");
		for( var i = 0; i < len; i++){
			var id = response.data[i]['id'];
			var nombre = response.data[i]['nombre'];
			$("#cboCategoria").append("<option value='"+id+"'>"+nombre+"</option>");
		}
		//		
	})
	.fail (function(jqXHR, textStatus, errorThrown) {
		//alert("Error " + errorThrown);
		swal("Algo salio mal.!","LlenarCategoria()", "error");
	})
	.always (function(jqXHROrData, textStatus, jqXHROrErrorThrown) {
		$("#cargando").css("display", "none");
		$("#cboCategoria").prop("disabled", false);
	});
	//
	
}
//
// <!-- CATEGORIA -->
//
$("#cboCategoria").on("change", function() {
	// Fill combo area + Zona + canal + Fabricante + Marca + Segmento + Tamaño + producto + indicadores + semanas
	//debugger;    
	let opcion   =0;
	let multiple =1;
	let combo    ="";
	let optCat   = $("#cboCategoria option:selected").val();
	let idCli    = sessionStorage.getItem("idCliente");	
	$("#Cat").val(optCat);
	//	
	opcion = 2;
	combo="#cboArea";
	fillAllCombos(opcion, optCat, combo, multiple, idCli);
	//
	opcion = 3;
	combo="#cboZona";
	fillAllCombos(opcion, optCat, combo, multiple, idCli);
	//
	opcion = 4;
	combo="#cboCanal";	
	fillAllCombos(opcion, optCat, combo, multiple, idCli);
	//
	opcion = 5;
	combo="#cboFabricante";	
	fillAllCombos(opcion, optCat, combo, multiple, idCli);
	//
	opcion = 6;
	combo="#cboMarca";
	fillAllCombos(opcion, optCat, combo, multiple, idCli);
	//	
	opcion = 7;
	combo="#cboSegmento";
	fillAllCombos(opcion, optCat, combo, multiple, idCli);
	//
	opcion = 8;
	combo="#cboTamano";
	fillAllCombos(opcion, optCat, combo, multiple, idCli);
	//
	opcion = 9;
	combo="#cboProducto";
	fillAllCombos(opcion, optCat, combo, multiple, idCli);
	//			
	opcion = 10;
	combo="#cboIndicadores";
	fillAllCombos(opcion, optCat, combo, multiple, idCli);
	//
	opcion = 11;
	combo="#cboSemanas";
	fillAllCombos(opcion, optCat, combo, multiple, idCli);
});
//
// <!-- FUNCIONES -->
//
function fillAllCombos(opc, idcat, cmb, mtp, idCli) {
	// debugger;
	$(cmb).prop("disabled", true);
	$("#cargando").css("display", "block");
	let ajax = { opcion: opc, idCat: idcat, idCli: idCli };
	//
	$.ajax({
		url: "RetSem_llenar_cmb.asp",
		type: "GET",
		dataType: 'json',
		data:  ajax,
		beforeSend: function(){
			$("#cargando").css("display", "block");
		}
	})
	.done (function(response, textStatus, jqXHR) {
		console.log(response);

		if (mtp == 0) {
			
			var len = response.data.length;
			$(cmb).empty();
			$(cmb).append("<option selected disabled value='0'>-- Seleccione --</option>");
			for( var i = 0; i < len; i++){
				var id = response.data[i]['id'];
				var nombre = response.data[i]['nombre'];
				$(cmb).append("<option value='"+id+"'>"+nombre+"</option>");
			}
			
		}else{
			
			$(cmb).multiselect('destroy');
			var len = response.data.length;
			$(cmb).empty();			
			for( var i = 0; i < len; i++){
				var id = response.data[i]['id'];
				var nombre = response.data[i]['nombre'];
				$(cmb).append("<option value='"+id+"'>"+nombre+"</option>");
			}

			$(cmb).multiselect({
  				nonSelectedText: '-- Seleccione --',
  				buttonWidth: '285px',
				includeSelectAllOption: true,
            	//enableFiltering: true
 			});
			
		}
	})
	.fail (function(jqXHR, textStatus, errorThrown) {
		swal("Algo salio mal.!",errorThrown, "error");
	})
	.always (function(jqXHROrData, textStatus, jqXHROrErrorThrown) {
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
		url: "mConv_llenar_cmb_x_Cat-Fab.asp",
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
		url: "mConv_llenar_cmb_x_Cat-Fab-Mar.asp",
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
		url: "mConv_llenar_cmb_x_Cat-Fab-Mar-Seg.asp",
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