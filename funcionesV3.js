//
// funcionesV3.js // 12jul21 - 15jul21
//
function Reset() {
	//
	$("#DivRetailScanningSem").css("display", "none");		
	$("#DivRetailScanningSem").html("");		
	//
	$("#cboArea").find('option:not(:first)').remove();
	$("#cboArea").prop("selectedIndex", 0);
	$("#cboZona").find('option:not(:first)').remove();
	$("#cboZona").prop("selectedIndex", 0);
	$("#cboCanal").find('option:not(:first)').remove();
	$("#cboCanal").prop("selectedIndex", 0);
	$("#cboFabricante").find('option:not(:first)').remove();
	$("#cboFabricante").prop("selectedIndex", 0);
	$("#cboMarca").find('option:not(:first)').remove();
	$("#cboMarca").prop("selectedIndex", 0);
	//
	$("#cboSegmento").find('option:not(:first)').remove();
	$("#cboSegmento").prop("selectedIndex", 0);
	$("#cboTamano").find('option:not(:first)').remove();
	$("#cboTamano").prop("selectedIndex", 0);
	$("#cboProducto").find('option:not(:first)').remove();
	$("#cboProducto").prop("selectedIndex", 0);
	$("#cboIndicadores").find('option:not(:first)').remove();
	$("#cboIndicadores").prop("selectedIndex", 0);
	$("#cboSemanas").find('option:not(:first)').remove();
	$("#cboSemanas").prop("selectedIndex", 0);	
	//
	$("#cboCategoria").prop("selectedIndex", 0);
	$("#cboCategoria").focus();
}

//
function ValidarCliente(){
	
	debugger;
	
	let ajax = { opcion: 12, idCli: sessionStorage.getItem("idCliente"), };
	//
	$.ajax({
		url: "RetSem_llenar_cmb1.asp",
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
			LlenarCategoria();
			return true;			
		}else{
			$("#cboCategoria").empty();			
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
// <!-- CATEGORIA -->
//
function LlenarCategoria() {
	//
	$("#cboCategoria").prop("disabled", true);	
	let ajax = { opcion: 1,	idCli: sessionStorage.getItem("idCliente"), };
	//
	$.ajax({
		url: "RetSem_llenar_cmb1.asp",
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
		$("#cboCategoria").multiselect('destroy');
		$("#cboCategoria").empty();
		$("#cboCategoria").append("<option selected disabled value='0'>-- Seleccione --</option>");
		for( var i = 0; i < len; i++){
			var id = response.data[i]['id'];
			var nombre = response.data[i]['nombre'];
			$("#cboCategoria").append("<option value='"+id+"'>"+nombre+"</option>");
		}
		$("#cboCategoria").multiselect({
  				nonSelectedText: '-- Seleccione --',
  				buttonWidth: '285px',
				enableFiltering: true,
				filterPlaceholder: 'Buscar...',
				includeFilterClearBtn: true,
				enableCaseInsensitiveFiltering: true,
				maxHeight: 300,				
 			});
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
$("#cboCategoria").on("change", function() {
	// Fill combo area + Zona + canal + Fabricante + Marca + Segmento + Tamaño + producto + indicadores + semanas
	//debugger;    
	let opcion   =0;
	let multiple =0;
	let combo    ="";
	let cambio	 =0;
	let optCat   = $("#cboCategoria option:selected").val();
	let idCli    = sessionStorage.getItem("idCliente");	
	$("#Cat").val(optCat);
	//	
	opcion = 2;
	combo="#cboArea";
	multiple =1;
	cambio	 =2;
	fillAllCombos(opcion, optCat, combo, multiple, idCli, cambio);
	//
	opcion = 3;
	combo="#cboZona";
	multiple =1;
	cambio	 =3;
	fillAllCombos(opcion, optCat, combo, multiple, idCli, cambio);
	//
	opcion = 4;
	combo="#cboCanal";	
	multiple =1;
	cambio	 =4;
	fillAllCombos(opcion, optCat, combo, multiple, idCli, cambio);
	//
	opcion = 5;
	combo="#cboFabricante";	
	multiple =1;
	cambio	 =5;
	fillAllCombos(opcion, optCat, combo, multiple, idCli, cambio);
	//
	opcion = 6;
	combo="#cboMarca";
	multiple =1;
	cambio	 =6;
	fillAllCombos(opcion, optCat, combo, multiple, idCli, cambio);
	//	
	opcion = 7;
	combo="#cboSegmento";
	multiple =1;
	cambio	 =7;
	fillAllCombos(opcion, optCat, combo, multiple, idCli, cambio);
	//
	opcion = 8;
	combo="#cboTamano";
	multiple =1;
	cambio	 =8;
	fillAllCombos(opcion, optCat, combo, multiple, idCli, cambio);
	//
	opcion = 9;
	combo="#cboProducto";
	multiple =0;
	cambio	 =0;
	fillAllCombos(opcion, optCat, combo, multiple, idCli, cambio);
	//			
	opcion = 10;
	combo="#cboIndicadores";
	multiple =0;
	cambio	 =0;
	fillAllCombos(opcion, optCat, combo, multiple, idCli, cambio);
	//
	opcion = 11;
	combo="#cboSemanas";
	multiple =0;
	cambio	 =0;
	fillAllCombos(opcion, optCat, combo, multiple, idCli, cambio);
});
//
// <!-- FUNCIONES -->
//
function fillAllCombos(opc, idcat, cbo, mtp, idCli, cmb) {
	
	//$(cbo).prop("disabled", true);
	//$(cbo).multiselect('disable');
	
	$("#cargando").css("display", "block");
	let ajax = { opcion: opc, idCat: idcat, idCli: idCli };
	//
	$.ajax({
		url: "RetSem_llenar_cmb1.asp",
		type: "GET",
		dataType: 'json',
		data:  ajax,
		beforeSend: function(){
			$("#cargando").css("display", "block");
		}
	})
	.done (function(response, textStatus, jqXHR) {
		//console.log(response);
		//debugger;
		if (mtp == 0) {			
						
			var len = response.data.length;
			$(cbo).multiselect('destroy');			
			$(cbo).empty();						
			//$(cbo).append("<option selected disabled value='0'>-- Seleccione --</option>");
			for( var i = 0; i < len; i++){
				var id = response.data[i]['id'];
				var nombre = response.data[i]['nombre'];				  
				$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
			}
			$(cbo).multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, });
			$(cbo).multiselect('refresh');
			
		}else{			
			$(cbo).multiselect('destroy');
			var len = response.data.length;
			$(cbo).empty();			
			for( var i = 0; i < len; i++){
				var id = response.data[i]['id'];
				var nombre = response.data[i]['nombre'];
				$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
			}
			$(cbo).multiselect({ 
				nonSelectedText: '-- Seleccione --',
				disableIfEmpty: true,
  				buttonWidth: '285px',
				buttonHeight: '30px',
				includeSelectAllOption: true,
				enableFiltering: true,
				filterPlaceholder: 'Buscar...',
				includeFilterClearBtn: true,
				enableCaseInsensitiveFiltering: true,
				maxHeight: 300,
				onDropdownHide: function(event) {        			
					//GetCambioCombo(cmb,jQuery(cbo).val());					
					GetCambioCombo(cmb);					
				}
			});
			$(cbo).multiselect('refresh');

		}
	})
	.fail (function(jqXHR, textStatus, errorThrown) {
		swal("Algo salio mal.!",errorThrown, "error");
	})
	.always (function(jqXHROrData, textStatus, jqXHROrErrorThrown) {
		$("#cargando").css("display", "none");		
	});
}
//

