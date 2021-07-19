//
// funcionesV3.js // 12jul21 - 15jul21
//
function Reset() {
	$("#detallesMaestro").css("display", "none");
	$("#detalleTotalHogares").css("display", "none");
	$("#tablaResultados").css("display", "none");
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
function clear() {
	$("#detallesMaestro").css("display", "none");
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
				//buttonContainer: '<div class="btn-group w-100" />',
            	//dropUp: true
				//includeSelectAllOption: true,
            	//enableFiltering: true
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
// Promise.all([totalHogares(cboPeriodo, cboArea), ejecutar_A(categoria_A,fabricante_A,marca_A,segmento_A,rangotam_A,cboArea,cboPeriodo), ejecutar_B(categoria_B,fabricante_B,marca_B,segmento_B,rangotam_B,cboArea,cboPeriodo)]).then(() => { // try removing ajax 1 or replacing with ajax2
    	// //alert('All Ajax done with success!');
		// console.log('All Ajax done with success!');
		// Totalizar();
  	// }).catch((response) => {
    	// alert('All Ajax done: some failed!');
		// console.log('All Ajax some failed!');
  	// })	
//
$("#cboCategoria").on("change", function() {
	// Fill combo area + Zona + canal + Fabricante + Marca + Segmento + Tamaño + producto + indicadores + semanas
	//debugger;    
	let opcion   =0;
	let multiple =0;
	let combo    ="";
	let optCat   = $("#cboCategoria option:selected").val();
	let idCli    = sessionStorage.getItem("idCliente");	
	$("#Cat").val(optCat);
	//	
	opcion = 2;
	combo="#cboArea";
	multiple =0;
	fillAllCombos(opcion, optCat, combo, multiple, idCli);
	//
	opcion = 3;
	combo="#cboZona";
	multiple =0;
	fillAllCombos(opcion, optCat, combo, multiple, idCli);
	//
	opcion = 4;
	combo="#cboCanal";	
	multiple =0;
	fillAllCombos(opcion, optCat, combo, multiple, idCli);
	//
	opcion = 5;
	combo="#cboFabricante";	
	multiple =1;
	fillAllCombos(opcion, optCat, combo, multiple, idCli);
	//
	opcion = 6;
	combo="#cboMarca";
	multiple =1;
	fillAllCombos(opcion, optCat, combo, multiple, idCli);
	//	
	opcion = 7;
	combo="#cboSegmento";
	multiple =1;
	fillAllCombos(opcion, optCat, combo, multiple, idCli);
	//
	opcion = 8;
	combo="#cboTamano";
	multiple =1;
	fillAllCombos(opcion, optCat, combo, multiple, idCli);
	//
	opcion = 9;
	combo="#cboProducto";
	multiple =0;
	fillAllCombos(opcion, optCat, combo, multiple, idCli);
	//			
	opcion = 10;
	combo="#cboIndicadores";
	multiple =0;
	fillAllCombos(opcion, optCat, combo, multiple, idCli);
	//
	opcion = 11;
	combo="#cboSemanas";
	multiple =0;
	fillAllCombos(opcion, optCat, combo, multiple, idCli);
});
//
// <!-- FUNCIONES -->
//
function fillAllCombos(opc, idcat, cmb, mtp, idCli) {
	// debugger;
	$(cmb).prop("disabled", true);
	$(cmb).multiselect('disable');
	
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
		if (mtp == 0) {			
			//var len = response.data.length;
			// $(cmb).empty();
			// $(cmb).append("<option selected disabled value='0'>-- Seleccione --</option>");
			// for( var i = 0; i < len; i++){
				// var id = response.data[i]['id'];
				// var nombre = response.data[i]['nombre'];
				// $(cmb).append("<option value='"+id+"'>"+nombre+"</option>");
			// }
			
			var len = response.data.length;
			$(cmb).multiselect('destroy');			
			$(cmb).empty();						
			//$(cmb).append("<option selected disabled value='0'>-- Seleccione --</option>");
			for( var i = 0; i < len; i++){
				var id = response.data[i]['id'];
				var nombre = response.data[i]['nombre'];
				$(cmb).append("<option value='"+id+"'>"+nombre+"</option>");
			}

			$(cmb).multiselect('refresh');
			
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
					FillCmbxFabricante(jQuery("#cboFabricante").val());					
				}
		});
			
			
			
			
			
			
			
			
			

			//$(cmb).multiselect('setOptions', options)
			//$(cmb).multiselect('refresh');
			
			// $(cmb).multiselect({
  				// nonSelectedText: '-- Seleccione --',
  				// buttonWidth: '285px',
				// buttonHeight: '30px',
				// includeSelectAllOption: true,
				// enableFiltering: true,
				// filterPlaceholder: 'Buscar...',
				// includeFilterClearBtn: true,
				// enableCaseInsensitiveFiltering: true,
				// maxHeight: 300,
			//enableFiltering: true
			/*
				onDropdownHide: function(event) {
        			//alert('Dropdown closed.');					
					FillCmbxFabricante(idVal)
					alert(jQuery(cmb).val());					
				}
			*/
		//});
		
		
		
			
		}
	})
	.fail (function(jqXHR, textStatus, errorThrown) {
		swal("Algo salio mal.!",errorThrown, "error");
	})
	.always (function(jqXHROrData, textStatus, jqXHROrErrorThrown) {
		$("#cargando").css("display", "none");
		$(cmb).prop("disabled", false);
		$(cmb).multiselect('enable');
	});
	//
}

function multiselect (obj) {// Initialization Method
	$(obj).multiselect({
        includeSelectAllOption: true,
        enableClickableOptGroups: true,
        enableCollapsibleOptGroups: true,
        buttonWidth: 195,
        maxHeight: 300,
    });
}
//
function FillCmbxFabricante(idVal){
	alert('valores'+ idVal);
}

/*
jQuery('#cboFabricante').multiselect({
    onDropdownHide: function(event) {
        //FillDropdown2(jQuery('#cboFabricante').val());
		alert(jQuery('#cboFabricante').val());
    }
});
*/

//
// $("#cboFabricante").on('focusout blur',function(){
    // //will return an array of the values for the selected options
    // var myselectedvalues = $(this).val();
	// let fab  = $("#cboFabricante :selected").map((_,e) => e.value).get();
// });
//
//
// $('#cboFabricante').blur(function() {    
	// debugger;
    // $('option:selected', $(this)).each(function() {
		
		// var myselectedvalues = $(this).val();
		// let fab  = $("#cboFabricante :selected").map((_,e) => e.value).get();
        // alert('Fabricante cerrado.');
      
    // });
  // });
  
// $(document).ready(function() {
	// $('#cboFabricante').multiselect({
		// nonSelectedText: '-- Seleccione --',
		// disableIfEmpty: true,
		// buttonWidth: '285px',
		// includeSelectAllOption: true, 	
		// onDropdownHidden: function(event) {
			// debugger;
			// var myselectedvalues = $(this).val();
			// let fab  = $("#cboFabricante :selected").map((_,e) => e.value).get();
			// alert('Fabricante cerrado.');
		// }
	// });
// });
//
// $(document).ready(function() {
	// $('#cboFabricante').multiselect({
		// enableFiltering: true,
		// onChange: function(option, checked) {
			// alert('onChange!');
		// },
		// onDropdownHide: function(event) {
			// alert('onDropdownHide!');
		// }
	// });
// });
// $(document).ready(function() {
    // $('#cboFabricante').multiselect({
        // onDropdownHide: function(event) {
            // alert('Dropdown closed.');
            // // to reload the page
            // location.reload();
        // }
    // });
// });
//
function getOptions(node, isFilter) {
  var isChanged = false;
  return {
    enableCaseInsensitiveFiltering: isFilter,
    includeSelectAllOption: true,
    filterPlaceholder: 'Search ...',
    nonSelectedText: node,
    numberDisplayed: 1,
    buttonWidth: '100%',
    maxHeight: 400,
    onChange: function() {
      alert('Changes');
      isChanged = true;
    },
    onSelectAll: function() {
      alert("SELECT ALL");
      isChanged = true;
    },
    onDropdownHide: function(event) {
      if (isChanged) {
        filterData(node);
        isChanged = false;
      }

    }
  }
}
//$('#myselect').multiselect(getOptions('myselect', true));
//

$("#cboFabricante2").on("change", function() {
	// Fill combo Marca + Segmento + Tamaño + producto + indicadores + semanas
	debugger;    
	let opcion   = 0;
	let multiple = 1;
	let combo    = "";
	let idCat    = $("#cboCategoria option:selected").val();
	let idFab    = $("#Fabricante :selected").map((_,e) => e.value).get();
	let idCli    = sessionStorage.getItem("idCliente");	
	$("#Cat").val(idCat);
	//	
	//opcion = 2;
	//combo="#cboArea";
	//fillAllCombos(opcion, optCat, combo, multiple, idCli);
	//
	//opcion = 3;
	//combo="#cboZona";
	//fillAllCombos(opcion, optCat, combo, multiple, idCli);
	//
	//opcion = 4;
	//combo="#cboCanal";	
	//fillAllCombos(opcion, optCat, combo, multiple, idCli);
	//
	//opcion = 5;
	//combo="#cboFabricante";	
	//fillAllCombos(opcion, optCat, combo, multiple, idCli);
	//
	opcion = 6;
	combo="#cboMarca";
	fillAllCombos2(opcion, combo, idCli, idCat, idFab );
	//	
	opcion = 7;
	combo="#cboSegmento";
	fillAllCombos2(opcion, combo, idCli, idCat, idFab );
	//
	opcion = 8;
	combo="#cboTamano";
	fillAllCombos2(opcion, combo, idCli, idCat, idFab );
	//
	opcion = 9;
	combo="#cboProducto";
	fillAllCombos2(opcion, combo, idCli, idCat, idFab );
	//			
	//opcion = 10;
	//combo="#cboIndicadores";
	//fillAllCombos2(opcion, combo, idCli, idCat, idFab );
	//
	//opcion = 11;
	//combo="#cboSemanas";
	//fillAllCombos2(opcion, combo, idCli, idCat, idFab );
});
function fillAllCombos2(opc, cmb, iCli, iCat, iFab ) {
	//debugger;
	let ajax = { opcion: opc, idCat: iCat, idFab: iFab, };
	//
	$.ajax({
		url: "RetSem_llenar_cmb2.asp",
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