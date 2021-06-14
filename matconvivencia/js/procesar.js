// funcion procesar.js - 14jun21
//
$("#BtnValidarProceso").click(function() {
	event.preventDefault();
	debugger;
	//
	var cboCategoria_A = $("#cboCategoria_A :selected").map((_,e) => e.value).get();
	if (cboCategoria_A == null || cboCategoria_A == undefined || cboCategoria_A.length== 0 || cboCategoria_A == 0) {
		swal("Seleccione una Categoria ","Equipo A","error");
		return false;
	};
	var cboCategoria_B = $("#cboCategoria_B :selected").map((_,e) => e.value).get();
	if (cboCategoria_B == null || cboCategoria_B == undefined || cboCategoria_B.length== 0 || cboCategoria_B == 0) {
		swal("Seleccione una Categoria ","Equipo B","error");
		return false;
	};
	//
	var cboArea    = $("#cboArea :selected").map((_,e) => e.value).get();
	if (cboArea == null || cboArea == undefined || cboArea.length== 0) {
		swal("Seleccione un Area","","error");
		return false;
	};
	var cboPeriodo = $("#cboPeriodo :selected").map((_,e) => e.value).get();
	//
	if (cboPeriodo == null || cboPeriodo == undefined || cboPeriodo.length== 0) {
		swal("Seleccione un Periodo","","error");
		return false;
	}
	//
	var categoria_A = $("#cboCategoria_A").val();
	var fabricante_A = ($("#cboFabricante_A").val() == null) ?  0 : $("#cboFabricante_A").val();
	var marca_A = ($("#cboMarca_A").val() == null) ?  0 : $("#cboMarca_A").val();
	var segmento_A = ($("#cboSegmento_A").val() == null) ?  0 : $("#cboSegmento_A").val();
	var rangotam_A = ($("#cboRangTamanoA").val() == null) ?  0 : $("#cboRangTamanoA").val();
	//
	var categoria_B = $("#cboCategoria_B").val();
	var fabricante_B = ($("#cboFabricante_B").val() == null) ?  0 : $("#cboFabricante_B").val();
	var marca_B = ($("#cboMarca_B").val() == null) ?  0 : $("#cboMarca_B").val();
	var segmento_B = ($("#cboSegmento_B").val() == null) ?  0 : $("#cboSegmento_B").val();
	var rangotam_B = ($("#cboRangTamanoB").val() == null) ?  0 : $("#cboRangTamanoB").val();
	//
	var formData = $('#frmConvivencia').serialize();
	formData = formData + "&" + 'cboArea=' + cboArea + "&" + 'cboPeriodo=' + cboPeriodo;
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


});

