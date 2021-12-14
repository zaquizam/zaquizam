//
// funciones Notificaciones Liberaciones
// 
function fillData(){
	Promise.all([findCliente()]).then(() => { // try removing ajax 1 or replacing with ajax2
		console.log('findCliente() done with success!');
		$('#cargando').css('display', 'none');
  	}).catch((response) => {
		console.log('All Ajax some failed!');
		$('#cargando').css('display', 'none');
		swal("Algo salio mal.!","findCliente()", "error");
  	});	
}
//
function findCliente() {
	//debugger;
    return $.ajax({      	
	  	url: sessionStorage.getItem("urlApi") + "getAllClienteNotificar",
      	type: 'get',
      	success: function(response) {
        	//console.log(response);
			//debugger;
			var len = response.data.length;
			$("#cboCliente").empty();			
			$("#cboCliente").append("<option selected disabled value='0'>-- Seleccione --</option>");
			for( var i = 0; i < len; i++){
				var id = response.data[i]['id'];
				var nombre = response.data[i]['nombre'];
				$("#cboCliente").append("<option value='" + id+ "'>" + nombre + "</option>");
			}
			//			 		
      },
      error: function(jqXHR, textStatus, errorThrown) {
	        console.log("Fallo findCliente()");
      }
    });
}
//
function findSemanas() {
	//debugger;
    return $.ajax({      	
	  	url: sessionStorage.getItem("urlApi") + "getAllSemanaNotificar",
      	type: 'get',
      	success: function(response) {
        	//console.log(response);
			//debugger;
			var len = response.data.length;
			$("#cboTiempo").multiselect('destroy');			
			$("#cboTiempo").empty();
			$("#cboTiempo").append("<option selected disabled value='0'>-- Seleccione --</option>");
			for( var i = 0; i < len; i++){
				var id = response.data[i]['id'];
				var nombre = response.data[i]['nombre'];
				$("#cboTiempo").append("<option value='" + id + "'>" + nombre + "</option>");
			}
			//
			$("#cboTiempo").multiselect({
				nonSelectedText: '-- Seleccione --',
				buttonWidth:'450px',
				includeSelectAllOption: true,
				dropRight: true,
				maxHeight: 200
				//enableFiltering: true
			});
      },
      error: function(jqXHR, textStatus, errorThrown) {
	        swal("Algo salio mal.!","findSemanas()", "error");
      }
    });
}
//
function findPeriodo() {
	//debugger;
    return $.ajax({      	
	  	url: sessionStorage.getItem("urlApi") + "getAllMesesNotificar",
      	type: 'get',
      	success: function(response) {
        	//console.log(response);
			//debugger;			
			$("#cboTiempo").multiselect('destroy');			
			$("#cboTiempo").empty();			
			var len = response.data.length;
			$("#cboTiempo").append("<option selected disabled value='0'>-- Seleccione --</option>");
			for( var i = 0; i < len; i++){
				var id = response.data[i]['id'];
				var nombre = response.data[i]['nombre'];
				$("#cboTiempo").append("<option value='" + id+ "'>" + nombre + "</option>");
			}
			//
			$("#cboTiempo").multiselect({
				nonSelectedText: '-- Seleccione --',
				buttonWidth:'450px',
				includeSelectAllOption: true,
				dropRight: true,
				maxHeight: 200
				//enableFiltering: true
			});			
      },
      error: function(jqXHR, textStatus, errorThrown) {
	        swal("Algo salio mal.!","findPeriodo()", "error");
      }
    });
}
//
$("#cboCliente").change(function(){
	//
	//debugger;
	event.preventDefault();		
	let idCli = $('#cboCliente').val();
	$("#cboCategoria").prop("disabled", true);
    return $.ajax({      	
	  	url: sessionStorage.getItem("urlApi")+"getAllCategoriaClienteNotificar/"+ idCli + '',
      	type: 'get',
      	success: function(response) {
        	//console.log(response);
			//debugger;
			var len = response.data.length;
			$("#cboCategoria").multiselect('destroy');			
			$("#cboCategoria").empty();
			//$("#cboCategoria").append("<option selected disabled value='0'>-- Seleccione --</option>");
			for( var i = 0; i < len; i++){
				var id = response.data[i]['id'];
				var nombre = response.data[i]['nombre'];
				$("#cboCategoria").append("<option value='" + id + "'>" + nombre + "</option>");
			}
			//			
			$("#cboCategoria").multiselect({
				nonSelectedText: '-- Seleccione --',
				includeSelectAllOption: true,
				enableFiltering: true,
				buttonWidth:'450px',				
				maxHeight: 350
			});
			 $('#cboCategoria').multiselect('disable');
      },
      error: function(jqXHR, textStatus, errorThrown) {
	        alert("Fallo findCliente()");
      }
    });	
});
//
$("#cboReporte").change(function(){
	//debugger;			
	if ($("#cboReporte").val()=='Semanal'){
		findSemanas();
		sessionStorage.setItem("semanal", 1);
		sessionStorage.setItem("mensual", 0);
		$("#lblTiempo").html("<i class='fas fa-calendar-week'></i>&nbsp;Semana:");        
    }else{
		findPeriodo();
		sessionStorage.setItem("semanal", 0);
		sessionStorage.setItem("mensual", 1);
        $("#lblTiempo").html("<i class='fas fa-calendar-alt'></i>&nbsp;Periodo:");
    }
});
//
$("#cboTreporte").change(function(){
    if($("#cboTreporte").val()=="Detallado"){        
		$('#cboCategoria').multiselect('enable');        
    }else{		
		$('#cboCategoria option:selected').each(function() {
            $(this).prop('selected', false);
        }); 
		$('#cboCategoria').multiselect('disable');	
        $('#cboCategoria').multiselect('refresh');        
    }
});
//
$("#btnReset").click(function(e) {
	e.preventDefault();	
	//
	sessionStorage.removeItem("semanal");
	sessionStorage.removeItem("mensual");
	//
	$("#cargando").css("display", "none");
	$("#cboCliente").prop("selectedIndex", 0);	
	$("#cboTreporte").prop("selectedIndex", 0);
	$("#cboReporte").prop("selectedIndex", 0);
	$("#cboTiempo").multiselect('destroy');			
	$("#cboTiempo").empty();			
	$("#cboTiempo").multiselect({
		nonSelectedText: '-- Seleccione --',
		buttonWidth:'450px',
		includeSelectAllOption: true,
		disableIfEmpty: true
		//enableFiltering: true
	});		
	$("#cboCategoria").multiselect('destroy');			
	$("#cboCategoria").empty();			
	$("#cboCategoria").multiselect({
		nonSelectedText: '-- Seleccione --',
		buttonWidth:'450px',
		includeSelectAllOption: true,
		disableIfEmpty: true
		//enableFiltering: true
	});				
	//	
});
//
$("#btnEnviar").click(function(e) {
	e.preventDefault();
	$('#envioEmail').css('display', 'block');
	debugger;
	let Cliente = $('#cboCliente').val();
	let tipoReporte = $('#cboTreporte').val();
	let repSemMen = $('#cboReporte').val();
	let semana  = $('#cboTiempo').val();	
	let Categoria = $('#cboCategoria').val();
	//
	if(Categoria.length=="0"){
		Categoria="0";
	}
	if (sessionStorage.getItem("mensual")=='0'){		
		repMes="0"		
	}else{
		repMes = semana;	
	}
	if (sessionStorage.getItem("semanal")=='0'){		
		repSemana="0"
	}else{
		repSemana = semana;
	}
	//
	//return false
	
	return $.ajax({
	  url: sessionStorage.getItem("urlApi") + 'SendNotificacionesEmailCategoriaHp/'+ Cliente +'/'+ tipoReporte +'/'+ repSemMen +'/'+ repSemana +'/'+ repMes +'/'+ Categoria +'',
      type: 'get',
      success: function(response) {
        //
		//Resultado del envio
		setTimeout(function () {
			$('#envioEmail').css('display', 'none');			
			console.log('Send Mail() done with success!');
			$("#btnReset").click();
		}, 6000);					
		
      },
      error: function(jqXHR, textStatus, errorThrown) {
        swal("Algo salio mal.!","Envio email()", "error");
      }
    });	
	//				
});

