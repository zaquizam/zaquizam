// funcion procesar.js - 16jun21
//
$("#BtnValidarProceso").click(function() {
	event.preventDefault();	
	//	
	var arrayCatA = Array[0]; 
	var arrayCatB = Array[0]; 
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
	var fabricante_A = ($("#cboFabricante_A").val() == null) ? 0 : $("#cboFabricante_A").val();
	var marca_A = ($("#cboMarca_A").val() == null) ? 0 : $("#cboMarca_A").val();
	var segmento_A = ($("#cboSegmento_A").val() == null) ? 0 : $("#cboSegmento_A").val();
	var rangotam_A = ($("#cboRangTamanoA").val() == null) ? 0 : $("#cboRangTamanoA").val();
	//
	var categoria_B = $("#cboCategoria_B").val();
	var fabricante_B = ($("#cboFabricante_B").val() == null) ? 0 : $("#cboFabricante_B").val();
	var marca_B = ($("#cboMarca_B").val() == null) ? 0 : $("#cboMarca_B").val();
	var segmento_B = ($("#cboSegmento_B").val() == null) ? 0 : $("#cboSegmento_B").val();
	var rangotam_B = ($("#cboRangTamanoB").val() == null) ? 0 : $("#cboRangTamanoB").val();
	//debugger;
		
	Promise.all([totalHogares(cboPeriodo), ejecutar_A(categoria_A,fabricante_A,marca_A,segmento_A,rangotam_A,cboArea,cboPeriodo), ejecutar_B(categoria_B,fabricante_B,marca_B,segmento_B,rangotam_B,cboArea,cboPeriodo)]).then(() => { // try removing ajax 1 or replacing with ajax2
    	//alert('All Ajax done with success!');
		console.log('All Ajax done with success!');
		Totalizar();
  	}).catch((response) => {
    	alert('All Ajax done: some failed!');
		console.log('All Ajax some failed!');
  	})	
	//	
});
//
function totalHogares(cboPeriodo) {
	
    return $.ajax({
      //url: "http://localhost:3000/api/getTotalHogaresPeriodo/"+cboPeriodo+"",
	  url: sessionStorage.getItem("urlApi")+"getTotalHogaresPeriodo/"+cboPeriodo+"",
      type: 'get',
      success: function(response) {
        console.log("totalHogares ok");
		sessionStorage.setItem("totalHogares", parseInt(response.recordsTotal));			
      },
      error: function(jqXHR, textStatus, errorThrown) {
        alert("Fallo () TotalHogares");
      }
    });
}
//
function ejecutar_A(categoria_A,fabricante_A,marca_A,segmento_A,rangotam_A,cboArea,cboPeriodo){
	//	
    return $.ajax({
	  url: sessionStorage.getItem("urlApi")+"getTotalConvivencia/"+categoria_A+"/"+fabricante_A+"/"+marca_A+"/"+segmento_A+"/"+rangotam_A+"/"+cboArea+"/"+cboPeriodo+"",
      type: 'get',
      success: function(response) {
        console.log("totalHogares A ok");
		$("#totalHogaresA").html(parseInt(response.recordsTotal).toLocaleString("es-ES", { minimumFractionDigits: 0 }));
		sessionStorage.setItem("totalHogaresA", parseInt(response.recordsTotal));			
		let arrayCatA = new Array(parseInt(response.recordsTotal)); 
		//
		for (let i = 0; i < response.data.length; i++) {
			arrayCatA[i] = response.data[i].id_hogar;				
		}
		sessionStorage.setItem("arrayCatA", arrayCatA);
      },
      error: function(jqXHR, textStatus, errorThrown) {
        alert("Fallo () Run-A");
      }
    });
}
//
function ejecutar_B(categoria_B,fabricante_B,marca_B,segmento_B,rangotam_B,cboArea,cboPeriodo){
	//
	return $.ajax({
	  url: sessionStorage.getItem("urlApi")+"getTotalConvivencia/"+categoria_B+"/"+fabricante_B+"/"+marca_B+"/"+segmento_B+"/"+rangotam_B+"/"+cboArea+"/"+cboPeriodo+"",	  
      type: 'get',
      success: function(response) {        
		//
		debugger;
		console.log("totalHogares B ok");
		totalHogaresB = response.recordsTotal;
		sessionStorage.setItem("totalHogaresB", parseInt(response.recordsTotal));			
		$("#totalHogaresB").html(parseInt(totalHogaresB).toLocaleString("es-ES", { minimumFractionDigits: 0 }));
		//
		let arrayCatB = new Array(parseInt(response.recordsTotal)); 			
		for (let i = 0; i < response.data.length; i++) {
			arrayCatB[i] = response.data[i].id_hogar;				
		}
		sessionStorage.setItem("arrayCatB", arrayCatB);
		//						
		var categ_A  = sessionStorage.getItem("arrayCatA").split(',') ;
		var categ_B  = sessionStorage.getItem("arrayCatB").split(',') ;
		//
		var conviven = categ_A.filter(x => categ_B.indexOf(x) !== -1);
		sessionStorage.setItem("totalConviven", conviven.length);
		console.log("los hogares conviven son are: " + conviven);
		console.log("total hogares conviven son: " + conviven.length );		
		//
		//let arrayDifference = arr1.filter(x => arr2.indexOf(x) === -1);
		let exclusivo_categ_A = categ_A.filter(x => !categ_B.includes(x));
		let exclusivo_categ_B = categ_B.filter(x => !categ_A.includes(x));
		console.log("Hogares Exclusivos A: " + exclusivo_categ_A);
		console.log("total Hogares Exclusivos A: " + exclusivo_categ_A.length );
		sessionStorage.setItem("totalExcl_A", exclusivo_categ_A.length);		
		console.log("Hogares Exclusivos B: " + exclusivo_categ_B);
		console.log("total Hogares Exclusivos B: " + exclusivo_categ_B.length );		
		sessionStorage.setItem("totalExcl_B", exclusivo_categ_B.length);
      },
      error: function(jqXHR, textStatus, errorThrown) {
        alert("Fallo () Run-B");
      }
    });
	
}
//
function Totalizar(){
	debugger;
	// Totalizando Hogares	
	total=parseInt(sessionStorage.getItem("totalHogaresA")) + parseInt(totalHogaresB);
	$("#totalHogaresAB").html(parseInt(total).toLocaleString("es-ES", { minimumFractionDigits: 0 }));
	console.log("Totalizando Hogares ok");
	//
	total = ( parseInt(sessionStorage.getItem("totalHogaresA")) / parseInt(sessionStorage.getItem("totalHogaresA")) ) * 100;		
	total = (parseFloat(total).toLocaleString("es-ES", { maximumFractionDigits: 2, minimumFractionDigits: 2 }) );
	$("#total_AA").html(total+" %");
	total = ( parseInt(sessionStorage.getItem("totalConviven")) / parseInt(sessionStorage.getItem("totalHogaresB")) ) * 100;
	total = (parseFloat(total).toLocaleString("es-ES", { maximumFractionDigits: 2, minimumFractionDigits: 2 }) );
	$("#total_AB").html(total+" %");
	console.log("Totalizando Hogares A ok");
	//
	total = ( parseInt(sessionStorage.getItem("totalConviven")) / parseInt(sessionStorage.getItem("totalHogaresA")) ) * 100 ;
	total = ( parseFloat(total).toLocaleString("es-ES", { maximumFractionDigits: 2, minimumFractionDigits: 2 }) ) ;
	$("#total_BA").html(total+" %");
	total = ( parseInt(sessionStorage.getItem("totalHogaresB")) / parseInt(sessionStorage.getItem("totalHogaresB")) ) * 100 ;
	total = ( parseFloat(total).toLocaleString("es-ES", { maximumFractionDigits: 2, minimumFractionDigits: 2 }) );
	$("#total_BB").html(total+" %");
	console.log("Totalizando Hogares B ok");
	//Penetracion		
	var total =0;
	total = (( parseInt(sessionStorage.getItem("totalHogaresA")) + parseInt(sessionStorage.getItem("totalHogaresB")) ) / parseInt(sessionStorage.getItem("totalHogares")))  * 100;	
	$("#penetracion_AB").html(parseFloat(total).toLocaleString("es-ES", { maximumFractionDigits: 2, minimumFractionDigits: 2 }) + " %");
	//Penetracion A
	total = ( parseInt(sessionStorage.getItem("totalHogaresA")) / parseInt(sessionStorage.getItem("totalHogares")) ) * 100;	
	$("#penetracion_A").html(parseFloat(total).toLocaleString("es-ES", { maximumFractionDigits: 2, minimumFractionDigits: 2 }) + " %");
	//Penetracion B
	total = ( parseInt(sessionStorage.getItem("totalHogaresB")) / parseInt(sessionStorage.getItem("totalHogares")) ) * 100;	
	$("#penetracion_B").html(parseFloat(total).toLocaleString("es-ES", { maximumFractionDigits: 2, minimumFractionDigits: 2 }) + " %");	
	//Convivencia
	$("#totalConvivencia").html(parseInt(sessionStorage.getItem("totalConviven")).toLocaleString("es-ES", { minimumFractionDigits: 0 }));
	//exclusividad
	total = ( parseInt(sessionStorage.getItem("totalExcl_A")) / parseInt(sessionStorage.getItem("totalHogares")) ) * 100;	
	$("#exclusivo_A").html(parseFloat(total).toLocaleString("es-ES", { maximumFractionDigits: 2, minimumFractionDigits: 2 }));	
	total = ( parseInt(sessionStorage.getItem("totalExcl_B")) / parseInt(sessionStorage.getItem("totalHogares")) ) * 100;	
	$("#exclusivo_B").html(parseFloat(total).toLocaleString("es-ES", { maximumFractionDigits: 2, minimumFractionDigits: 2 }));		
	//
	//Show
	$("#detalleTotalHogares").css("display", "block");
	$("#tablaResultados").css("display", "block");
	//Alerta
	if( parseInt(sessionStorage.getItem("totalHogaresA")) <= 30 ){
		swal("Total Hogares A","No son validos para muestra","info");
	}
	if(parseInt(sessionStorage.getItem("totalHogaresB"))<=30){
		swal("Total Hogares B","No son validos para muestra","error");
	}
		
}

