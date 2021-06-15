// funcion procesar.js - 14jun21
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
	//
	// var settings ="";
	// settings = {
		// "url": "http://localhost:3000/api/getTotalHogaresPeriodo/"+cboPeriodo+"",	        
        // "method": "get"
    // }	
	//totalHogares(settings);
	totalHogares();
	// settings = {
		// "url": "http://localhost:3000/api/getTotalConvivencia/"+categoria_A+"/"+fabricante_A+"/"+marca_A+"/"+segmento_A+"/"+rangotam_A+"/"+cboArea+"/"+cboPeriodo+"",	        
        // "method": "get"
    // }
	//ejecutar(settings,1);
	ejecutar_A(categoria_A,fabricante_A,marca_A,segmento_A,rangotam_A,cboArea,cboPeriodo);
	// settings = {
		// "url": "http://localhost:3000/api/getTotalConvivencia/"+categoria_B+"/"+fabricante_B+"/"+marca_B+"/"+segmento_B+"/"+rangotam_B+"/"+cboArea+"/"+cboPeriodo+"",
        // "method": "get"
    // }
	// ejecutar(settings,2);
	ejecutar_B(categoria_B,fabricante_B,marca_B,segmento_B,rangotam_B,cboArea,cboPeriodo);
	//		
	Totalizar();
	//
	$("#detalleTotalHogares").css("display", "block");
	//	
});

function totalHogares(settings){
	//
	var settings = {
		"url": "http://localhost:3000/api/getTotalHogaresPeriodo/"+cboPeriodo+"",	        
        "method": "get"
    }	
	$.ajax(settings).done(function(response) {
        debugger;
		sessionStorage.setItem("totalHogares", parseInt(response.recordsTotal));			
		//			
    });
	//	
}

//function ejecutar_A(settings, opcion){
function ejecutar_A(categoria_A,fabricante_A,marca_A,segmento_A,rangotam_A,cboArea,cboPeriodo){
	//
	settings = {
		"url": "http://localhost:3000/api/getTotalConvivencia/"+categoria_A+"/"+fabricante_A+"/"+marca_A+"/"+segmento_A+"/"+rangotam_A+"/"+cboArea+"/"+cboPeriodo+"",	        
        "method": "get"
    }
	$.ajax(settings).done(function(response) {
        //debugger;
        var total = 0;
		$("#totalHogaresA").html(parseInt(response.recordsTotal).toLocaleString("es-ES", { minimumFractionDigits: 0 }));
		sessionStorage.setItem("totalHogaresA", parseInt(response.recordsTotal));			
		let arrayCatA = new Array(parseInt(response.recordsTotal)); 
		//
		for (let i = 0; i < response.data.length; i++) {
			arrayCatA[i] = response.data[i].id_hogar;				
		}
		sessionStorage.setItem("arrayCatA", arrayCatA);
    });

}
//
function ejecutar_B(categoria_B,fabricante_B,marca_B,segmento_B,rangotam_B,cboArea,cboPeriodo){
	//
	settings = {
		"url": "http://localhost:3000/api/getTotalConvivencia/"+categoria_B+"/"+fabricante_B+"/"+marca_B+"/"+segmento_B+"/"+rangotam_B+"/"+cboArea+"/"+cboPeriodo+"",	        
        "method": "get"
    }
	$.ajax(settings).done(function(response) {
        debugger;
        var total = 0;
		var totalHogaresB=0;
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
		sessionStorage.setItem("totalConviven", conviven);
		//			
		console.log("los hogares conviven son are: " + conviven);
		// alert("los hogares conviven son : " + conviven );
		// alert("total hogares conviven son: " + conviven.length );
		
    });

}
//
function Totalizar(){
		
		debugger;
		total=parseInt(sessionStorage.getItem("totalHogaresA")) + parseInt(totalHogaresB);
		$("#totalHogaresAB").html(parseInt(total).toLocaleString("es-ES", { minimumFractionDigits: 0 }));
		// Totalizando Matriz
		//
		total = ( parseInt(sessionStorage.getItem("totalHogaresA")) / parseInt(sessionStorage.getItem("totalHogaresA")) ) * 100;		
		total = (parseFloat(total).toLocaleString("es-ES", { maximumFractionDigits: 2, minimumFractionDigits: 2 }) );
		$("#total_AA").html(total+" %");
		total = ( parseInt(sessionStorage.getItem("totalConviven")) / parseInt(sessionStorage.getItem("totalHogaresB")) ) * 100;
		total = (parseFloat(total).toLocaleString("es-ES", { maximumFractionDigits: 2, minimumFractionDigits: 2 }) );
		$("#total_AB").html(total+" %");
		//
		total = ( parseInt(sessionStorage.getItem("totalConviven")) / parseInt(sessionStorage.getItem("totalHogaresA")) ) * 100 ;
		total = ( parseFloat(total).toLocaleString("es-ES", { maximumFractionDigits: 2, minimumFractionDigits: 2 }) ) ;
		$("#total_BA").html(total+" %");
		total = ( parseInt(sessionStorage.getItem("totalHogaresB")) / parseInt(sessionStorage.getItem("totalHogaresB")) ) * 100 ;
		total = ( parseFloat(total).toLocaleString("es-ES", { maximumFractionDigits: 2, minimumFractionDigits: 2 }) );
		$("#total_BB").html(total+" %");	
	
}

function ejecutar(categoria_A,fabricante_A,marca_A,segmento_A,rangotam_A,cboArea,cboPeriodo){
	//
	settings = {
		"url": "http://localhost:3000/api/getTotalConvivencia/"+categoria_A+"/"+fabricante_A+"/"+marca_A+"/"+segmento_A+"/"+rangotam_A+"/"+cboArea+"/"+cboPeriodo+"",	        
        "method": "get"
    }
	$.ajax(settings).done(function(response) {
        //debugger;
        var total = 0;
		var totalHogaresB=0;
		//
		if(opcion==1){
			$("#totalHogaresA").html(parseInt(response.recordsTotal).toLocaleString("es-ES", { minimumFractionDigits: 0 }));
			sessionStorage.setItem("totalHogaresA", parseInt(response.recordsTotal));			
			let arrayCatA = new Array(parseInt(response.recordsTotal)); 
			//
			for (let i = 0; i < response.data.length; i++) {
				arrayCatA[i] = response.data[i].id_hogar;				
			}
			sessionStorage.setItem("arrayCatA", arrayCatA);
		}else if (opcion==2){		
			//debugger;
			totalHogaresB = response.recordsTotal;
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
			sessionStorage.setItem("totalConviven", conviven);
			//			
			console.log("los hogares conviven son are: " + conviven);
			// alert("los hogares conviven son : " + conviven );
			// alert("total hogares conviven son: " + conviven.length );
									
		}		
		debugger;
		total=parseInt(sessionStorage.getItem("totalHogaresA")) + parseInt(totalHogaresB);
		$("#totalHogaresAB").html(parseInt(total).toLocaleString("es-ES", { minimumFractionDigits: 0 }));
		// Totalizando Matriz
		//
		total = ( parseInt(sessionStorage.getItem("totalHogaresA")) / parseInt(sessionStorage.getItem("totalHogaresA")) ) * 100;		
		total = (parseFloat(total).toLocaleString("es-ES", { minimumFractionDigits: 2 }) );
		$("#total_AA").html(total+" %");
		total = ( parseInt(sessionStorage.getItem("totalConviven")) / parseInt(sessionStorage.getItem("totalHogaresB")) ) * 100;
		total = (parseFloat(total).toLocaleString("es-ES", { minimumFractionDigits: 2 }) );
		$("#total_AB").html(total+" %");
		//
		total = ( parseInt(sessionStorage.getItem("totalConviven")) / parseInt(sessionStorage.getItem("totalHogaresA")) ) * 100 ;
		total = ( parseFloat(total).toLocaleString("es-ES", { minimumFractionDigits: 2 }) ) ;
		$("#total_BA").html(total+" %");
		total = ( parseInt(sessionStorage.getItem("totalHogaresB")) / parseInt(sessionStorage.getItem("totalHogaresB")) ) * 100 ;
		total = ( parseFloat(total).toLocaleString("es-ES", { minimumFractionDigits: 2 }) );
		$("#total_BB").html(total+" %");
		
    });

}
//
function CrearJson(){
	var obj = new Object();
	obj.aa = (parseInt(sessionStorage.getItem("totalHogaresA"))/parseInt(sessionStorage.getItem("totalHogaresA")))*100 ;
	obj.ab = (parseInt(sessionStorage.getItem("totalConviven"))/parseInt(sessionStorage.getItem("totalHogaresB")))*100 ;
	//convert object to json string
	var string = JSON.stringify(obj);
	//convert string to Json Object
	console.log(JSON.parse(string)); // this is your requirement.	
}