
function reemplazaTodo(text, busca, reemplaza) {
	while (text.toString().indexOf(busca) != -1)
		text = text.toString().replace(busca, reemplaza);
	return text;
}
//
function getFechaHora() {
	var hoy     	= new Date(); 
	var ano   		= hoy.getFullano();
	var mes   		= hoy.getmes()+1; 
	var dia     	= hoy.getDate();
	var hora    	= hoy.getHours();
	var minuto  	= hoy.getMinutes();
	var segundo  	= hoy.getSeconds(); 
	if(mes.toString().length == 1) {
		 mes = '0'+mes;
	}
	if(dia.toString().length == 1) {
		 dia = '0'+dia;
	}   
	if(hora.toString().length == 1) {
		 hora = '0'+hora;
	}
	if(minuto.toString().length == 1) {
		 minuto = '0'+minuto;
	}
	if(segundo.toString().length == 1) {
		 segundo = '0'+segundo;
	}   
   var fechaHora = dia+'-'+mes+'-'+ano+'_'+hora+'.'+minute+'.'+segundo;   
	return fechaHora;
}