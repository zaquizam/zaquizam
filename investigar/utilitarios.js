//
// Utilitarios.js // 17ene21 -
//
$("#next").click(function() {
    var esUltimoElementoSeleccionado = $('#cboIdConsumo > option:selected').index() == $('#cboIdConsumo > option').length -1;
    if (!esUltimoElementoSeleccionado) {     
        $('#cboIdConsumo > option:selected').removeAttr('selected').next('option').attr('selected', 'selected');
		$("#cboIdConsumo").change();
    } else {
        $('#cboIdConsumo > option:selected').removeAttr('selected');
        $('#cboIdConsumo > option').first().attr('selected', 'selected');
		Reset();
    }   
});

$("#prev").click(function() {
    var esPrimerElementoSeleccionado = $('#cboIdConsumo > option:selected').index() == 0;
    if (!esPrimerElementoSeleccionado) {
       	$('#cboIdConsumo > option:selected').removeAttr('selected').prev('option').attr('selected', 'selected');
		$("#cboIdConsumo").change();
    } else {
       	$('#cboIdConsumo > option:selected').removeAttr('selected');
       	$('#cboIdConsumo > option').last().attr('selected', 'selected'); 
		Reset();
	}
});
//
function formatMonto(amount, decimalCount = 2, decimal = ",", thousands = ".") {
	debugger;	
	try {
		decimalCount = Math.abs(decimalCount);
		decimalCount = isNaN(decimalCount) ? 2 : decimalCount;

		const negativeSign = amount < 0 ? "-" : "";

		let i = parseInt(amount = Math.abs(Number(amount) || 0).toFixed(decimalCount)).toString();
		let j = (i.length > 3) ? i.length % 3 : 0;
		
		var monto = negativeSign + (j ? i.substr(0, j) + thousands : '') + i.substr(j).replace(/(\d{3})(?=\d)/g, "$1" + thousands) + (decimalCount ? decimal + Math.abs(amount - i).toFixed(decimalCount).slice(2) : "");

		$("#totalFactura").val(monto);
		
		// return negativeSign + (j ? i.substr(0, j) + thousands : '') + i.substr(j).replace(/(\d{3})(?=\d)/g, "$1" + thousands) + (decimalCount ? decimal + Math.abs(amount - i).toFixed(decimalCount).slice(2) : "");
		
		return true;
		
	} catch (e) {
		console.log(e)
	}	
};
//