<%@language=vbscript%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file="Conexion.asp"-->
<script type="text/javascript">
	//**Inicio Buscar Edad
	function buscar_tiempohogarpanel(){
		//debugger;
		//LR
		//alert("Llego Tiempo Hogar Panel:=" + document.getElementById("FechaRegistro").value);
		debugger;
		//let fecha = $("#FechaRegistro").val();
		let fecha = document.getElementById("FechaRegistro").value;
		//alert("Llego Edad1");
		  if (
			fecha == null ||
			fecha == "" ||
			fecha == undefined ||
			fecha.length == 0 ||
			!isNaN(fecha)
		  ) {
			$("#fechaErr").html(
			  '<span style="color:red;">Introduzca fecha valida!</span>'
			);
			//$("#FechaRegistro").focus();
			document.getElementById("FechaRegistro").focus();
			return false;
		  } else {
			if (!validarFormatoFechaHogar(fecha)) {
			  //$("#fechaErr").html(
			  document.getElementById("FechaRegistro").html(
				'<span style="color:red;">Introduzca fecha valida!</span>'
			  );
			  //$("#FechaRegistro").focus();
			  document.getElementById("FechaRegistro").focus();
			  return false;
			} else {
			  //debugger;
			  let hoy = new Date();
			  let fechaFormulario = new Date(fecha);
			  hoy.setHours(0, 0, 0, 0); // Lo iniciamos a 00:00 horas
			  if (hoy <= fechaFormulario) {
				//$("#fechaErr").html(
				document.getElementById("FechaRegistro").html(
				  '<span style="color:red;">Fecha ingreso posterior a Hoy!</span>'
				);
				return false;
			  } else {
				//$("#fechaErr").html("");
				//document.getElementById("FechaRegistro").html("");
			  }
			}
		  }
		  //alert("Llego Edad2");
		  // Si la fecha es correcta, calculamos la edad
		  let values = fecha.split("/");
		  let dia = values[0];
		  let mes = values[1];
		  let ano = values[2];
		  // tomamos los valores actuales
		  let fecha_hoy = new Date();
		  let ahora_ano = fecha_hoy.getYear();
		  let ahora_mes = fecha_hoy.getMonth() + 1;
		  let ahora_dia = fecha_hoy.getDate();
		  // realizamos el calculo
		  let edad = ahora_ano + 1900 - ano;
		  if (ahora_mes < mes) {
			edad--;
		  }
		  if (mes == ahora_mes && ahora_dia < dia) {
			edad--;
		  }
		  if (edad > 1900) {
			edad -= 1900;
		  }
		  // calculamos los meses
		  let meses = 0;
		  if (ahora_mes > mes && dia > ahora_dia) meses = ahora_mes - mes - 1;
		  else if (ahora_mes > mes) meses = ahora_mes - mes;
		  if (ahora_mes < mes && dia < ahora_dia) meses = 12 - (mes - ahora_mes);
		  else if (ahora_mes < mes) meses = 12 - (mes - ahora_mes + 1);
		  if (ahora_mes == mes && dia > ahora_dia) meses = 11;
		  // calculamos los dias
		  let dias = 0;
		  if (ahora_dia > dia) dias = ahora_dia - dia;
		  if (ahora_dia < dia) {
			ultimoDiaMes = new Date(ahora_ano, ahora_mes - 1, 0);
			dias = ultimoDiaMes.getDate() - (dia - ahora_dia);
		  }
		  let tiempo = edad + " años, " + meses + " meses y " + dias + " días";
		  //let tiempo = edad + " años ";
		  //edad + " años";
		  //$("#Edad").val(tiempo);
		  document.getElementById("TiempoHogarPanel").value = tiempo;

		  return; //edad + " años, " + meses + " meses y " + dias + " días";
	}	
	//**Fin Buscar Edad
	
	function validarFormatoFechaHogar(campo) {
	  let temp = campo.split("/");
	  //let fecha = temp[2] + "/" + temp[1] + "/" + temp[0];
	  let fecha = temp[0] + "/" + temp[1] + "/" + temp[2];
	  //alert("paso");
	  //alert(fecha);
	  let RegExPattern = /^\d{1,2}\/\d{1,2}\/\d{2,4}$/;
	  if (fecha.match(RegExPattern)) {
		return true;
	  } else {
		return false;
	  }
	}
	
</script>
<%
Session.LCID = 8202 
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	ynum=Request.QueryString("num")
	yOpc = "0"

	dim gDatosSol
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Fec_Registro "
	sql = sql & " FROM "
	sql = sql & " PH_PanelHogar "
	sql = sql & " WHERE "
	sql = sql & " Id_PanelHogar = " & ynum
	'response.write "<br>220 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	'response.write "<br> Linea 223 " &
	'response.end
	iExiste = 0
	if rsx1.eof then
		iExiste = 0
	else
		gDatosSol = rsx1.GetRows
		rsx1.close
		iExiste = 1
	end if
	Fecha=cstr(gDatosSol(0,0))
	%>
	<div id="DivClaseSocial"> 
		<input type="text" name="FechaRegistro" id="FechaRegistro" disabled value="<%=Fecha%>" align="right" size=8>
		<input type="text" name="TiempoHogarPanel" id="TiempoHogarPanel" disabled value="" align="right" size=30>
		<script>buscar_tiempohogarpanel()</script>
	</div> 
	<%
	
%>