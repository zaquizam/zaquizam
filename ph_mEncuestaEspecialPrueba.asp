<!DOCTYPE HTML>
<html >
<head>
	<title>Prueba Encuesta</title>
	<meta charset="utf-8">
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />	
	<meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
	<link href="css/sweetalert.css" rel="stylesheet" type="text/css" />	
</head>
<body topmargin="0">
<!--#include file="estiloscss.asp"-->
<!--#include file="encabezado.asp"-->
<!--#include file="nn_subN.asp"-->
<!--#include file="in_DataEN.asp"-->
<%
' 09dic20
'==========================================================================================
' Variables y Constantes
'==========================================================================================
    Apertura
  
'==========================================================================================
' Parámetros del Manteniemiento
'==========================================================================================
    'LeePar
	sPar=""      
    if ed_iPas<>4 then 
        Encabezado
    end if    
	' 	
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	
	dim arrEncuesta
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	QrySql = vbnullstring
    QrySql = QrySql & " SELECT"
    QrySql = QrySql & " PH_EncuestaEspecial.Id_EncuestaEspecial AS id,"
    QrySql = QrySql & " PH_EncuestaEspecial.EncuestaEspecial AS nombre"
    QrySql = QrySql & " FROM"
    QrySql = QrySql & " PH_EncuestaEspecial"
    QrySql = QrySql & " WHERE"
    QrySql = QrySql & " PH_EncuestaEspecial.Ind_Activo = 1"
    QrySql = QrySql & " ORDER BY"
    QrySql = QrySql & " PH_EncuestaEspecial.Id_EncuestaEspecial DESC"
	'
    rsx1.Open QrySql, conexion
	'	
	iExiste = 0
	if rsx1.eof then
		iExiste = 0
	else
		arrEncuesta = rsx1.GetRows()
		rsx1.close
		iExiste = 1
	end if
	
%>		
		<div class="container-fluid">  
		
			<br>									
			<div class="col-sm-1">				
				<div class="form-group">
					<label for="usr">Hogar:</label>
					
					<select class="form-control" title="Seleccionar Hogar" name="cboHogar" id="cboHogar"/>
						<option value="0" selected>...</option>
						<option value="313">Alexander</option>
						<option value="197">Alcira</option>
						<option value="1366">Cristina</option>
						<option value="3227">Pruebas</option>
					 </select>					
				</div>				
			</div>
			
			<div class="col-sm-4">
				<div class="form-group">				
					<label>Seleccione la Encuesta a Procesar:</label><span id="loader"></span>	
					<select class="form-control" title="Seleccionar Encuesta" name="cboEncuestas" id="cboEncuestas" onchange="checkEncuesta();" />
						<option value="0" select>-- Seleccionar --</option> 
						<%							
						for iReg = 0 to ubound(arrEncuesta,2)								
							Response.write "<option value=" &  arrEncuesta(0,iReg) &">" & arrEncuesta(1,iReg) & "</option>"
						next
						%>
					</select>
				</div>
			</div>
			
			<div class="col-sm-2">
				<div class="form-group">
					<label>Enviar al Movil:</label>
					<button id="generar" type="submit" title="Enviar encuesta al movil" class="btn btn-block btn-sm btn-success" onclick="genEncuesta();"><i class="fas fa-check fa-2x"></i></button>
				</div>
			</div>
						
			<div class="col-sm-2">
				<div class="form-group">
					<label>Resultados Encuesta:</label>
					<button id="mostrar" type="submit" title="Mostrar Resultados Encuesta" class="btn btn-block btn-sm btn-primary" onclick="showResultados();"><i class="fas fa-eye fa-2x"></i></button>
				</div>
			</div>
			
			<div class="col-sm-2">
				<div class="form-group">
					<label>Borrar la Encuesta:</label>
					<button id="borrar" type="submit" class="btn btn-block btn-sm btn-danger" onclick="borrarEncuesta();"><i class="fas fa-times fa-2x"></i></button>
				</div>
			</div>
			
			<div class="col-sm-1">				
				<div class="form-group">
					<label for="usr">Reset</label>
					<button id="borrar" type="submit" class="btn btn-block btn-sm btn-info" onclick="Reset();"><i class="fas fa-recycle fa-2x"></i></button>
				</div>				
			</div>
										
		</div>        
		<hr>
		<div class="container-fluid">  				
			<div class="table-responsive" id="tabla-resultados">
				<!-- // ** // -->
				<!-- Matriz de Datos Resultados -->
				<!-- // ** // -->
				...				 
			</div>				
		</div>        
	 
<%
conexion.close
%>
</body>
</html>
<script src="https://kit.fontawesome.com/9d7cfbccc5.js" crossorigin="anonymous"></script>
<script src="js/sweetalert.min.js"></script>

<script>

	$('#generar').attr('disabled','disabled');
	$('#mostrar').attr('disabled','disabled');
	$('#borrar').attr('disabled','disabled');
	
	function checkEncuesta() {
		//		
		var idHome=$("#cboHogar").val(); //$("#idHogar").val();
		if (idHome == null || idHome == 0) {
			swal("Aviso..!", "Seleccione un Hogar", "error");
			$("#cboHogar").focus();
			return false;
		}
		var idEncuesta = $("#cboEncuestas").val();
		//
		if (idEncuesta == null || idEncuesta == 0) {			
			swal("Aviso..!", "Seleccione una Encuesta", "error");
			$("#cboEncuestas").focus();
			Reset();
			return false;
		}else{
			$('#generar').removeAttr('disabled');
			$('#mostrar').removeAttr('disabled');
			$('#borrar').removeAttr('disabled');			
		}		
	}
			
	function genEncuesta() {
		//debugger;
		var idHome=$("#cboHogar").val(); //$("#idHogar").val();
		if (idHome == null || idHome == 0) {
			swal("Aviso..!", "Seleccione un Hogar", "error");
			$("#cboHogar").focus();
			return false;
		}
		var idEncuesta = $("#cboEncuestas").val();
		var title = $("#cboEncuestas option:selected").text();
		//
		if (idEncuesta == null || idEncuesta == 0) {
			swal("Aviso..!", "Seleccione una Encuesta", "error");
			$("#cboEncuestas").focus();
			return false;
		}
		//
		$.ajax({				
			url:"g_check_encuestas.asp?idhogar=" + idHome + "&idencuesta=" + idEncuesta,
			cache: false,		
			beforeSend: function(objeto){
				$('#loader').html('<img src="./images/ajax-small.gif"> Espere generando Resultados...!');
			},
			success:function(data){				
				//debugger;
				console.log(data);
				$('#loader').html('');
				if (data=="True"){					
					var msg ="Encuesta:\n" + title + "\nenviada al Movil!"; 
					swal("Aviso..!",  msg , "success");
					Reset();
				}else{					
					var msg ="Encuesta:\n" + title + "\nya fue enviada al Movil con Anterioridad!"; 
					swal("Aviso..!",  msg , "error");
					Reset();
				}					
								
			}
		})			
		
	}
	
	function showResultados() {
		debugger;
		var idHome=$("#cboHogar").val(); //$("#idHogar").val();
		if (idHome == null || idHome == 0) {
			swal("Aviso..!", "Seleccione un Hogar", "error");
			$("#cboHogar").focus();
			return false;
		}
		var idEncuesta = $("#cboEncuestas").val();		
		if (idEncuesta == null || idEncuesta == 0) {
			swal("Aviso..!", "Seleccione una Encuesta", "error");
			$("#cboEncuestas").focus();
			return false;
		}
		//
		$.ajax({
			url:"g_MostrarResultados_Encuesta.asp?idhogar=" + idHome + "&idencuesta=" + idEncuesta,				
			cache: false,		
			beforeSend: function(objeto){
				$('#loader').html('<img src="./images/ajax-small.gif"> Espere generando Resultados...!');
			},
			success:function(data){				
				debugger;
				console.log(data);
				$('#loader').html('');
				$("#tabla-resultados").html(data).fadeIn("slow");
				$('#generar').attr('disabled','disabled');
				$('#mostrar').attr('disabled','disabled');
				$('#borrar').attr('disabled','disabled');				
				//Reset();
			}
		})			
		
	}		
	
	function borrarEncuesta() {
		//debugger;
		var idHome=$("#cboHogar").val(); //$("#idHogar").val();
		if (idHome == null || idHome == 0) {
			swal("Aviso..!", "Seleccione un Hogar", "error");
			$("#cboHogar").focus();
			return false;
		}
		var idEncuesta = $("#cboEncuestas").val();
		var title = $("#cboEncuestas option:selected").text();
		//
		if (idEncuesta == null || idEncuesta == 0) {
			swal("Aviso..!", "Seleccione una Encuesta", "error");
			$("#cboEncuestas").focus();
			return false;
		}
		//
		$.ajax({				
			url:"g_Borrar_Encuestas.asp?idhogar=" + idHome + "&idencuesta=" + idEncuesta,
			cache: false,		
			beforeSend: function(objeto){
				$('#loader').html('<img src="./images/ajax-small.gif"> Espere borrando los Resultados...!');
			},
			success:function(data){				
				//debugger;
				console.log(data);
				$('#loader').html('');
				if (data=="True"){
					var msg ="Encuesta:\n" + title + "\nEliminada...!"; 
					swal("Aviso..!",  msg , "error");
					Reset();					
				}else{
					var msg ="Algo salio mal, Intente de nuevo!"; 
					swal("Aviso..!",  msg , "error");
					Reset();					
				}					
								
			}
		})			
	}
			
	function Reset() {
		//
		
		$('#generar').attr('disabled','disabled');
		$('#mostrar').attr('disabled','disabled');
		$('#borrar').attr('disabled','disabled');
		$("#tabla-resultados").html("");
		//
		var doc = document;			
		var mySelect = doc.getElementById('cboEncuestas');
		mySelect.selectedIndex = 0;
		var mySelect2 = doc.getElementById('cboHogar');
		mySelect2.selectedIndex = 0;
		$("#cboEncuestas").focus();			
	}
	

</script>  