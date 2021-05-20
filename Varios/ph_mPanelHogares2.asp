<!DOCTYPE HTML>
<html >
<head>
	<title>Panel de Hogares</title>
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
	<meta charset="utf-8">
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
	<link rel="icon" href="favicon.ico" type="image/x-icon"> 
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css">
  <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>
  <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/js/bootstrap.min.js"></script>
	
</head>
<body topmargin="0">
<!--#include file="estiloscss.asp"-->
<!--#include file="meta.asp"-->
<!--#include file="encabezado.asp"-->
<!--#include file="nn_subN.asp"-->
<!--#include file="in_DataEN.asp"-->

<%

  
'==========================================================================================
' Variables y Constantes
'==========================================================================================


    Apertura
%>
	<script>

$(document).ready(function(){
  $(".nav-tabs a").click(function(){
    $(this).tab('show');
  });
});
	
	</script>   
<%

   
'==========================================================================================
' Parámetros del Manteniemiento
'==========================================================================================
    LeePar
  
    
    if ed_iPas<>4 then 
        Encabezado
    end if    
	sExcel = request.Form("bus")

	'response.write "llego1"
	'response.end
    

%>
	<div style="width:98%">
	<!--General Inicio-->

		<div class="container">
			<h2>Socio Demografico</h2>
			<ul class="nav nav-pills">
				<li class="active"><a data-toggle="pill" href="#Tab01">DATOS DE IDENTIFICACION</a></li>
				<li><a data-toggle="pill" href="#Tab02">CARACTERISTICAS DE LA VIVIENDA</a></li>
				<li><a data-toggle="pill" href="#Tab03">SERVICIOS PUBLICOS</a></li>
				<li><a data-toggle="pill" href="#Tab04">SERVICIOS Y EQUIPAMIENTO DEL  HOGAR</a></li>
				<li><a data-toggle="pill" href="#Tab05">TELEVISORES EN EL HOGAR</a></li>
				<li><a data-toggle="pill" href="#Tab06">VEHICULOS</a></li>
				<li><a data-toggle="pill" href="#Tab07">TENENCIA  DE LA VIVIENDA</a></li>
				<li><a data-toggle="pill" href="#Tab08">COMPOSICION DEL HOGAR</a></li>
			</ul>
			<div class="tab-content">
				<div id="Tab01" class="tab-pane fade in active">
				  <h3>Titulo 01.-1</h3>
				  <p>Preguntas hacia abajo</p>
				</div>
				<div id="Tab02" class="tab-pane fade">
				  <h3>Titulo 02.-1</h3>
				  <p>Preguntas hacia abajo</p>
				</div>
				<div id="Tab03" class="tab-pane fade">
				  <h3>Titulo 03.-1</h3>
				  <p>Preguntas hacia abajo</p>
				</div>
				<div id="Tab04" class="tab-pane fade">
				  <h3>Titulo 04.-1</h3>
				  <p>Preguntas hacia abajo</p>
				</div>
				<div id="Tab05" class="tab-pane fade">
				  <h3>Titulo 05.-1</h3>
				  <p>Preguntas hacia abajo</p>
				</div>
				<div id="Tab06" class="tab-pane fade">
				  <h3>Titulo 06.-1</h3>
				  <p>Preguntas hacia abajo</p>
				</div>
				<div id="Tab07" class="tab-pane fade">
				  <h3>Titulo 07.-1</h3>
				  <p>Preguntas hacia abajo</p>
				</div>
				<div id="Tab08" class="tab-pane fade">
				  <h3>Titulo 08.-1</h3>
				  <p>Preguntas hacia abajo</p>
				</div>
			</div>
		</div>

	
	<!--General Fin-->
	</div>
	</center>

    <%conexion.close%>
	


</body>
</html>