<!DOCTYPE HTML>
<html >
<head>
	<title>Panel de Hogares</title>
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
	<meta charset="utf-8">
    <link href="01.css" rel="stylesheet" type="text/css" media="screen" />
	<link href="//netdna.bootstrapcdn.com/bootstrap/3.2.0/css/bootstrap.min.css" rel="stylesheet" id="bootstrap-css">
	<script src="//netdna.bootstrapcdn.com/bootstrap/3.2.0/js/bootstrap.min.js"></script>
	<script src="//cdnjs.cloudflare.com/ajax/libs/jquery/3.2.1/jquery.min.js"></script>
	<!------ Include the above in your HEAD tag ---------->

	<link href="//netdna.bootstrapcdn.com/bootstrap/3.2.0/css/bootstrap.min.css" rel="stylesheet" id="bootstrap-css">
	<script src="//netdna.bootstrapcdn.com/bootstrap/3.2.0/js/bootstrap.min.js"></script>
	<script src="//code.jquery.com/jquery-1.11.1.min.js"></script>
	<!------ Include the above in your HEAD tag ---------->
	
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
			<div class="page-header">
				<h1>Socio Demografico<span class="pull-right label label-default">:)</span></h1>
			</div>
			<div class="row">
			
				<div class="col-md-12">
					<div class="panel with-nav-tabs panel-primary">
						<div class="panel-heading">
								<ul class="nav nav-tabs">
									<li class="active"><a href="#tab1primary" data-toggle="tab">DATOS DE IDENTIFICACION</a></li>
									<li><a href="#tab2primary" data-toggle="tab">CARACTERISTICAS DE LA VIVIENDA</a></li>
									<li><a href="#tab3primary" data-toggle="tab">SERVICIOS PUBLICOS</a></li>
									<li><a href="#tab4primary" data-toggle="tab">SERVICIOS Y EQUIPAMIENTO DEL HOGAR</a></li>
									<li><a href="#tab5primary" data-toggle="tab">TELEVISORES EN EL HOGAR</a></li>
									<li><a href="#tab6primary" data-toggle="tab">VEHICULOS</a></li>
									<li><a href="#tab7primary" data-toggle="tab">TENENCIA  DE LA VIVIENDA</a></li>
									<li><a href="#tab8primary" data-toggle="tab">COMPOSICION DEL HOGAR</a></li>
								</ul>
						</div>
						<div class="panel-body">
							<div class="tab-content">
								<div class="tab-pane fade in active" id="tab1primary">
								<!--#include file="ph_mPanelHogaresP01.asp"-->
								</div>
								<div class="tab-pane fade" id="tab2primary">Primary 2</div>
								<div class="tab-pane fade" id="tab3primary">Primary 3</div>
								<div class="tab-pane fade" id="tab4primary">Primary 4</div>
								<div class="tab-pane fade" id="tab5primary">Primary 5</div>
								<div class="tab-pane fade" id="tab6primary">Primary 6</div>
								<div class="tab-pane fade" id="tab7primary">Primary 7</div>
								<div class="tab-pane fade" id="tab8primary">Primary 8</div>
							</div>
						</div>
					</div>
				</div>
			</div>
		</div>
		<br/>
	<!--General Fin-->
	</div>
	</center>

    <%conexion.close%>
	


</body>
</html>