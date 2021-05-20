<!DOCTYPE HTML>
<html >
<head>
	<title>Panel de Hogares</title>
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
	<meta http-equiv="refresh" content="60" />
    <link href="css/tabsPanelHogar.css" rel="stylesheet" type="text/css" media="screen" />
	
	<!--<link href="//netdna.bootstrapcdn.com/bootstrap/3.0.0/css/bootstrap.min.css" rel="stylesheet" id="bootstrap-css">-->
	<link href="tabs/bootstrap.min.css" rel="stylesheet" id="bootstrap-css">
	
	<!--<script src="//netdna.bootstrapcdn.com/bootstrap/3.0.0/js/bootstrap.min.js"></script>-->
	<script src="tabs/bootstrap.min.js"></script>

	<!--<script src="//code.jquery.com/jquery-1.11.1.min.js"></script>-->
	<script src="tabs/jquery-1.11.1.min.js"></script>
	
	<script type="text/javascript" src="js/tabsPanelHogar.js"></script>
	<!------ Include the above in your HEAD tag ---------->

	<!--<link href="//netdna.bootstrapcdn.com/font-awesome/4.0.3/css/font-awesome.css" rel="stylesheet">-->
	<link href="tabs/font-awesome.css" rel="stylesheet">
</head>
<body topmargin="0">
<!--#include file="estiloscss.asp"-->
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

	</script>   
<%

   
'==========================================================================================
' Parámetros del Manteniemiento
'==========================================================================================
    'LeePar
	sPar=""
  
    
    if ed_iPas<>4 then 
        Encabezado
    end if    
	sExcel = request.Form("bus")

	'response.write "llego1"
	'response.end
    

%>
	<!--General Inicio-->
	<div style="width:98%">
	

		<div class="container">
			<div class="panel-group" id="accordion">
				<!-- BLOQUE 1 -->
				<div class="panel panel-default">
					<div class="panel-heading">
						<h4 class="panel-title">
							<a data-toggle="collapse" data-parent="#accordion" href="#collapse0">
								<div class="row">
									<div class="col-md-1"><div class="step s0">1</div></div>
									<div class="col-md-11 step-text">Datos de Identificación</div>
								</div>
							</a>
						</h4>
					</div>
					<div id="collapse0" class="panel-collapse collapse in">
						<div class="panel-body">
							<div class="line-wizard l1"></div>
							<div class="row">
							<!--#include file="ph_mPanelHogaresP01.asp"-->
							</div>
						</div>
					</div>
				</div>
				<!-- BLOQUE 2 -->
				<div class="panel panel-default">
					<div class="panel-heading">
						<h4 class="panel-title">
							<a data-toggle="collapse" data-parent="#accordion" href="#collapse1">
								<div class="row">
									<div class="col-md-1"><div class="step s1">2</div></div>
									<div class="col-md-11 step-text">Características de la Vivienda</div>
								</div>
							</a>
						</h4>
					</div>
					<div id="collapse1" class="panel-collapse collapse">
						<div class="panel-body">
							<div class="line-wizard l2"></div>
							<div class="row">
								Preguntas Bloque 2
							</div>
						</div>
					</div>
				</div>
				<!-- BLOQUE 3 -->
				<div class="panel panel-default">
					<div class="panel-heading">
						<h4 class="panel-title">
							<a data-toggle="collapse" data-parent="#accordion" href="#collapse2">
								<div class="row">
									<div class="col-md-1"><div class="step s2">3</div></div>
									<div class="col-md-11 step-text">Servicios Públicos</div>
								</div>
							</a>
						</h4>
					</div>
					<div id="collapse2" class="panel-collapse collapse">
						<div class="panel-body">
							<div class="line-wizard l3"></div>
							<div class="row">
								Preguntas Bloque 3
							</div>
						</div>
					</div>
				</div>
				<!-- BLOQUE 4 -->
				<div class="panel panel-default">
					<div class="panel-heading">
						<h4 class="panel-title">
							<a data-toggle="collapse" data-parent="#accordion" href="#collapse3">
								<div class="row">
									<div class="col-md-1"><div class="step s3">4</div></div>
									<div class="col-md-11 step-text">Servicios y Equipamiento del Hogar</div>
								</div>
							</a>
						</h4>
					</div>
					<div id="collapse3" class="panel-collapse collapse">
						<div class="panel-body">
							<div class="line-wizard l3"></div>
							<div class="row">
								Preguntas Bloque 4
							</div>
						</div>
					</div>
				</div>
				<!-- BLOQUE 5 -->
				<div class="panel panel-default">
					<div class="panel-heading">
						<h4 class="panel-title">
							<a data-toggle="collapse" data-parent="#accordion" href="#collapse4">
								<div class="row">
									<div class="col-md-1"><div class="step s4">5</div></div>
									<div class="col-md-11 step-text">Televisores</div>
								</div>
							</a>
						</h4>
					</div>
					<div id="collapse4" class="panel-collapse collapse">
						<div class="panel-body">
							<div class="line-wizard l3"></div>
							<div class="row">
								Preguntas Bloque 5
							</div>
						</div>
					</div>
				</div>
				<!-- BLOQUE 6 -->
				<div class="panel panel-default">
					<div class="panel-heading">
						<h4 class="panel-title">
							<a data-toggle="collapse" data-parent="#accordion" href="#collapse5">
								<div class="row">
									<div class="col-md-1"><div class="step s5">6</div></div>
									<div class="col-md-11 step-text">Vehículos</div>
								</div>
							</a>
						</h4>
					</div>
					<div id="collapse5" class="panel-collapse collapse">
						<div class="panel-body">
							<div class="line-wizard l3"></div>
							<div class="row">
								Preguntas Bloque 6
							</div>
						</div>
					</div>
				</div>
				<!-- BLOQUE 7 -->
				<div class="panel panel-default">
					<div class="panel-heading">
						<h4 class="panel-title">
							<a data-toggle="collapse" data-parent="#accordion" href="#collapse6">
								<div class="row">
									<div class="col-md-1"><div class="step s6">7</div></div>
									<div class="col-md-11 step-text">Tenencia de la Vivienda</div>
								</div>
							</a>
						</h4>
					</div>
					<div id="collapse6" class="panel-collapse collapse">
						<div class="panel-body">
							<div class="line-wizard l3"></div>
							<div class="row">
								Preguntas Bloque 7
							</div>
						</div>
					</div>
				</div>
				<!-- BLOQUE 8 -->
				<div class="panel panel-default">
					<div class="panel-heading">
						<h4 class="panel-title">
							<a data-toggle="collapse" data-parent="#accordion" href="#collapse7">
								<div class="row">
									<div class="col-md-1"><div class="step s7">8</div></div>
									<div class="col-md-11 step-text">Composición del Hogar</div>
								</div>
							</a>
						</h4>
					</div>
					<div id="collapse7" class="panel-collapse collapse">
						<div class="panel-body">
							<div class="line-wizard l3"></div>
							<div class="row">
								Preguntas Bloque 8
							</div>
						</div>
					</div>
				</div>
				<!-- BLOQUE 9 -->
				<div class="panel panel-default">
					<div class="panel-heading">
						<h4 class="panel-title">
							<a data-toggle="collapse" data-parent="#accordion" href="#collapse8">
								<div class="row">
									<div class="col-md-1"><div class="step s8">9</div></div>
									<div class="col-md-11 step-text">Encargado de Compras</div>
								</div>
							</a>
						</h4>
					</div>
					<div id="collapse8" class="panel-collapse collapse">
						<div class="panel-body">
							<div class="line-wizard l3"></div>
							<div class="row">
								Preguntas Bloque 9
							</div>
						</div>
					</div>
				</div>
				<!-- BLOQUE 10 -->
				<div class="panel panel-default">
					<div class="panel-heading">
						<h4 class="panel-title">
							<a data-toggle="collapse" data-parent="#accordion" href="#collapse9">
								<div class="row">
									<div class="col-md-1"><div class="step s9">10</div></div>
									<div class="col-md-11 step-text">Titular de Cuenta - Beneficiario Incentivo</div>
								</div>
							</a>
						</h4>
					</div>
					<div id="collapse9" class="panel-collapse collapse">
						<div class="panel-body">
							<div class="line-wizard l3"></div>
							<div class="row">
								Preguntas Bloque 10
							</div>
						</div>
					</div>
				</div>



			</div>
		</div>	


		
		<br/>
	</div>
	<!--General Fin-->
	</center>

    <%conexion.close%>
	


</body>
</html>