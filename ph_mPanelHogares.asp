<!DOCTYPE HTML>
<html >
<head>
	<title>Panel de Hogares</title>
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
	<!--<meta http-equiv="refresh" content="240" />-->
	<meta name="viewport" content="width=device-width, initial-scale=1">
	<meta charset="utf-8">
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<link href="modal.css" rel="stylesheet" type="text/css" media="screen" />	
	<script type="text/javascript" src="js/sweetalert.min.js"></script>
	<script type="text/javascript" src="js/bootstrap.min.js"></script>
	<link href="css/bootstrap.min.css" rel="stylesheet" type="text/css" media="screen" />	
		
<style>
@keyframes showSweetAlert {
  0% {
    transform: scale(0.7);
  }
  45% {
    transform: scale(1.05);
  }
  80% {
    transform: scale(0.95);
  }
  100% {
    transform: scale(1);
  }
}
@keyframes hideSweetAlert {
  0% {
    transform: scale(1);
  }
  100% {
    transform: scale(0.5);
  }
}
@keyframes slideFromTop {
  0% {
    top: 0%;
  }
  100% {
    top: 50%;
  }
}
@keyframes slideToTop {
  0% {
    top: 50%;
  }
  100% {
    top: 0%;
  }
}
@keyframes slideFromBottom {
  0% {
    top: 70%;
  }
  100% {
    top: 50%;
  }
}
@keyframes slideToBottom {
  0% {
    top: 50%;
  }
  100% {
    top: 70%;
  }
}
.showSweetAlert {
  animation: showSweetAlert 0.3s;
}
.showSweetAlert[data-animation=none] {
  animation: none;
}
.showSweetAlert[data-animation=slide-from-top] {
  animation: slideFromTop 0.3s;
}
.showSweetAlert[data-animation=slide-from-bottom] {
  animation: slideFromBottom 0.3s;
}
.hideSweetAlert {
  animation: hideSweetAlert 0.3s;
}
.hideSweetAlert[data-animation=none] {
  animation: none;
}
.hideSweetAlert[data-animation=slide-from-top] {
  animation: slideToTop 0.3s;
}
.hideSweetAlert[data-animation=slide-from-bottom] {
  animation: slideToBottom 0.3s;
}
@keyframes animateSuccessTip {
  0% {
    width: 0;
    left: 1px;
    top: 19px;
  }
  54% {
    width: 0;
    left: 1px;
    top: 19px;
  }
  70% {
    width: 50px;
    left: -8px;
    top: 37px;
  }
  84% {
    width: 17px;
    left: 21px;
    top: 48px;
  }
  100% {
    width: 25px;
    left: 14px;
    top: 45px;
  }
}
@keyframes animateSuccessLong {
  0% {
    width: 0;
    right: 46px;
    top: 54px;
  }
  65% {
    width: 0;
    right: 46px;
    top: 54px;
  }
  84% {
    width: 55px;
    right: 0px;
    top: 35px;
  }
  100% {
    width: 47px;
    right: 8px;
    top: 38px;
  }
}
@keyframes rotatePlaceholder {
  0% {
    transform: rotate(-45deg);
  }
  5% {
    transform: rotate(-45deg);
  }
  12% {
    transform: rotate(-405deg);
  }
  100% {
    transform: rotate(-405deg);
  }
}
.animateSuccessTip {
  animation: animateSuccessTip 0.75s;
}
.animateSuccessLong {
  animation: animateSuccessLong 0.75s;
}
.sa-icon.sa-success.animate::after {
  animation: rotatePlaceholder 4.25s ease-in;
}
@keyframes animateErrorIcon {
  0% {
    transform: rotateX(100deg);
    opacity: 0;
  }
  100% {
    transform: rotateX(0deg);
    opacity: 1;
  }
}
.animateErrorIcon {
  animation: animateErrorIcon 0.5s;
}
@keyframes animateXMark {
  0% {
    transform: scale(0.4);
    margin-top: 26px;
    opacity: 0;
  }
  50% {
    transform: scale(0.4);
    margin-top: 26px;
    opacity: 0;
  }
  80% {
    transform: scale(1.15);
    margin-top: -6px;
  }
  100% {
    transform: scale(1);
    margin-top: 0;
    opacity: 1;
  }
}
.animateXMark {
  animation: animateXMark 0.5s;
}
@keyframes pulseWarning {
  0% {
    border-color: #F8D486;
  }
  100% {
    border-color: #F8BB86;
  }
}
.pulseWarning {
  animation: pulseWarning 0.75s infinite alternate;
}
@keyframes pulseWarningIns {
  0% {
    background-color: #F8D486;
  }
  100% {
    background-color: #F8BB86;
  }
}
.pulseWarningIns {
  animation: pulseWarningIns 0.75s infinite alternate;
}
@keyframes rotate-loading {
  0% {
    transform: rotate(0deg);
  }
  100% {
    transform: rotate(360deg);
  }
}
body.stop-scrolling {
  height: 100%;
  overflow: hidden;
}
.sweet-overlay {
  background-color: rgba(0, 0, 0, 0.4);
  position: fixed;
  left: 0;
  right: 0;
  top: 0;
  bottom: 0;
  display: none;
  z-index: 1040;
}
.sweet-alert {
  background-color: #ffffff;
  width: 478px;
  padding: 17px;
  border-radius: 5px;
  text-align: center;
  position: fixed;
  left: 50%;
  top: 50%;
  margin-left: -256px;
  margin-top: -200px;
  overflow: hidden;
  display: none;
  z-index: 2000;
}
@media all and (max-width: 767px) {
  .sweet-alert {
    width: auto;
    margin-left: 0;
    margin-right: 0;
    left: 15px;
    right: 15px;
  }
}
.sweet-alert .form-group {
  display: none;
}
.sweet-alert .form-group .sa-input-error {
  display: none;
}
.sweet-alert.show-input .form-group {
  display: block;
}
.sweet-alert .sa-confirm-button-container {
  display: inline-block;
  position: relative;
}
.sweet-alert .la-ball-fall {
  position: absolute;
  left: 50%;
  top: 50%;
  margin-left: -27px;
  margin-top: -9px;
  opacity: 0;
  visibility: hidden;
}
.sweet-alert button[disabled] {
  opacity: .6;
  cursor: default;
}
.sweet-alert button.confirm[disabled] {
  color: transparent;
}
.sweet-alert button.confirm[disabled] ~ .la-ball-fall {
  opacity: 1;
  visibility: visible;
  transition-delay: 0s;
}
.sweet-alert .sa-icon {
  width: 80px;
  height: 80px;
  border: 4px solid gray;
  border-radius: 50%;
  margin: 20px auto;
  position: relative;
  box-sizing: content-box;
}
.sweet-alert .sa-icon.sa-error {
  border-color: #d43f3a;
}
.sweet-alert .sa-icon.sa-error .sa-x-mark {
  position: relative;
  display: block;
}
.sweet-alert .sa-icon.sa-error .sa-line {
  position: absolute;
  height: 5px;
  width: 47px;
  background-color: #d9534f;
  display: block;
  top: 37px;
  border-radius: 2px;
}
.sweet-alert .sa-icon.sa-error .sa-line.sa-left {
  transform: rotate(45deg);
  left: 17px;
}
.sweet-alert .sa-icon.sa-error .sa-line.sa-right {
  transform: rotate(-45deg);
  right: 16px;
}
.sweet-alert .sa-icon.sa-warning {
  border-color: #eea236;
}
.sweet-alert .sa-icon.sa-warning .sa-body {
  position: absolute;
  width: 5px;
  height: 47px;
  left: 50%;
  top: 10px;
  border-radius: 2px;
  margin-left: -2px;
  background-color: #f0ad4e;
}
.sweet-alert .sa-icon.sa-warning .sa-dot {
  position: absolute;
  width: 7px;
  height: 7px;
  border-radius: 50%;
  margin-left: -3px;
  left: 50%;
  bottom: 10px;
  background-color: #f0ad4e;
}
.sweet-alert .sa-icon.sa-info {
  border-color: #46b8da;
}
.sweet-alert .sa-icon.sa-info::before {
  content: "";
  position: absolute;
  width: 5px;
  height: 29px;
  left: 50%;
  bottom: 17px;
  border-radius: 2px;
  margin-left: -2px;
  background-color: #5bc0de;
}
.sweet-alert .sa-icon.sa-info::after {
  content: "";
  position: absolute;
  width: 7px;
  height: 7px;
  border-radius: 50%;
  margin-left: -3px;
  top: 19px;
  background-color: #5bc0de;
}
.sweet-alert .sa-icon.sa-success {
  border-color: #4cae4c;
}
.sweet-alert .sa-icon.sa-success::before,
.sweet-alert .sa-icon.sa-success::after {
  content: '';
  border-radius: 50%;
  position: absolute;
  width: 60px;
  height: 120px;
  background: #ffffff;
  transform: rotate(45deg);
}
.sweet-alert .sa-icon.sa-success::before {
  border-radius: 120px 0 0 120px;
  top: -7px;
  left: -33px;
  transform: rotate(-45deg);
  transform-origin: 60px 60px;
}
.sweet-alert .sa-icon.sa-success::after {
  border-radius: 0 120px 120px 0;
  top: -11px;
  left: 30px;
  transform: rotate(-45deg);
  transform-origin: 0px 60px;
}
.sweet-alert .sa-icon.sa-success .sa-placeholder {
  width: 80px;
  height: 80px;
  border: 4px solid rgba(92, 184, 92, 0.2);
  border-radius: 50%;
  box-sizing: content-box;
  position: absolute;
  left: -4px;
  top: -4px;
  z-index: 2;
}
.sweet-alert .sa-icon.sa-success .sa-fix {
  width: 5px;
  height: 90px;
  background-color: #ffffff;
  position: absolute;
  left: 28px;
  top: 8px;
  z-index: 1;
  transform: rotate(-45deg);
}
.sweet-alert .sa-icon.sa-success .sa-line {
  height: 5px;
  background-color: #5cb85c;
  display: block;
  border-radius: 2px;
  position: absolute;
  z-index: 2;
}
.sweet-alert .sa-icon.sa-success .sa-line.sa-tip {
  width: 25px;
  left: 14px;
  top: 46px;
  transform: rotate(45deg);
}
.sweet-alert .sa-icon.sa-success .sa-line.sa-long {
  width: 47px;
  right: 8px;
  top: 38px;
  transform: rotate(-45deg);
}
.sweet-alert .sa-icon.sa-custom {
  background-size: contain;
  border-radius: 0;
  border: none;
  background-position: center center;
  background-repeat: no-repeat;
}
.sweet-alert .btn-default:focus {
  border-color: #cccccc;
  outline: 0;
  -webkit-box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(204, 204, 204, 0.6);
  box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(204, 204, 204, 0.6);
}
.sweet-alert .btn-success:focus {
  border-color: #4cae4c;
  outline: 0;
  -webkit-box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(76, 174, 76, 0.6);
  box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(76, 174, 76, 0.6);
}
.sweet-alert .btn-info:focus {
  border-color: #46b8da;
  outline: 0;
  -webkit-box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(70, 184, 218, 0.6);
  box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(70, 184, 218, 0.6);
}
.sweet-alert .btn-danger:focus {
  border-color: #d43f3a;
  outline: 0;
  -webkit-box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(212, 63, 58, 0.6);
  box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(212, 63, 58, 0.6);
}
.sweet-alert .btn-warning:focus {
  border-color: #eea236;
  outline: 0;
  -webkit-box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(238, 162, 54, 0.6);
  box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(238, 162, 54, 0.6);
}
.sweet-alert button::-moz-focus-inner {
  border: 0;
}
/*!
 * Load Awesome v1.1.0 (http://github.danielcardoso.net/load-awesome/)
 * Copyright 2015 Daniel Cardoso <@DanielCardoso>
 * Licensed under MIT
 */
.la-ball-fall,
.la-ball-fall > div {
  position: relative;
  -webkit-box-sizing: border-box;
  -moz-box-sizing: border-box;
  box-sizing: border-box;
}
.la-ball-fall {
  display: block;
  font-size: 0;
  color: #fff;
}
.la-ball-fall.la-dark {
  color: #333;
}
.la-ball-fall > div {
  display: inline-block;
  float: none;
  background-color: currentColor;
  border: 0 solid currentColor;
}
.la-ball-fall {
  width: 54px;
  height: 18px;
}
.la-ball-fall > div {
  width: 10px;
  height: 10px;
  margin: 4px;
  border-radius: 100%;
  opacity: 0;
  -webkit-animation: ball-fall 1s ease-in-out infinite;
  -moz-animation: ball-fall 1s ease-in-out infinite;
  -o-animation: ball-fall 1s ease-in-out infinite;
  animation: ball-fall 1s ease-in-out infinite;
}
.la-ball-fall > div:nth-child(1) {
  -webkit-animation-delay: -200ms;
  -moz-animation-delay: -200ms;
  -o-animation-delay: -200ms;
  animation-delay: -200ms;
}
.la-ball-fall > div:nth-child(2) {
  -webkit-animation-delay: -100ms;
  -moz-animation-delay: -100ms;
  -o-animation-delay: -100ms;
  animation-delay: -100ms;
}
.la-ball-fall > div:nth-child(3) {
  -webkit-animation-delay: 0ms;
  -moz-animation-delay: 0ms;
  -o-animation-delay: 0ms;
  animation-delay: 0ms;
}
.la-ball-fall.la-sm {
  width: 26px;
  height: 8px;
}
.la-ball-fall.la-sm > div {
  width: 4px;
  height: 4px;
  margin: 2px;
}
.la-ball-fall.la-2x {
  width: 108px;
  height: 36px;
}
.la-ball-fall.la-2x > div {
  width: 20px;
  height: 20px;
  margin: 8px;
}
.la-ball-fall.la-3x {
  width: 162px;
  height: 54px;
}
.la-ball-fall.la-3x > div {
  width: 30px;
  height: 30px;
  margin: 12px;
}
/*
 * Animation
 */
@-webkit-keyframes ball-fall {
  0% {
    opacity: 0;
    -webkit-transform: translateY(-145%);
    transform: translateY(-145%);
  }
  10% {
    opacity: .5;
  }
  20% {
    opacity: 1;
    -webkit-transform: translateY(0);
    transform: translateY(0);
  }
  80% {
    opacity: 1;
    -webkit-transform: translateY(0);
    transform: translateY(0);
  }
  90% {
    opacity: .5;
  }
  100% {
    opacity: 0;
    -webkit-transform: translateY(145%);
    transform: translateY(145%);
  }
}
@-moz-keyframes ball-fall {
  0% {
    opacity: 0;
    -moz-transform: translateY(-145%);
    transform: translateY(-145%);
  }
  10% {
    opacity: .5;
  }
  20% {
    opacity: 1;
    -moz-transform: translateY(0);
    transform: translateY(0);
  }
  80% {
    opacity: 1;
    -moz-transform: translateY(0);
    transform: translateY(0);
  }
  90% {
    opacity: .5;
  }
  100% {
    opacity: 0;
    -moz-transform: translateY(145%);
    transform: translateY(145%);
  }
}
@-o-keyframes ball-fall {
  0% {
    opacity: 0;
    -o-transform: translateY(-145%);
    transform: translateY(-145%);
  }
  10% {
    opacity: .5;
  }
  20% {
    opacity: 1;
    -o-transform: translateY(0);
    transform: translateY(0);
  }
  80% {
    opacity: 1;
    -o-transform: translateY(0);
    transform: translateY(0);
  }
  90% {
    opacity: .5;
  }
  100% {
    opacity: 0;
    -o-transform: translateY(145%);
    transform: translateY(145%);
  }
}
@keyframes ball-fall {
  0% {
    opacity: 0;
    -webkit-transform: translateY(-145%);
    -moz-transform: translateY(-145%);
    -o-transform: translateY(-145%);
    transform: translateY(-145%);
  }
  10% {
    opacity: .5;
  }
  20% {
    opacity: 1;
    -webkit-transform: translateY(0);
    -moz-transform: translateY(0);
    -o-transform: translateY(0);
    transform: translateY(0);
  }
  80% {
    opacity: 1;
    -webkit-transform: translateY(0);
    -moz-transform: translateY(0);
    -o-transform: translateY(0);
    transform: translateY(0);
  }
  90% {
    opacity: .5;
  }
  100% {
    opacity: 0;
    -webkit-transform: translateY(145%);
    -moz-transform: translateY(145%);
    -o-transform: translateY(145%);
    transform: translateY(145%);
  }
}

.accordion {
  background-color: #eee;
  color: #444;
  cursor: pointer;
  padding: 18px;
  width: 100%;
  border: none;
  text-align: left;
  outline: none;
  font-size: 20px;
  transition: 0.4s;
}

.active, .accordion:hover {
  background-color: #ccc;
}

.accordion:after {
  content: '\002B';
  color: #777;
  font-weight: bold;
  float: right;
  margin-left: 5px;
}

.active:after {
  content: "\2212";
}

.panel {
  padding: 0 18px;
  background-color: white;
  max-height: 0;
  overflow: hidden;
  transition: max-height 0.2s ease-out;
}
</style>
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
	<!--hidden-button-->
	<!--PANTALLAS MODALES-->
	
	<div id="DetallePauta">
		<div id="1" class="modalmask">
			<div class="modalbox movedown">
				<a href="#close" title="Cerrar" class="close">x</a>
				<fieldset class="det">
				<legend class="det">Incluya el Numero de Identificacion del Hogar o Palabra Clave y presione la imagen de Buscar</legend>
					Buscar:<input type="text" name="BuscarHogar" id="BuscarHogar" align="right" size=10>
					<img src="images/Buscar.png"  style="margin-left:0px;" alt="Buscar" width="20px" onclick="BuscarHogares();"/>
					<br>
					<br>
					<span id="loader2"></span>
					<div id="DivBuscarHogares" style="width:140px; float:left;">
					</div>
				</fieldset>
			</div> <!--modalbox rotate-->
		</div> <!--modal4-->

	</div> <!--DetallePauta1-->
	<div style="width:98%">
		<div class="container-fluid">        
			<div class="row">
				<!--Contenido General-->			
				<div class="container">
					<div class="col-md-8 col-sm-8 col-xs-12">
						<div class="pull-right">
							<img src="images/Nuevo.jpg"  style="margin-left:0px;" title="Crear Hogar Nuevo" alt="Nuevo" width="50px" onclick="NuevoHogar()"/>
						</div>
						<div class="pull-right">
							<a href="#1" title="Buscar Hogar por Lista">
								<img src="images/BuscarLista.jpg"  style="margin-left:0px;" alt="BuscarLista" width="55px" onclick=""/>
							</a>
						</div>
					</div>
				</div>
			</div>
		</div>
	</div>
	<!--General Inicio-->
	<center>
	<div> 
	
		<h2>Panel de Hogares-Socio Demogr&aacute;fico</h2>
		<h4>N&uacute;mero de Identificaci&oacute;n del Hogar</h4>
		<h4>
			<!--hidden -->
			
			<div id="DivHogar"> 
				<!--<input type="text" name="Hogar" disabled id="Hogar" value="136" size=6 style="text-align:right; background-color:#d1d1d1;">
				<input type="text" name="CodigoHogar" disabled id="CodigoHogar" value="0200100120" size=10 style="text-align:right; background-color:#d1d1d1;">-->
				<input type="text" name="Hogar" disabled id="Hogar" value="" size=6 style="text-align:right; background-color:#d1d1d1;">
				<input type="text" name="CodigoHogar" disabled id="CodigoHogar" value="" size=10 style="text-align:right; background-color:#d1d1d1;">
			</div>
		</h4>
		<h4 style=" text-align:right;">
			<div id="DivPreguntaObigatorias"> 
				Total Preguntas Obligatorias
				<input type="text" name="TotalPreguntas" disabled id="TotalPreguntas" value="17" size=3 style="text-align:right; background-color:#d1d1d1;">
			</div>
		</h4>
		<h4 style=" text-align:right;">
			<div id="DivPreguntaObigatorias"> 
				Validas
				<input type="text" name="TotalValidas" disabled id="TotalValidas" value="" size=3 style="text-align:right; background-color:#d1d1d1;">
			</div>
		</h4>
		
		<p>haga clic sobre el bloque que desea expandir, para incluir la informaci&oacute;n del Hogar a ser Registrado.</p>
		<!--<button type="button" onclick="alert('Nuevo Registro')">Nuevo Registro</button>-->


		<!-- BLOQUE 0 -->
		<button class="accordion">0-. Datos de Identificaci&oacute;n del Hogar</button>
		<div class="panel">
			<!--#include file="ph_mPanelHogaresP01.asp"-->		  
		</div>

		<!-- BLOQUE 1 -->
		<button class="accordion">1.- Responsable del Panel</button>
		<div class="panel">
			<!--#include file="ph_mPanelHogaresP00.asp"-->		  
		</div>

		<!-- BLOQUE 8 -->
		<button class="accordion">2.- Composici&oacute;n del Hogar</button>
		<div class="panel">
			<!--#include file="ph_mPanelHogaresP08.asp"-->
		</div>


		<!-- BLOQUE 2 -->
		<button class="accordion">3.- Caracter&iacute;sticas de la Vivienda</button>
		<div class="panel">
			<!--#include file="ph_mPanelHogaresP02.asp"-->
		</div>

		<!-- BLOQUE 7 -->
		<button class="accordion">4.- Tenencia de la Vivienda</button>
		<div class="panel">
			<!--#include file="ph_mPanelHogaresP07.asp"-->
		</div>
		
		<!-- BLOQUE 3 -->
		<button class="accordion">5.- Servicios P&uacute;blicos</button>
		<div class="panel">
			<!--#include file="ph_mPanelHogaresP03.asp"-->
		</div>
		
		<!-- BLOQUE 4 -->
		<button class="accordion">6.- Servicios y Equipamiento del Hogar</button>
		<div class="panel">
			<!--#include file="ph_mPanelHogaresP04.asp"-->
		</div>

		<!-- BLOQUE 5 -->
		<button class="accordion">7.- Televisores</button>
		<div class="panel">
			<!--#include file="ph_mPanelHogaresP05.asp"-->
		</div>

		<!-- BLOQUE 6 -->
		<button class="accordion">8.- Veh&iacute;culos</button>
		<div class="panel">
			<!--#include file="ph_mPanelHogaresP06.asp"-->
		</div>

		<!-- BLOQUE 10 -->
		<button class="accordion">9.- Otros</button>
		<div class="panel">
			<!--#include file="ph_mPanelHogaresP10.asp"-->
		</div>

		<!-- BLOQUE 11 -->
		<button class="accordion">10.- Informacion General</button>
		<div class="panel">
			<!--#include file="ph_mPanelHogaresP11.asp"-->
		</div>
		

		<script>
			var acc = document.getElementsByClassName("accordion");
			var i;

			for (i = 0; i < acc.length; i++) {
			  acc[i].addEventListener("click", function() {
				this.classList.toggle("active");
				var panel = this.nextElementSibling;
				if (panel.style.maxHeight) {
				  panel.style.maxHeight = null;
				} else {
				  panel.style.maxHeight = panel.scrollHeight + "px";
				} 
			  });
			}
		</script>

		<br/>
	</div>
	<!--General Fin-->
	</center>

    <%conexion.close%>
	


</body>
</html>
<script>

	//**Inicio Buscar Ciudad Total
	function buscar_ciudadtotal(){
		//alert("Llego Ciudad0");
		var sx = "g_BuscarCiudadTotal.asp";
		$.ajax({
			url:sx,
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
			},
			success:function(data){
				
				$('#loader2').html('');
				console.log(data);
				$('#DivCiudad').html(data);
				//alert("paso");
			}
		})
	}	
	//**Fin Buscar Ciudad Total

	//**Inicio Buscar Hogares
	function SelHogar(id){
		BlanquearCampos();
		//30Dic
		//alert("SelHogar");
		//debugger;
		
		var value= id;
		//BLOQUE 0
		let ajax = {
			id: id						
		  };
		  $.ajax({
			url: "g_BuscarBloque0.asp",
			type: "POST",
			cache: false,
			data: ajax,			
			dataType:"JSON",			
			success: function (data, textStatus, jqXHR) {				
				console.log(data);
				//debugger;				
				//					
				//buscar_municipio();
				$("#Hogar").val(id);
				$("#CodigoHogar").val(data[0].CodigoHogar);
				$("#Pais").val(data[0].Id_Pais).change();
				$("#Estado").val(data[0].Id_Estado).change();
				$("#Ciudad").val(data[0].Id_Ciudad);
				$("#Municipio").val(data[0].Id_Municipio);
				$("#Parroquia").val(data[0].Id_Parroquia);
				$("#Calle").val(data[0].Calle);
				$("#Edificio").val(data[0].Edificio);
				$("#Casa").val(data[0].Casa);
				$("#Escalera").val(data[0].Escalera);
				$("#Piso").val(data[0].Piso);
				$("#Apartamento").val(data[0].Apto);
				$("#Barrio").val(data[0].Barrio);
				$("#Referencia").val(data[0].Referencia);
				$("#TelefonoLocal").val(data[0].TelefonoLocal);
				//let iciu = data[0].Id_Ciudad;
				//alert("Ciu:=" + iciu);
				//document.getElementById("Ciudad").value = 2;
			},
			error: function (request, textStatus, errorThrown) {
			  alert("Error "+ request.responseJSON.message + "error");
			},
		  });
		SelHogar1(id);
	}	

	function SelHogar1(id){
		//debugger;		
		//BLOQUE 1
		//alert("SelHogar1");
		var value= id;
		let ajax = {
			id: id						
		  };
		  $.ajax({
			url: "g_BuscarBloque1.asp",
			type: "POST",
			cache: false,
			data: ajax,			
			dataType:"JSON",			
			success: function (data1, textStatus, jqXHR) {				
				console.log(data1);
				//debugger;				
				//					
				$("#Hogar").val(id);
				$("#PrimerNombre").val(data1[0].Nombre1);
				$("#SegundoNombre").val(data1[0].Nombre2);
				$("#PrimerApellido").val(data1[0].Apellido1);
				$("#SegundoApellido").val(data1[0].Apellido2);
				$("#Nacionalidad").val(data1[0].Id_Nacionalidad);
				$("#Cedula").val(data1[0].Cedula);
				$("#Celular").val(data1[0].Celular);
				$("#CelularAdicional").val(data1[0].CelularAdicional);
				$("#NumeroCortesia").val(data1[0].NumeroCortesia);
				$("#Correo").val(data1[0].Correo);
				$("#CorreoAlterno").val(data1[0].CorreoAlterno);
				$("#Titular").val(data1[0].Titular);
				$("#CedulaTitular").val(data1[0].CedulaTitular);
				$("#Banco").val(data1[0].Id_Banco);
				$("#Cuenta").val(data1[0].NumeroCuenta);
				$("#PagoRapido").val(data1[0].Id_PagoRapido);
				$("#Parentesco").val(data1[0].Id_Parentesco);
				$("#EstadoCivil").val(data1[0].Id_EstadoCivil);
				$("#FechaNacimiento").val(data1[0].Fec_Nacimiento);
				$("#Sexo").val(data1[0].Id_Sexo);
				$("#Educacion").val(data1[0].Id_Educacion);
				$("#FrecuenciaCompra").val(data1[0].Id_FrecuenciaCompra);
				$("#TipoIngreso").val(data1[0].Id_TipoIngreso);
				$("#NumeroPersonas").val(data1[0].CantidadPersonas);
				buscar_edad();
			},
			error: function (request, textStatus, errorThrown) {
			  alert("Error "+ request.responseJSON.message + "error");
			},
		  });

		  SelHogar3(id);
	
	}	

	function SelHogar3(id){
		//debugger;		
		//BLOQUE 3
		//alert("SelHogar3");
		var value= id;
		let ajax = {
			id: id						
		  };
		  $.ajax({
			url: "g_BuscarBloque3.asp",
			type: "POST",
			cache: false,
			data: ajax,			
			dataType:"JSON",			
			success: function (data3, textStatus, jqXHR) {				
				console.log(data3);
				//debugger;				
				//					
				$("#Hogar").val(id);
				$("#TipoVivienda").val(data3[0].Id_TipoVivienda);
				$("#Explique").val(data3[0].OtroTipoVivienda);
				$("#MetrosVivienda").val(data3[0].Id_Metros);
				$("#TotalAmbientes").val(data3[0].NumeroAmbientes);
				$("#TotalBanos").val(data3[0].NumeroBanos);
				$("#PuntosLuz").val(data3[0].id_PuntosLuz);
			},
			error: function (request, textStatus, errorThrown) {
			  alert("Error "+ request.responseJSON.message + "error");
			},
		  });

		  SelHogar4(id);
	
	}	

	function SelHogar4(id){
		//debugger;		
		//BLOQUE 4
		//alert("SelHogar4");
		var value= id;
		let ajax = {
			id: id						
		  };
		  $.ajax({
			url: "g_BuscarBloque4.asp",
			type: "POST",
			cache: false,
			data: ajax,			
			dataType:"JSON",			
			success: function (data4, textStatus, jqXHR) {				
				console.log(data4);
				//debugger;				
				//					
				$("#Hogar").val(id);
				$("#OcupacionVivienda").val(data4[0].Id_OcupacionVivienda);
				$("#ExpliqueOcupacion").val(data4[0].OtroOcupacionVivienda);
				$("#MontoVivienda").val(data4[0].Id_MontoVivienda);
			},
			error: function (request, textStatus, errorThrown) {
			  alert("Error "+ request.responseJSON.message + "error");
			},
		  });

		  SelHogar5(id);
	
	}	
	
	function SelHogar5(id){
		//debugger;		
		//BLOQUE 5
		//alert("SelHogar5");
		var value= id;
		let ajax = {
			id: id						
		  };
		  $.ajax({
			url: "g_BuscarBloque5.asp",
			type: "POST",
			cache: false,
			data: ajax,			
			dataType:"JSON",			
			success: function (data5, textStatus, jqXHR) {				
				console.log(data5);
				//debugger;				
				//					
				$("#Hogar").val(id);
				$("#AguasBlancas").val(data5[0].Id_AguasBlancas);
				$("#AguasNegras").val(data5[0].Id_AguasNegras);
				$("#AseoUrbano").val(data5[0].Id_AseoUrbano);
				$("#Electricidad").val(data5[0].Id_ServicioElectricidad);
				$("#Telefono").val(data5[0].Id_ServicioTelefono);
			},
			error: function (request, textStatus, errorThrown) {
			  alert("Error "+ request.responseJSON.message + "error");
			},
		  });

		  SelHogar6(id);
	
	}	

	function SelHogar6(id){
		//debugger;		
		//BLOQUE 6
		//alert("SelHogar6");
		var value= id;
		let ajax = {
			id: id						
		  };
		  $.ajax({
			url: "g_BuscarBloque6.asp",
			type: "POST",
			cache: false,
			data: ajax,			
			dataType:"JSON",			
			success: function (data6, textStatus, jqXHR) {				
				console.log(data6);
				//debugger;				
				//					
				$("#Hogar").val(id);
				$("#DomesticaFija").val(data6[0].Id_DomesticaFija);
				$("#PersonalLabores").val(data6[0].Id_PersonalLabores);
				$("#DomesticaDia").val(data6[0].Id_DomesticaDia);
				$("#Conexion1").val(data6[0].id_ConexionInternet1);
				$("#Conexion2").val(data6[0].id_ConexionInternet2);
				$("#Conexion3").val(data6[0].id_ConexionInternet3);
				$("#TelefonoCelular").val(data6[0].id_CelularJefe);
				$("#Seguro1").val(data6[0].id_SeguroHCMParticular);
				$("#Seguro2").val(data6[0].id_SeguroHCMColectivo);
				$("#Seguro3").val(data6[0].id_SeguroHCMSS);
				$("#Aire").val(data6[0].Id_AireAcondicionado);
				$("#CalentadorElectrico").val(data6[0].Id_Calentador1);
				$("#CalentadorGas").val(data6[0].Id_Calentador2);
				$("#ComputadorPc").val(data6[0].Id_Computador1);
				$("#ComputadorLaptop").val(data6[0].Id_Computador2);
				$("#Dvd").val(data6[0].Id_DVD);
				$("#Home").val(data6[0].Id_HomeTheater);
				$("#Juego").val(data6[0].Id_JuegosVodeo);
				$("#Horno").val(data6[0].Id_HornoMicro);
				$("#Secadora").val(data6[0].Id_Secadora);
				$("#LavadoraAutomatica").val(data6[0].Id_Lavadora1);
				$("#LavadoraSemi").val(data6[0].Id_Lavadora2);
				$("#LavadoraRodillo").val(data6[0].Id_Lavadora3);
				$("#Nevera").val(data6[0].Id_Nevera);
				$("#Freezer").val(data6[0].Id_Freezer);
				$("#CocinaElectrica").val(data6[0].Id_Cocina1);
				$("#CocinaBombona").val(data6[0].Id_Cocina2);
				$("#CocinaGas").val(data6[0].Id_Cocina3);
				$("#CocinaKerosene").val(data6[0].Id_Cocina4);
				$("#Lavaplatos").val(data6[0].Id_LavaPlato);
			},
			error: function (request, textStatus, errorThrown) {
			  alert("Error "+ request.responseJSON.message + "error");
			},
		  });

		  SelHogar7(id);
	
	}	
	
	function SelHogar7(id){
		//debugger;		
		//BLOQUE 7
		//alert("SelHogar7");
		var value= id;
		let ajax = {
			id: id						
		  };
		  $.ajax({
			url: "g_BuscarBloque7.asp",
			type: "POST",
			cache: false,
			data: ajax,			
			dataType:"JSON",			
			success: function (data7, textStatus, jqXHR) {				
				console.log(data7);
				//debugger;				
				//					
				$("#Hogar").val(id);
				$("#NumeroTeletisores").val(data7[0].Id_Televisores);
				$("#TipoTelevisores").val(data7[0].Id_TipoTelevisores);
				$("#Senal").val(data7[0].Id_Senal);
				$("#Cableras1").val(data7[0].Id_Cablera1);
				$("#Cableras2").val(data7[0].Id_Cablera2);
				$("#TvOnline1").val(data7[0].Id_TelevisionOnline1);
				$("#TvOnline2").val(data7[0].Id_TelevisionOnline2);
			},
			error: function (request, textStatus, errorThrown) {
			  alert("Error "+ request.responseJSON.message + "error");
			},
		  });

		  SelHogar8(id);
	
	}	

	function SelHogar8(id){
		//debugger;		
		//BLOQUE 8
		//alert("SelHogar8");
		var value= id;
		let ajax = {
			id: id						
		  };
		  $.ajax({
			url: "g_BuscarBloque8.asp",
			type: "POST",
			cache: false,
			data: ajax,			
			dataType:"JSON",			
			success: function (data8, textStatus, jqXHR) {				
				console.log(data8);
				//debugger;				
				//					
				$("#Hogar").val(id);
				$("#Autos").val(data8[0].Id_Autos);
				$("#Moto").val(data8[0].Id_Moto);
				$("#SeguroCasco").val(data8[0].Id_SeguroCasco);
			},
			error: function (request, textStatus, errorThrown) {
			  alert("Error "+ request.responseJSON.message + "error");
			},
		  });

		  SelHogar9(id);
	
	}	

	function SelHogar9(id){
		//debugger;		
		//BLOQUE 9
		//alert("SelHogar9");
		var value= id;
		let ajax = {
			id: id						
		  };
		  $.ajax({
			url: "g_BuscarBloque9.asp",
			type: "POST",
			cache: false,
			data: ajax,			
			dataType:"JSON",			
			success: function (data9, textStatus, jqXHR) {				
				console.log(data9);
				//debugger;				
				//					
				$("#Hogar").val(id);
				$("#Mascotas").val(data9[0].Id_Mascotas);
				
				//$("#Perro").val(data9[0].Ind_Perro);
				//$("#Gato").val(data9[0].Ind_Gato);
				//$("#Pez").val(data9[0].Ind_Pez);
				//$("#Ave").val(data9[0].Ind_Ave);
				//$("#Roedor").val(data9[0].Ind_Roedor);
				//$("#Otro").val(data9[0].Ind_Otro);
				let iperro = data9[0].Ind_Perro;
				let igato = data9[0].Ind_Gato;
				let ipez = data9[0].Ind_Pez;
				let iave = data9[0].Ind_Ave;
				let iroedor = data9[0].Ind_Roedor;
				let iotro = data9[0].Ind_Otro;
				if (iperro == 'True')
				{
					document.getElementById("Perro").checked = true;
				}
				if (igato == 'True')
				{
					document.getElementById("Gato").checked = true;
				}
				if (ipez == 'True')
				{
					document.getElementById("Pez").checked = true;
				}
				if (iave == 'True')
				{
					document.getElementById("Ave").checked = true;
				}
				if (iroedor  == 'True')
				{
					document.getElementById("Roedor").checked = true;
				}
				if (iotro  == 'True')
				{
					document.getElementById("Otro").checked = true;
				}
			},
			error: function (request, textStatus, errorThrown) {
			  alert("Error "+ request.responseJSON.message + "error");
			},
		  });
		  //buscar_panelistas();
		  document.getElementById('DivBuscarPanelistas').innerHTML = '';
		  buscar_clasesocial(id);
		  buscar_edadhogar(id);
		  //buscar_tiempohogarpanel();
		  CierraModal(); 
	}	
	
	//**Inicio Buscar Clase Social
	function buscar_clasesocial(id){
		//alert("Llego Clase Social");
		var mreg = id;
		//alert(mreg);
		$.ajax({
			url:'g_BuscarClaseSocial.asp?num='+mreg,
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivClaseSocial').html(data);
				
			}
		})
	}	
	//**Fin Buscar Clase Social
	
	//**Inicio Buscar Edad Hogar
	function buscar_edadhogar(id){
		//alert("Llego Edad Hogar");
		var mreg = id;
		//alert(mreg);
		$.ajax({
			url:'g_BuscarEdadHogar.asp?num='+mreg,
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#EdadHogar').html(data);
				
			}
		})
	}	
	//**Fin Buscar Edad Hogar


	
	//**Inicio Buscar Hogares
	function BuscarHogares(){
		//debugger;
		//alert("Llego Buscar Hogares");
		num = document.getElementById("BuscarHogar").value;
		//alert("Llego Buscar Hogares");
		var stodo = "num=" + num;
		//alert("Todo:=" + stodo);
		$.ajax({
			url:'g_BuscarHogares.asp?'+stodo,
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivBuscarHogares').html(data);
				//alert("Llego Ciudad2");
				
			}
		})
	}	
	//**Fin Registrar


	//**Inicio Cerrar Modal 
	//function CierraModal() {
		//return;
		//alert("paso");
		//window.open("http://localhost/atenas.pricetrack.com.ve/ph_mPanelHogares.asp?x=1&smenu=Hogares%20/%20Socio%20Demografico&edpas=1#close","_parent");
	//	window.open("http://atenas.pricetrack.com.ve/ph_mPanelHogares.asp?x=1&smenu=Hogares%20/%20Socio%20Demografico&edpas=1#close","_parent");
	//}	
	//**Fin Cerrar Modal

	//**Inicio Nuevo Hogar
	function NuevoHogar() {
	
	
		swal({
                title: "Desea Crear un Hogar Nuevo ?",
                text: "",
                type: "warning",
                showCancelButton: true,
                confirmButtonClass: "btn-primary",
                confirmButtonText: "Si",
                cancelButtonText: "No",
                closeOnConfirm: false,
                showLoaderOnConfirm: true
            },
            function() {
                //
                window.open("?edpas=1&smenu=Hogares%20/%20Socio%20Demografico","_parent");
                
                
            });
	
	
	
	
	
	/*
		var txt;
		var r = confirm("Desea Crear un Hogar Nuevo?");
		if (r == true) {
		//txt = "You pressed OK!";
		window.open("?edpas=1&smenu=Hogares%20/%20Socio%20Demografico","_parent");
		} else {
		//txt = "You pressed Cancel!";
		}
		*/
		return;
	}
	//**Fin Nuevo Pedido

	//**Inicio Cerrar Modal 
	function CierraModal() {
		//return;
		//alert("paso");
		//window.open("http://localhost/atenas.pricetrack.com.ve/ph_mPanelHogares.asp?x=1&smenu=Hogares%20/%20Socio%20Demografico&edpas=1#close","_parent");
		//window.open("http://atenas.pricetrack.com.ve/ph_mPanelHogares.asp?x=1&smenu=Hogares%20/%20Socio%20Demografico&edpas=1#close","_parent");
		//window.open("ph_mPanelHogares.asp?x=1&smenu=Hogares%20/%20Socio%20Demografico&edpas=1#close","_parent");
		window.location.href = "#close";
	}	
	//**Fin Cerrar Modal

	//**Inicio BlanquearCampos
	function BlanquearCampos() {
		//alert("llego a BlanquearCampos");
		//Bloque 0
		//30Dic
		document.getElementById("Estado").value = 0;
		document.getElementById("Ciudad").value = 0;
		document.getElementById("Municipio").value = 0;
		document.getElementById("Parroquia").value = 0;
		document.getElementById("Calle").value = "";
		document.getElementById("Edificio").value = "";
		document.getElementById("Casa").value = "";
		document.getElementById("Escalera").value = "";
		document.getElementById("Piso").value = "";
		document.getElementById("Apartamento").value = "";
		document.getElementById("Barrio").value = "";
		document.getElementById("Referencia").value = "";
		document.getElementById("TelefonoLocal").value = "";
		//Bloque 1
		debugger;
		document.getElementById("PrimerNombre").value = "";
		document.getElementById("SegundoNombre").value = "";
		document.getElementById("PrimerApellido").value = "";
		document.getElementById("SegundoApellido").value = ""; 
		document.getElementById("Nacionalidad").value = 0;
		document.getElementById("Cedula").value = "";
		document.getElementById("Celular").value = "";
		document.getElementById("CelularAdicional").value = "";
		document.getElementById("NumeroCortesia").value = "";
		document.getElementById("Correo").value = "";
		document.getElementById("CorreoAlterno").value = "";
		document.getElementById("Parentesco").value = 0;
		document.getElementById("EstadoCivil").value = 0;
		document.getElementById("FechaNacimiento").value = "";
		document.getElementById("Edad").value = "";
		document.getElementById("Sexo").value = 0;
		document.getElementById("Educacion").value = 0;
		document.getElementById("TipoIngreso").value = 0;
		document.getElementById("NumeroPersonas").value = 0;
		document.getElementById("FrecuenciaCompra").value = 0;
		document.getElementById("Titular").value = "";
		document.getElementById("CedulaTitular").value = "";
		document.getElementById("Banco").value = 0;
		document.getElementById("Cuenta").value = "";
		document.getElementById("PagoRapido").value = 0;
		//return;
		//Bloque 3
		document.getElementById("TipoVivienda").value = 0;
		document.getElementById("Explique").value = "";
		document.getElementById("MetrosVivienda").value = 0;
		document.getElementById("TotalAmbientes").value = "";
		document.getElementById("TotalBanos").value = "";
		document.getElementById("PuntosLuz").value = 0;
		//Bloque 4
		document.getElementById("OcupacionVivienda").value = 0;
		document.getElementById("ExpliqueOcupacion").value = "";
		document.getElementById("MontoVivienda").value = 0;
		//Bloque 5
		document.getElementById("AguasBlancas").value = 0;
		document.getElementById("AguasNegras").value = 0;
		document.getElementById("AseoUrbano").value = 0;
		document.getElementById("Electricidad").value = 0;
		document.getElementById("Telefono").value = 0;
		//Bloque 6
		document.getElementById("DomesticaFija").value = 0;
		document.getElementById("PersonalLabores").value = 0;
		document.getElementById("DomesticaDia").value = 0;
		document.getElementById("Conexion1").value = 0;
		document.getElementById("Conexion2").value = 0;
		document.getElementById("Conexion3").value = 0;
		document.getElementById("TelefonoCelular").value = 0;
		document.getElementById("Seguro1").value = 0;
		document.getElementById("Seguro2").value = 0;
		document.getElementById("Seguro3").value = 0;
		document.getElementById("Aire").value = 0;
		document.getElementById("CalentadorElectrico").value = 0;
		document.getElementById("CalentadorGas").value = 0;
		document.getElementById("ComputadorPc").value = 0;
		document.getElementById("ComputadorLaptop").value = 0;
		document.getElementById("Dvd").value = 0;
		document.getElementById("Home").value = 0;
		document.getElementById("Juego").value = 0;
		document.getElementById("Horno").value = 0;
		document.getElementById("Secadora").value = 0;
		document.getElementById("LavadoraAutomatica").value = 0;
		document.getElementById("LavadoraSemi").value = 0;
		document.getElementById("LavadoraRodillo").value = 0;
		document.getElementById("Nevera").value = 0;
		document.getElementById("Freezer").value = 0;
		document.getElementById("CocinaElectrica").value = 0;
		document.getElementById("CocinaBombona").value = 0;
		document.getElementById("CocinaGas").value = 0;
		document.getElementById("CocinaKerosene").value = 0;
		document.getElementById("Lavaplatos").value = 0;
		//Bloque 7
		document.getElementById("NumeroTeletisores").value = 0;
		document.getElementById("TipoTelevisores").value = 0;
		document.getElementById("Senal").value = 0;
		document.getElementById("Cableras1").value = 0;
		document.getElementById("Cableras2").value = 0;
		document.getElementById("TvOnline1").value = 0;
		document.getElementById("TvOnline2").value = 0;
		//Bloque 8
		document.getElementById("Autos").value = 0;
		document.getElementById("Moto").value = 0;
		document.getElementById("SeguroCasco").value = 0;
		//Bloque 9
		//debugger;
		document.getElementById("Mascotas").value = 0;
		//alert("llego1");
		
		//$("#Perro").checked = false; 		//No da error pero no lo destilda
		//$("#Perro").checked = 'false';	//No da error pero no lo destilda
		//$("#Perro").checked = falso;		//Da error
		//$("#Perro").checked = 'falso';	//No da error pero no lo destilda
		//$("#Perro").checked = 0;			//No da error pero no lo destilda
		//$("#Perro").checked = "0";		//No da error pero no lo destilda
		//$("#Perro").checked = null;		//No da error pero no lo destilda
		
		//$("Perro").prop("checked"),false;		//Da error
		//$("Perro").prop("checked"),falso;		//Da error
		//$("Perro").prop("checked") = falso;	//Da error
		//$("Perro").prop("checked") = false;	//Da error
		//$("Perro").prop("checked") = 0;		//Da error
		//$("Perro").prop("checked") = null;	//Da error
		
		//document.getElementById("Perro").checked = 0:			//Da error
		//document.getElementById("Perro").checked = false:		//Da error
		//document.getElementById("Perro").checked = 'false':	//Da error
		//document.getElementById("Perro").checked = 'falso':	//Da error
		//document.getElementById("Perro").checked = null:		//Da error
		
		//document.getElementById("Perro").value = false; 		//No da error pero no lo destilda
		//document.getElementById("Perro").value = 'false'; 	//No da error pero no lo destilda
		//document.getElementById("Perro").value = 'falso'; 	//No da error pero no lo destilda
		//document.getElementById("Perro").value = falso; 		//Da error
		//document.getElementById("Perro").value = null; 		//Da error
		
		//document.getElementById('Perro').style.display = "none"; //Lo elimina
		
		//document.getElementById('Perro').value = '';

		//document.querySelectorAll('.text input[name="Perro"]')[0].checked = false; document.querySelector('.text input[name="Perro"]').checked = false;	//Da error
		
		//$('.text input[name="Perro"]').prop('checked', false); //Da error
		
		document.getElementById("Perro").checked = false; 
		document.getElementById("Gato").checked = false; 
		document.getElementById("Pez").checked = false; 
		document.getElementById("Ave").checked = false; 
		document.getElementById("Roedor").checked = false; 
		document.getElementById("Otro").checked = false; 
		
		//alert("llego2");
		//document.getElementById("Perro").checked = 0:
		//document.getElementById("Gato").checked = 0:
		//document.getElementById("Pez").checked = 0:
		//document.getElementById("Ave").checked = 0:
		//document.getElementById("Roedor").checked = 0:
		//document.getElementById("Otro").checked = 0:
		
		
	}
	//**Fin BlanquearCampos

	
</script>