<!DOCTYPE HTML>
<html >
<head>
	<title>Home Pantry Mensual</title>
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
	<link rel="icon" href="favicon.ico" type="image/x-icon"> 
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<script type="text/javascript" src="js/sweetalert.min.js"></script>
	<link href="css/sweetalert.css" rel="stylesheet" type="text/css" media="screen" />	

</head>
<script type="text/javascript">
	function GenerarExcel()
	{
		//alert("Generar Excel");
		num = document.getElementById("Excel").value;
		//alert("Generar Excel:="+ num);
		window.open("g_CteHomePartyMenLRExcel.asp?" + num,"_blank");
	}

		function Mensaje(){
			swal("Atenas Grupo Consultor","Servicio No Contratado","info");
			return;
		}	

	//**Inicio Buscar Fabricante
	function buscar_fabricante(){
		//return;
		//alert("Llego Fabricante");
		$("#Fabricante").prop("selectedIndex", 0);
		categoria = document.getElementById("Cat").value;
		marca = document.getElementById("Mar").value;
		var stodo = "cat=" + categoria;
		stodo = stodo + "&mar=" + marca;
		
		$.ajax({
			url:'g_CteFabricantes.asp?'+stodo,
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivFabricante').html(data);
			}
		})
	}	
	//**Fin Buscar Fabricante

	//**Inicio Buscar Marca
	function buscar_marca(){
		//return;
		//alert("Llego Marca");
		$("#Marca").prop("selectedIndex", 0);
		categoria = document.getElementById("Cat").value;
		fabricante = document.getElementById("Fab").value;
		var stodo = "cat=" + categoria;
		stodo = stodo + "&fab=" + fabricante;
		//alert("stodo"+stodo);
		$.ajax({
			url:'g_CteMarca.asp?'+stodo,
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivMarca').html(data);
			}
		})
	}	
	//**Fin Buscar Marca


</script>
	
<body topmargin="0">
<!--#include file="estiloscss.asp"-->
<!--#include file="encabezado.asp"-->
<!--#include file="nn_subN.asp"-->
<!--#include file="in_DataEN.asp"-->
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
<%
'==========================================================================================
' Variables y Constantes
'==========================================================================================
    Apertura
	dim idCliente
	
	idCliente = Session("idCliente")
	

	dim idCategoria
	dim idFabricante
	dim idMarca
	dim idArea
	dim idRango
	dim idTamano
	dim strSemana

	dim gProductos
	dim gCategoria
	dim gArea
	dim gFabricante
	dim gMarca
	dim gSegmento
	dim gRango
	dim gTamano
	dim gIndicadores

			
	dim gDatos1
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	dim gDatos2
	dim rsx2
	set rsx2 = CreateObject("ADODB.Recordset")
	rsx2.CursorType = adOpenKeyset 
	rsx2.LockType = 2 'adLockOptimistic 

Sub Combos
 
	'response.write "<br>372 Combo1:=" & ed_sPar(1,0)
	'response.write " Combo2:=" & ed_sPar(2,0)
	'response.write " Combo3:=" & ed_sPar(3,0)
	'response.write " Combo3:=" & ed_sPar(4,0)
	'response.write " Combo3:=" & ed_sPar(5,0)
    ed_iCombo = 1
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " ss_ClienteCategoria.Id_Categoria, "
	sql = sql & " PH_CB_Categoria.Categoria "
	sql = sql & " FROM ss_ClienteCategoria INNER JOIN PH_CB_Categoria ON ss_ClienteCategoria.Id_Categoria = PH_CB_Categoria.id_Categoria "
	if idCliente <> 1 then
		sql = sql & " WHERE "
		sql = sql & " ss_ClienteCategoria.Id_Cliente = " & idCliente
		sql = sql & " and ss_ClienteCategoria.Ind_Activo = 1"
		sql = sql & " and ss_ClienteCategoria.Ind_Mensual = 1"
	end if
	sql = sql & " GROUP BY "
	sql = sql & " ss_ClienteCategoria.Id_Categoria, "
	sql = sql & " PH_CB_Categoria.Categoria "
	sql = sql & " ORDER BY "
	sql = sql & " PH_CB_Categoria.Categoria "
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(1,0)="Categoria"
    ed_sCombo(1,1)=sql 
    'ed_sCombo(1,2)="Sin Selección"
	
End Sub

Sub DataCombos
	'exit sub
	'response.write "<br>372 Combo1:=" & ed_sPar(1,0)
	if ed_sPar(1,0) = "" then
		sql = ""
		sql = sql & " SELECT "
		sql = sql & " ss_ClienteCategoria.Id_Categoria, "
		sql = sql & " PH_CB_Categoria.Categoria "
		sql = sql & " FROM ss_ClienteCategoria INNER JOIN PH_CB_Categoria ON ss_ClienteCategoria.Id_Categoria = PH_CB_Categoria.id_Categoria "
		if idCliente <> 1 then
			sql = sql & " WHERE "
			sql = sql & " ss_ClienteCategoria.Id_Cliente = " & idCliente
			sql = sql & " and ss_ClienteCategoria.Ind_Activo = 1"
		end if
		sql = sql & " GROUP BY "
		sql = sql & " ss_ClienteCategoria.Id_Categoria, "
		sql = sql & " PH_CB_Categoria.Categoria "
		sql = sql & " ORDER BY "
		sql = sql & " PH_CB_Categoria.Categoria "
		rsx1.Open sql ,conexion
		if rsx1.eof then
			rsx1.close
		else
			gCat = rsx1.GetRows
			rsx1.close
			ed_sPar(1,0) = gCat(0,0)
		end if
	end if
	
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " ss_ClienteCategoria.Id_Categoria, "
	sql = sql & " rtrim(PH_CB_Categoria.Categoria) "
	sql = sql & " FROM ss_ClienteCategoria INNER JOIN PH_CB_Categoria ON ss_ClienteCategoria.Id_Categoria = PH_CB_Categoria.id_Categoria "
	if idCliente <> 1 then
		sql = sql & " WHERE "
		sql = sql & " ss_ClienteCategoria.Id_Cliente = " & idCliente
		sql = sql & " and ss_ClienteCategoria.Ind_Activo = 1"
	end if
	sql = sql & " GROUP BY "
	sql = sql & " ss_ClienteCategoria.Id_Categoria, "
	sql = sql & " PH_CB_Categoria.Categoria "
	sql = sql & " ORDER BY "
	sql = sql & " PH_CB_Categoria.Categoria "
	'response.write "<br>372 Combo1:=" & sql
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gCategoria = rsx1.GetRows
		rsx1.close
	end if

	sql = "" 
	sql = sql & " SELECT "
	sql = sql & " Id_Area, "
	sql = sql & " rtrim(Area) "
	sql = sql & " FROM "
	sql = sql & " PH_DataCrudaMensual "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria = " & ed_sPar(1,0)
	sql = sql & " GROUP BY "
	sql = sql & " Id_Area, "
	sql = sql & " Area "
	sql = sql & " HAVING "
	sql = sql & " Id_Area <> 0 "
	sql = sql & " ORDER BY "
	sql = sql & " Area "
	'response.write "<br>372 Combo1:=" & sql
	'response.end
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gArea = rsx1.GetRows
		rsx1.close
	end if

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Fabricante, "
	sql = sql & " rtrim(Fabricante)"
	sql = sql & " FROM "
	sql = sql & " PH_DataCrudaMensual "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria = " & ed_sPar(1,0)
	sql = sql & " GROUP BY "
	sql = sql & " Id_Fabricante, "
	sql = sql & " Fabricante "
	sql = sql & " HAVING "
	sql = sql & " Id_Fabricante <> 0 "
	sql = sql & " ORDER BY "
	sql = sql & " Fabricante "
	'response.write "<br>372 sql1:=" & sql
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gFabricante = rsx1.GetRows
		rsx1.close
	end if

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Marca, "
	sql = sql & " rtrim(Marca) "
	sql = sql & " FROM "
	sql = sql & " PH_DataCrudaMensual "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria =  " & ed_sPar(1,0)
	sql = sql & " GROUP BY "
	sql = sql & " Id_Marca, "
	sql = sql & " Marca "
	sql = sql & " HAVING "
	sql = sql & " Id_Marca <> 0 "
	sql = sql & " ORDER BY "
	sql = sql & " Marca "
	'response.write "<br>372 sql2:=" & sql
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gMarca = rsx1.GetRows
		rsx1.close
	end if

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Segmento, "
	sql = sql & " rtrim(Segmento) "
	sql = sql & " FROM "
	sql = sql & " PH_DataCrudaMensual "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria =  " & ed_sPar(1,0)
	sql = sql & " GROUP BY "
	sql = sql & " Id_Segmento, "
	sql = sql & " Segmento "
	sql = sql & " HAVING "
	sql = sql & " Id_Segmento <> 0 "
	sql = sql & " ORDER BY "
	sql = sql & " Segmento "
	'response.write "<br>372 sql3:=" & sql
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gSegmento = rsx1.GetRows
		rsx1.close
	end if

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_RangoTamano, "
	sql = sql & " rtrim(RangoTamano) "
	sql = sql & " FROM "
	sql = sql & " PH_DataCrudaMensual "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria =  " & ed_sPar(1,0)
	sql = sql & " GROUP BY "
	sql = sql & " Id_RangoTamano, "
	sql = sql & " RangoTamano "
	sql = sql & " HAVING "
	sql = sql & " Id_RangoTamano <> 0 "
	sql = sql & " ORDER BY "
	sql = sql & " RangoTamano "
	'response.write "<br>372 sql4:=" & sql
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gRango = rsx1.GetRows
		rsx1.close
	end if

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Tamano, "
	sql = sql & " rtrim(Tamano) "
	sql = sql & " FROM "
	sql = sql & " PH_DataCrudaMensual "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria =  " & ed_sPar(1,0)
	sql = sql & " GROUP BY "
	sql = sql & " Id_Tamano, "
	sql = sql & " Tamano "
	sql = sql & " HAVING "
	sql = sql & " Id_Tamano <> 0 "
	sql = sql & " ORDER BY "
	sql = sql & " Tamano "
	'response.write "<br>372 sql4:=" & sql
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gTamano = rsx1.GetRows
		rsx1.close
	end if
	
	'response.write "<br>372 PASO" 
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Indicador, "
	sql = sql & " Abreviatura, "
	sql = sql & " Ind_Activo " 
	sql = sql & " FROM "
	sql = sql & " PH_Indicadores "
	sql = sql & " WHERE "
	sql = sql & " Ind_Atenas = 1 " 
	sql = sql & " ORDER BY "
	sql = sql & " Id_Indicador "
	'response.write "<br>372 sql5:=" & sql
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gIndicadores = rsx1.GetRows
		rsx1.close
	end if
	'response.write "<br>372 SALIO" 
End Sub

sub VerificarData
	exit sub
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Cliente, "
	sql = sql & " Ind_Mensual "
	sql = sql & " FROM "
	sql = sql & " ss_ClienteCategoria "
	sql = sql & " WHERE "
	sql = sql & " Id_Cliente = " & idCliente
	sql = sql & " AND Ind_Mensual = 1 "
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
		%>
		<script language="JavaScript" type="text/javascript">
			Mensaje()
		</script>
		<%
		response.end
	else
		rsx1.close
	end if

end sub
   
    LeePar
      
    if ed_iPas<>4 then 
        Encabezado
    end if
	VerificarData
	Combos
	DataCombos
	'response.write "<br>327 Cliente:= " & Session("idCliente") & "<br>"
	
%>
	
	<!--hidden-->
	
	<input type="hidden" name="Filtro" id="Filtro" align="right" size=200>
	<input type="hidden" name="Cliente" id="Cliente" align="right" size=4 value="<%=Session("idCliente")%>">
	<input type="hidden" name="Cat" id="Cat" align="right" size=4 value="<%=ed_sPar(1,0)%>">
	<input type="text" name="Fab" id="Fab" align="right" size=20 value="">	
	<input type="text" name="Mar" id="Mar" align="right" size=20 value="">
	
	<link rel="stylesheet" href="https://netdna.bootstrapcdn.com/bootstrap/3.3.2/css/bootstrap.min.css">
	<script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.3/jquery.min.js"></script>
	<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.2/js/bootstrap.min.js"></script>
	<link rel="stylesheet" href="css/bootstrap-multiselect.css" type="text/css">
	
	<!--=============================================================================================-->
	
	<link rel="stylesheet" href="css/homePantry.css" type="text/css">
	<script type="text/javascript" src="js/bootstrap-multiselect.js"></script>				
	<!--===============================================================================================-->
	<link rel="stylesheet" type="text/css" href="css/perfect-scrollbar.css">
	<link rel="stylesheet" type="text/css" href="css/util.css">
	<link rel="stylesheet" type="text/css" href="css/main1.css">	
	
	<div class="container-fluid" id="grad1">  
			
			<div class="col-sm-4">
											
				<div class="form-group">
					<!--Categoria-->	
					 <label for="categoria"><i class="fas fa-shapes"></i>&nbsp;Categoría:</label>
					<%
					ed_vCombo1
					%>
				</div>
				
				<div class="form-group">
					<!--area-->	
					 <label for="area"><i class="	fas fa-globe-americas"></i>&nbsp;Área:</label>
					 <select id="Area" multiple="multiple">
						<option value="0">TOTAL NACIONAL</option>
						<% for iAre = 0 to  ubound(gArea,2) %>
							<option value="<%=gArea(0,iAre)%>"><%=gArea(1,iAre)%></option>
						<% next %>
					 </select>							                              					
				</div>

				<div class="form-group">
					<!--Fabricante-->	
					 <label for="fabricante"><i class="fas fa-industry"></i>&nbsp;Fabricante:</label>
					 <select id="Fabricante" multiple="multiple">
						<!--<option value="0">TOTAL CATEGORIA</option>-->
						<% for iFra = 0 to  ubound(gFabricante,2) %>
							<option value="<%=gFabricante(0,iFra)%>"><%=gFabricante(1,iFra)%></option>
						<% next %>
					 </select>							                              					
				</div>
				

				<!--NUEVA MARCA-->
				<div id="DivMarca">
					<div class="form-group">
						<script>buscar_marca()</script>
					</div>	
				</div>
												
			</div>  <!-- class="col-sm-6"> -->
			
			<div class="col-sm-6">
			
				<div class="form-group">
					<!--Segmento-->
				 	<label for="segmento"><i class="fas fa-project-diagram"></i>&nbsp;Segmento:</label>
				 	<select id="Segmento" multiple="multiple">
						<%	for iSeg = 0 to  ubound(gSegmento,2) %>
							<option value="<%=gSegmento(0,iSeg)%>"><%=gSegmento(1,iSeg)%></option>
						<% next %>
					</select>			 
				</div>
				<div class="form-group">
					<!--Rango Tamaño-->
				 	<label for="rango"><i class="fas fa-gopuram"></i>&nbsp;Rango Tamaño:</label>
				 	<select id="Rango" multiple="multiple">
						<%	for iRan = 0 to  ubound(gRango,2) %>
							<option value="<%=gRango(0,iRan)%>"><%=gRango(1,iRan)%></option>
						<% next %>
					</select>			 
				</div>
				<%
				
				'idCliente = Session("idCliente")
				'response.write  idCliente
				if (cint(idCliente) = 3 or cint(idCliente) = 1) and ed_sPar(1,0) = 1 then
					
				%>

				<div class="form-group">
					<!--Tamaño-->
				 	<label for="tamano"><i class="fas fa-ruler-combined"></i>&nbsp;Tamaño:</label>
				 	<select id="Tamano" multiple="multiple">
						<%	
						
						for iTam = 0 to  ubound(gTamano,2) %>
							<option value="<%=gTamano(0,iTam)%>"><%=gTamano(1,iTam)%></option>
						<% 
						next 
						
						%>
					</select>			 
				</div>
				<%
				end if
				%>
				
				<div class="form-group">
					<!--Indicadores-->
				 	<label for="indicadores"><i class="fas fa-tachometer-alt"></i>&nbsp;Indicadores:</label>
				 	<select id="Indicadores" multiple="multiple">
						<%	for iInd = 0 to  ubound(gIndicadores,2) : sx = gIndicadores(1,iInd) %>
							<option value="<%=gIndicadores(0,iInd)%>"><%=sx%></option>
						<% next %>
					</select> 
				</div>
				
				<div class="form-group">
					
					<div class="col-sm-4">				
						<!--Borrar Filtros-->
						<button type="button" title="Borrar Pantalla"  class="btn btn-block btn-sm btn-danger" onclick="BorrarFiltros()">BORRAR FILTROS&nbsp;&nbsp;<i class="fas fa-times-circle"></i></button>
					</div>
					
					<div class="col-sm-4">				
						<!--Ejecutar-->
						<button type="button" title="Aplicar Selección" class="btn btn-block btn-sm btn-success" id="submit">APLICAR SELECCIÓN&nbsp;&nbsp;<i class="fas fa-check"></i></button>
						</div>
					
					<div class="col-sm-4">				
						<!--Exportar-->
						<button type="button" title="Exportar a Excel" class="btn btn-block btn-sm btn-primary" onclick="GenerarExcel();">EXPORTAR EXCEL&nbsp;&nbsp;<i class="fas fa-download"></i></button>
						<!--hidden-->
						<input type="hidden" name="Excel" id="Excel" align="right" size=0 value='<%=sExcel%>'>
					</div>
					
				</div>
							
			</div>  <!-- class="col-sm-6"> -->
			<div class="col-sm-2">
				<img alt="Logo de la Empresa" src="images/logo/LogoHomePantry.png" style = "width:128px;  " class="img-responsive center-block" >
			</div>
	
	</div> <!-- class="container-fluid" id=grad1 --> 
	
	<div class="container-fluid text-center text-primary" id="cargando" style="display:none;" >
		<span ><img src="images/ajax-loader7.gif"><strong>&nbsp;Procesando...., Espere!</strong></span>
	</div>
	<div id="DivHomePartyMen">
	</div>
	
	<% conexion.close %>
	
</body>
</html>
<!--================================================================================-->
<script src="https://kit.fontawesome.com/9d7cfbccc5.js" crossorigin="anonymous"></script>
<!--===============================================================================================-->
<script src="js/perfect-scrollbar.min.js"></script>
<script>
	$('.js-pscroll').each(function(){
		var ps = new PerfectScrollbar(this);
		$(window).on('resize', function(){
			ps.update();
		})
	});	
</script>
<script src="js/main.js"></script>
<!--===============================================================================================-->


<script type="text/javascript">
	$(document).ready(function() {
		//
		//debugger;
		//$('#Categoria').multiselect();
		$('#Area').multiselect();
		//$('#Fabricante').multiselect();
		$("#Fabricante").multiselect('destroy');
		$('#Fabricante').multiselect
			(
				{
				onChange: function(element, checked) 
					{
						var fabricante = $("#Fabricante :selected").map((_,e) => e.value).get();
						document.getElementById("Fab").value = fabricante;
						//alert("fabricante:" + fabricante);
						buscar_marca();
					}
				}
			);		
		$("#Marca").multiselect('destroy');
		$('#Marca').multiselect();
		$('#Segmento').multiselect();
		$('#Rango').multiselect();
		$('#Tamano').multiselect();
		$('#Indicadores').multiselect();
		
		
		$('#submit').click(function() {
			//debugger;
			var categ = document.getElementById("Cat").value;
			//alert(categ);
			$("#cargando").css("display", "block");		
			var area = $("#Area :selected").map((_,e) => e.value).get();
			//var categoria = $("#Categoria :selected").map((_,e) => e.value).get();
			var fabricante = $("#Fabricante :selected").map((_,e) => e.value).get();
			var marca = $("#Marca :selected").map((_,e) => e.value).get();
			var segmento = $("#Segmento :selected").map((_,e) => e.value).get();
			var rango = $("#Rango :selected").map((_,e) => e.value).get();
			var tamano = $("#Tamano :selected").map((_,e) => e.value).get();
			var indicadores = $("#Indicadores :selected").map((_,e) => e.value).get();
			
			
			//alert(categoria);
			//alert("fabricante:" + fabricante);
			//alert("marca:" + marca);
			//alert("segmento:" + segmento);
			//return;
			//alert(indicadores);
			//var stodo = "cat=" + categoria;
			//
			var stodo = "cat=" + categ;
			stodo = stodo + "&are=" + area;
			stodo = stodo + "&fab=" + fabricante;
			stodo = stodo + "&mar=" + marca;
			stodo = stodo + "&seg=" + segmento;
			stodo = stodo + "&ran=" + rango;
			stodo = stodo + "&ind=" + indicadores;
			//08Mar2021 - 1
			stodo = stodo + "&tam=" + tamano;
			document.getElementById("Filtro").value = "g_CteHomePartyMenLR.asp?" + stodo;
			document.getElementById("Excel").value = stodo;
			//return;
			$('#DivHomePartyMen').html("");
			$.ajax({
				url:'g_CteHomePartyMenLR.asp?'+stodo,
				beforeSend: function(objeto){
					//$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');						
				},
				success:function(data){
					//debugger;
					//$('#loader2').html('');
					console.log(data);
					$('#DivHomePartyMen').html(data);
					$("#cargando").css("display", "none");		
					//alert("Registrado");
					//swal("Datos de Identificacion del Hogar","Registrado","success");
				}
			})

		});
	});
	
	function BorrarFiltros() {
		swal({
                title: "Desea Borrar los Filtros ?",
                text: "",
                type: "warning",
                showCancelButton: true,
                confirmButtonClass: "btn-primary",
                confirmButtonText: "Si",
                cancelButtonText: "No",
                closeOnConfirm: true,
                showLoaderOnConfirm: true
            },
            function() {
                //
                window.open("?x=1&smenu=Reporte%20Mensual","_parent");				
				/*
				$("#Categoria").prop("selectedIndex", 0);
				$("#Fabricante").prop("selectedIndex", 0);
				$("#Marca").prop("selectedIndex", 0);
				$("#Segmento").prop("selectedIndex", 0);
				$("#Indicadores").prop("selectedIndex", 0);
				$('#DivHomePartyMen').html("");				
				*/
            });
		return;
	}

</script>
