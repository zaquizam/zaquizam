<!DOCTYPE HTML>
<html >
<head>
	<title>A donde se fue mi Marca</title>
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
	<script type="text/javascript" src="js/sweetalert.min.js"></script>
	<link rel="icon" href="favicon.ico" type="image/x-icon"> 
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
</head>
<body topmargin="0">
<!--#include file="estiloscss.asp"-->
<!--#include file="encabezado.asp"-->
<!--#include file="nn_subN.asp"-->
<!--#include file="in_DataEN.asp"-->

<%

%>
	<script>
	//**Inicio Procesar
	function Procesar(){
		debugger; 
		Cat =document.getElementById("Categoria").value;
		Mes =document.getElementById("Mes").value;
		Marca =document.getElementById("Marca").value;
		//alert("Llego Cat:=" + Cat);
		//alert("Mes:=" + Mes);
		//alert("Marca:=" + Marca);
		//return;
		var stodo = "cat=" + Cat + "&mes=" + Mes + "&mar=" + Marca;
		document.getElementById("Programa").value = "g_BuscarDondeMarca.asp?" + stodo;
		//return;
		
		swal({
                title: "Desea Buscar a Donde se fue esta Marca?",
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

				$.ajax({
					url:'g_BuscarDondeMarca.asp?'+stodo,
					beforeSend: function(objeto){
						$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
						
					},
					success:function(data){
						//debugger;
						$('#loader2').html('');
						console.log(data); 
						$('#DivData').html(data);
						swal("Busqueda Realizada","Proceso Culminado","success");
						//window.location.reload();
					}
				})
                
                
            });
	}	
	//**Fin Procesar




	</script>   
<%
  
'==========================================================================================
' Variables y Constantes
'==========================================================================================


    Apertura
	
	
	dim gDatosSol
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	dim gDatosSol2
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
    ed_iCombo = 3
	'if ed_sPar(1,0) <> "" and ed_sPar(1,0) <> "Seleccionar" then ed_sPar(1,0) = 0

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " ss_Periodo.idPeriodo, "
	sql = sql & " ss_Periodo.Periodo "
	sql = sql & " FROM ss_Periodo "
	sql = sql & " WHERE "
	sql = sql & " ss_Periodo.Semanas Is Not Null "
	sql = sql & " and ss_Periodo.idPeriodo > 24253 "
	sql = sql & " ORDER BY "
	sql = sql & " ss_Periodo.IdPeriodo DESC "
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(1,0)="Mes"
    ed_sCombo(1,1)=sql 
    ed_sCombo(1,2)="Seleccionar"
 
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Categoria, "
	sql = sql & " Categoria "
	sql = sql & " FROM "
	sql = sql & " PH_DataCruda "
	sql = sql & " GROUP BY "
	sql = sql & " Id_Categoria, "
	sql = sql & " Categoria "
	sql = sql & " Order by Categoria "
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(2,0)="Categoria"
    ed_sCombo(2,1)=sql 
    ed_sCombo(2,2)="Seleccionar"

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Marca, "
	sql = sql & " Marca "
	sql = sql & " FROM "
	sql = sql & " PH_DataCruda "
	if ed_sPar(2,0) <> "" and ed_sPar(2,0) <> "Seleccionar" then
		sql = sql & " WHERE "
		sql = sql & " Id_Categoria = " & cint(ed_sPar(2,0))
	else
		ed_iCombo = 2
	end if
	sql = sql & " GROUP BY "
	sql = sql & " Id_Marca, "
	sql = sql & " Marca "
	sql = sql & " HAVING "
	sql = sql & " Id_Marca<>0 "
	sql = sql & " ORDER BY "
	sql = sql & " Marca "
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(3,0)="Marca"
    ed_sCombo(3,1)=sql 
    ed_sCombo(3,2)="Seleccionar"
	
End Sub
	

   
'==========================================================================================
' Parámetros del Manteniemiento
'==========================================================================================
    LeePar
  
    if ed_iPas<>4 then 
        Encabezado
    end if    

	'response.write "llego1"
	'response.end
	'if ed_sPar(1,0) = "" or ed_sPar(1,0) = "Seleccionar" then ed_sPar(1,0) = 17
    Combos
%>
		
	<br>
	<div style="width:98%">
	<%
	
	%></div></center>
	<table border="0" align="right">
		<tr>
			<td>
				<%
				ed_vCombo
				%>
			</td>
		</tr>
	</table>
	</br>
	</br>
	</br>
	</br>
	</br>
	<%
	idMes = trim(ed_sPar(1,0))
	idCategoria = trim(ed_sPar(2,0))
	idMarca = trim(ed_sPar(3,0))
	'response.write "<br> Combo1:=Mes"  & "==>" & idMes
	'response.write "<br> Combo2:=Categoria" & "==>" & idCategoria
	'response.write "<br> Combo3:=Marca" & "==>" & idMarca
	'hidden 
	if idMes <> "Seleccionar" and  idCategoria <> "Seleccionar" and  idMarca <> "Seleccionar" then
		
		'sPro=Request.ServerVariables("HTTP_REFERER")
		'response.write "pro:=" & sPro
		'hidden
		'response.write "pro:=" & sPar
		%>
		<input type="hidden" name="Mes" id="Mes" value="<%=idMes%>" align="right" size=10>
		<input type="hidden" name="Categoria" id="Categoria" value="<%=idCategoria%>" align="right" size=10>
		<input type="hidden" name="Marca" id="Marca" value="<%=idMarca%>" align="right" size=10>
		<input type="hidden" name="Programa" id="Programa" value="" align="right" size=50>
		</br>
		</br>
			<center>
			<img src="images/Procesamiento.jpg"  style="margin-left:0px;" alt="Procesar" width="80px"' onclick="Procesar()"/>
			</center>
		</br>
		<div id="DivData"> 
		</div>
		<%
	end if
	
	conexion.close
	%>

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

</body>
</html>