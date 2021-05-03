<!DOCTYPE HTML>
<html >
<head>
	<title>Enviar a Movil</title>
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="mensaje.css" rel="stylesheet" type="text/css" media="screen" />
	<link rel="icon" href="favicon.ico" type="image/x-icon"> 
	<script type="text/javascript" src="js/sweetalert.min.js"></script>
	<script type="text/javascript" src="js/jquery-1.12.4.min.js"></script>

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
	   
	dim envMensaje
	dim envCelular
%>
<script type="text/javascript">
	function alerta(total) 
	{
		swal("Se Enviaron Encuestas " + total + " Hogares ","Enviado","success");
		//window.open("?edpas=1&smenu=?x=1&smenu=Envio%20SMS%20Bienvenida%20y%20Link","_parent");
	}
</script>
<%

Sub Combos
 
	'response.write "<br>372 Combo1:=" & ed_sPar(1,0)
	'response.write " Combo2:=" & ed_sPar(2,0)
	'response.write " Combo3:=" & ed_sPar(3,0)
	'response.write " Combo3:=" & ed_sPar(4,0)
	'response.write " Combo3:=" & ed_sPar(5,0)
    ed_iCombo = 2 
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_EncuestaEspecial, "
	sql = sql & " EncuestaEspecial "
	sql = sql & " FROM PH_EncuestaEspecial "
	sql = sql & " WHERE "
	sql = sql & " Ind_Activo = 1 "
	sql = sql & " Order By "
	sql = sql & " EncuestaEspecial "
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(1,0)="Encuesta Especial"
    ed_sCombo(1,1)=sql 
    ed_sCombo(1,2)="Seleccionar"


	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_SiNo, "
	sql = sql & " SiNo "
	sql = sql & " FROM PH_SiNo "
	sql = sql & " Order By "
	sql = sql & " Id_SiNo "
	'response.write "<br>372 Combo2:=" & sql
    ed_sCombo(2,0)="Enviar"
    ed_sCombo(2,1)=sql 
    ed_sCombo(2,2)="Seleccionar"
	
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
	'ParDat
%>
		
	<br>
	<br>
	<br>
	<div style="width:98%"></div></center>
<%
	Combos
	'response.write "paso"

%>

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

<%
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	dim gDatosSol1
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	dim gDatosSol2
	dim rsx2
	set rsx2 = CreateObject("ADODB.Recordset")
	rsx2.CursorType = adOpenKeyset 
	rsx2.LockType = 2 'adLockOptimistic 


	if ed_sPar(1,0) <> "Seleccionar" and ed_sPar(2,0) <> "Seleccionar"  then
		if cint(ed_sPar(2,0)) = 1 then
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Id_PanelHogar "
			sql = sql & " FROM "
			sql = sql & " PH_PanelHogar "
			sql = sql & " WHERE "
			sql = sql & " Ind_Activo = 1 "
			sql = sql & " and Id_PanelHogar > 1 "
			'response.write "<br>36 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			if rsx1.eof then
			else
				gDatosSol1 = rsx1.GetRows
				rsx1.close
				'Response.write "<br>" & cint(ed_sPar(1,0))
				Encuesta = cint(ed_sPar(1,0))
				for iReg = 0 to ubound(gDatosSol1,2)
					'Response.write "<br>" & gDatosSol1(0,iReg) 
					Hogar = cint(gDatosSol1(0,iReg))
					sql = ""
					sql = sql & " SELECT "
					sql = sql & " Id_EncuestaEspecial, "
					sql = sql & " Id_Hogar "
					sql = sql & " FROM "
					sql = sql & " PH_EncuestaHogar "
					sql = sql & " WHERE "
					sql = sql & " Id_EncuestaEspecial = " & Encuesta
					sql = sql & " and Id_Hogar  = " & Hogar
					'response.write "<br>36 sql:=" & sql
					'response.end
					rsx2.Open sql ,conexion
					
					if rsx2.eof then
						'response.write "<br> no  Existe"
						rsx2.close
						dim rsx3
						set rsx3 = CreateObject("ADODB.Recordset")
						rsx3.CursorType = 0
						rsx3.LockType = 3
						sql = ""
						sql = sql & " SELECT * "
						sql = sql & " FROM "
						sql = sql & " PH_EncuestaHogar "
						'response.write "<br>57 sql:=" & sql
						'response.end
						rsx3.Open sql ,conexion
						rsx3.addNew
						rsx3("Id_EncuestaEspecial") = Encuesta
						rsx3("Id_Hogar") = Hogar
						rsx3("Ind_Activo") = 1
						rsx3.Update
						rsx3.Close 
						set rsx3 = nothing 
					else
						'response.write "<br> si  Existe"
						rsx2.close
					end if
					
				next 
			end if
			response.write "<br> Encuesta Enviada al Movil"
			
			
		end if
	end if
	'response.end 
%>
<style>


</style>	

    <%conexion.close%>
	
</body>
</html>