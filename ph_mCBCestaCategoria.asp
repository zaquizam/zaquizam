<!DOCTYPE HTML>
<html >
<head>
	<title>Cesta Categoria</title>
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
	<link rel="icon" href="favicon.ico" type="image/x-icon"> 
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

		//**Inicio Generar PDF
		function GenerarExcel(){
			//alert("Bus:= "+ document.getElementById("Bus").value );
			//alert("Buscar:= "+ document.getElementById("Excel").value );
			var sBus = document.getElementById("Excel").value
			window.open('Sys_mUsuarioExcel.asp?bus='+sBus,'_blank');
		}	
		//**Fin Generar PDF
	
	</script>   
<%

   
'==========================================================================================
' Parámetros del Manteniemiento
'==========================================================================================
Sub ParDat
	if ed_sPar(1,0) = "" or ed_sPar(1,0) = "Seleccionar" then
		ed_sPar(1,0) = 6
	end if
	ed_Bot(4)="disabled"
	ed_Bot(1)="disabled"
	ed_iNumCam		=18					' Numero de campos en la pantalla principal
	ed_iRegPag		=25					' Numero de registros por pagina
	
	ed_sNomTab		="PH_CB_CestaCategoria"
	ed_sNomInd		="Id_CestaCategoria"
	ed_cCol		=1	' Columna a Ordenar
	ed_cOrd		=0	' Orden 0=ascendente 1=descendente
	ed_iRan		=0	' Presentar ranking de columnas
	ed_iRep=0
'ed_ides=1
	SqlCla = " SELECT * FROM "  & ed_sNomTab
	sqlcla = sqlcla & " WHERE  (fec_inactivo is null)"
	sqlcla = sqlcla & " and id_Cesta = " & ed_sPar(1,0)
	'response.write "<br>47 Perfil:= " & Session("idPerfil")
	

' Titulo	
   ed_sCampo(00,0)="#"
   ed_sCampo(01,0)="Cesta"
   ed_sCampo(02,0)="Categoria"
   ed_sCampo(03,0)="Activo?"
   'ed_sCampo(04,0)="Desde"
   'ed_sCampo(05,0)="Hasta"
   'ed_sCampo(06,0)="Ano"
   'ed_sCampo(07,0)="Mes"
   

   ' Tools Tip
    'ed_sCampo(08,5)="000000=NINGUNA / 000001=Caracas"
	
' Valor Default
	'ed_sCampo(04,1)=1
	'ed_sCampo(06,1)=1
	'ed_sCampo(08,1)="000001"
	'ed_sCampo(14,1)=0
	'ed_sCampo(17,1)=2
	'ed_sCampo(18,1)=3
'	ed_sCampo(13,1)=false
'	ed_sCampo(14,1)=false	
'	ed_sCampo(15,1)=false		
'	ed_sCampo(17,1)=false

' Obligatorio
	'ed_sCampo(01,4)=1
	'ed_sCampo(02,4)=1
	'ed_sCampo(03,4)=1
	'ed_sCampo(05,4)=1
	'ed_sCampo(07,4)=1
	'ed_sCampo(08,4)=1
	
' No Presentar	
	'ed_sCampo(08,2)="1"
	'ed_sCampo(04,2)="1"
	'ed_sCampo(06,2)="1"
	'ed_sCampo(08,2)="1"
	'ed_sCampo(09,2)="1"
	'ed_sCampo(10,2)="1"
	'ed_sCampo(11,2)="1"	
	'ed_sCampo(12,2)="1"
    'ed_sCampo(13,2)="1"
	'ed_sCampo(14,2)="1"
	'ed_sCampo(15,2)="1"
	'ed_sCampo(16,2)="1"
	'ed_sCampo(17,2)="1"
	'ed_sCampo(18,2)="1"
  '  ed_sCampo(08,2)="1"
  '  ed_sCampo(10,2)="1"
  '  ed_sCampo(12,2)="1"
'	ed_sCampo(13,2)="1"
'	ed_sCampo(14,2)="1"
	
' Copiar
 '   ed_sCampo(3,8)=2
  '  ed_sCampo(3,8)=2
     
	
	ed_sQue(1,0)=  " SELECT Id_Cesta, Cesta FROM  PH_CB_Cesta WHERE Fec_Inactivo is Null and id_Cesta = " & ed_sPar(1,0) & " order by Cesta "
	ed_sQue(2,0)=  " SELECT Id_Categoria, Categoria FROM  PH_CB_Categoria WHERE Fec_Inactivo is Null order by Categoria "
	'ed_sQue(5,0)=  " SELECT Id_PerfilUsuario, PerfilUsuario FROM  ss_PerfilUsuario WHERE Fec_Inactivo is Null And Id_PerfilUsuario >1"
	'ed_sQue(6,0)=  " SELECT id_Cliente, Cliente FROM  Syn_Cliente WHERE Fec_Inactivo is Null"

	ed_Formato(00,0)="w3-col l1  w3-left w3-padding "
	ed_Formato(01,0)="w3-col l3  w3-left w3-padding "
	ed_Formato(02,0)="w3-col l3  w3-left w3-padding "
	ed_Formato(03,0)="w3-col l1  w3-left w3-padding "
End Sub
Sub Combos
 
	'response.write "<br>372 Combo1:=" & ed_sPar(1,0)
	'response.write " Combo2:=" & ed_sPar(2,0)
	'response.write " Combo3:=" & ed_sPar(3,0)
	'response.write " Combo3:=" & ed_sPar(4,0)
	'response.write " Combo3:=" & ed_sPar(5,0)
    ed_iCombo = 1
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Cesta, "
	sql = sql & " Cesta "
	sql = sql & " FROM PH_CB_Cesta "
	sql = sql & " Order By "
	sql = sql & " Cesta "
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(1,0)="Cesta"
    ed_sCombo(1,1)=sql 
    'ed_sCombo(1,2)="Seleccionar"
	'ed_sCombo(1,2)=""
	if ed_sPar(1,0) = "" or ed_sPar(1,0) = "Seleccionar" then
		ed_sPar(1,0) = 6
	end if
	
End Sub
    LeePar
  
    ParDat
    
    if ed_iPas<>4 then 
        Encabezado
    end if    
	sExcel = request.Form("bus")

	'response.write "llego1"
	'response.end
	Combos

%>
	<div style="width:98%">
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
	<div style="width:98%"><%ed_Main %></div></center>

    <%conexion.close%>
	


</body>
</html>