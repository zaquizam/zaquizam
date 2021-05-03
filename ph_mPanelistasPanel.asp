<!DOCTYPE HTML>
<html >
<head>
	<title>Panelistas/Panel</title>
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
	ed_Bot(4)="disabled"
	ed_Bot(1)="disabled"
	ed_iNumCam		=18					' Numero de campos en la pantalla principal
	ed_iRegPag		=25					' Numero de registros por pagina
	
	ed_sNomTab		="PH_PanelistaPanel"
	ed_sNomInd		="Id_PanelistaPanel"
	ed_cCol		=1	' Columna a Ordenar
	ed_cOrd		=0	' Orden 0=ascendente 1=descendente
	ed_iRan		=0	' Presentar ranking de columnas
	ed_iRep=0
'ed_ides=1
	SqlCla = " SELECT * FROM "  & ed_sNomTab
	sqlcla = sqlcla & " WHERE  (fec_inactivo is null)"
	'response.write "<br>47 Perfil:= " & Session("idPerfil")
	

' Titulo	
   ed_sCampo(00,0)="#"
   ed_sCampo(01,0)="Panelista"
   ed_sCampo(02,0)="Panel"
   'ed_sCampo(02,0)="1er Nombre"
   'ed_sCampo(03,0)="2do Nombre"
   'ed_sCampo(04,0)="1er Apellido"
   'ed_sCampo(05,0)="2do Apellido"
   'ed_sCampo(06,0)="Cedula Identidad"
   'ed_sCampo(07,0)="Fec. Nacimiento"
   'ed_sCampo(08,0)="Sexo"
   'ed_sCampo(09,0)="Profesion"
   'ed_sCampo(10,0)="Correo"
   'ed_sCampo(11,0)="Activo?"
   'ed_sCampo(12,0)="Filtro1"
   'ed_sCampo(13,0)="Filtro2"
   'ed_sCampo(14,0)="Filtro3"
   'ed_sCampo(15,0)="Filtro4"
   

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
	'ed_sCampo(12,2)="1"
	'ed_sCampo(13,2)="1"
	'ed_sCampo(14,2)="1"
	'ed_sCampo(15,2)="1"
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
    
	sql = ""
	sql = sql & " Select "
	sql = sql & " Id_Panelista, "
	sql = sql & " Nombre1+' '+Nombre2+' '+Apellido1+' '+Apellido2 as Panelista "
	sql = sql & " FROM  PH_Panelistas "
	sql = sql & " WHERE Fec_Inactivo is Null "
	ed_sQue(1,0)=  sql
	ed_sQue(2,0)=  " SELECT Id_Panel, Panel FROM  PH_Paneles WHERE Fec_Inactivo is Null "

	ed_Formato(00,0)="w3-col l1  w3-left w3-padding "
	ed_Formato(01,0)="w3-col l3  w3-left w3-padding "
	ed_Formato(02,0)="w3-col l3  w3-left w3-padding "
	ed_Formato(03,0)="w3-col l1  w3-left w3-padding "
	'ed_Formato(04,0)="w3-col l2  w3-left w3-padding "
	'ed_Formato(05,0)="w3-col l2  w3-left w3-padding "
	'ed_Formato(06,0)="w3-col l2  w3-left w3-padding "
	'ed_Formato(07,0)="w3-col l2  w3-left w3-padding "
	'ed_Formato(08,0)="w3-col l2  w3-left w3-padding "
	'ed_Formato(09,0)="w3-col l3  w3-left w3-padding "
	'ed_Formato(10,0)="w3-col l3  w3-left w3-padding "
	'ed_Formato(11,0)="w3-col l1  w3-left w3-padding "
	'ed_Formato(12,0)="w3-col l2  w3-left w3-padding "
	'ed_Formato(13,0)="w3-col l2  w3-left w3-padding "
	'ed_Formato(14,0)="w3-col l2  w3-left w3-padding "
	'ed_Formato(15,0)="w3-col l2  w3-left w3-padding "
End Sub
    LeePar
  
    ParDat
    
    if ed_iPas<>4 then 
        Encabezado
    end if    
	sExcel = request.Form("bus")

	'response.write "llego1"
	'response.end
    

%>
	<div style="width:98%">
	<div class="container-fluid">        
		<div class="row">
			<!--Contenido General-->			
			<div class="container">
				<div class="col-md-8 col-sm-8 col-xs-12">
					<div class="pull-right">
						<!--<img src="images/Excel.png"  style="margin-left:0px;" title="Generar Excel" alt="PDF" width="70px" onclick="GenerarExcel()"/>
						<input type="hidden" name="Excel" id="Excel" align="right" size=0 value='<%=sExcel%>'>-->
					</div>
				</div>
			</div>
		</div>
	</div>
		
	<br>
	<div style="width:98%"><%ed_Main %></div></center>

    <%conexion.close%>
	


</body>
</html>