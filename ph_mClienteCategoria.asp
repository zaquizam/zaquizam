<!DOCTYPE HTML>
<html >
<head>
	<title>Cliente Categoria</title>
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
	if ed_sPar(1,0) = "" then ed_sPar(1,0) = 0
	ed_Bot(4)="disabled"
	ed_Bot(1)="disabled"
	ed_iNumCam		=18					' Numero de campos en la pantalla principal
	ed_iRegPag		=25					' Numero de registros por pagina
	
	ed_sNomTab		="ss_ClienteCategoria"
	ed_sNomInd		="Id_ClienteCategoria"
	ed_cCol		=1	' Columna a Ordenar
	ed_cOrd		=0	' Orden 0=ascendente 1=descendente
	ed_iRan		=0	' Presentar ranking de columnas
	ed_iRep=0
'ed_ides=1
	SqlCla = " SELECT * FROM "  & ed_sNomTab
	sqlcla = sqlcla & " WHERE  (fec_inactivo is null)"
	sqlcla = sqlcla & " And id_Cliente = " & int(ed_sPar(1,0))	
	'response.write "<br>47 Perfil:= " & Session("idPerfil")
	

' Titulo	
   ed_sCampo(00,0)="#"
   ed_sCampo(01,0)="Cliente"
   ed_sCampo(02,0)="Categoria"
   ed_sCampo(03,0)="Semanal?"
   ed_sCampo(04,0)="Mensual?"
   ed_sCampo(05,0)="Semana Desde"
   ed_sCampo(06,0)="Semana Hasta"
   ed_sCampo(07,0)="Semana Publicada"
   ed_sCampo(08,0)="Mes Desde"
   ed_sCampo(09,0)="Mes Hasta"
   ed_sCampo(10,0)="Mes Publicado"
   ed_sCampo(11,0)="Correo Notificacion"
   ed_sCampo(12,0)="Activa?"
   

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
	ed_sCampo(11,2)="2"	
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
    'ed_sCampo(1,8)=2
    'ed_sCampo(3,8)=2
    
	
	ed_sQue(1,0)=  " SELECT Id_Cliente, Cliente FROM  ss_Cliente WHERE Fec_Inactivo is Null " & " And id_Cliente = " & int(ed_sPar(1,0)) & " order by Cliente "
	ed_sQue(2,0)=  " SELECT Id_Categoria, Categoria FROM  PH_CB_Categoria WHERE Fec_Inactivo is Null order by Categoria"
	ed_sQue(5,0)=  " SELECT idSemana, Semana FROM ss_Semana WHERE Fec_Inactivo is Null order by idSemana desc"
	ed_sQue(6,0)=  " SELECT idSemana, Semana FROM ss_Semana WHERE Fec_Inactivo is Null order by idSemana desc"
	ed_sQue(7,0)=  " SELECT idSemana, Semana FROM ss_Semana WHERE Fec_Inactivo is Null order by idSemana desc"
	ed_sQue(8,0)=  " SELECT idPeriodo, Periodo FROM ss_Periodo WHERE Fec_Inactivo is Null order by idPeriodo desc"
	ed_sQue(9,0)=  " SELECT idPeriodo, Periodo FROM ss_Periodo WHERE Fec_Inactivo is Null order by idPeriodo desc"
	ed_sQue(10,0)=  " SELECT idPeriodo, Periodo FROM ss_Periodo WHERE Fec_Inactivo is Null order by idPeriodo desc"

	ed_Formato(00,0)="w3-col l1  w3-left w3-padding "
	ed_Formato(01,0)="w3-col l2  w3-left w3-padding "
	ed_Formato(02,0)="w3-col l2  w3-left w3-padding "
	ed_Formato(03,0)="w3-col l1  w3-left w3-padding "
	ed_Formato(04,0)="w3-col l1  w3-left w3-padding "
	ed_Formato(05,0)="w3-col l3  w3-left w3-padding "
	ed_Formato(06,0)="w3-col l3  w3-left w3-padding "
	ed_Formato(07,0)="w3-col l3  w3-left w3-padding "
	ed_Formato(08,0)="w3-col l2  w3-left w3-padding "
	ed_Formato(09,0)="w3-col l2  w3-left w3-padding "
	ed_Formato(10,0)="w3-col l2  w3-left w3-padding "
	ed_Formato(11,0)="w3-col l10  w3-left w3-padding "
	ed_Formato(12,0)="w3-col l1  w3-left w3-padding "
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
	sql = sql & " ss_Cliente.Id_Cliente, "
	sql = sql & " ss_Cliente.Cliente "
	sql = sql & " FROM ss_ClienteCategoria RIGHT JOIN ss_Cliente ON ss_ClienteCategoria.Id_Cliente = ss_Cliente.Id_Cliente "
	sql = sql & " GROUP BY "
	sql = sql & " ss_Cliente.Id_Cliente, "
	sql = sql & " ss_Cliente.Cliente "
	sql = sql & " ORDER BY "
	sql = sql & " ss_Cliente.Cliente "
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(1,0)="Cliente"
    ed_sCombo(1,1)=sql 
    ed_sCombo(1,2)="Seleccionar"
	
End Sub


    LeePar
  
    ParDat
    
    if ed_iPas<>4 then 
        Encabezado
    end if    
	sExcel = request.Form("bus")

	'if ed_sPar(1,0) = "" then ed_sPar(1,0) = ""
	'response.write "llego1"
	'response.end
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

<%
    

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