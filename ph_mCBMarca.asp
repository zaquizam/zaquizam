<!DOCTYPE HTML>
<html >
<head>
	<title>Marcas</title>
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
	'exit sub
	ed_Bot(4)="disabled"
	'ed_Bot(1)="disabled"
	ed_iNumCam		=18					' Numero de campos en la pantalla principal
	ed_iRegPag		=25					' Numero de registros por pagina
	
	ed_sNomTab		="PH_CB_Marca"
	ed_sNomInd		="Id_Marca"
	ed_cCol		=2	' Columna a Ordenar
	ed_cOrd		=0	' Orden 0=ascendente 1=descendente
	ed_iRan		=0	' Presentar ranking de columnas
	ed_iRep=0
'ed_ides=1
	SqlCla = " SELECT * FROM "  & ed_sNomTab
	sqlcla = sqlcla & " WHERE  (fec_inactivo is null)"
	if ed_sPar(1,0) <> "" and ed_sPar(1,0) <> "Seleccionar" then
		sqlcla = sqlcla & " and id_categoria = " & ed_sPar(1,0)
	end if
	'response.write "<br>47 Categoria:= " & ed_sPar(1,0) 
	'response.end
	

' Titulo	
   ed_sCampo(00,0)="#"
   ed_sCampo(01,0)="Categoria"
   ed_sCampo(02,0)="Fabricante"
   ed_sCampo(03,0)="Marca"
   ed_sCampo(04,0)="Abreviatura"
   ed_sCampo(05,0)="Activo?"
   ed_sCampo(06,0)="Registrar Marca en Consumo"
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
	ed_sCampo(01,4)=1
	ed_sCampo(02,4)=1
	ed_sCampo(03,4)=1
	ed_sCampo(05,4)=1
	if ed_sPar(1,0) <> "" and ed_sPar(1,0) <> "Seleccionar" then
		if cint(ed_sPar(1,0)) = 7 or cint(ed_sPar(1,0)) = 93 then
			'ed_sCampo(06,4)=1
		end if
	end if
	ed_sCampo(07,0)="Medicina?"
	'ed_sCampo(06,4)=1
	'ed_sCampo(08,4)=1
	
' No Presentar	
	if ed_sPar(1,0) <> "" and ed_sPar(1,0) <> "Seleccionar" then
		if cint(ed_sPar(1,0)) <> 7 and cint(ed_sPar(1,0)) <> 93 then
			ed_sCampo(06,2)="1"
			'response.write "<br>115 Pasoooo"
		else
		end if
	end if
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
    
	
	ed_sQue(1,0)=  " SELECT Id_Categoria, Categoria FROM  PH_CB_Categoria WHERE Fec_Inactivo is Null and id_Categoria = " & ed_sPar(1,0)
	ed_sQue(2,0)=  " SELECT Id_Fabricante, Fabricante FROM  PH_CB_Fabricante WHERE Fec_Inactivo is Null and id_categoria = " & ed_sPar(1,0) & " Order by Fabricante "
	'ed_sQue(5,0)=  " SELECT Id_PerfilUsuario, PerfilUsuario FROM  ss_PerfilUsuario WHERE Fec_Inactivo is Null And Id_PerfilUsuario >1"
	'ed_sQue(6,0)=  " SELECT id_Cliente, Cliente FROM  Syn_Cliente WHERE Fec_Inactivo is Null"

	ed_Formato(00,0)="w3-col l1  w3-left w3-padding "
	ed_Formato(01,0)="w3-col l3  w3-left w3-padding "
	ed_Formato(02,0)="w3-col l3  w3-left w3-padding "
	ed_Formato(03,0)="w3-col l3  w3-left w3-padding "
	ed_Formato(04,0)="w3-col l1  w3-left w3-padding "
	ed_Formato(05,0)="w3-col l1  w3-left w3-padding "
	
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
	sql = sql & " Id_Categoria, "
	sql = sql & " Categoria "
	sql = sql & " FROM PH_CB_Categoria "
	sql = sql & " Order By "
	sql = sql & " Categoria "
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(1,0)="Categoria"
    ed_sCombo(1,1)=sql 
    ed_sCombo(1,2)="Seleccionar"
	
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
		
	<br>
	<div style="width:98%"><%ed_Main %></div></center>

    <%conexion.close%>
	


</body>
</html>