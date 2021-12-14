<!DOCTYPE HTML>
<html >
<head>
	<title>Producto Medicina</title>
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
	<link rel="icon" href="favicon.ico" type="image/x-icon"> 
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
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

	dim idUsuario 
	idUsuario = Session("idUsu")
	'response.write "<br>29 idUsuario:= " & idUsuario
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
		ed_sPar(1,0) = 127
	end if
	ed_Bot(4)="disabled"
	'ed_Bot(1)="disabled"
	ed_iNumCam		=14					' Numero de campos en la pantalla principal
	ed_iRegPag		=25					' Numero de registros por pagina
	
	ed_sNomTab		="PH_CB_Producto"
	ed_sNomInd		="Id_Producto"
	ed_cCol		=3	' Columna a Ordenar
	ed_cOrd		=0	' Orden 0=ascendente 1=descendente
	ed_iRan		=0	' Presentar ranking de columnas
	ed_iRep=0
	'if idUsuario <> 170 and idUsuario <> 1 and idUsuario <> 76 then  ed_iRep=1
'ed_ides=1
	SqlCla = " SELECT * FROM "  & ed_sNomTab
	'sqlcla = sqlcla & " WHERE  (fec_inactivo is null)"
	if ed_sPar(1,0) <> "" and ed_sPar(1,0) <> "Seleccionar" then
		sqlcla = sqlcla & " WHERE id_categoria = " & ed_sPar(1,0)
		'sqlcla = sqlcla & " and  ind_medicina =  1 "
	end if
	'response.write "<br>65 sqlcla:= " & sqlcla
	'response.end 
	

' Titulo	
   ed_sCampo(00,0)="#"
   ed_sCampo(01,0)="Categoria"
   ed_sCampo(02,0)="Codigo Barra"
   ed_sCampo(03,0)="Producto"
   ed_sCampo(04,0)="Segmento"
   ed_sCampo(05,0)="Fabricante"
   ed_sCampo(06,0)="Marca"
   ed_sCampo(07,0)="Tamano"
   ed_sCampo(08,0)="Tamano Rango"
   ed_sCampo(09,0)="Unidad Medida"
   ed_sCampo(10,0)="Fecha Alta"
   ed_sCampo(11,0)="Fragmentacion"
   ed_sCampo(12,0)="Activo?"
   ed_sCampo(13,0)="Pendiente?"
   ed_sCampo(14,0)="CONTENIDO DE SUSTANCIA ACTIVA"
   ed_sCampo(15,0)="UNIDAD DE SUSTANCIA ACTIVA"
   ed_sCampo(16,0)="PRESENTACIÓN"
   ed_sCampo(17,0)="CORPORATIVO"
   ed_sCampo(18,0)="SEGMENTO"
   ed_sCampo(19,0)="ATC 1"
   ed_sCampo(20,0)=""
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
	'ed_sCampo(13,1)=true

' Obligatorio
	ed_sCampo(01,4)=1
	ed_sCampo(02,4)=1
	ed_sCampo(03,4)=1
	ed_sCampo(05,4)=1
	ed_sCampo(06,4)=1
	ed_sCampo(07,4)=1
	ed_sCampo(08,4)=1
	ed_sCampo(09,4)=1
	ed_sCampo(12,4)=1
	'ed_sCampo(13,4)=1
	
' No Presentar	
	'ed_sCampo(04,2)="2"
	'ed_sCampo(05,2)="1"
	'ed_sCampo(06,2)="1"
	'ed_sCampo(07,2)="2"
	'ed_sCampo(08,2)="2"
	'ed_sCampo(09,2)="2"
	'ed_sCampo(10,2)="2"	
	'ed_sCampo(11,2)="2"
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
	'ed_sCampo(13,2)="1"
	
' Copiar
 '   ed_sCampo(3,8)=2
  '  ed_sCampo(3,8)=2
    
	
	ed_sQue(1,0)=  " SELECT Id_Categoria, Categoria FROM  PH_CB_Categoria WHERE Fec_Inactivo is Null and ind_activo = 1 and id_Categoria = " & ed_sPar(1,0)
	ed_sQue(4,0)=  " SELECT Id_Segmento, Segmento FROM  PH_CB_Segmento WHERE Fec_Inactivo is Null and ind_activo = 1 and id_categoria = " & ed_sPar(1,0) & " Order by Segmento "
	ed_sQue(5,0)=  " SELECT Id_Fabricante, Fabricante FROM  PH_CB_Fabricante WHERE Fec_Inactivo is Null and ind_activo = 1 and id_categoria = 0 " & " Order by Fabricante "
	
	sql = " SELECT Id_Marca, Marca FROM  PH_CB_Marca WHERE Fec_Inactivo is Null and ind_activo = 1 and id_categoria = 0 and ind_medicina = 1 " &  " Order by Marca "
	
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " PH_CB_Marca.Id_Marca, "
	'sql = sql & " PH_CB_Marca.Marca, "
	sql = sql & " PH_CB_Marca.Marca+'('+PH_CB_Fabricante.Fabricante+')' as Marca "
	'sql = sql & " PH_CB_Fabricante.Fabricante "
	sql = sql & " FROM PH_CB_Marca INNER JOIN PH_CB_Fabricante ON PH_CB_Marca.Id_Fabricante = PH_CB_Fabricante.id_Fabricante "
	sql = sql & " WHERE (((PH_CB_Marca.Ind_Medicina)=1) AND ((PH_CB_Marca.Fec_Inactivo) Is Null) AND ((PH_CB_Marca.Ind_Activo)=1) AND ((PH_CB_Marca.Id_Categoria)=0)) "
	sql = sql & " ORDER BY PH_CB_Marca.Marca "

	ed_sQue(6,0)=  sql
	
	ed_sQue(7,0)=  " SELECT Id_Tamano, Tamano FROM  PH_CB_Tamano WHERE Fec_Inactivo is Null and ind_activo = 1 and id_categoria = " & ed_sPar(1,0) & " Order by Tamano "
	ed_sQue(8,0)=  " SELECT Id_TamanoRango, TamanoRango FROM  PH_CB_TamanoRango WHERE Fec_Inactivo is Null and ind_activo = 1 and id_categoria = " & ed_sPar(1,0) & " Order by TamanoRango "
	ed_sQue(9,0)=  " SELECT Id_UnidadMedida, UnidadMedida FROM  PH_CB_UnidadMedida WHERE Fec_Inactivo is Null and ind_activo = 1 and id_categoria = " & cint(ed_sPar(1,0))
	ed_sQue(14,0)=  " SELECT Id_Atributo1, Atributo1 FROM  PH_CB_Atributo1 WHERE Fec_Inactivo is Null and ind_activo = 1 and id_categoria = " & cint(ed_sPar(1,0))
	ed_sQue(15,0)=  " SELECT Id_Atributo2, Atributo2 FROM  PH_CB_Atributo2 WHERE Fec_Inactivo is Null and ind_activo = 1 and id_categoria = " & cint(ed_sPar(1,0))
	ed_sQue(16,0)=  " SELECT Id_Atributo3, Atributo3 FROM  PH_CB_Atributo3 WHERE Fec_Inactivo is Null and ind_activo = 1 and id_categoria = " & cint(ed_sPar(1,0))
	ed_sQue(17,0)=  " SELECT Id_Atributo4, Atributo4 FROM  PH_CB_Atributo4 WHERE Fec_Inactivo is Null and ind_activo = 1 and id_categoria = " & cint(ed_sPar(1,0))
	ed_sQue(18,0)=  " SELECT Id_Atributo5, Atributo5 FROM  PH_CB_Atributo5 WHERE Fec_Inactivo is Null and ind_activo = 1 and id_categoria = " & cint(ed_sPar(1,0))
	ed_sQue(19,0)=  " SELECT Id_Atributo6, Atributo6 FROM  PH_CB_Atributo6 WHERE Fec_Inactivo is Null and ind_activo = 1 and id_categoria = " & cint(ed_sPar(1,0))
	ed_sQue(20,0)=  " SELECT Id_Atributo7, Atributo7 FROM  PH_CB_Atributo7 WHERE Fec_Inactivo is Null and ind_activo = 1 and id_categoria = " & cint(ed_sPar(1,0))
	

	ed_Formato(00,0)="w3-col l1  w3-left w3-padding "
	ed_Formato(01,0)="w3-col l2  w3-left w3-padding "
	ed_Formato(02,0)="w3-col l4  w3-left w3-padding "
	ed_Formato(03,0)="w3-col l6  w3-left w3-padding "
	ed_Formato(04,0)="w3-col l3  w3-left w3-padding "
	ed_Formato(05,0)="w3-col l3  w3-left w3-padding "
	ed_Formato(06,0)="w3-col l2  w3-left w3-padding "
	ed_Formato(07,0)="w3-col l2  w3-left w3-padding "
	ed_Formato(08,0)="w3-col l2  w3-left w3-padding "
	ed_Formato(09,0)="w3-col l2  w3-left w3-padding "
	ed_Formato(10,0)="w3-col l2  w3-left w3-padding "
	ed_Formato(11,0)="w3-col l4  w3-left w3-padding "
	ed_Formato(12,0)="w3-col l1  w3-left w3-padding "
	ed_Formato(13,0)="w3-col l1  w3-left w3-padding "
	ed_Formato(14,0)="w3-col l4  w3-left w3-padding "
	ed_Formato(15,0)="w3-col l4  w3-left w3-padding "
	ed_Formato(16,0)="w3-col l4  w3-left w3-padding "
	ed_Formato(17,0)="w3-col l4  w3-left w3-padding "
	ed_Formato(18,0)="w3-col l4  w3-left w3-padding "
	ed_Formato(19,0)="w3-col l4  w3-left w3-padding "
	ed_Formato(20,0)="w3-col l4  w3-left w3-padding "
	'ed_Formato(13,0)="w3-col l1  w3-left w3-padding "
	
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
	sql = sql & " where ind_medicina = 1 "
	sql = sql & " and id_Categoria > 0 "
	sql = sql & " Order By "
	sql = sql & " Categoria "
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(1,0)="Categoria"
    ed_sCombo(1,1)=sql 
    'ed_sCombo(1,2)="Seleccionar"
	'ed_sCombo(1,2)=""
	
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
	if ed_sPar(1,0) = "" or ed_sPar(1,0) = "Seleccionar" then
		ed_sPar(1,0) = 127
	end if
    

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