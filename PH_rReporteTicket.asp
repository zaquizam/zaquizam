<!DOCTYPE HTML>
<html >
<head>
	<title>Reporte del Tiket</title>
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

 'response.write "<br> perfil:= " & Session("perusu")
 Perfil = Session("perusu")
 dim idEstado
 dim idArea
 dim idSemana
 

Sub Combos
	'response.write "<br>372 Combo1:=" & ed_sPar(1,0)
	'response.write " Combo2:=" & ed_sPar(2,0)
	'response.write " Combo3:=" & ed_sPar(3,0)

    ed_iCombo = 3

	if ed_sPar(1,0) = "" or isnull(ed_sPar(1,0)) or ed_sPar(1,0) = "Seleccionar" then
		ed_sPar(1,0) = 0
	end if
	if ed_sPar(2,0) = "" or isnull(ed_sPar(2,0)) or ed_sPar(2,0) = "Seleccionar" then
		ed_sPar(2,0) = 0
	end if
	if ed_sPar(3,0) = "" or isnull(ed_sPar(3,0)) or ed_sPar(3,0) = "Seleccionar" then
		ed_sPar(3,0) = 0
	end if
	
	'response.write "<br>33 Paso"
	'response.end
	
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id, "
	sql = sql & " Semana "
	sql = sql & " FROM ss_Semana "
	'if ed_sPar(1,0) <> 0 then
	'	sql = sql & " WHERE Id = " & ed_sPar(1,0)
	'end if
	'response.write "<br>33 Paso"
	'response.end
	sql = sql & " Order By "
	sql = sql & " id desc "
	'response.write "<br>372 Combo1:=" & sql
	'response.end
    ed_sCombo(1,0)="Semana"
    ed_sCombo(1,1)=sql 
    ed_sCombo(1,2)="Seleccionar"

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Area, "
	sql = sql & " Area "
	sql = sql & " FROM PH_GArea "
	'if ed_sPar(2,0) <> 0 then
	'	sql = sql & " WHERE Id_Area = " & ed_sPar(2,0)
	'end if
	'response.write "<br>33 Paso"
	'response.end
	sql = sql & " Order By "
	sql = sql & " Area "
	'response.write "<br>372 Combo1:=" & sql
	'response.end
    ed_sCombo(2,0)="Area"
    ed_sCombo(2,1)=sql 
    ed_sCombo(2,2)="Seleccionar"


	sql = ""
	sql = sql & " SELECT "
	sql = sql & " ss_Estado.Id_Estado, "
	sql = sql & " ss_Estado.Estado "
	sql = sql & " FROM PH_GAreaEstado INNER JOIN ss_Estado ON PH_GAreaEstado.Id_Estado = ss_Estado.Id_Estado "
	if ed_sPar(2,0) <> 0 then
		sql = sql & " WHERE PH_GAreaEstado.Id_Area = " & ed_sPar(2,0)
	end if
	'response.write "<br>33 Paso"
	'response.end
	sql = sql & " Order By "
	sql = sql & " ss_Estado.Estado "
	'response.write "<br>372 Combo1:=" & sql
	'response.end
    ed_sCombo(3,0)="Estado"
    ed_sCombo(3,1)=sql 
    ed_sCombo(3,2)="Seleccionar"

	
End Sub

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
	'response.write "<br>372 Combo1:=" & ed_sPar(1,0)
	'response.write " Combo2:=" & ed_sPar(2,0)
	'response.write " Combo3:=" & ed_sPar(3,0)
	'response.write "<br>372 Combo1:=" & idSemana
	'response.write " Combo2:=" & idArea
	'response.write " Combo3:=" & idEstado
	
	if ed_sPar(1,0) = "" or isnull(ed_sPar(1,0)) or ed_sPar(1,0) = "Seleccionar" then
		idSemana = 0 
	else
		idSemana = ed_sPar(1,0)
	end if
	
	if ed_sPar(2,0) = "" or isnull(ed_sPar(2,0)) or ed_sPar(2,0) = "Seleccionar" then
		idArea= 0 
	else
		idArea = ed_sPar(2,0)
	end if
	
	if ed_sPar(3,0) = "" or isnull(ed_sPar(3,0)) or ed_sPar(3,0) = "Seleccionar" then
		idEstado= 0 
	else
		idEstado = ed_sPar(3,0)
	end if
	
	if idSemana = 0 then exit sub
	if idArea = 0 then exit sub
	
	ed_Bot(2)="disabled"
	ed_Bot(3)="disabled"
	ed_Bot(4)="disabled"
	
	ed_iNumCam		=18					' Numero de campos en la pantalla principal
	ed_iRegPag		=25					' Numero de registros por pagina
	
	ed_sNomTab		="ss_meses"
	ed_sNomInd		="Id_Mes"
	ed_cCol		=0	' Columna a Ordenar
	ed_cOrd		=0	' Orden 0=ascendente 1=descendente
	ed_iRan		=0	' Presentar ranking de columnas
	ed_iRep=0
	
	'ed_ides=1
	'SqlCla = " SELECT * FROM "  & ed_sNomTab
	'sqlcla = sqlcla & " WHERE  (fec_inactivo is null)"

	SqlCla = ""
	SqlCla = SqlCla & " SELECT "
	SqlCla = SqlCla & " PH_Consumo.Id_Hogar as IdHogar, "
	SqlCla = SqlCla & " PH_PanelHogar.CodigoHogar, "
	SqlCla = SqlCla & " PH_GArea.Area, "
	SqlCla = SqlCla & " ss_Estado.Estado, "
	SqlCla = SqlCla & " PH_Consumo.Id_Consumo as Consumo, "
	SqlCla = SqlCla & " PH_Medio.Medio, "
	SqlCla = SqlCla & " PH_Moneda.Moneda, "
	SqlCla = SqlCla & " PH_Consumo.fecha_consumo, "
	SqlCla = SqlCla & " PH_FormaPago.FormaPago, "
	SqlCla = SqlCla & " PH_Consumo.Total_items, "
	SqlCla = SqlCla & " PH_Consumo.Total_Compra, "
	SqlCla = SqlCla & " PH_Canal.Canal, "
	SqlCla = SqlCla & " PH_Cadena.Cadena, "
	SqlCla = SqlCla & " PH_PanelHogar.Ind_Activo, "
	SqlCla = SqlCla & " PH_Consumo.IP, PH_Consumo.USR, PH_Consumo.Fec_Ult_Mod, PH_Consumo.Fec_Inactivo, PH_Consumo.IdSession "
	SqlCla = SqlCla & " FROM ((((((((PH_Consumo INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) INNER JOIN PH_GAreaEstado ON ss_Estado.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) INNER JOIN PH_Medio ON PH_Consumo.Id_Medio = PH_Medio.Id_Medio) INNER JOIN PH_Moneda ON PH_Consumo.Id_Moneda = PH_Moneda.Id_Moneda) INNER JOIN PH_FormaPago ON PH_Consumo.Id_FomaPago = PH_FormaPago.Id_FormaPago) INNER JOIN PH_Canal ON PH_Consumo.Id_Canal = PH_Canal.Id_Canal) INNER JOIN PH_Cadena ON PH_Consumo.Id_Cadena = PH_Cadena.Id_Cadena "
	SqlCla = SqlCla & " WHERE "
	SqlCla = SqlCla & " PH_Consumo.Tiene_Factura = 1 "
	SqlCla = SqlCla & " AND PH_Consumo.Status_registro='G' "
	SqlCla = SqlCla & " AND PH_Consumo.Id_Semana > 0 "
	SqlCla = SqlCla & " AND PH_Consumo.Id_Semana = " & idSemana
	if idArea <> 0 then
		SqlCla = SqlCla & " AND PH_GAreaEstado.Id_Area = "  & idArea
	end if
	if idEstado <> 0 then
		SqlCla = SqlCla & " AND PH_GAreaEstado.Id_Estado = " & idEstado
	end if
	'response.write "<br> Semana:= " & idSemana
	'response.write "<br> Area:= " & idArea
	'response.write "<br> Estado:= " & idEstado
	'response.write "<br> SqlCla:= " & SqlCla

' Titulo	
	'ed_sCampo(00,0)="#"
	'ed_sCampo(01,0)="Hogar"
	'ed_sCampo(02,0)="Fecha"   
	'ed_sCampo(03,0)="Bitacora"   

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
	'ed_sCampo(16,2)="1"
	'ed_sCampo(17,2)="1"
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
	'sql = ""
	'sql = sql & " Select "
	'sql = sql & " Id_PanelHogar, "
	'sql = sql & " CONVERT(CHAR(10), Id_PanelHogar)+' '+CodigoHogar as Hogar "
	'sql = sql & " FROM  PH_PanelHogar "
	'sql = sql & " WHERE Fec_Inactivo is Null "
	'response.write "<br>132 sql:= " & sql
	'ed_sQue(1,0)=  sql
    
	'ed_sQue(1,0)=  " SELECT Id_PanelHogar, Id_PanelHogar&' '&CodigoHogar FROM  PH_PanelHogar WHERE Fec_Inactivo is Null And Id_PanelHogar >1"
	'ed_sQue(4,0)=  " SELECT Id_PerfilUsuario, PerfilUsuario FROM  Syn_PerfilUsuario WHERE Fec_Inactivo is Null And Id_PerfilUsuario >1"
	'ed_sQue(5,0)=  " SELECT Id_PerfilUsuario, PerfilUsuario FROM  ss_PerfilUsuario WHERE Fec_Inactivo is Null And Id_PerfilUsuario >1"
	'ed_sQue(2,0)=  " SELECT id_Pais, Pais FROM  ss_Pais "
	'ed_sQue(4,0)=  " SELECT id_Estado, Estado FROM  ss_Estado "
	'ed_sQue(4,0)=  " SELECT id_Ciudad, Ciudad FROM  PH_Ciudad "
	'ed_sQue(5,0)=  " SELECT id_Municipio, Municipio FROM  ss_Municipio "
	'ed_sQue(6,0)=  " SELECT id_Parroquia, Parroquia FROM  ss_Parroquia "
	'ed_sQue(17,0)=  " SELECT id_TipoVivienda, TipoVivienda FROM  PH_TipoVivienda "
	'ed_sQue(19,0)=  " SELECT id_MetrosVivienda, MetrosVivienda FROM  PH_MetrosVivienda "
	'ed_sQue(22,0)=  " SELECT id_PuntosLuz, PuntosLuz FROM  PH_PuntosLuz "
	'ed_sQue(23,0)=  " SELECT id_OcupacionVivienda, OcupacionVivienda FROM  PH_OcupacionVivienda "
	'ed_sQue(25,0)=  " SELECT id_MontoVivienda, MontoVivienda FROM  PH_MontoVivienda "
	'ed_sQue(26,0)=  " SELECT id_AguasBlancas, AguasBlancas FROM  PH_AguasBlancas "
	'ed_sQue(27,0)=  " SELECT id_AguasNegras, AguasNegras FROM  PH_AguasNegras "
	'ed_sQue(28,0)=  " SELECT id_AseoUrbano, AseoUrbano FROM  PH_AseoUrbano "
	'ed_sQue(29,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(30,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(31,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(32,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(33,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(34,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(35,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(36,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(37,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(38,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(39,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(40,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(41,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(42,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(43,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(44,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(45,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(46,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(47,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(48,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(49,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(50,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(51,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(52,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(53,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(54,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(55,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(56,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(57,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(58,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(59,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(60,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(61,0)=  " SELECT id_Televisores, Televisores FROM  PH_Televisores "
	'ed_sQue(62,0)=  " SELECT id_TipoTelevisores, TipoTelevisores FROM  PH_TipoTelevisores "
	'ed_sQue(63,0)=  " SELECT id_Senal, Senal FROM  PH_Senal "
	'ed_sQue(64,0)=  " SELECT id_OperadoraCable, OperadoraCable FROM  PH_OperadoraCable "
	'ed_sQue(65,0)=  " SELECT id_OperadoraCable, OperadoraCable FROM  PH_OperadoraCable "
	'ed_sQue(66,0)=  " SELECT Id_TvOnline, Id_TvOnline FROM  PH_TelevisionOnline "
	'ed_sQue(67,0)=  " SELECT Id_TvOnline, Id_TvOnline FROM  PH_TelevisionOnline "
	'ed_sQue(68,0)=  " SELECT id_Autos, Autos FROM  PH_Autos "
	'ed_sQue(69,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(70,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(71,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(79,0)=  " SELECT id_Usuario, Usuario FROM  ss_Usuarios " 
	
	'ed_Formato(00,0)="w3-col l1  w3-left w3-padding "
	'ed_Formato(03,0)="w3-col l12  w3-left w3-padding "
	'ed_Formato(02,0)="w3-col l1  w3-left w3-padding "
	'ed_Formato(03,0)="w3-col l1  w3-left w3-padding "
	'ed_Formato(84,0)="w3-col l6  w3-left w3-padding "
	
End Sub
    LeePar
  
    ParDat
	
    if ed_iPas<>4 then 
        Encabezado
    end if    
	sExcel = request.Form("bus")

	'response.write "<br>372 Combo1:=" & ed_sPar(1,0)
	'response.write "<br> Combo2:=" & ed_sPar(2,0)
	'response.write "<br>"
	if ed_sPar(1,0) = "" or isnull(ed_sPar(1,0)) or ed_sPar(1,0) = "Seleccionar" then
		idSemana = 0 
	else
		idSemana = ed_sPar(1,0)
	end if
	
	if ed_sPar(2,0) = "" or isnull(ed_sPar(2,0)) or ed_sPar(2,0) = "Seleccionar" then
		idArea= 0 
	else
		idArea = ed_sPar(2,0)
	end if
	
	if ed_sPar(3,0) = "" or isnull(ed_sPar(3,0)) or ed_sPar(3,0) = "Seleccionar" then
		idEstado= 0 
	else
		idEstado = ed_sPar(3,0)
	end if
	'response.write "<br>372 idEstado:=" & idEstado
	'response.write "<br> idHogar:=" & idHogar
	'response.write "<br>"

	Combos

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
	<table border="0" align="right">
		<tr>
			<td>
				<%
				ed_vCombo
				%>
			</td>
		</tr>
	</table>
		
	<br>
	<div style="width:98%"><%
	ed_Main 
	%></div></center>
	</br>
	</br>
	</br>

    <%conexion.close%>
	


</body>
</html>