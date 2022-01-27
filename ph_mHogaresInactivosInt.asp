<!DOCTYPE HTML>
<html >
<head>
	<title>Hogares Inactivos y Parnelistas</title>
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
' Perfil = Session("perusu")

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
	ed_Bot(2)="disabled"
	ed_Bot(4)="disabled"
	'ed_Bot(1)="disabled"
	if Perfil = 2 then
		ed_Bot(3)="disabled"
	end if
	
	ed_iNumCam		=17					' Numero de campos en la pantalla principal
	ed_iRegPag		=25					' Numero de registros por pagina
	
	ed_sNomTab		="PH_PanelHogar"
	ed_sNomInd		="Id_PanelHogar"
	ed_cCol		=0	' Columna a Ordenar
	ed_cOrd		=0	' Orden 0=ascendente 1=descendente
	ed_iRan		=0	' Presentar ranking de columnas
	ed_iRep=1
'ed_ides=1
	'SqlCla = " SELECT * FROM "  & ed_sNomTab
	'sqlcla = sqlcla & " WHERE  (fec_inactivo is null)"
	'sqlcla = sqlcla & " and  (ind_activo = 0)"
	SqlCla = ""
	SqlCla = SqlCla & " SELECT "
	SqlCla = SqlCla & " PH_PanelHogar.Id_PanelHogar, "
	SqlCla = SqlCla & " PH_PanelHogar.CodigoHogar, "
	SqlCla = SqlCla & " PH_GArea.Area, "
	SqlCla = SqlCla & " ss_Estado.Estado, "
	SqlCla = SqlCla & " ss_Municipio.Municipio, "
	SqlCla = SqlCla & " ss_Parroquia.Parroquia, "
	SqlCla = SqlCla & " PH_PanelHogar.Barrio, "
	SqlCla = SqlCla & " PH_PanelHogar.Referencia, "
	SqlCla = SqlCla & " PH_PanelHogar.TotalPersonas, "
	SqlCla = SqlCla & " PH_PanelHogar.TelefonoLocal, "
	SqlCla = SqlCla & " PH_Panelistas.Nombre1, "
	SqlCla = SqlCla & " PH_Panelistas.Apellido1, "
	SqlCla = SqlCla & " PH_Panelistas.ResponsablePanel, "
	SqlCla = SqlCla & " PH_Panelistas.Cedula, "
	SqlCla = SqlCla & " PH_Panelistas.Correo, "
	SqlCla = SqlCla & " PH_Panelistas.Celular, "
	SqlCla = SqlCla & " PH_Panelistas.Titular, "
	SqlCla = SqlCla & " PH_PanelHogar.IP, "
	SqlCla = SqlCla & " PH_PanelHogar.USR, "
	SqlCla = SqlCla & " PH_PanelHogar.Fec_Ult_Mod, "
	SqlCla = SqlCla & " PH_PanelHogar.Fec_Inactivo, "
	SqlCla = SqlCla & " PH_PanelHogar.IdSession "
	SqlCla = SqlCla & " FROM ((((PH_PanelHogar INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) INNER JOIN (PH_GArea INNER JOIN PH_GAreaEstado ON PH_GArea.Id_Area = PH_GAreaEstado.Id_Area) ON ss_Estado.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN ss_Municipio ON PH_PanelHogar.Id_Municipio = ss_Municipio.Id_Municipio) INNER JOIN ss_Parroquia ON PH_PanelHogar.Id_Parroquia = ss_Parroquia.Id_Parroquia) LEFT JOIN PH_Panelistas ON PH_PanelHogar.Id_PanelHogar = PH_Panelistas.Id_Hogar "
	SqlCla = SqlCla & " WHERE "
	SqlCla = SqlCla & " PH_PanelHogar.Ind_Activo = 0 "
	'SqlCla = SqlCla & " AND PH_PanelHogar.Id_PanelHogar = 972 "
	'SqlCla = SqlCla & " Order By "
	'SqlCla = SqlCla & " PH_GArea.Area, "
	'SqlCla = SqlCla & " ss_Estado.Estado "
	'SqlCla = SqlCla & " PH_PanelHogar.Id_PanelHogar "

	'response.write "<br>47 SqlCla:= " & SqlCla

' Titulo	
	ed_sCampo(00,0)="#"
	'ed_sCampo(02,0)="Pais"
	'ed_sCampo(03,0)="Estado"   
	'ed_sCampo(04,0)="Ciudad"   
	'ed_sCampo(05,0)="Municipio"   
	'ed_sCampo(06,0)="Parroquia"   

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
    
	
	'ed_sQue(4,0)=  " SELECT Id_PerfilUsuario, PerfilUsuario FROM  Syn_PerfilUsuario WHERE Fec_Inactivo is Null And Id_PerfilUsuario >1"
	'ed_sQue(5,0)=  " SELECT Id_PerfilUsuario, PerfilUsuario FROM  ss_PerfilUsuario WHERE Fec_Inactivo is Null And Id_PerfilUsuario >1"
	'ed_sQue(2,0)=  " SELECT id_Pais, Pais FROM  ss_Pais "
	'ed_sQue(3,0)=  " SELECT id_Estado, Estado FROM  ss_Estado "
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
	'ed_Formato(01,0)="w3-col l3  w3-left w3-padding "
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
	<div style="width:98%"><%
	ed_Main 
	%></div></center>

    <%conexion.close%>
	


</body>
</html>