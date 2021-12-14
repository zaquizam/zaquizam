<!DOCTYPE HTML>
<html >
<head>
	<title>Reporte General</title>
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
	ed_Bot(2)="disabled"
	ed_Bot(4)="disabled"
	'ed_Bot(1)="disabled"
	ed_iNumCam		=40					' Numero de campos en la pantalla principal
	ed_iRegPag		=25					' Numero de registros por pagina
	
	ed_sNomTab		="PH_PanelHogar"
	ed_sNomInd		="Id_PanelHogar"
	ed_cCol		=0	' Columna a Ordenar
	ed_cOrd		=0	' Orden 0=ascendente 1=descendente
	ed_iRan		=0	' Presentar ranking de columnas
	ed_iRep=1
'ed_ides=1
	SqlCla = " SELECT * FROM "  & ed_sNomTab
	sqlcla = sqlcla & " WHERE  (fec_inactivo is null)"
	'response.write "<br>47 Perfil:= " & Session("idPerfil")
	sqlcla = ""
	sqlcla = sqlcla & " SELECT PH_PanelHogar.Id_PanelHogar, PH_PanelHogar.CodigoHogar, ss_Estado.Estado, PH_Ciudad.Ciudad, ss_Municipio.Municipio, ss_Parroquia.Parroquia, PH_PanelHogar.Barrio, PH_Panelistas.Nombre1, PH_Panelistas.Apellido1, PH_Parentesco.Parentesco, PH_Panelistas.Fec_Nacimiento, PH_Sexo.Sexo, PH_Educacion.Educacion, PH_TipoIngreso.TipoIngreso, PH_Panelistas.Correo, PH_Panelistas.Celular, PH_Panelistas.Titular, PH_Panelistas.CedulaTitular, PH_Banco.Banco, PH_Panelistas.NumeroCuenta, PH_Panelistas.CantidadPersonas, PH_SiNo.SiNo AS TieneMascota, ss_Usuarios.Usuario, PH_PanelHogar.Fec_Registro, PH_PanelHogar.IP, PH_PanelHogar.USR, PH_PanelHogar.Fec_Ult_Mod, PH_PanelHogar.Fec_Inactivo, PH_PanelHogar.IdSession "
	sqlcla = sqlcla & " FROM (((((((((((((PH_PanelHogar INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) INNER JOIN PH_Ciudad ON PH_PanelHogar.Id_Ciudad = PH_Ciudad.Id_Ciudad) INNER JOIN ss_Municipio ON PH_PanelHogar.Id_Municipio = ss_Municipio.Id_Municipio) INNER JOIN ss_Parroquia ON PH_PanelHogar.Id_Parroquia = ss_Parroquia.Id_Parroquia) LEFT JOIN PH_Panelistas ON PH_PanelHogar.Id_PanelHogar = PH_Panelistas.Id_Hogar) LEFT JOIN PH_Nacionalidad ON PH_Panelistas.Id_Nacionalidad = PH_Nacionalidad.Id_Nacionalidad) LEFT JOIN PH_Sexo ON PH_Panelistas.Id_Sexo = PH_Sexo.Id_Sexo) LEFT JOIN PH_Educacion ON PH_Panelistas.Id_Educacion = PH_Educacion.Id_Educacion) LEFT JOIN PH_TipoIngreso ON PH_Panelistas.Id_TipoIngreso = PH_TipoIngreso.Id_TipoIngreso) LEFT JOIN PH_Banco ON PH_Panelistas.Id_Banco = PH_Banco.Id_Banco) LEFT JOIN PH_FrecuenciaCompra ON PH_Panelistas.Id_FrecuenciaCompra = PH_FrecuenciaCompra.Id_FrecuenciaCompra) LEFT JOIN ss_Usuarios ON PH_PanelHogar.Id_Usuario = ss_Usuarios.Id_Usuario) LEFT JOIN PH_Parentesco ON PH_Panelistas.Id_Parentesco = PH_Parentesco.Id_Parentesco) LEFT JOIN PH_SiNo ON PH_PanelHogar.Id_Mascotas = PH_SiNo.Id_SiNo "
	sqlcla = sqlcla & " WHERE PH_PanelHogar.Id_PanelHogar>1 "
	sqlcla = sqlcla & " and PH_PanelHogar.Ind_Activo = 1 "
	'response.write "<br>69 sqlcla:= " & sqlcla


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
	'ed_sCampo(10,2)="2"
	'ed_sCampo(11,2)="2"
	'ed_sCampo(12,2)="2"
	'ed_sCampo(13,2)="2"
	'ed_sCampo(14,2)="1"
	'ed_sCampo(15,2)="1"
	'ed_sCampo(16,2)="2"
	'ed_sCampo(17,2)="2"
	'ed_sCampo(18,2)="2"
	'ed_sCampo(19,2)="2"
	'ed_sCampo(20,2)="2"
	'ed_sCampo(22,2)="1"
	'ed_sCampo(30,2)="1"
	'ed_sCampo(31,2)="1"
	'ed_sCampo(32,2)="1"
	'ed_sCampo(33,2)="1"
	
' Copiar
 '   ed_sCampo(3,8)=2
  '  ed_sCampo(3,8)=2
    
	
	'ed_sQue(1,0)=  " SELECT Id_PanelHogar, CodigoHogar FROM  PH_PanelHogar "
	'ed_sQue(7,0)=  " SELECT Id_Nacionalidad, Nacionalidad FROM  PH_Nacionalidad "
	'ed_sQue(9,0)=  " SELECT Id_Parentesco, Parentesco FROM  PH_Parentesco "
	'ed_sQue(10,0)=  " SELECT Id_EstadoCivil, EstadoCivil FROM  PH_EstadoCivil "
	'ed_sQue(12,0)=  " SELECT Id_Sexo, Sexo FROM  PH_Sexo "
	'ed_sQue(13,0)=  " SELECT Id_Educacion, Educacion FROM  PH_Educacion "
	'ed_sQue(14,0)=  " SELECT Id_CondicionLaboral, CondicionLaboral FROM  PH_CondicionLaboral "
	'ed_sQue(15,0)=  " SELECT Id_Ocupacion, Ocupacion FROM  PH_Ocupacion "
	'ed_sQue(16,0)=  " SELECT Id_TipoIngreso, TipoIngreso FROM  PH_TipoIngreso "
	'ed_sQue(22,0)=  " SELECT Id_Profesion, Profesion FROM  PH_Profesion "
	'ed_sQue(25,0)=  " SELECT Id_Banco, Banco FROM  PH_Banco "
	ed_sQue(14,0)=  " SELECT id_SiNo, SiNo FROM  PH_SiNo "
	'ed_sQue(29,0)=  " SELECT id_FrecuenciaCompra, FrecuenciaCompra FROM  PH_FrecuenciaCompra "
	
	ed_Formato(00,0)="w3-col l1  w3-left w3-padding "
	'ed_Formato(01,0)="w3-col l3  w3-left w3-padding "
	'ed_Formato(02,0)="w3-col l1  w3-left w3-padding "
	'ed_Formato(03,0)="w3-col l1  w3-left w3-padding "
	
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