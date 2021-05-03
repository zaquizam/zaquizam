<!DOCTYPE HTML>
<html >
<head>
	<title>Reporte Total</title>
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
	<link rel="icon" href="favicon.ico" type="image/x-icon"> 

</head>
<body topmargin="0">
<!--#include file="Conexion.asp"-->

<%

  
'==========================================================================================
' Variables y Constantes
'==========================================================================================

    'Apertura
	   
%>
<script type="text/javascript">
</script>
<%

   
'==========================================================================================
' Parámetros del Manteniemiento
'==========================================================================================
  
    

	'response.write "llego1"
	'response.end
	'ParDat
%>
		
	<br>
	<br>
	<br>
	<div style="width:98%"></div></center>
<%
	'response.write "paso"

%>


<%
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	dim gDatosSol0
	dim rsx0
	set rsx0 = CreateObject("ADODB.Recordset")
	rsx0.CursorType = adOpenKeyset 
	rsx0.LockType = 2 'adLockOptimistic 
	'Buscar Area
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " PH_GAreaEstado.Id_Estado, "
	sql = sql & " PH_GArea.Area "
	sql = sql & " FROM PH_GAreaEstado INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area "
	sql = sql & " ORDER "
	sql = sql & " BY PH_GAreaEstado.Id_Estado "
	'response.write "<br>36 sql:=" & sql
	'response.end
	rsx0.Open sql ,conexion
	if rsx0.eof then
		rsx0.close
	else
		gDatosSol0 = rsx0.GetRows
		rsx0.close
	end if
	dim gArea(24)
	for iReg = 0 to ubound(gDatosSol0,2)
		iEstado = gDatosSol0(0,iReg)
		sArea = gDatosSol0(1,iReg)
		gArea(iEstado) = sArea
	next 
	
	dim gDatosSol
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	sql = ""
    sql = sql & " SELECT "
	sql = sql & " PH_PanelHogar.Id_PanelHogar, "
	sql = sql & " PH_PanelHogar.CodigoHogar, "
	sql = sql & " ss_Estado.Estado, "
	sql = sql & " PH_Ciudad.Ciudad, "
	sql = sql & " ss_Municipio.Municipio, "
	sql = sql & " ss_Parroquia.Parroquia, "
	sql = sql & " PH_PanelHogar.Calle, "
	sql = sql & " PH_PanelHogar.Edificio, "
	sql = sql & " PH_PanelHogar.Casa, "
	sql = sql & " PH_PanelHogar.Escalera, "
	sql = sql & " PH_PanelHogar.Piso, "
	sql = sql & " PH_PanelHogar.Apto, "
	sql = sql & " PH_PanelHogar.Barrio, "
	sql = sql & " PH_PanelHogar.Referencia, "
	sql = sql & " PH_PanelHogar.TelefonoLocal, "
	sql = sql & " PH_PanelHogar.TotalPersonas, "
	sql = sql & " PH_TipoVivienda.TipoVivienda, "
	sql = sql & " PH_PanelHogar.OtroTipoVivienda, "
	sql = sql & " PH_MetrosVivienda.MetrosVivienda, "
	sql = sql & " PH_PanelHogar.NumeroAmbientes, "
	sql = sql & " PH_PanelHogar.NumeroBanos, "
	sql = sql & " PH_PuntosLuz.PuntosLuz, "
	sql = sql & " PH_OcupacionVivienda.OcupacionVivienda, "
	sql = sql & " PH_PanelHogar.OtroOcupacionVivienda, "
	sql = sql & " PH_MontoVivienda.MontoVivienda, "
	sql = sql & " PH_AguasBlancas.AguasBlancas, "
	sql = sql & " PH_AguasNegras.AguasNegras, "
	sql = sql & " PH_AseoUrbano.AseoUrbano, "
	sql = sql & " PH_PanelHogar.Id_ServicioElectricidad, "
	sql = sql & " PH_PanelHogar.Id_ServicioTelefono, "
	sql = sql & " PH_PanelHogar.Id_DomesticaFija, "
	sql = sql & " PH_PanelHogar.Id_PersonalLabores, "
	sql = sql & " PH_PanelHogar.Id_DomesticaDia, "
	sql = sql & " PH_PanelHogar.id_ConexionInternet1, "
	sql = sql & " PH_PanelHogar.id_ConexionInternet2, "
	sql = sql & " PH_PanelHogar.id_ConexionInternet3, "
	sql = sql & " PH_PanelHogar.id_CelularJefe, "
	sql = sql & " PH_PanelHogar.id_SeguroHCMParticular, "
	sql = sql & " PH_PanelHogar.id_SeguroHCMColectivo, "
	sql = sql & " PH_PanelHogar.id_SeguroHCMSS, "
	sql = sql & " PH_PanelHogar.Id_AireAcondicionado, "
	sql = sql & " PH_PanelHogar.Id_Calentador1, "
	sql = sql & " PH_PanelHogar.Id_Calentador2, "
	sql = sql & " PH_PanelHogar.Id_Computador1, "
	sql = sql & " PH_PanelHogar.Id_Computador2, "
	sql = sql & " PH_PanelHogar.Id_DVD, "
	sql = sql & " PH_PanelHogar.Id_HomeTheater, "
	sql = sql & " PH_PanelHogar.Id_JuegosVodeo, "
	sql = sql & " PH_PanelHogar.Id_HornoMicro, "
	sql = sql & " PH_PanelHogar.Id_Secadora, "
	sql = sql & " PH_PanelHogar.Id_Lavadora1, "
	sql = sql & " PH_PanelHogar.Id_Lavadora2, "
	sql = sql & " PH_PanelHogar.Id_Lavadora3, "
	sql = sql & " PH_PanelHogar.Id_Nevera, "
	sql = sql & " PH_PanelHogar.Id_Freezer, "
	sql = sql & " PH_PanelHogar.Id_Cocina1, "
	sql = sql & " PH_PanelHogar.Id_Cocina2, "
	sql = sql & " PH_PanelHogar.Id_Cocina3, "
	sql = sql & " PH_PanelHogar.Id_Cocina4, "
	sql = sql & " PH_PanelHogar.Id_LavaPlato, "
	sql = sql & " PH_Televisores.Televisores, "
	sql = sql & " PH_TipoTelevisores.TipoTelevisores, "
	sql = sql & " PH_Senal.Senal, "
	sql = sql & " PH_OperadoraCable_1.OperadoraCable, "
	sql = sql & " PH_OperadoraCable.OperadoraCable AS Expr1, "
	sql = sql & " PH_TelevisionOnline_1.TvOnline, "
	sql = sql & " PH_TelevisionOnline.TvOnline AS Expr2, "
	sql = sql & " PH_Autos.Autos, "
	sql = sql & " PH_PanelHogar.Id_Moto, "
	sql = sql & " PH_PanelHogar.Id_SeguroCasco, "
	sql = sql & " PH_PanelHogar.Id_Mascotas, "
	sql = sql & " PH_PanelHogar.Ind_Perro, "
	sql = sql & " PH_PanelHogar.Ind_Gato, "
	sql = sql & " PH_PanelHogar.Ind_Pez, "
	sql = sql & " PH_PanelHogar.Ind_Ave, "
	sql = sql & " PH_PanelHogar.Ind_Roedor, "
	sql = sql & " PH_PanelHogar.Ind_Otro, "
	sql = sql & " PH_PanelHogar.ClaseSocial, "
	sql = sql & " ss_Estado.Id_Estado, "
	sql = sql & " ss_Usuarios.Usuario, "
	sql = sql & " PH_PanelHogar.Ind_Activo "
	sql = sql & " FROM  "
	sql = sql & " PH_AguasNegras RIGHT OUTER JOIN "
	sql = sql & " PH_Televisores RIGHT OUTER JOIN "
	sql = sql & " PH_PanelHogar INNER JOIN "
	sql = sql & " ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado INNER JOIN "
	sql = sql & " PH_Ciudad ON PH_PanelHogar.Id_Ciudad = PH_Ciudad.Id_Ciudad INNER JOIN "
	sql = sql & " ss_Municipio ON PH_PanelHogar.Id_Municipio = ss_Municipio.Id_Municipio INNER JOIN "
	sql = sql & " ss_Parroquia ON PH_PanelHogar.Id_Parroquia = ss_Parroquia.Id_Parroquia LEFT OUTER JOIN "
	sql = sql & " PH_Autos ON PH_PanelHogar.Id_Autos = PH_Autos.Id_Autos LEFT OUTER JOIN "
	sql = sql & " PH_TelevisionOnline ON PH_PanelHogar.Id_TelevisionOnline2 = PH_TelevisionOnline.Id_TvOnline LEFT OUTER JOIN "
	sql = sql & " PH_TelevisionOnline AS PH_TelevisionOnline_1 ON PH_PanelHogar.Id_TelevisionOnline1 = PH_TelevisionOnline_1.Id_TvOnline LEFT OUTER JOIN "
	sql = sql & " PH_OperadoraCable ON PH_PanelHogar.Id_Cablera2 = PH_OperadoraCable.Id_OperadoraCable LEFT OUTER JOIN "
	sql = sql & " PH_OperadoraCable AS PH_OperadoraCable_1 ON PH_PanelHogar.Id_Cablera1 = PH_OperadoraCable_1.Id_OperadoraCable LEFT OUTER JOIN "
	sql = sql & " PH_Senal ON PH_PanelHogar.Id_Senal = PH_Senal.Id_Senal LEFT OUTER JOIN "
	sql = sql & " PH_TipoTelevisores ON PH_PanelHogar.Id_TipoTelevisores = PH_TipoTelevisores.Id_TipoTelevisores ON  "
	sql = sql & " PH_Televisores.Id_Televisores = PH_PanelHogar.Id_Televisores LEFT OUTER JOIN "
	sql = sql & " PH_AseoUrbano ON PH_PanelHogar.Id_AseoUrbano = PH_AseoUrbano.Id_AseoUrbano ON  "
	sql = sql & " PH_AguasNegras.Id_AguasNegras = PH_PanelHogar.Id_AguasNegras LEFT OUTER JOIN "
	sql = sql & " PH_AguasBlancas ON PH_PanelHogar.Id_AguasBlancas = PH_AguasBlancas.Id_AguasBlancas LEFT OUTER JOIN "
	sql = sql & " PH_MontoVivienda ON PH_PanelHogar.Id_MontoVivienda = PH_MontoVivienda.Id_MontoVivienda LEFT OUTER JOIN "
	sql = sql & " PH_OcupacionVivienda ON PH_PanelHogar.Id_OcupacionVivienda = PH_OcupacionVivienda.Id_OcupacionVivienda LEFT OUTER JOIN "
	sql = sql & " PH_PuntosLuz ON PH_PanelHogar.id_PuntosLuz = PH_PuntosLuz.Id_PuntosLuz LEFT OUTER JOIN "
	sql = sql & " PH_MetrosVivienda ON PH_PanelHogar.id_Metros = PH_MetrosVivienda.Id_MetrosVivienda LEFT OUTER JOIN "
	sql = sql & " PH_TipoVivienda ON PH_PanelHogar.Id_TipoVivienda = PH_TipoVivienda.Id_TipoVivienda "
	sql = sql & " LEFT OUTER JOIN ss_Usuarios ON PH_PanelHogar.Id_Usuario = cacevedo_atenas.ss_Usuarios.Id_Usuario "
	sql = sql & " WHERE "
	sql = sql & " PH_PanelHogar.Ind_Activo = 1 "
	'sql = sql & " and PH_PanelHogar.Id_PanelHogar = 706 "
	'sql = sql & " ss_Estado.Id_Estado in (6,9,15,17,20) "
	sql = sql & " Order By "
	sql = sql & " PH_PanelHogar.Id_PanelHogar "
	'response.write "<br>36 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	iExiste = 0
	if rsx1.eof then
	else
		Response.ContentType = "application/vnd.ms-excel"
		Response.AddHeader "Content-disposition","attachment; filename=tem.xls"
		gDatosSol = rsx1.GetRows
		rsx1.close
		Response.write "<table border=1>"
			Response.write "<tr>"
				Response.write "<td>idHogar</td>"
				Response.write "<td>CodigoHogar</td>"
				Response.write "<td>Estado</td>"
				Response.write "<td>Ciudad</td>"
				Response.write "<td>Municipio</td>"
				Response.write "<td>Parroquia</td>"
				Response.write "<td>Calle</td>"
				Response.write "<td>Edificio</td>"
				Response.write "<td>Casa</td>"
				Response.write "<td>Escalera</td>"
				Response.write "<td>Piso</td>"
				Response.write "<td>Apto</td>"
				Response.write "<td>Barrio</td>"
				Response.write "<td>Referencia</td>"
				Response.write "<td>TelefonoLocal</td>"
				Response.write "<td>TotalPersonas</td>"
				Response.write "<td>TipoVivienda</td>"
				Response.write "<td>OtroTipoVivienda</td>"
				Response.write "<td>MetrosVivienda</td>"
				Response.write "<td>NumeroAmbientes</td>"
				Response.write "<td>NumeroBanos</td>"
				Response.write "<td>PuntosLuz</td>"
				Response.write "<td>OcupacionVivienda</td>"
				Response.write "<td>OtrosOcupacionVivienda</td>"
				Response.write "<td>MontoVivienda</td>"
				Response.write "<td>AguasBlancas</td>"
				Response.write "<td>AguasNegras</td>"
				Response.write "<td>AseoUrbano</td>"
				Response.write "<td>ServicioElectricidad</td>"
				Response.write "<td>ServicioTelefono</td>"
				Response.write "<td>DomesticaFija</td>"
				Response.write "<td>PersonalLabores</td>"
				Response.write "<td>DomesticaDia</td>"
				Response.write "<td>InternetDialUp</td>"
				Response.write "<td>InternetBandaAncha</td>"
				Response.write "<td>InternetMovil</td>"
				Response.write "<td>CelularJefe</td>"
				Response.write "<td>SeguroParticular</td>"
				Response.write "<td>SeguroColectivo</td>"
				Response.write "<td>SeguroSocial</td>"
				Response.write "<td>AireAcondicionado</td>"
				Response.write "<td>CalentadorElectrico</td>"
				Response.write "<td>CalectadorGas</td>"
				Response.write "<td>PC</td>"
				Response.write "<td>Laptop</td>"
				Response.write "<td>DVD</td>"
				Response.write "<td>HomeTheater</td>"
				Response.write "<td>VideoJuego</td>"
				Response.write "<td>HornoMicro</td>"
				Response.write "<td>Secadora</td>"
				Response.write "<td>LavadoraAuto</td>"
				Response.write "<td>LavadoraSemi</td>"
				Response.write "<td>LavadoraRodillo</td>"
				Response.write "<td>Nevera</td>"
				Response.write "<td>Freezer</td>"
				Response.write "<td>CocinaElectrica</td>"
				Response.write "<td>CocinaBombona</td>"
				Response.write "<td>ConinaGasDirecto</td>"
				Response.write "<td>CocinaKerosene</td>"
				Response.write "<td>Lavaplatos</td>"
				Response.write "<td>Televisores</td>"
				Response.write "<td>TipoTelevisores</td>"
				Response.write "<td>Señal</td>"
				Response.write "<td>Cablera1</td>"
				Response.write "<td>Cablera2</td>"
				Response.write "<td>TvOnLine1</td>"
				Response.write "<td>TvOnLine2</td>"
				Response.write "<td>Vehiculos</td>"
				Response.write "<td>Motos</td>"
				Response.write "<td>SeguroVehiculo</td>"
				Response.write "<td>Mascotas</td>"
				Response.write "<td>Perro</td>"
				Response.write "<td>Gato</td>"
				Response.write "<td>Pez</td>"
				Response.write "<td>Ave</td>"
				Response.write "<td>Roedor</td>"
				Response.write "<td>Otro</td>"
				Response.write "<td>Nombre1</td>"
				Response.write "<td>Nombre2</td>"
				Response.write "<td>Apellido1</td>"
				Response.write "<td>Apellido2</td>"
				Response.write "<td>Nacionalidad</td>"
				Response.write "<td>Cedula</td>"
				Response.write "<td>Parentesco</td>"
				Response.write "<td>EstadoCivil</td>"
				Response.write "<td>Fec_Nacimiento</td>"
				Response.write "<td>Sexo</td>"
				Response.write "<td>Educacion</td>"
				Response.write "<td>TipoIngreso</td>"				
				Response.write "<td>Correo</td>"
				Response.write "<td>CorreoAlterno</td>"
				Response.write "<td>Celular</td>"
				Response.write "<td>CelularAdicional</td>"
				Response.write "<td>NumeroCortesia</td>"
				Response.write "<td>Titular</td>"
				Response.write "<td>CedulaTitular</td>"
				Response.write "<td>Banco</td>"
				Response.write "<td>NumeroCuenta</td>"
				Response.write "<td>PagoRapido</td>"
				Response.write "<td>FrecuenciaCompra</td>"
				Response.write "<td>ClaseSocial</td>"
				Response.write "<td>Area</td>"
				Response.write "<td>Usuario</td>"
				'Response.write "<td>Activo</td>"
			Response.write "</tr>"
		
		for iReg = 0 to ubound(gDatosSol,2)
			Response.write "<tr>"
				Response.write "<td>" & gDatosSol(0,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(1,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(2,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(3,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(4,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(5,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(6,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(7,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(8,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(9,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(10,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(11,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(12,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(13,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(14,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(15,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(16,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(17,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(18,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(19,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(20,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(21,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(22,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(23,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(24,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(25,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(26,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(27,iReg) & "</td>"
				iY = 28
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 29
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 30
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 31
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 32
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 33
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 34
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 35
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 36
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 37
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 38
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 39
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 40
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 41
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 42
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 43
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 44
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 45
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 46
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 47
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 48
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 49
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 50
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 51
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 52
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 53
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 54
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 55
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 56
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 57
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 58
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 59
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				Response.write "<td>" & gDatosSol(60,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(61,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(62,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(63,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(64,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(65,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(66,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(67,iReg) & "</td>"
				iY = 68
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 69
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 70
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(iY,iReg))
					if ix = 1 then Valor = "Si"
					if ix = 2 then Valor = "No"
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 71
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					if gDatosSol(iY,iReg) then 
						Valor = "Si"
					else
						Valor = "No"
					end if
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 72
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					if gDatosSol(iY,iReg) then 
						Valor = "Si"
					else
						Valor = "No"
					end if
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 73
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					if gDatosSol(iY,iReg) then 
						Valor = "Si"
					else
						Valor = "No"
					end if
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 74
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					if gDatosSol(iY,iReg) then 
						Valor = "Si"
					else
						Valor = "No"
					end if
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 75
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					if gDatosSol(iY,iReg) then 
						Valor = "Si"
					else
						Valor = "No"
					end if
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 76
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = 0
				else
					if gDatosSol(iY,iReg) then 
						Valor = "Si"
					else
						Valor = "No"
					end if
				end if
				Response.write "<td>" & Valor & "</td>"

				dim gDatosSol2
				dim rsx2
				set rsx2 = CreateObject("ADODB.Recordset")
				rsx2.CursorType = adOpenKeyset 
				rsx2.LockType = 2 'adLockOptimistic 
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " PH_Panelistas.Nombre1, "
				sql = sql & " PH_Panelistas.Nombre2, "
				sql = sql & " PH_Panelistas.Apellido1, "
				sql = sql & " PH_Panelistas.Apellido2, "
				sql = sql & " PH_Nacionalidad.Nacionalidad, "
				sql = sql & " PH_Panelistas.Cedula, "
				sql = sql & " PH_Parentesco.Parentesco, "
				sql = sql & " PH_EstadoCivil.EstadoCivil, "
				sql = sql & " PH_Panelistas.Fec_Nacimiento, "
				sql = sql & " PH_Sexo.Sexo, "
				sql = sql & " PH_Educacion.Educacion, "
				sql = sql & " PH_TipoIngreso.TipoIngreso, "
				sql = sql & " PH_Panelistas.Correo, "
				sql = sql & " PH_Panelistas.CorreoAlterno, "
				sql = sql & " PH_Panelistas.Celular, "
				sql = sql & " PH_Panelistas.CelularAdicional, "
				sql = sql & " PH_Panelistas.NumeroCortesia, "
				sql = sql & " PH_Panelistas.Titular, "
				sql = sql & " PH_Panelistas.CedulaTitular, "
				sql = sql & " PH_Banco.Banco, "
				sql = sql & " PH_Panelistas.NumeroCuenta, "
				sql = sql & " PH_SiNo.SiNo, "
				sql = sql & " PH_FrecuenciaCompra.FrecuenciaCompra "
				sql = sql & " FROM "
				sql = sql & " PH_Panelistas LEFT OUTER JOIN "
				sql = sql & " PH_FrecuenciaCompra ON PH_Panelistas.Id_FrecuenciaCompra = PH_FrecuenciaCompra.Id_FrecuenciaCompra LEFT OUTER JOIN "
				sql = sql & " PH_SiNo ON PH_Panelistas.Id_PagoRapido = PH_SiNo.Id_SiNo LEFT OUTER JOIN "
				sql = sql & " PH_Banco ON PH_Panelistas.Id_Banco = PH_Banco.Id_Banco LEFT OUTER JOIN "
				sql = sql & " PH_TipoIngreso ON PH_Panelistas.Id_TipoIngreso = PH_TipoIngreso.Id_TipoIngreso LEFT OUTER JOIN "
				sql = sql & " PH_Educacion ON PH_Panelistas.Id_Educacion = PH_Educacion.Id_Educacion LEFT OUTER JOIN "
				sql = sql & " PH_Sexo ON PH_Panelistas.Id_Sexo = PH_Sexo.Id_Sexo LEFT OUTER JOIN "
				sql = sql & " PH_EstadoCivil ON PH_Panelistas.Id_EstadoCivil = PH_EstadoCivil.Id_EstadoCivil LEFT OUTER JOIN "
				sql = sql & " PH_Parentesco ON PH_Panelistas.Id_Parentesco = PH_Parentesco.Id_Parentesco LEFT OUTER JOIN "
				sql = sql & " PH_Nacionalidad ON PH_Panelistas.Id_Nacionalidad = PH_Nacionalidad.Id_Nacionalidad "
				sql = sql & " WHERE "
				sql = sql & " PH_Panelistas.ResponsablePanel = 1 "
				sql = sql & " AND PH_Panelistas.Ind_Activo = 1 "
				sql = sql & " AND PH_Panelistas.Id_Hogar = " & gDatosSol(0,iReg)
				'response.write "<br>36 sql:=" & sql
				'response.end
				rsx2.Open sql ,conexion
				if rsx2.eof then
				else
					gDatosSol2 = rsx2.GetRows
					rsx2.close
					for iReg2 = 0 to 19
						Response.write "<td>"
						if isnull(gDatosSol2(iReg2,0)) or gDatosSol2(iReg2,0) = "" then
						else
							Response.write gDatosSol2(iReg2,0)
						end if
						Response.write "</td>"
					next 
					Response.write "<td>'" & gDatosSol2(20,0) & "</td>"
					Response.write "<td>" & gDatosSol2(21,0) & "</td>"
					Response.write "<td>" & gDatosSol2(22,0) & "</td>"
				end if
				iY = 77
				Valor = ""
				if isnull(gDatosSol(iY,iReg)) or gDatosSol(iY,iReg) = "" then
					Valor = ""
				else
					Valor = gDatosSol(iY,iReg)
				end if
				Response.write "<td>" & Valor & "</td>"
				iY = 78
				idEstado = gDatosSol(iY,iReg)
				Response.write "<td>" & gArea(idEstado) & "</td>"
				iY = 79
				Usuario = gDatosSol(iY,iReg)
				Response.write "<td>" & Usuario & "</td>"
				'iY = 80
				'Activo = gDatosSol(iY,iReg)
				'if Activo = true then
				'	Response.write "<td>Si</td>"
				'else
				'	Response.write "<td>No</td>"
				'end if
			Response.write "</tr>"
		next
		Response.write "</table>"
	end if
%>
<style>


</style>	

    <%conexion.close%>
	
</body>
</html>