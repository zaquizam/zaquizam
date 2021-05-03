<!DOCTYPE HTML>
<html >
<head>
	<title>NSE</title>
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
	dim IdHogar

	
	dim gDatosSol
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	sql = ""
    sql = sql & " SELECT "
	sql = sql & " PH_PanelHogar.Id_PanelHogar AS idHogar, "'00
	sql = sql & " PH_PanelHogar.CodigoHogar, "'01
	sql = sql & " ss_Estado.Estado, "'02
	sql = sql & " PH_Ciudad.Ciudad, "'03
	sql = sql & " ss_Municipio.Municipio, "'04
	sql = sql & " ss_Parroquia.Parroquia, "'05
	sql = sql & " PH_TipoVivienda.TipoVivienda, "'06
	sql = sql & " PH_TipoVivienda.Peso AS PesoTipoVivienda, "'07
	sql = sql & " PH_MetrosVivienda.MetrosVivienda, "'08
	sql = sql & " PH_MetrosVivienda.Peso AS PesoMetrosVivienda, "'09
	sql = sql & " PH_PanelHogar.NumeroAmbientes, "'10
	sql = sql & " 0 AS PesoNumAmbientes, "'11
	sql = sql & " PH_PanelHogar.TotalPersonas, "'12
	sql = sql & " 0 AS PesoTotalPersonas, "'13
	sql = sql & " PH_PanelHogar.NumeroBanos, "'14
	sql = sql & " 0 AS PesoNumBanos, "'15
	sql = sql & " PH_OcupacionVivienda.OcupacionVivienda, "'16
	sql = sql & " PH_OcupacionVivienda.Peso AS PesoOcupacionVivienda, "'17
	sql = sql & " PH_MontoVivienda.MontoVivienda, "'18
	sql = sql & " PH_MontoVivienda.Peso AS PesoMontoVivienda, "'19
	sql = sql & " PH_PuntosLuz.PuntosLuz, "'20
	sql = sql & " PH_PuntosLuz.Peso AS PesoPuntoLuz, "'21
	sql = sql & " PH_AguasBlancas.AguasBlancas, "'22
	sql = sql & " PH_AguasBlancas.Peso AS PesoAguasBlancas, "'23
	sql = sql & " PH_AguasNegras.AguasNegras, "'24
	sql = sql & " PH_AguasNegras.Peso AS PesoAguasNegras, "'25
	sql = sql & " PH_AseoUrbano.AseoUrbano, "'26
	sql = sql & " PH_AseoUrbano.Peso as PesoAseoUrbano, "'27
	sql = sql & " PH_PanelHogar.Id_ServicioElectricidad, "'28
	sql = sql & " 0 AS PesoServicioElectricidad, "'29
	sql = sql & " PH_PanelHogar.Id_DomesticaFija, "'30
	sql = sql & " PH_PanelHogar.Id_PersonalLabores, "'31
	sql = sql & " PH_PanelHogar.Id_DomesticaDia, "'32
	sql = sql & " 0 AS PesoPeronalAseo, "'33
	sql = sql & " PH_PanelHogar.id_ConexionInternet2 AS BandaAncha, "'34
	sql = sql & " 0 AS PesoBandaAncha, "'35
	sql = sql & " PH_PanelHogar.id_ConexionInternet3 AS Movil, "'36
	sql = sql & " 0 AS PesoMovil, "'37
	sql = sql & " PH_PanelHogar.id_ConexionInternet1 AS DialUp, "'38
	sql = sql & " 0 AS PesoDialUp, "'39
	sql = sql & " PH_PanelHogar.id_CelularJefe AS CelularJefe, "'40
	sql = sql & " 0 AS PesoCelularJefe, "'41
	sql = sql & " PH_PanelHogar.Id_AireAcondicionado AS AireAcondicionado, "'42
	sql = sql & " 0 AS PesoAireAcondicionado, "'43
	sql = sql & " PH_PanelHogar.Id_Calentador1 AS CalentadorElectrico, "'44
	sql = sql & " 0 AS PesoCalentadorElectrico, "'45
	sql = sql & " PH_PanelHogar.Id_Calentador2 AS CalectadorGas, "'46
	sql = sql & " 0 AS PesoCalentadorGas, "'47
	sql = sql & " PH_PanelHogar.Id_Computador1 AS PC, "'48
	sql = sql & " 0 AS PesoPC, "'49
	sql = sql & " PH_PanelHogar.Id_Computador2 AS Laptop, "'50
	sql = sql & " 0 AS PesoLaptop, "'51
	sql = sql & " PH_PanelHogar.Id_DVD AS DVD, "'52
	sql = sql & " 0 AS PesoDVD, "'53
	sql = sql & " PH_PanelHogar.Id_HomeTheater AS HomeTheater, "'54
	sql = sql & " 0 AS PesoHomeTheater, "'54
	sql = sql & " PH_PanelHogar.Id_JuegosVodeo AS VideoJuego, "'56
	sql = sql & " 0 AS PesoVideoJuego, "'57
	sql = sql & " PH_PanelHogar.Id_HornoMicro AS HornoMicro, "'58
	sql = sql & " 0 AS PesoHornoMicro, "'59
	sql = sql & " PH_PanelHogar.Id_Lavadora1 AS LavadoraAuto, "'60
	sql = sql & " 0 AS PesoLavadoraAuto, "'61
	sql = sql & " PH_PanelHogar.Id_Lavadora2 AS LavadoraSemi, "'62
	sql = sql & " 0 AS PesoLavadoraSemi, "'63
	sql = sql & " PH_PanelHogar.Id_Lavadora3 AS LavadoraRodillo, "'64
	sql = sql & " 0 AS PesoLavadoraRodillo, "'65
	sql = sql & " PH_PanelHogar.Id_Secadora AS Secadora, "'66
	sql = sql & " 0 AS PesoSecadora, "'67
	sql = sql & " PH_PanelHogar.Id_Nevera AS Nevera, "'68
	sql = sql & " 0 AS PesoNevera, "'69
	sql = sql & " PH_PanelHogar.Id_Freezer AS Freezer, "'70
	sql = sql & " 0 AS PesoFreezer, "'71
	sql = sql & " PH_PanelHogar.Id_Cocina1 AS CocinaElectrica, "'72
	sql = sql & " 0 AS PesoCocinaElectrica, "'73
	sql = sql & " PH_PanelHogar.Id_Cocina2 AS CocinaBombona, "'74
	sql = sql & " 0 AS PesoCocinaBombona, "'75
	sql = sql & " PH_PanelHogar.Id_Cocina3 AS ConinaGasDirecto, "'76
	sql = sql & " 0 AS PesoCocinaGasDirecto, "'77
	sql = sql & " PH_PanelHogar.Id_Cocina4 AS CocinaKerosene, "'78
	sql = sql & " 0 AS PesoCocinaKerosene, "'79
	sql = sql & " PH_PanelHogar.Id_LavaPlato AS Lavaplatos, "'80
	sql = sql & " 0 AS PesoLavaplatos, "'81
	sql = sql & " PH_Autos.Autos AS Vehiculos, "'82
	sql = sql & " PH_Autos.Peso AS PesoVehiculo, "'83
	sql = sql & " PH_Televisores.Televisores, "'84
	sql = sql & " PH_Televisores.Peso AS PesoTelevisores, "'85
	sql = sql & " PH_PanelHogar.id_SeguroHCMParticular AS SeguroParticular, "'86
	sql = sql & " 0 AS PesoSeguroParticular, "'87
	sql = sql & " PH_PanelHogar.id_SeguroHCMColectivo AS SeguroColectivo, "'88
	sql = sql & " 0 AS PesoSeguroColectivo, "'89
	sql = sql & " PH_PanelHogar.id_SeguroHCMSS AS SeguroSocial, "'90
	sql = sql & " 0 AS PesoSeguroSocial, "'91
	sql = sql & " PH_PanelHogar.Id_SeguroCasco AS SeguroVehiculo, "'92
	sql = sql & " 0 AS PesoSeguroVehiculo "'93
	sql = sql & " FROM PH_PanelHogar INNER JOIN "
	sql = sql & " PH_Ciudad ON PH_PanelHogar.Id_Ciudad = PH_Ciudad.Id_Ciudad INNER JOIN "
	sql = sql & " ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado INNER JOIN "
	sql = sql & " ss_Municipio ON PH_PanelHogar.Id_Municipio = ss_Municipio.Id_Municipio INNER JOIN "
	sql = sql & " ss_Parroquia ON PH_PanelHogar.Id_Parroquia = ss_Parroquia.Id_Parroquia LEFT OUTER JOIN "
	sql = sql & " PH_Televisores ON PH_PanelHogar.Id_Televisores = PH_Televisores.Id_Televisores LEFT OUTER JOIN "
	sql = sql & " PH_Autos ON PH_PanelHogar.Id_Autos = PH_Autos.Id_Autos LEFT OUTER JOIN "
	sql = sql & " PH_AguasNegras ON PH_PanelHogar.Id_AguasNegras = PH_AguasNegras.Id_AguasNegras LEFT OUTER JOIN "
	sql = sql & " PH_AseoUrbano ON PH_PanelHogar.Id_AseoUrbano = PH_AseoUrbano.Id_AseoUrbano LEFT OUTER JOIN "
	sql = sql & " PH_AguasBlancas ON PH_PanelHogar.Id_AguasBlancas = PH_AguasBlancas.Id_AguasBlancas LEFT OUTER JOIN "
	sql = sql & " PH_PuntosLuz ON PH_PanelHogar.id_PuntosLuz = PH_PuntosLuz.Id_PuntosLuz LEFT OUTER JOIN "
	sql = sql & " PH_MontoVivienda ON PH_PanelHogar.Id_MontoVivienda = PH_MontoVivienda.Id_MontoVivienda LEFT OUTER JOIN "
	sql = sql & " PH_OcupacionVivienda ON PH_PanelHogar.Id_OcupacionVivienda = PH_OcupacionVivienda.Id_OcupacionVivienda LEFT OUTER JOIN "
	sql = sql & " PH_MetrosVivienda ON PH_PanelHogar.id_Metros = PH_MetrosVivienda.Id_MetrosVivienda LEFT OUTER JOIN "
	sql = sql & " PH_TipoVivienda ON PH_PanelHogar.Id_TipoVivienda = PH_TipoVivienda.Id_TipoVivienda "
	sql = sql & " WHERE "
	sql = sql & " PH_PanelHogar.Ind_Activo = 1 "
	'sql = sql & " and PH_PanelHogar.Id_PanelHogar = 609 "
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
				Response.write "<td>TipoVivienda</td>"
				Response.write "<td>PesoTipoVivienda</td>"
				Response.write "<td>MetrosVivienda</td>"
				Response.write "<td>PesoMetrosVivienda</td>"
				Response.write "<td>NumeroAmbientes</td>"
				Response.write "<td>PesoNumAmbientes</td>"
				Response.write "<td>TotalPersonas</td>"
				Response.write "<td>PesoTotalPersonas</td>"
				Response.write "<td>NumeroBanos</td>"
				Response.write "<td>PesoNumBanos</td>"
				Response.write "<td>OcupacionVivienda</td>"
				Response.write "<td>Peso AS PesoOcupacionVivienda</td>"
				Response.write "<td>MontoVivienda</td>"
				Response.write "<td>PesoMontoVivienda</td>"
				Response.write "<td>PuntosLuz</td>"
				Response.write "<td>PesoPuntoLuz</td>"
				Response.write "<td>AguasBlancas</td>"
				Response.write "<td>PesoAguasBlancas</td>"
				Response.write "<td>AguasNegras</td>"
				Response.write "<td>PesoAguasNegras</td>"
				Response.write "<td>AseoUrbano</td>"
				Response.write "<td>PesoAseoUrbano</td>"
				Response.write "<td>Id_ServicioElectricidad</td>"
				Response.write "<td>PesoServicioElectricidad</td>"
				Response.write "<td>Id_DomesticaFija</td>"
				Response.write "<td>Id_PersonalLabores</td>"
				Response.write "<td>Id_DomesticaDia</td>"
				Response.write "<td>PesoPeronalAseo</td>"
				Response.write "<td>BandaAncha</td>"
				Response.write "<td>PesoBandaAncha</td>"
				Response.write "<td>Movil</td>"
				Response.write "<td>PesoMovil</td>"
				Response.write "<td>DialUp</td>"
				Response.write "<td>PesoDialUp</td>"
				Response.write "<td>CelularJefe</td>"
				Response.write "<td>PesoCelularJefe</td>"
				Response.write "<td>AireAcondicionado</td>"
				Response.write "<td>PesoAireAcondicionado</td>"
				Response.write "<td>CalentadorElectrico</td>"
				Response.write "<td>PesoCalentadorElectrico</td>"
				Response.write "<td>CalectadorGas</td>"
				Response.write "<td>PesoCalentadorGas</td>"
				Response.write "<td>PC</td>"
				Response.write "<td>PesoPC</td>"
				Response.write "<td>Laptop</td>"
				Response.write "<td>PesoLaptop</td>"
				Response.write "<td>DVD</td>"
				Response.write "<td>PesoDVD</td>"
				Response.write "<td>HomeTheater</td>"
				Response.write "<td>PesoHomeTheater</td>"
				Response.write "<td>VideoJuego</td>"
				Response.write "<td>PesoVideoJuego</td>"
				Response.write "<td>HornoMicro</td>"
				Response.write "<td>PesoHornoMicro</td>"
				Response.write "<td>LavadoraAuto</td>"
				Response.write "<td>PesoLavadoraAuto</td>"
				Response.write "<td>LavadoraSemi</td>"
				Response.write "<td>PesoLavadoraSemi</td>"
				Response.write "<td>LavadoraRodillo</td>"
				Response.write "<td>PesoLavadoraRodillo</td>"
				Response.write "<td>Secadora</td>"
				Response.write "<td>PesoSecadora</td>"
				Response.write "<td>Nevera</td>"
				Response.write "<td>PesoNevera</td>"
				Response.write "<td>Freezer</td>"
				Response.write "<td>PesoFreezer</td>"
				Response.write "<td>CocinaElectrica</td>"
				Response.write "<td>PesoCocinaElectrica</td>"
				Response.write "<td>CocinaBombona</td>"
				Response.write "<td>PesoCocinaBombona</td>"
				Response.write "<td>ConinaGasDirecto</td>"
				Response.write "<td>PesoCocinaGasDirecto</td>"
				Response.write "<td>CocinaKerosene</td>"
				Response.write "<td>PesoCocinaKerosene</td>"
				Response.write "<td>Lavaplatos</td>"
				Response.write "<td>PesoLavaplatos</td>"
				Response.write "<td>Vehiculos</td>"
				Response.write "<td>PesoVehiculo</td>"
				Response.write "<td>Televisores</td>"
				Response.write "<td>PesoTelevisores</td>"
				Response.write "<td>SeguroParticular</td>"
				Response.write "<td>PesoSeguroParticular</td>"
				Response.write "<td>SeguroColectivo</td>"
				Response.write "<td>PesoSeguroColectivo</td>"
				Response.write "<td>SeguroSocial</td>"
				Response.write "<td>PesoSeguroSocial</td>"
				Response.write "<td>SeguroVehiculo</td>"
				Response.write "<td>PesoSeguroVehiculo</td>"
				Response.write "<td>Educacion</td>"
				Response.write "<td>PesoEducacion</td>"
				Response.write "<td>TipoIngreso</td>"
				Response.write "<td>PesoTipoIngreso</td>"
				Response.write "<td>TotalPeso</td>"
				Response.write "<td>ClaseSocial</td>"
			Response.write "</tr>"
		
		for iReg = 0 to ubound(gDatosSol,2)
			idHogar = cint(gDatosSol(0,iReg))
			TotalClase = 0
			Response.write "<tr>"
				Response.write "<td>" & gDatosSol(0,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(1,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(2,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(3,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(4,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(5,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(6,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(7,iReg) & "</td>"
				if isnull(gDatosSol(7,iReg)) then
					iPeso = 0
				else
					iPeso = gDatosSol(7,iReg)
				end if
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(8,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(9,iReg) & "</td>"
				if isnull(gDatosSol(9,iReg)) then
					iPeso = 0
				else
					iPeso = gDatosSol(9,iReg)
				end if
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(10,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(10,iReg)) then
					Valor = 0
				else
					ix = int(gDatosSol(10,iReg))
					if ix <= 2 then Valor = "0,8"
					if ix >= 3 and ix <= 4 then Valor = "1,6"
					if ix >= 5 and ix <= 6 then Valor = "2,4"
					if ix >= 7 and ix <= 8 then Valor = "3,2"
					if ix >= 9 then Valor = "4"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(12,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(12,iReg)) then
					Valor = 0
				else
					ix = int(gDatosSol(12,iReg))
					if ix <= 2 then Valor = "0,8"
					if ix >= 3 and ix <= 4 then Valor = "1,6"
					if ix >= 5 and ix <= 6 then Valor = "2,4"
					if ix >= 7 and ix <= 8 then Valor = "3,2"
					if ix >= 9 then Valor = "4"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(14,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(14,iReg)) or gDatosSol(14,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(14,iReg))
					if ix = 1 then Valor = "0,8"
					if ix = 2 then Valor = "1,6"
					if ix = 3 then Valor = "2,4"
					if ix = 4 then Valor = "3,2"
					if ix > 4 then Valor = "4"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(16,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(17,iReg) & "</td>"
				if isnull(gDatosSol(17,iReg)) then
					iPeso = 0
				else
					iPeso = gDatosSol(17,iReg)
				end if
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(18,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(19,iReg) & "</td>"
				if isnull(gDatosSol(19,iReg)) then
					iPeso = 0
				else
					iPeso = gDatosSol(19,iReg)
				end if
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(20,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(21,iReg) & "</td>"
				if isnull(gDatosSol(21,iReg)) then
					iPeso = 0
				else
					iPeso = gDatosSol(21,iReg)
				end if
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(22,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(23,iReg) & "</td>"
				if isnull(gDatosSol(23,iReg)) then
					iPeso = 0
				else
					iPeso = gDatosSol(23,iReg)
				end if
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(24,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(25,iReg) & "</td>"
				if isnull(gDatosSol(25,iReg)) then
					iPeso = 0
				else
					iPeso = gDatosSol(25,iReg)
				end if
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(26,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(27,iReg) & "</td>"
				if isnull(gDatosSol(27,iReg)) then
					iPeso = 0
				else
					iPeso = gDatosSol(27,iReg)
				end if
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(28,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(28,iReg)) or gDatosSol(28,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(28,iReg))
					if ix = 1 then Valor = "0,5"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(30,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(31,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(32,iReg) & "</td>"
				Valor = 0
				Valor1 = 0
				Valor2 = 0
				Valor3 = 0
				if isnull(gDatosSol(30,iReg)) or gDatosSol(30,iReg) = "" then
					Valor1 = 0
				else
					ix = int(gDatosSol(30,iReg))
					if ix = 1 then Valor1 = "1"
					if ix = 2 then Valor1 = "0"
				end if
				if isnull(gDatosSol(31,iReg)) or gDatosSol(31,iReg) = "" then
					Valor2 = 0
				else
					ix = int(gDatosSol(31,iReg))
					if ix = 1 then Valor2 = "1"
					if ix = 2 then Valor2 = "0"
				end if
				if isnull(gDatosSol(32,iReg)) or gDatosSol(32,iReg) = "" then
					Valor3 = 0
				else
					ix = int(gDatosSol(32,iReg))
					if ix = 1 then Valor3 = "1"
					if ix = 2 then Valor3 = "0"
				end if
				Valor = 0 
				ValorT = int(Valor1) + int(Valor2) + int(Valor3)
				if ValorT = 3 Then  Valor = "3"
				if ValorT = 2 Then  Valor = "1,5"
				if ValorT = 1 Then  Valor = "1,5"
				if ValorT = 0 Then  Valor = "0"
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(34,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(34,iReg)) or gDatosSol(34,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(34,iReg))
					if ix = 1 then Valor = "1"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(36,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(36,iReg)) or gDatosSol(36,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(36,iReg))
					if ix = 1 then Valor = "0,7"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(38,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(38,iReg)) or gDatosSol(38,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(38,iReg))
					if ix = 1 then Valor = "1"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(40,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(40,iReg)) or gDatosSol(40,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(40,iReg))
					if ix = 1 then Valor = "2"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(42,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(42,iReg)) or gDatosSol(42,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(42,iReg))
					if ix = 1 then Valor = "0,5"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(44,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(44,iReg)) or gDatosSol(44,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(44,iReg))
					if ix = 1 then Valor = "0,5"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(46,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(46,iReg)) or gDatosSol(46,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(46,iReg))
					if ix = 1 then Valor = "0,3"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(48,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(48,iReg)) or gDatosSol(48,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(48,iReg))
					if ix = 1 then Valor = "0,5"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(50,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(50,iReg)) or gDatosSol(50,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(50,iReg))
					if ix = 1 then Valor = "0,6"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(52,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(52,iReg)) or gDatosSol(52,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(52,iReg))
					if ix = 1 then Valor = "0,3"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(54,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(54,iReg)) or gDatosSol(54,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(54,iReg))
					if ix = 1 then Valor = "0,1"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(56,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(56,iReg)) or gDatosSol(56,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(56,iReg))
					if ix = 1 then Valor = "0,3"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(58,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(58,iReg)) or gDatosSol(58,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(58,iReg))
					if ix = 1 then Valor = "0,3"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(60,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(60,iReg)) or gDatosSol(60,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(60,iReg))
					if ix = 1 then Valor = "0,4"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(62,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(62,iReg)) or gDatosSol(62,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(62,iReg))
					if ix = 1 then Valor = "0,3"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(64,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(64,iReg)) or gDatosSol(64,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(64,iReg))
					if ix = 1 then Valor = "0,1"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(66,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(66,iReg)) or gDatosSol(66,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(66,iReg))
					if ix = 1 then Valor = "0,6"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(68,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(68,iReg)) or gDatosSol(68,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(68,iReg))
					if ix = 1 then Valor = "0,3"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(70,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(70,iReg)) or gDatosSol(70,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(70,iReg))
					if ix = 1 then Valor = "0,9"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(72,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(72,iReg)) or gDatosSol(72,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(72,iReg))
					if ix = 1 then Valor = "0,9"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(74,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(74,iReg)) or gDatosSol(74,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(74,iReg))
					if ix = 1 then Valor = "0,3"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(76,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(76,iReg)) or gDatosSol(76,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(76,iReg))
					if ix = 1 then Valor = "0,6"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(78,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(78,iReg)) or gDatosSol(78,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(78,iReg))
					if ix = 1 then Valor = "0"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(80,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(80,iReg)) or gDatosSol(80,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(80,iReg))
					if ix = 1 then Valor = "0,6"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(82,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(83,iReg) & "</td>"
				if isnull(gDatosSol(83,iReg)) then
					iPeso = 0
				else
					iPeso = gDatosSol(83,iReg)
				end if
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(84,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(85,iReg) & "</td>"
				if isnull(gDatosSol(85,iReg)) then
					iPeso = 0
				else
					iPeso = gDatosSol(85,iReg)
				end if
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(86,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(86,iReg)) or gDatosSol(86,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(86,iReg))
					if ix = 1 then Valor = "5"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(88,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(88,iReg)) or gDatosSol(88,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(88,iReg))
					if ix = 1 then Valor = "3,3"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(90,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(90,iReg)) or gDatosSol(90,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(90,iReg))
					if ix = 1 then Valor = "1,7"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				Response.write "<td>" & gDatosSol(92,iReg) & "</td>"
				Valor = 0
				if isnull(gDatosSol(92,iReg)) or gDatosSol(92,iReg) = "" then
					Valor = 0
				else
					ix = int(gDatosSol(92,iReg))
					if ix = 1 then Valor = "3"
					if ix = 2 then Valor = "0"
				end if
				Response.write "<td>" & Valor & "</td>"
				iPeso = replace(Valor,",",".")
				iPeso = Valor
				TotalClase = TotalClase + cDbl(iPeso)
				dim gDatosSol2
				dim rsx2
				set rsx2 = CreateObject("ADODB.Recordset")
				rsx2.CursorType = adOpenKeyset 
				rsx2.LockType = 2 'adLockOptimistic 
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " PH_Educacion.Educacion, " '94 0
				sql = sql & " PH_Educacion.Peso, " '95 1
				sql = sql & " PH_TipoIngreso.TipoIngreso, " '96 2
				sql = sql & " PH_TipoIngreso.Peso " '97 3
				sql = sql & " FROM (PH_Panelistas LEFT JOIN PH_Educacion ON PH_Panelistas.Id_Educacion = PH_Educacion.Id_Educacion) LEFT JOIN PH_TipoIngreso ON PH_Panelistas.Id_TipoIngreso = PH_TipoIngreso.Id_TipoIngreso "
				sql = sql & " WHERE "
				sql = sql & " PH_Panelistas.Id_Hogar = " & gDatosSol(0,iReg)
				sql = sql & " AND PH_Panelistas.Id_Parentesco = 1 "
				'response.write "<br>36 sql:=" & sql
				'response.end
				rsx2.Open sql ,conexion
				if rsx2.eof then
					Response.write "<td></td>"
					Response.write "<td></td>"
					Response.write "<td></td>"
					Response.write "<td></td>"
				else
					gDatosSol2 = rsx2.GetRows
					rsx2.close
					Response.write "<td>" & gDatosSol2(0,0) & "</td>"
					Response.write "<td>" & gDatosSol2(1,0) & "</td>"
					if isnull(gDatosSol2(1,0)) then
						iPeso = 0
					else
						iPeso = gDatosSol2(1,0)
					end if
					TotalClase = TotalClase + cDbl(iPeso)
					Response.write "<td>" & gDatosSol2(2,0) & "</td>"
					Response.write "<td>" & gDatosSol2(3,0) & "</td>"
					if isnull(gDatosSol2(3,0)) then
						iPeso = 0
					else
						iPeso = gDatosSol2(3,0)
					end if
					TotalClase = TotalClase + cDbl(iPeso)
				end if
				Response.write "<td>" & TotalClase & "</td>"
				Clase = ""
				if TotalClase >= 56 then Clase = "ABC+"
				if TotalClase => 40 and TotalClase < 56 then Clase = "C"
				if TotalClase > 20 and TotalClase < 40 then Clase = "D"
				if TotalClase <= 20 then Clase = "E"
				Response.write "<td>" & Clase & "</td>"
				'Response.write "<td>" & TotalClase & "</td>"
				
				'Actualizar Clase Social
				dim rsx3
				set rsx3 = CreateObject("ADODB.Recordset")
				rsx3.CursorType = 0
				rsx3.LockType = 3
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Id_PanelHogar, "
				sql = sql & " ClaseSocial "
				sql = sql & " FROM "
				sql = sql & " PH_PanelHogar "
				sql = sql & " WHERE "
				sql = sql & " Id_PanelHogar = " & idHogar
				'response.write "<br>57 sql:=" & sql
				'response.end
				rsx3.Open sql ,conexion
				rsx3("ClaseSocial") = Clase
				rsx3.Update
				rsx3.Close 
				set rsx3 = nothing 
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