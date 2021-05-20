<!DOCTYPE HTML>
<html >
<head>
	<title>Gastos por Hogar</title>
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
	Server.ScriptTimeout=1000
	Response.buffer = true
  
	dim gDatosSol44
	dim rsx44
	dim rsx45
	dim rsx46
	dim rsx47
    

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


sub VerData	
	dim gDatosSol
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	sql = ""
    sql = sql & " SELECT "
	sql = sql & " PH_PanelHogar.Id_PanelHogar AS idHogar, "
	sql = sql & " PH_PanelHogar.CodigoHogar, "
	sql = sql & " ss_Estado.Estado, "
	sql = sql & " PH_Ciudad.Ciudad, "
	sql = sql & " ss_Municipio.Municipio, "
	sql = sql & " ss_Parroquia.Parroquia, "
	sql = sql & " PH_PanelHogar.ClaseSocial, "
	sql = sql & " PH_GArea.Area, "
	sql = sql & " PH_GArea.Id_Area "
	sql = sql & " FROM PH_PanelHogar INNER JOIN "
	sql = sql & " PH_Ciudad ON PH_PanelHogar.Id_Ciudad = PH_Ciudad.Id_Ciudad INNER JOIN "
	sql = sql & " ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado INNER JOIN "
	sql = sql & " ss_Municipio ON PH_PanelHogar.Id_Municipio = ss_Municipio.Id_Municipio INNER JOIN "
	sql = sql & " ss_Parroquia ON PH_PanelHogar.Id_Parroquia = ss_Parroquia.Id_Parroquia INNER JOIN "
	sql = sql & " PH_GAreaEstado ON ss_Estado.Id_Estado = PH_GAreaEstado.Id_Estado INNER JOIN "
	sql = sql & " PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area "
	sql = sql & " WHERE "
	sql = sql & " PH_PanelHogar.Ind_Activo = 1 "
	'sql = sql & " and PH_PanelHogar.Id_PanelHogar = 706 "
	sql = sql & " and PH_GArea.Id_Area  = 1 "
	'sql = sql & " and PH_PanelHogar.Id_PanelHogar = 313 "
	sql = sql & " Order By "
	sql = sql & " PH_PanelHogar.Id_PanelHogar "
	'response.write "<br>36 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	iExiste = 0
	if rsx1.eof then
	else
		'Response.ContentType = "application/vnd.ms-excel"
		'Response.AddHeader "Content-disposition","attachment; filename=tem.xls"
		gDatosSol = rsx1.GetRows
		rsx1.close
		'set rsx1 = nothing
		Response.write "<table border=1>"
			Response.write "<tr>"
				Response.write "<td>Reg.</td>"
				Response.write "<td>idHogar</td>"
				Response.write "<td>CodigoHogar</td>"
				Response.write "<td>Estado</td>"
				Response.write "<td>Ciudad</td>"
				Response.write "<td>Municipio</td>"
				Response.write "<td>Parroquia</td>"
				Response.write "<td>ClaseSocial</td>"
				Response.write "<td>Gasto Ene21</td>"
				Response.write "<td>Gasto Feb21</td>"
				Response.write "<td>Gasto Mar21</td>"
				Response.write "<td>Gasto Abr21</td>" 
			Response.write "</tr>"
		
		for iReg = 0 to ubound(gDatosSol,2)
			idHogar = cint(gDatosSol(0,iReg))
			TotalClase = 0
			Response.write "<tr>"
				Response.write "<td>" & iReg+1 & "/" & ubound(gDatosSol,2)+1 &  "</td>"
				Response.write "<td>" & gDatosSol(0,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(1,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(2,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(3,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(4,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(5,iReg) & "</td>"
				Response.write "<td>" & gDatosSol(6,iReg) & "</td>"
				'Gastos Enero
				set rsx44 = CreateObject("ADODB.Recordset")
				rsx44.CursorType = adOpenKeyset 
				rsx44.LockType = 2 'adLockOptimistic 
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " PH_Consumo.Id_Semana, "
				sql = sql & " PH_Consumo.Id_Hogar, "
				sql = sql & " PH_Consumo.Id_Moneda, "
				sql = sql & " PH_Consumo.fecha_consumo, "
				sql = sql & " PH_Consumo.Total_Compra, "
				sql = sql & " ss_Semana_1.Dolar, "
				sql = sql & " ss_Semana_1.Euro, "
				sql = sql & " ss_Semana_1.Petro, "
				sql = sql & " ss_Semana_1.Peso "
				sql = sql & " FROM (ss_Semana INNER JOIN PH_Consumo ON ss_Semana.IdSemana = PH_Consumo.Id_Semana) INNER JOIN ss_Semana AS ss_Semana_1 ON PH_Consumo.Id_Semana = ss_Semana_1.IdSemana "
				sql = sql & " WHERE "
				sql = sql & " PH_Consumo.Id_Semana in(16,17,18,19) "
				sql = sql & " AND PH_Consumo.Id_Hogar = " & idHogar
				'response.write "<br>36 sql:=" & sql
				'response.end
				rsx44.Open sql ,conexion
				if rsx44.eof then
					Response.write "<td></td>"
				else
					gDatosSol44 = rsx44.GetRows
					rsx44.close
					TotGasto = 0
					for iReg44 = 0 to ubound(gDatosSol44,2)
						idMoneda = cint(gDatosSol44(2,iReg44))
						idCompra = cdbl(gDatosSol44(4,iReg44))
						idDolar = cdbl(gDatosSol44(5,iReg44))
						idEuro = cdbl(gDatosSol44(6,iReg44))
						idPetro = cdbl(gDatosSol44(7,iReg44))
						idPeso = cdbl(gDatosSol44(8,iReg44))
						select case idMoneda
							case 1 'Dolar
								TotGasto = TotGasto + idCompra
							case 2 'Bs
								TotGasto = TotGasto + (idCompra/idDolar)
							case 3 'Petro
								TotGasto = TotGasto + (idCompra/idPetro)
							case 4 'Euro
								TotGasto = TotGasto + (idCompra/idEuro)
							case 5 'Peso
								TotGasto = TotGasto + (idCompra/idPeso)
						end select 
					next
					Response.write "<td>" & formatnumber(TotGasto,2) & "</td>"
				end if
				set rsx44 = nothing
				set rsx45 = CreateObject("ADODB.Recordset")
				rsx45.CursorType = adOpenKeyset 
				rsx45.LockType = 2 'adLockOptimistic 
				'Gastos Febrero
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " PH_Consumo.Id_Semana, "
				sql = sql & " PH_Consumo.Id_Hogar, "
				sql = sql & " PH_Consumo.Id_Moneda, "
				sql = sql & " PH_Consumo.fecha_consumo, "
				sql = sql & " PH_Consumo.Total_Compra, "
				sql = sql & " ss_Semana_1.Dolar, "
				sql = sql & " ss_Semana_1.Euro, "
				sql = sql & " ss_Semana_1.Petro, "
				sql = sql & " ss_Semana_1.Peso "
				sql = sql & " FROM (ss_Semana INNER JOIN PH_Consumo ON ss_Semana.IdSemana = PH_Consumo.Id_Semana) INNER JOIN ss_Semana AS ss_Semana_1 ON PH_Consumo.Id_Semana = ss_Semana_1.IdSemana "
				sql = sql & " WHERE "
				sql = sql & " PH_Consumo.Id_Semana in(20,21,22,23) "
				sql = sql & " AND PH_Consumo.Id_Hogar = " & idHogar
				'response.write "<br>36 sql:=" & sql
				'response.end
				rsx45.Open sql ,conexion
				if rsx45.eof then
					Response.write "<td></td>"
				else
					gDatosSol45 = rsx45.GetRows
					rsx45.close
					TotGasto = 0
					for iReg45 = 0 to ubound(gDatosSol45,2)
						idMoneda = cint(gDatosSol45(2,iReg45))
						idCompra = cdbl(gDatosSol45(4,iReg45))
						idDolar = cdbl(gDatosSol45(5,iReg45))
						idEuro = cdbl(gDatosSol45(6,iReg45))
						idPetro = cdbl(gDatosSol45(7,iReg45))
						idPeso = cdbl(gDatosSol45(8,iReg45))
						select case idMoneda
							case 1 'Dolar
								TotGasto = TotGasto + idCompra
							case 2 'Bs
								TotGasto = TotGasto + (idCompra/idDolar)
							case 3 'Petro
								TotGasto = TotGasto + (idCompra/idPetro)
							case 4 'Euro
								TotGasto = TotGasto + (idCompra/idEuro)
							case 5 'Peso
								TotGasto = TotGasto + (idCompra/idPeso)
						end select 
					next
					Response.write "<td>" & formatnumber(TotGasto,2) & "</td>"
				end if
				set rsx45 = nothing
				'Gastos Marzo
				set rsx46 = CreateObject("ADODB.Recordset")
				rsx46.CursorType = adOpenKeyset 
				rsx46.LockType = 2 'adLockOptimistic 
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " PH_Consumo.Id_Semana, "
				sql = sql & " PH_Consumo.Id_Hogar, "
				sql = sql & " PH_Consumo.Id_Moneda, "
				sql = sql & " PH_Consumo.fecha_consumo, "
				sql = sql & " PH_Consumo.Total_Compra, "
				sql = sql & " ss_Semana_1.Dolar, "
				sql = sql & " ss_Semana_1.Euro, "
				sql = sql & " ss_Semana_1.Petro, "
				sql = sql & " ss_Semana_1.Peso "
				sql = sql & " FROM (ss_Semana INNER JOIN PH_Consumo ON ss_Semana.IdSemana = PH_Consumo.Id_Semana) INNER JOIN ss_Semana AS ss_Semana_1 ON PH_Consumo.Id_Semana = ss_Semana_1.IdSemana "
				sql = sql & " WHERE "
				sql = sql & " PH_Consumo.Id_Semana in(24,25,26,27,28) "
				sql = sql & " AND PH_Consumo.Id_Hogar = " & idHogar
				'response.write "<br>36 sql:=" & sql
				'response.end
				rsx46.Open sql ,conexion
				if rsx46.eof then
					Response.write "<td></td>"
				else
					gDatosSol46 = rsx46.GetRows
					rsx46.close
					TotGasto = 0
					for iReg46 = 0 to ubound(gDatosSol46,2)
						idMoneda = cint(gDatosSol46(2,iReg46))
						idCompra = cdbl(gDatosSol46(4,iReg46))
						idDolar = cdbl(gDatosSol46(5,iReg46))
						idEuro = cdbl(gDatosSol46(6,iReg46))
						idPetro = cdbl(gDatosSol46(7,iReg46))
						idPeso = cdbl(gDatosSol46(8,iReg46))
						select case idMoneda
							case 1 'Dolar
								TotGasto = TotGasto + idCompra
							case 2 'Bs
								TotGasto = TotGasto + (idCompra/idDolar)
							case 3 'Petro
								TotGasto = TotGasto + (idCompra/idPetro)
							case 4 'Euro
								TotGasto = TotGasto + (idCompra/idEuro)
							case 5 'Peso
								TotGasto = TotGasto + (idCompra/idPeso)
						end select 
					next
					Response.write "<td>" & formatnumber(TotGasto,2) & "</td>"
				end if
				set rsx46 = nothing
				'Gastos Abril
				set rsx47 = CreateObject("ADODB.Recordset")
				rsx47.CursorType = adOpenKeyset 
				rsx47.LockType = 2 'adLockOptimistic 
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " PH_Consumo.Id_Semana, "
				sql = sql & " PH_Consumo.Id_Hogar, "
				sql = sql & " PH_Consumo.Id_Moneda, "
				sql = sql & " PH_Consumo.fecha_consumo, "
				sql = sql & " PH_Consumo.Total_Compra, "
				sql = sql & " ss_Semana_1.Dolar, "
				sql = sql & " ss_Semana_1.Euro, "
				sql = sql & " ss_Semana_1.Petro, "
				sql = sql & " ss_Semana_1.Peso "
				sql = sql & " FROM (ss_Semana INNER JOIN PH_Consumo ON ss_Semana.IdSemana = PH_Consumo.Id_Semana) INNER JOIN ss_Semana AS ss_Semana_1 ON PH_Consumo.Id_Semana = ss_Semana_1.IdSemana "
				sql = sql & " WHERE "
				sql = sql & " PH_Consumo.Id_Semana in(29,30,31,32) "
				sql = sql & " AND PH_Consumo.Id_Hogar = " & idHogar
				'response.write "<br>36 sql:=" & sql
				'response.end
				rsx47.Open sql ,conexion
				if rsx47.eof then
					Response.write "<td></td>"
				else
					gDatosSol47 = rsx47.GetRows
					rsx47.close
					TotGasto = 0
					for iReg44 = 0 to ubound(gDatosSol47,2)
						idMoneda = cint(gDatosSol47(2,iReg47))
						idCompra = cdbl(gDatosSol47(4,iReg47))
						idDolar = cdbl(gDatosSol47(5,iReg47))
						idEuro = cdbl(gDatosSol47(6,iReg47))
						idPetro = cdbl(gDatosSol47(7,iReg47))
						idPeso = cdbl(gDatosSol47(8,iReg47))
						select case idMoneda
							case 1 'Dolar
								TotGasto = TotGasto + idCompra
							case 2 'Bs
								TotGasto = TotGasto + (idCompra/idDolar)
							case 3 'Petro
								TotGasto = TotGasto + (idCompra/idPetro)
							case 4 'Euro
								TotGasto = TotGasto + (idCompra/idEuro)
							case 5 'Peso
								TotGasto = TotGasto + (idCompra/idPeso)
						end select 
					next
					'Response.write "<td>" & icontador & "-" & formatnumber(TotGasto,2) & "</td>"
					Response.write "<td>" & formatnumber(TotGasto,2) & "</td>"
				end if
				set rsx47 = nothing
				
			Response.write "</tr>"
			icontador = icontador + 1
			if icontador > 100 then
				Response.flush 
				icontador = 0
			end if
				
		next
		Response.write "</table>"
	end if
end sub 


%>
<style>


</style>	

    <%conexion.close%>
	
</body>
</html>