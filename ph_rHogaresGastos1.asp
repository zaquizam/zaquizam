<!DOCTYPE HTML>
<html >
<head>
	<title>Gastos por Hogar</title>
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="mensaje.css" rel="stylesheet" type="text/css" media="screen" />
	<link rel="icon" href="favicon.ico" type="image/x-icon"> 
	<script type="text/javascript" src="js/sweetalert.min.js"></script>
	<script type="text/javascript" src="js/jquery-1.12.4.min.js"></script>
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
	Server.ScriptTimeout=1000
	Response.buffer = true


    Apertura
	   
	dim gDatosSol44
	dim gDatosSol45
	dim gDatosSol46
	dim gDatosSol47
	dim rsx44
	dim rsx45
	dim rsx46
	dim rsx47

	dim IdArea


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
	'sql = sql & " and PH_PanelHogar.Id_PanelHogar = 1322 "
	sql = sql & " and PH_GArea.Id_Area  = " & IdArea
	if ed_sPar(2,0) <> "Seleccionar" and ed_sPar(2,0) <> "" then
		idEstado = cint(ed_sPar(2,0))
		sql = sql & " and ss_Estado.Id_Estado =  " & idEstado
		Response.ContentType = "application/vnd.ms-excel"
		Response.AddHeader "Content-disposition","attachment; filename=tem.xls"
	else
		if IdArea = 2 then
			response.end
		end if
	end if
	'sql = sql & " and PH_PanelHogar.Id_PanelHogar = 313 "
	sql = sql & " Order By "
	sql = sql & " PH_PanelHogar.Id_PanelHogar "
	'response.write "<br>36 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	iExiste = 0
	if rsx1.eof then
	else
		if IdArea <> 2 then 
			'Response.ContentType = "application/vnd.ms-excel"
			'Response.AddHeader "Content-disposition","attachment; filename=tem.xls"
		end if
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
				iReg2 = cint(ubound(gDatosSol,2)+1)
				Response.write "<td>" & iReg+1 & "/" & iReg2 &  "</td>"
				Response.write "<td>" & idHogar & "</td>"
				Codhogar = gDatosSol(1,iReg)
				Estado = gDatosSol(2,iReg)
				Ciudad = gDatosSol(3,iReg)
				Municipio = gDatosSol(4,iReg)
				Parroquia = gDatosSol(5,iReg)
				Clase =  gDatosSol(6,iReg)
				Response.write "<td>" & Codhogar & "</td>"
				Response.write "<td>" & Estado & "</td>"
				Response.write "<td>" & Ciudad & "</td>"
				Response.write "<td>" & Municipio & "</td>"
				Response.write "<td>" & Parroquia & "</td>"
				Response.write "<td>" & Clase & "</td>"
			
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
<script type="text/javascript">
	function GenerarExcel()
	{
		num = document.getElementById("Excel").value;
		//alert("Generar Excel:="+ num);
		window.open("ph_rEncuestasRealizadasExcel.asp?num=" +num,"_blank");
	}

	function GenerarExcel1()
	{
		num = document.getElementById("Excel").value;
		//alert("Generar Excel1:="+ num);
		//return;
		window.open("ph_rEncuestasTotalesExcel.asp?num=" +num,"_blank");
	}

	function GenerarExcelFaltantes()
	{
		num = document.getElementById("Excel").value;
		//alert("Generar Excel1:="+ num);
		//return;
		window.open("ph_rEncuestasHogaresFaltanteExcel.asp?num=" +num,"_blank");
	}
	
	function alerta(total) 
	{
		swal("Se Enviaron Encuestas " + total + " Hogares ","Enviado","success");
		//window.open("?edpas=1&smenu=?x=1&smenu=Envio%20SMS%20Bienvenida%20y%20Link","_parent");
	}
</script>
<%

Sub Combos
 
	'response.write "<br>372 Combo1:=" & ed_sPar(1,0)
	'response.write " Combo2:=" & ed_sPar(2,0)
	'response.write " Combo3:=" & ed_sPar(3,0)
	'response.write " Combo3:=" & ed_sPar(4,0)
	'response.write " Combo3:=" & ed_sPar(5,0)
    ed_iCombo = 1
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Area, "
	sql = sql & " Area "
	sql = sql & " FROM PH_GArea "
	sql = sql & " Order By "
	sql = sql & " Area "
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(1,0)="Area"
    ed_sCombo(1,1)=sql 
    ed_sCombo(1,2)="Seleccionar"

	if ed_sPar(1,0) <> "Seleccionar" and ed_sPar(1,0) <> "" then
		ed_iCombo = 2
		sql = ""
		sql = sql & " SELECT "
		sql = sql & " ss_Estado.Id_Estado, "
		sql = sql & " ss_Estado.Estado "
		sql = sql & " FROM (PH_GArea INNER JOIN PH_GAreaEstado ON PH_GArea.Id_Area = PH_GAreaEstado.Id_Area) INNER JOIN ss_Estado ON PH_GAreaEstado.Id_Estado = ss_Estado.Id_Estado "
		sql = sql & " WHERE "
		IdArea = cint(ed_sPar(1,0))
		if IdArea <>2 then ed_iCombo = 1
		sql = sql & " PH_GAreaEstado.Id_Area = " & IdArea
		sql = sql & " Order By "
		sql = sql & " ss_Estado.Estado "
		'response.write "<br>372 Combo1:=" & sql
		ed_sCombo(2,0)="Estado"
		ed_sCombo(2,1)=sql 
		ed_sCombo(2,2)="Seleccionar"
	end if
End Sub

   
'==========================================================================================
' Parámetros del Manteniemiento
'==========================================================================================
    LeePar
  
    
    if ed_iPas<>4 then 
        Encabezado
    end if    

	'response.write "llego1"
	'response.end
	'ParDat
%>
		
	<br>
	<br>
	<br>
	<div style="width:98%"></div></center>
<%
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
	</br>
	</br>

<%
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	if ed_sPar(1,0) <> "Seleccionar" then
		IdArea = cint(ed_sPar(1,0))
		VerData
	end if
	'response.end 
%>
<style>


</style>	

    <%conexion.close%>
	
</body>
</html>