<!DOCTYPE HTML>
<html >
<head>
	<title>PH Pen x Cat</title>
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
	<link rel="icon" href="favicon.ico" type="image/x-icon"> 
	<script type="text/javascript" src="js/sweetalert.min.js"></script>
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
		function Mensaje(){
			swal("Atenas Grupo Consultor","Servicio No Contratado","info");
			return;
		}	
		
	function GenerarExcel()
	{
		//alert("Generar Excel");
		//num = document.getElementById("Excel").value;
		//alert("Generar Excel:="+ num);
		window.open("ph_rHomePantryPenCatExcel.asp","_blank");
	}
		

	</script>   
<%

   
    LeePar
  
    
    if ed_iPas<>4 then 
        Encabezado
    end if    

	'response.write "llego1"
	'response.end
    
	dim gDatos1
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	dim gCategorias
	dim gCategoriasEne
	dim gCategoriasFeb
	dim gCategoriasMar
	dim gCategoriasAbr
	dim gCategoriasMay
	dim gCategoriasJun
	dim gCategoriasAcu
	
	dim gHogaresTotalEne
	dim gHogaresTotalFeb
	dim gHogaresTotalMar
	dim gHogaresTotalAbr
	dim gHogaresTotalMay
	dim gHogaresTotalJun
	dim gHogaresTotalAcu
	
	dim gHogaresCategoriaEne
	dim gHogaresCategoriaFeb
	dim gHogaresCategoriaMar
	dim gHogaresCategoriaAbr
	dim gHogaresCategoriaMay
	dim gHogaresCategoriaJun
	dim gHogaresCategoriaAcu
	
	Dim TotalHogaresEne
	Dim TotalHogaresFeb
	Dim TotalHogaresMar
	Dim TotalHogaresAbr
	Dim TotalHogaresMay
	Dim TotalHogaresJun
	Dim TotalHogaresAcu
	
	Dim TotalHogaresCatEne
	Dim TotalHogaresCatFeb
	Dim TotalHogaresCatMar
	Dim TotalHogaresCatAbr
	Dim TotalHogaresCatMay
	Dim TotalHogaresCatJun
	Dim TotalHogaresCatAcu
	
	TotalHogaresEne = 0
	TotalHogaresFeb = 0
	TotalHogaresMar = 0
	TotalHogaresAbr = 0
	TotalHogaresMay = 0
	TotalHogaresJun = 0
	TotalHogaresAcu = 0
	
	'TotalHogaresEne
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " ss_Semana.IdMes, "
	sql = sql & " ss_Semana.IdAno "
	sql = sql & " FROM ss_Semana INNER JOIN PH_DataCruda ON ss_Semana.IdSemana = PH_DataCruda.Id_Semana "
	sql = sql & " WHERE "
	sql = sql & " PH_DataCruda.Id_Fabricante <> 0 "
	sql = sql & " GROUP BY "
	sql = sql & " ss_Semana.IdMes, "
	sql = sql & " ss_Semana.IdAno, "
	sql = sql & " PH_DataCruda.Id_Hogar "
	sql = sql & " HAVING "
	sql = sql & " ss_Semana.IdMes = 1 "
	sql = sql & " AND ss_Semana.IdAno=2021 "
	'response.write "<br>75 sql:=" & sql
	'response.end
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
		TotalHogaresEne = 0
	else
		gHogaresTotalEne = rsx1.GetRows
		rsx1.close
		TotalHogaresEne = ubound(gHogaresTotalEne,2) + 1
	end if

	'TotalHogaresFeb
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " ss_Semana.IdMes, "
	sql = sql & " ss_Semana.IdAno "
	sql = sql & " FROM ss_Semana INNER JOIN PH_DataCruda ON ss_Semana.IdSemana = PH_DataCruda.Id_Semana "
	sql = sql & " WHERE "
	sql = sql & " PH_DataCruda.Id_Fabricante <> 0 "
	sql = sql & " GROUP BY "
	sql = sql & " ss_Semana.IdMes, "
	sql = sql & " ss_Semana.IdAno, "
	sql = sql & " PH_DataCruda.Id_Hogar "
	sql = sql & " HAVING "
	sql = sql & " ss_Semana.IdMes = 2 "
	sql = sql & " AND ss_Semana.IdAno=2021 "
	'response.write "<br>75 sql:=" & sql
	'response.end
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
		TotalHogaresFeb = 0
	else
		gHogaresTotalFeb = rsx1.GetRows
		rsx1.close
		TotalHogaresFeb = ubound(gHogaresTotalFeb,2) + 1
	end if

	'TotalHogaresMar
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " ss_Semana.IdMes, "
	sql = sql & " ss_Semana.IdAno "
	sql = sql & " FROM ss_Semana INNER JOIN PH_DataCruda ON ss_Semana.IdSemana = PH_DataCruda.Id_Semana "
	sql = sql & " WHERE "
	sql = sql & " PH_DataCruda.Id_Fabricante <> 0 "
	sql = sql & " GROUP BY "
	sql = sql & " ss_Semana.IdMes, "
	sql = sql & " ss_Semana.IdAno, "
	sql = sql & " PH_DataCruda.Id_Hogar "
	sql = sql & " HAVING "
	sql = sql & " ss_Semana.IdMes = 3 "
	sql = sql & " AND ss_Semana.IdAno=2021 "
	'response.write "<br>75 sql:=" & sql
	'response.end
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
		TotalHogaresMar = 0
	else
		gHogaresTotalMar = rsx1.GetRows
		rsx1.close
		TotalHogaresMar = ubound(gHogaresTotalMar,2) + 1
	end if

	'TotalHogaresAbr
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " ss_Semana.IdMes, "
	sql = sql & " ss_Semana.IdAno "
	sql = sql & " FROM ss_Semana INNER JOIN PH_DataCruda ON ss_Semana.IdSemana = PH_DataCruda.Id_Semana "
	sql = sql & " WHERE "
	sql = sql & " PH_DataCruda.Id_Fabricante <> 0 "
	sql = sql & " GROUP BY "
	sql = sql & " ss_Semana.IdMes, "
	sql = sql & " ss_Semana.IdAno, "
	sql = sql & " PH_DataCruda.Id_Hogar "
	sql = sql & " HAVING "
	sql = sql & " ss_Semana.IdMes = 4 "
	sql = sql & " AND ss_Semana.IdAno=2021 "
	'response.write "<br>75 sql:=" & sql
	'response.end
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
		TotalHogaresAbr = 0
	else
		gHogaresTotalAbr = rsx1.GetRows
		rsx1.close
		TotalHogaresAbr = ubound(gHogaresTotalAbr,2) + 1
	end if

	'TotalHogaresMay
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " ss_Semana.IdMes, "
	sql = sql & " ss_Semana.IdAno "
	sql = sql & " FROM ss_Semana INNER JOIN PH_DataCruda ON ss_Semana.IdSemana = PH_DataCruda.Id_Semana "
	sql = sql & " WHERE "
	sql = sql & " PH_DataCruda.Id_Fabricante <> 0 "
	sql = sql & " GROUP BY "
	sql = sql & " ss_Semana.IdMes, "
	sql = sql & " ss_Semana.IdAno, "
	sql = sql & " PH_DataCruda.Id_Hogar "
	sql = sql & " HAVING "
	sql = sql & " ss_Semana.IdMes = 5 "
	sql = sql & " AND ss_Semana.IdAno=2021 "
	'response.write "<br>75 sql:=" & sql
	'response.end
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
		TotalHogaresMay = 0
	else
		gHogaresTotalMay = rsx1.GetRows
		rsx1.close
		TotalHogaresMay = ubound(gHogaresTotalMay,2) + 1
	end if

	'TotalHogaresJun
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " ss_Semana.IdMes, "
	sql = sql & " ss_Semana.IdAno "
	sql = sql & " FROM ss_Semana INNER JOIN PH_DataCruda ON ss_Semana.IdSemana = PH_DataCruda.Id_Semana "
	sql = sql & " WHERE "
	sql = sql & " PH_DataCruda.Id_Fabricante <> 0 "
	sql = sql & " GROUP BY "
	sql = sql & " ss_Semana.IdMes, "
	sql = sql & " ss_Semana.IdAno, "
	sql = sql & " PH_DataCruda.Id_Hogar "
	sql = sql & " HAVING "
	sql = sql & " ss_Semana.IdMes = 6 "
	sql = sql & " AND ss_Semana.IdAno=2021 "
	'response.write "<br>75 sql:=" & sql
	'response.end
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
		TotalHogaresJun = 0
	else
		gHogaresTotalJun = rsx1.GetRows
		rsx1.close
		TotalHogaresJun = ubound(gHogaresTotalJun,2) + 1
	end if

	
	'TotalHogaresAcu
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " PH_DataCruda.Id_Hogar "
	sql = sql & " FROM ss_Semana INNER JOIN PH_DataCruda ON ss_Semana.IdSemana = PH_DataCruda.Id_Semana "
	sql = sql & " WHERE "
	sql = sql & " PH_DataCruda.Id_Fabricante <> 0 "
	sql = sql & " GROUP BY "
	sql = sql & " PH_DataCruda.Id_Hogar "

	'response.write "<br>75 sql:=" & sql
	'response.end
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
		TotalHogaresAcu = 0
	else
		gHogaresTotalAcu = rsx1.GetRows
		rsx1.close
		TotalHogaresAcu = ubound(gHogaresTotalAcu,2) + 1
	end if

	'Categorias
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " PH_DataCruda.Id_Categoria, "
	sql = sql & " PH_DataCruda.Categoria "
	sql = sql & " FROM PH_DataCruda "
	sql = sql & " WHERE "
	sql = sql & " PH_DataCruda.Id_Fabricante <> 0 "
	sql = sql & " GROUP BY "
	sql = sql & " PH_DataCruda.Id_Categoria, "
	sql = sql & " PH_DataCruda.Categoria "
	sql = sql & " ORDER BY "
	sql = sql & " PH_DataCruda.Categoria "
	'response.write "<br>104 sql:=" & sql
	'response.end
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gCategorias = rsx1.GetRows
		rsx1.close
	end if

%>
	<div class="container-fluid">        
		<div class="row">
			<!--Contenido Generalhidden-->			
			<div class="container">
				<div class="col-md-8 col-sm-8 col-xs-12">
					<div class="pull-right">
						<img src="images/Excel.png"  style="margin-left:0px;" title="Generar Excel" alt="PDF" width="70px" onclick="GenerarExcel()"/>
						<input type="hidden" name="Excel" id="Excel" align="right" size=0 value='<%=sExcel%>'>
					</div>
				</div>
			</div>
		</div>
	</div>
	<br>
	<div style="width:98%">
		<div id="DivHomePartySem">
			<div class="ex1">
				<table class="w3-table w3-striped w3-bordered w3-card-4 w3-small" style="width:1000px; margin-left:auto; margin-right:auto;margin-top:10px ">
					<thead>
						<tr class="w3-blue">
							<th>IdCategoria</th>
							<th>Categoria</th>
							<th>Penetración <br>Enero 2021</th>
							<th>Penetración <br>Febrero 2021</th>
							<th>Penetración <br>Marzo 2021</th>
							<th>Penetración <br>Abril 2021</th>
							<th>Penetración <br>Mayo 2021</th>
							<th>Penetración <br>Jun 2021</th>
							<th>Acumulado <br>2021</th>
						</tr>
					</thead>
					<%
					for iPro = 0 to  ubound(gCategorias,2)
						response.write "<tr>"
							response.write "<td>"
								response.write gCategorias(0,iPro)
							response.write "</td>"
							response.write "<td>"
								response.write gCategorias(1,iPro)
							response.write "</td>"
							iCat = gCategorias(0,iPro)
							
							'Penetración Enero
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " PH_DataCruda.Id_Hogar "
							sql = sql & " FROM PH_DataCruda INNER JOIN ss_Semana ON PH_DataCruda.Id_Semana = ss_Semana.IdSemana "
							sql = sql & " WHERE "
							sql = sql & " ss_Semana.IdMes = 1 "
							sql = sql & " AND ss_Semana.IdAno = 2021 "
							sql = sql & " AND PH_DataCruda.Id_Categoria = " & iCat
							sql = sql & " GROUP BY "
							sql = sql & " PH_DataCruda.Id_Hogar "
							'if iCat = 91 then
							'	response.write "<br>237 sql:=" & sql
							'end if
							'response.write "<br>190 sql:=" & sql
							'response.end
							rsx1.Open sql ,conexion
							if rsx1.eof then
								rsx1.close
								TotalHogaresCatEne = 0
							else
								gHogaresCategoriaEne = rsx1.GetRows
								rsx1.close
								TotalHogaresCatEne = ubound(gHogaresCategoriaEne,2) + 1
							end if
							PenetracionEne = 0
							response.write "<td>"
								PenetracionEne = (TotalHogaresCatEne * 100) / TotalHogaresEne
								PenetracionEne = FormatNumber(PenetracionEne,2)
								response.write PenetracionEne
								response.write "<br>(" & TotalHogaresCatEne & "-" & TotalHogaresEne & ")"
							response.write "</td>"
							
							'Penetración Febrero
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " PH_DataCruda.Id_Hogar "
							sql = sql & " FROM PH_DataCruda INNER JOIN ss_Semana ON PH_DataCruda.Id_Semana = ss_Semana.IdSemana "
							sql = sql & " WHERE "
							sql = sql & " ss_Semana.IdMes = 2 "
							sql = sql & " AND ss_Semana.IdAno = 2021 "
							sql = sql & " AND PH_DataCruda.Id_Categoria = " & iCat
							sql = sql & " GROUP BY "
							sql = sql & " PH_DataCruda.Id_Hogar "
							'if iCat = 91 then
							'	response.write "<br>269 sql:=" & sql
							'end if
							'response.write "<br>190 sql:=" & sql
							'response.end
							rsx1.Open sql ,conexion
							if rsx1.eof then
								rsx1.close
								TotalHogaresCatFeb = 0
							else
								gHogaresCategoriaFeb = rsx1.GetRows
								rsx1.close
								TotalHogaresCatFeb = ubound(gHogaresCategoriaFeb,2) + 1
							end if
							PenetracionFeb = 0
							response.write "<td>"
								PenetracionFeb = (TotalHogaresCatFeb * 100) / TotalHogaresFeb
								PenetracionFeb = FormatNumber(PenetracionFeb,2)
								response.write PenetracionFeb
								response.write "<br>(" & TotalHogaresCatFeb & "-" & TotalHogaresFeb & ")"
							response.write "</td>"

							'Penetración Marzo
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " PH_DataCruda.Id_Hogar "
							sql = sql & " FROM PH_DataCruda INNER JOIN ss_Semana ON PH_DataCruda.Id_Semana = ss_Semana.IdSemana "
							sql = sql & " WHERE "
							sql = sql & " ss_Semana.IdMes = 3 "
							sql = sql & " AND ss_Semana.IdAno = 2021 "
							sql = sql & " AND PH_DataCruda.Id_Categoria = " & iCat
							sql = sql & " GROUP BY "
							sql = sql & " PH_DataCruda.Id_Hogar "
							'if iCat = 43 then
							'	response.write "<br>408 sql:=" & sql
							'end if
							'response.write "<br>190 sql:=" & sql
							'response.end
							rsx1.Open sql ,conexion
							if rsx1.eof then
								rsx1.close
								TotalHogaresCatMar = 0
							else
								gHogaresCategoriaMar = rsx1.GetRows
								rsx1.close
								TotalHogaresCatMar = ubound(gHogaresCategoriaMar,2) + 1
							end if
							PenetracionMar = 0
							response.write "<td>"
								PenetracionMar = (TotalHogaresCatMar * 100) / TotalHogaresMar
								PenetracionMar = FormatNumber(PenetracionMar,2)
								response.write PenetracionMar
								response.write "<br>(" & TotalHogaresCatMar & "-" & TotalHogaresMar & ")"
							response.write "</td>"

							'Penetración Abril
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " PH_DataCruda.Id_Hogar "
							sql = sql & " FROM PH_DataCruda INNER JOIN ss_Semana ON PH_DataCruda.Id_Semana = ss_Semana.IdSemana "
							sql = sql & " WHERE "
							sql = sql & " ss_Semana.IdMes = 4 "
							sql = sql & " AND ss_Semana.IdAno = 2021 "
							sql = sql & " AND PH_DataCruda.Id_Categoria = " & iCat
							sql = sql & " GROUP BY "
							sql = sql & " PH_DataCruda.Id_Hogar "
							'if iCat = 91 then
							'	response.write "<br>269 sql:=" & sql
							'end if
							'response.write "<br>190 sql:=" & sql
							'response.end
							rsx1.Open sql ,conexion
							if rsx1.eof then
								rsx1.close
								TotalHogaresCatAbr = 0
							else
								gHogaresCategoriaAbr = rsx1.GetRows
								rsx1.close
								TotalHogaresCatAbr = ubound(gHogaresCategoriaAbr,2) + 1
							end if
							PenetracionAbr = 0
							response.write "<td>"
								PenetracionAbr = (TotalHogaresCatAbr * 100) / TotalHogaresAbr
								PenetracionAbr = FormatNumber(PenetracionAbr,2)
								response.write PenetracionAbr
								response.write "<br>(" & TotalHogaresCatAbr & "-" & TotalHogaresAbr & ")"
							response.write "</td>"

							'Penetración Mayo
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " PH_DataCruda.Id_Hogar "
							sql = sql & " FROM PH_DataCruda INNER JOIN ss_Semana ON PH_DataCruda.Id_Semana = ss_Semana.IdSemana "
							sql = sql & " WHERE "
							sql = sql & " ss_Semana.IdMes = 5 "
							sql = sql & " AND ss_Semana.IdAno = 2021 "
							sql = sql & " AND PH_DataCruda.Id_Categoria = " & iCat
							sql = sql & " GROUP BY "
							sql = sql & " PH_DataCruda.Id_Hogar "
							'if iCat = 91 then
							'	response.write "<br>269 sql:=" & sql
							'end if
							'response.write "<br>190 sql:=" & sql
							'response.end
							rsx1.Open sql ,conexion
							if rsx1.eof then
								rsx1.close
								TotalHogaresCatMay = 0
							else
								gHogaresCategoriaMay = rsx1.GetRows
								rsx1.close
								TotalHogaresCatMay = ubound(gHogaresCategoriaMay,2) + 1
							end if
							PenetracionMay = 0
							response.write "<td>"
								PenetracionMay = (TotalHogaresCatMay * 100) / TotalHogaresMay
								PenetracionMay = FormatNumber(PenetracionMay,2)
								response.write PenetracionMay
								response.write "<br>(" & TotalHogaresCatMay & "-" & TotalHogaresMay & ")"
							response.write "</td>"

							'Penetración Junio
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " PH_DataCruda.Id_Hogar "
							sql = sql & " FROM PH_DataCruda INNER JOIN ss_Semana ON PH_DataCruda.Id_Semana = ss_Semana.IdSemana "
							sql = sql & " WHERE "
							sql = sql & " ss_Semana.IdMes = 6 "
							sql = sql & " AND ss_Semana.IdAno = 2021 "
							sql = sql & " AND PH_DataCruda.Id_Categoria = " & iCat
							sql = sql & " GROUP BY "
							sql = sql & " PH_DataCruda.Id_Hogar "
							'if iCat = 91 then
							'	response.write "<br>269 sql:=" & sql
							'end if
							'response.write "<br>190 sql:=" & sql
							'response.end
							rsx1.Open sql ,conexion
							if rsx1.eof then
								rsx1.close
								TotalHogaresCatJun = 0
							else
								gHogaresCategoriaJun = rsx1.GetRows
								rsx1.close
								TotalHogaresCatJun = ubound(gHogaresCategoriaJun,2) + 1
							end if
							PenetracionJun = 0
							response.write "<td>"
								PenetracionJun = (TotalHogaresCatJun * 100) / TotalHogaresJun
								PenetracionJun = FormatNumber(PenetracionJun,2)
								response.write PenetracionJun
								response.write "<br>(" & TotalHogaresCatJun & "-" & TotalHogaresJun & ")"
							response.write "</td>"

							'Penetración Acumulado
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " PH_DataCruda.Id_Hogar "
							sql = sql & " FROM PH_DataCruda INNER JOIN ss_Semana ON PH_DataCruda.Id_Semana = ss_Semana.IdSemana "
							sql = sql & " WHERE "
							sql = sql & " PH_DataCruda.Id_Categoria = " & iCat
							sql = sql & " GROUP BY "
							sql = sql & " PH_DataCruda.Id_Hogar "
							'if iCat = 91 then
							'	response.write "<br>293 sql:=" & sql
							'end if
							'response.end
							rsx1.Open sql ,conexion
							if rsx1.eof then
								rsx1.close
								TotalHogaresCatAcu = 0
							else
								gHogaresCategoriaAcu = rsx1.GetRows
								rsx1.close
								TotalHogaresCatAcu = ubound(gHogaresCategoriaAcu,2) + 1
							end if
							PenetracionAcu = 0
							'response.write "<br>190 TotalHogaresCatAcu:=" & TotalHogaresCatAcu & "<br>"
							'response.end
							response.write "<td>"
								PenetracionAcu = (TotalHogaresCatAcu * 100) / TotalHogaresAcu
								PenetracionAcu = FormatNumber(PenetracionAcu,2)
								response.write PenetracionAcu
								response.write "<br>(" & TotalHogaresCatAcu & "-" & TotalHogaresAcu & ")"
							response.write "</td>"
							
						response.write "</tr>"
					next
					%>
				</table>
			</div>
		</div>
	<br>

    <%

	%>
	<!--<script language="JavaScript" type="text/javascript">
		Mensaje()
	</script>-->
	
	<%

	conexion.close
	%>
	


<style>
@keyframes showSweetAlert {
  0% {
    transform: scale(0.7);
  }
  45% {
    transform: scale(1.05);
  }
  80% {
    transform: scale(0.95);
  }
  100% {
    transform: scale(1);
  }
}
@keyframes hideSweetAlert {
  0% {
    transform: scale(1);
  }
  100% {
    transform: scale(0.5);
  }
}
@keyframes slideFromTop {
  0% {
    top: 0%;
  }
  100% {
    top: 50%;
  }
}
@keyframes slideToTop {
  0% {
    top: 50%;
  }
  100% {
    top: 0%;
  }
}
@keyframes slideFromBottom {
  0% {
    top: 70%;
  }
  100% {
    top: 50%;
  }
}
@keyframes slideToBottom {
  0% {
    top: 50%;
  }
  100% {
    top: 70%;
  }
}
.showSweetAlert {
  animation: showSweetAlert 0.3s;
}
.showSweetAlert[data-animation=none] {
  animation: none;
}
.showSweetAlert[data-animation=slide-from-top] {
  animation: slideFromTop 0.3s;
}
.showSweetAlert[data-animation=slide-from-bottom] {
  animation: slideFromBottom 0.3s;
}
.hideSweetAlert {
  animation: hideSweetAlert 0.3s;
}
.hideSweetAlert[data-animation=none] {
  animation: none;
}
.hideSweetAlert[data-animation=slide-from-top] {
  animation: slideToTop 0.3s;
}
.hideSweetAlert[data-animation=slide-from-bottom] {
  animation: slideToBottom 0.3s;
}
@keyframes animateSuccessTip {
  0% {
    width: 0;
    left: 1px;
    top: 19px;
  }
  54% {
    width: 0;
    left: 1px;
    top: 19px;
  }
  70% {
    width: 50px;
    left: -8px;
    top: 37px;
  }
  84% {
    width: 17px;
    left: 21px;
    top: 48px;
  }
  100% {
    width: 25px;
    left: 14px;
    top: 45px;
  }
}
@keyframes animateSuccessLong {
  0% {
    width: 0;
    right: 46px;
    top: 54px;
  }
  65% {
    width: 0;
    right: 46px;
    top: 54px;
  }
  84% {
    width: 55px;
    right: 0px;
    top: 35px;
  }
  100% {
    width: 47px;
    right: 8px;
    top: 38px;
  }
}
@keyframes rotatePlaceholder {
  0% {
    transform: rotate(-45deg);
  }
  5% {
    transform: rotate(-45deg);
  }
  12% {
    transform: rotate(-405deg);
  }
  100% {
    transform: rotate(-405deg);
  }
}
.animateSuccessTip {
  animation: animateSuccessTip 0.75s;
}
.animateSuccessLong {
  animation: animateSuccessLong 0.75s;
}
.sa-icon.sa-success.animate::after {
  animation: rotatePlaceholder 4.25s ease-in;
}
@keyframes animateErrorIcon {
  0% {
    transform: rotateX(100deg);
    opacity: 0;
  }
  100% {
    transform: rotateX(0deg);
    opacity: 1;
  }
}
.animateErrorIcon {
  animation: animateErrorIcon 0.5s;
}
@keyframes animateXMark {
  0% {
    transform: scale(0.4);
    margin-top: 26px;
    opacity: 0;
  }
  50% {
    transform: scale(0.4);
    margin-top: 26px;
    opacity: 0;
  }
  80% {
    transform: scale(1.15);
    margin-top: -6px;
  }
  100% {
    transform: scale(1);
    margin-top: 0;
    opacity: 1;
  }
}
.animateXMark {
  animation: animateXMark 0.5s;
}
@keyframes pulseWarning {
  0% {
    border-color: #F8D486;
  }
  100% {
    border-color: #F8BB86;
  }
}
.pulseWarning {
  animation: pulseWarning 0.75s infinite alternate;
}
@keyframes pulseWarningIns {
  0% {
    background-color: #F8D486;
  }
  100% {
    background-color: #F8BB86;
  }
}
.pulseWarningIns {
  animation: pulseWarningIns 0.75s infinite alternate;
}
@keyframes rotate-loading {
  0% {
    transform: rotate(0deg);
  }
  100% {
    transform: rotate(360deg);
  }
}
body.stop-scrolling {
  height: 100%;
  overflow: hidden;
}
.sweet-overlay {
  background-color: rgba(0, 0, 0, 0.4);
  position: fixed;
  left: 0;
  right: 0;
  top: 0;
  bottom: 0;
  display: none;
  z-index: 1040;
}
.sweet-alert {
  background-color: #ffffff;
  width: 478px;
  padding: 17px;
  border-radius: 5px;
  text-align: center;
  position: fixed;
  left: 50%;
  top: 50%;
  margin-left: -256px;
  margin-top: -200px;
  overflow: hidden;
  display: none;
  z-index: 2000;
}
@media all and (max-width: 767px) {
  .sweet-alert {
    width: auto;
    margin-left: 0;
    margin-right: 0;
    left: 15px;
    right: 15px;
  }
}
.sweet-alert .form-group {
  display: none;
}
.sweet-alert .form-group .sa-input-error {
  display: none;
}
.sweet-alert.show-input .form-group {
  display: block;
}
.sweet-alert .sa-confirm-button-container {
  display: inline-block;
  position: relative;
}
.sweet-alert .la-ball-fall {
  position: absolute;
  left: 50%;
  top: 50%;
  margin-left: -27px;
  margin-top: -9px;
  opacity: 0;
  visibility: hidden;
}
.sweet-alert button[disabled] {
  opacity: .6;
  cursor: default;
}
.sweet-alert button.confirm[disabled] {
  color: transparent;
}
.sweet-alert button.confirm[disabled] ~ .la-ball-fall {
  opacity: 1;
  visibility: visible;
  transition-delay: 0s;
}
.sweet-alert .sa-icon {
  width: 80px;
  height: 80px;
  border: 4px solid gray;
  border-radius: 50%;
  margin: 20px auto;
  position: relative;
  box-sizing: content-box;
}
.sweet-alert .sa-icon.sa-error {
  border-color: #d43f3a;
}
.sweet-alert .sa-icon.sa-error .sa-x-mark {
  position: relative;
  display: block;
}
.sweet-alert .sa-icon.sa-error .sa-line {
  position: absolute;
  height: 5px;
  width: 47px;
  background-color: #d9534f;
  display: block;
  top: 37px;
  border-radius: 2px;
}
.sweet-alert .sa-icon.sa-error .sa-line.sa-left {
  transform: rotate(45deg);
  left: 17px;
}
.sweet-alert .sa-icon.sa-error .sa-line.sa-right {
  transform: rotate(-45deg);
  right: 16px;
}
.sweet-alert .sa-icon.sa-warning {
  border-color: #eea236;
}
.sweet-alert .sa-icon.sa-warning .sa-body {
  position: absolute;
  width: 5px;
  height: 47px;
  left: 50%;
  top: 10px;
  border-radius: 2px;
  margin-left: -2px;
  background-color: #f0ad4e;
}
.sweet-alert .sa-icon.sa-warning .sa-dot {
  position: absolute;
  width: 7px;
  height: 7px;
  border-radius: 50%;
  margin-left: -3px;
  left: 50%;
  bottom: 10px;
  background-color: #f0ad4e;
}
.sweet-alert .sa-icon.sa-info {
  border-color: #46b8da;
}
.sweet-alert .sa-icon.sa-info::before {
  content: "";
  position: absolute;
  width: 5px;
  height: 29px;
  left: 50%;
  bottom: 17px;
  border-radius: 2px;
  margin-left: -2px;
  background-color: #5bc0de;
}
.sweet-alert .sa-icon.sa-info::after {
  content: "";
  position: absolute;
  width: 7px;
  height: 7px;
  border-radius: 50%;
  margin-left: -3px;
  top: 19px;
  background-color: #5bc0de;
}
.sweet-alert .sa-icon.sa-success {
  border-color: #4cae4c;
}
.sweet-alert .sa-icon.sa-success::before,
.sweet-alert .sa-icon.sa-success::after {
  content: '';
  border-radius: 50%;
  position: absolute;
  width: 60px;
  height: 120px;
  background: #ffffff;
  transform: rotate(45deg);
}
.sweet-alert .sa-icon.sa-success::before {
  border-radius: 120px 0 0 120px;
  top: -7px;
  left: -33px;
  transform: rotate(-45deg);
  transform-origin: 60px 60px;
}
.sweet-alert .sa-icon.sa-success::after {
  border-radius: 0 120px 120px 0;
  top: -11px;
  left: 30px;
  transform: rotate(-45deg);
  transform-origin: 0px 60px;
}
.sweet-alert .sa-icon.sa-success .sa-placeholder {
  width: 80px;
  height: 80px;
  border: 4px solid rgba(92, 184, 92, 0.2);
  border-radius: 50%;
  box-sizing: content-box;
  position: absolute;
  left: -4px;
  top: -4px;
  z-index: 2;
}
.sweet-alert .sa-icon.sa-success .sa-fix {
  width: 5px;
  height: 90px;
  background-color: #ffffff;
  position: absolute;
  left: 28px;
  top: 8px;
  z-index: 1;
  transform: rotate(-45deg);
}
.sweet-alert .sa-icon.sa-success .sa-line {
  height: 5px;
  background-color: #5cb85c;
  display: block;
  border-radius: 2px;
  position: absolute;
  z-index: 2;
}
.sweet-alert .sa-icon.sa-success .sa-line.sa-tip {
  width: 25px;
  left: 14px;
  top: 46px;
  transform: rotate(45deg);
}
.sweet-alert .sa-icon.sa-success .sa-line.sa-long {
  width: 47px;
  right: 8px;
  top: 38px;
  transform: rotate(-45deg);
}
.sweet-alert .sa-icon.sa-custom {
  background-size: contain;
  border-radius: 0;
  border: none;
  background-position: center center;
  background-repeat: no-repeat;
}
.sweet-alert .btn-default:focus {
  border-color: #cccccc;
  outline: 0;
  -webkit-box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(204, 204, 204, 0.6);
  box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(204, 204, 204, 0.6);
}
.sweet-alert .btn-success:focus {
  border-color: #4cae4c;
  outline: 0;
  -webkit-box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(76, 174, 76, 0.6);
  box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(76, 174, 76, 0.6);
}
.sweet-alert .btn-info:focus {
  border-color: #46b8da;
  outline: 0;
  -webkit-box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(70, 184, 218, 0.6);
  box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(70, 184, 218, 0.6);
}
.sweet-alert .btn-danger:focus {
  border-color: #d43f3a;
  outline: 0;
  -webkit-box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(212, 63, 58, 0.6);
  box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(212, 63, 58, 0.6);
}
.sweet-alert .btn-warning:focus {
  border-color: #eea236;
  outline: 0;
  -webkit-box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(238, 162, 54, 0.6);
  box-shadow: inset 0 1px 1px rgba(0,0,0,.075), 0 0 8px rgba(238, 162, 54, 0.6);
}
.sweet-alert button::-moz-focus-inner {
  border: 0;
}
/*!
 * Load Awesome v1.1.0 (http://github.danielcardoso.net/load-awesome/)
 * Copyright 2015 Daniel Cardoso <@DanielCardoso>
 * Licensed under MIT
 */
.la-ball-fall,
.la-ball-fall > div {
  position: relative;
  -webkit-box-sizing: border-box;
  -moz-box-sizing: border-box;
  box-sizing: border-box;
}
.la-ball-fall {
  display: block;
  font-size: 0;
  color: #fff;
}
.la-ball-fall.la-dark {
  color: #333;
}
.la-ball-fall > div {
  display: inline-block;
  float: none;
  background-color: currentColor;
  border: 0 solid currentColor;
}
.la-ball-fall {
  width: 54px;
  height: 18px;
}
.la-ball-fall > div {
  width: 10px;
  height: 10px;
  margin: 4px;
  border-radius: 100%;
  opacity: 0;
  -webkit-animation: ball-fall 1s ease-in-out infinite;
  -moz-animation: ball-fall 1s ease-in-out infinite;
  -o-animation: ball-fall 1s ease-in-out infinite;
  animation: ball-fall 1s ease-in-out infinite;
}
.la-ball-fall > div:nth-child(1) {
  -webkit-animation-delay: -200ms;
  -moz-animation-delay: -200ms;
  -o-animation-delay: -200ms;
  animation-delay: -200ms;
}
.la-ball-fall > div:nth-child(2) {
  -webkit-animation-delay: -100ms;
  -moz-animation-delay: -100ms;
  -o-animation-delay: -100ms;
  animation-delay: -100ms;
}
.la-ball-fall > div:nth-child(3) {
  -webkit-animation-delay: 0ms;
  -moz-animation-delay: 0ms;
  -o-animation-delay: 0ms;
  animation-delay: 0ms;
}
.la-ball-fall.la-sm {
  width: 26px;
  height: 8px;
}
.la-ball-fall.la-sm > div {
  width: 4px;
  height: 4px;
  margin: 2px;
}
.la-ball-fall.la-2x {
  width: 108px;
  height: 36px;
}
.la-ball-fall.la-2x > div {
  width: 20px;
  height: 20px;
  margin: 8px;
}
.la-ball-fall.la-3x {
  width: 162px;
  height: 54px;
}
.la-ball-fall.la-3x > div {
  width: 30px;
  height: 30px;
  margin: 12px;
}
/*
 * Animation
 */
@-webkit-keyframes ball-fall {
  0% {
    opacity: 0;
    -webkit-transform: translateY(-145%);
    transform: translateY(-145%);
  }
  10% {
    opacity: .5;
  }
  20% {
    opacity: 1;
    -webkit-transform: translateY(0);
    transform: translateY(0);
  }
  80% {
    opacity: 1;
    -webkit-transform: translateY(0);
    transform: translateY(0);
  }
  90% {
    opacity: .5;
  }
  100% {
    opacity: 0;
    -webkit-transform: translateY(145%);
    transform: translateY(145%);
  }
}
@-moz-keyframes ball-fall {
  0% {
    opacity: 0;
    -moz-transform: translateY(-145%);
    transform: translateY(-145%);
  }
  10% {
    opacity: .5;
  }
  20% {
    opacity: 1;
    -moz-transform: translateY(0);
    transform: translateY(0);
  }
  80% {
    opacity: 1;
    -moz-transform: translateY(0);
    transform: translateY(0);
  }
  90% {
    opacity: .5;
  }
  100% {
    opacity: 0;
    -moz-transform: translateY(145%);
    transform: translateY(145%);
  }
}
@-o-keyframes ball-fall {
  0% {
    opacity: 0;
    -o-transform: translateY(-145%);
    transform: translateY(-145%);
  }
  10% {
    opacity: .5;
  }
  20% {
    opacity: 1;
    -o-transform: translateY(0);
    transform: translateY(0);
  }
  80% {
    opacity: 1;
    -o-transform: translateY(0);
    transform: translateY(0);
  }
  90% {
    opacity: .5;
  }
  100% {
    opacity: 0;
    -o-transform: translateY(145%);
    transform: translateY(145%);
  }
}
@keyframes ball-fall {
  0% {
    opacity: 0;
    -webkit-transform: translateY(-145%);
    -moz-transform: translateY(-145%);
    -o-transform: translateY(-145%);
    transform: translateY(-145%);
  }
  10% {
    opacity: .5;
  }
  20% {
    opacity: 1;
    -webkit-transform: translateY(0);
    -moz-transform: translateY(0);
    -o-transform: translateY(0);
    transform: translateY(0);
  }
  80% {
    opacity: 1;
    -webkit-transform: translateY(0);
    -moz-transform: translateY(0);
    -o-transform: translateY(0);
    transform: translateY(0);
  }
  90% {
    opacity: .5;
  }
  100% {
    opacity: 0;
    -webkit-transform: translateY(145%);
    -moz-transform: translateY(145%);
    -o-transform: translateY(145%);
    transform: translateY(145%);
  }
}

.accordion {
  background-color: #eee;
  color: #444;
  cursor: pointer;
  padding: 18px;
  width: 100%;
  border: none;
  text-align: left;
  outline: none;
  font-size: 20px;
  transition: 0.4s;
}

.active, .accordion:hover {
  background-color: #ccc;
}

.accordion:after {
  content: '\002B';
  color: #777;
  font-weight: bold;
  float: right;
  margin-left: 5px;
}

.active:after {
  content: "\2212";
}

.panel {
  padding: 0 18px;
  background-color: white;
  max-height: 0;
  overflow: hidden;
  transition: max-height 0.2s ease-out;
}
</style>

</body>
</html>