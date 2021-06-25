<%@language=vbscript%>

<!--#include file="Conexion.asp"-->


<%

LCID = 1034  
'==========================================================================================
' Variables y Constantes
'==========================================================================================

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
	Response.AddHeader "Content-disposition","attachment; filename=tem.xls"
	Response.ContentType = "application/vnd.ms-excel"

%>
	<table>
		<tr>
			<td>IdCategoria</td>
			<td>Categoria</td>
			<td>Penetracion Enero 2021</td>
			<td>Hogares Enero 2021</td>
			<td>Penetracion Febrero 2021</td>
			<td>Hogares Febrero 2021</td>
			<td>Penetracion Marzo 2021</td>
			<td>Hogares Marzo 2021</td>
			<td>Penetracion Abril 2021</td>
			<td>Hogares Abril 2021</td>
			<td>Penetracion Mayo 2021</td>
			<td>Hogares Mayo 2021</td>
			<td>Penetracion Jun 2021</td>
			<td>Hogares Jun 2021</td>
			<td>Acumulado 2021</td>
			<td>Hogares Acumulado 2021</td>
		</tr>
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
					'PenetracionEne = cStr(PenetracionEne)
					'PenetracionEne = replace(PenetracionEne,",",".")
					response.write PenetracionEne
				response.write "</td>"
				response.write "<td>"
					PenetracionEne = (TotalHogaresCatEne * 100) / TotalHogaresEne
					PenetracionEne = FormatNumber(PenetracionEne,2)
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
					'PenetracionFeb = cStr(PenetracionFeb)
					'PenetracionFeb = replace(PenetracionFeb,",",".")
					response.write PenetracionFeb
				response.write "</td>"
				response.write "<td>"
					PenetracionFeb = (TotalHogaresCatFeb * 100) / TotalHogaresFeb
					PenetracionFeb = FormatNumber(PenetracionFeb,2)
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
					'PenetracionMar = cStr(PenetracionMar)
					'PenetracionMar = replace(PenetracionMar,",",".")
					response.write PenetracionMar
				response.write "</td>"
				response.write "<td>"
					PenetracionMar = (TotalHogaresCatMar * 100) / TotalHogaresMar
					PenetracionMar = FormatNumber(PenetracionMar,2)
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
					'PenetracionAbr = cStr(PenetracionAbr)
					'PenetracionAbr = replace(PenetracionAbr,",",".")
					response.write PenetracionAbr
				response.write "</td>"
				response.write "<td>"
					PenetracionAbr = (TotalHogaresCatAbr * 100) / TotalHogaresAbr
					PenetracionAbr = FormatNumber(PenetracionAbr,2)
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
					'PenetracionMay = cStr(PenetracionMay)
					'PenetracionMay = replace(PenetracionMay,",",".")
					response.write PenetracionMay
				response.write "</td>"
				response.write "<td>"
					PenetracionMay = (TotalHogaresCatMay * 100) / TotalHogaresMay
					PenetracionMay = FormatNumber(PenetracionMay,2)
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
					'PenetracionJun = cStr(PenetracionJun)
					'PenetracionJun = replace(PenetracionJun,",",".")
					response.write PenetracionJun
				response.write "</td>"
				response.write "<td>"
					PenetracionJun = (TotalHogaresCatJun * 100) / TotalHogaresJun
					PenetracionJun = FormatNumber(PenetracionJun,2)
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
				'response.write "<br>190 sql:=" & sql
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
					'PenetracionAcu = cStr(PenetracionAcu)
					'PenetracionAcu = replace(PenetracionAcu,",",".")
					response.write PenetracionAcu
				response.write "</td>"
				response.write "<td>"
					PenetracionAcu = (TotalHogaresCatAcu * 100) / TotalHogaresAcu
					PenetracionAcu = FormatNumber(PenetracionAcu,2)
					response.write "<br>(" & TotalHogaresCatAcu & "-" & TotalHogaresAcu & ")"
				response.write "</td>"
				
			response.write "</tr>"
		next
		%>
	</table>

    <%

	%>

	<!--<script language="JavaScript" type="text/javascript">
		Mensaje()
	</script>-->
	
	<%

	conexion.close
	%>
	


</body>
</html>