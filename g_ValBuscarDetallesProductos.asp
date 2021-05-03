 <%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	'09dic20
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'	
	Dim idhogar, idencuesta, rsx1, rsx2, arrRespuestas, arrDetalles
	'
	idhogar   = Request.QueryString("idHogar")
	idencuesta= Request.QueryString("idEncuesta")
	'
	' Buscar los detalles de la Encuestas
	'
	set rsx2 = CreateObject("ADODB.Recordset")
	rsx2.CursorType = adOpenKeyset 
	rsx2.LockType = 2 'adLockOptimistic 
	
	sql = vbnullstring
	sql = sql & " SELECT"
	'sql = sql & " PH_EncuestaEspecial.Fec_Desde,"  '0
	'sql = sql & " PH_EncuestaEspecial.Fec_Hasta,"  '1
	'sql = sql & " PH_EncuestaHogar.Ind_Acepto,"    '2
	'sql = sql & " PH_EncuestaHogar.Ind_Rechazada," '3
	'sql = sql & " PH_EncuestaHogar.Fec_Realizada"  '4
	'
	sql = sql & " FORMAT (PH_EncuestaEspecial.Fec_Desde, 'dd/MM/yyyy '),"
	sql = sql & " FORMAT (PH_EncuestaEspecial.Fec_Hasta, 'dd/MM/yyyy '),"		
	sql = sql & " CASE WHEN PH_EncuestaHogar.Ind_Acepto = 1 THEN 'Si' ELSE 'No' END,"	
	sql = sql & " CASE WHEN PH_EncuestaHogar.Ind_Rechazada = 1 THEN 'Si' ELSE 'No' END,"	
	sql = sql & " FORMAT (PH_EncuestaHogar.Fec_Realizada, 'dd/MM/yyyy ')"	
	sql = sql & " FROM"
	sql = sql & " PH_EncuestaHogar"
	sql = sql & " INNER JOIN cacevedo_atenas.PH_EncuestaEspecial ON cacevedo_atenas.PH_EncuestaHogar.Id_EncuestaEspecial = cacevedo_atenas.PH_EncuestaEspecial.Id_EncuestaEspecial"
	sql = sql & " WHERE"
	sql = sql & " PH_EncuestaHogar.Id_Hogar =" & idhogar
	sql = sql & " AND PH_EncuestaHogar.Id_EncuestaEspecial =" & idencuesta
	'
    rsx2.Open sql ,conexion
	'
	if not rsx2.eof then
		'
		arrDetalles = rsx2.GetRows()  ' Convert recordset to 2D Array
		fDesde		=arrDetalles(0,0)
		fHasta		=arrDetalles(1,0)
		sAcepto		=arrDetalles(2,0)
		sRechazo	=arrDetalles(3,0)
		fRealizada	=arrDetalles(4,0)
		'
		if isNull(fRealizada) then fRealizada="No Aplica"
		if sAcepto="Si" and sRechazo="No" then sStatus="Aceptada"
		if sAcepto="No" and sRechazo="Si" then sStatus="Rechazada"
		if sAcepto="No" and sRechazo="No" then sStatus="No Aplica"
		'		
	end if
		'
	rsx2.Close
	Set rsx2 = Nothing
	'
	'	
	' Buscar Resultados
	'
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 
	'		
	sql = vbnullstring
	sql = sql & " SET DATEFORMAT MDY"
	sql = sql & " SELECT"
	sql = sql & " PH_EncuestaEspecialDet.Orden,"
	sql = sql & " PH_EncuestaEspecialDet.Pregunta,"
	sql = sql & " PH_EncuestaEspecialResultados.RespuestaTexto,"
	sql = sql & " PH_EncuestaEspecialResultados.Id_Respuesta"
	sql = sql & " FROM"
	sql = sql & " PH_EncuestaEspecialResultados"
	sql = sql & " INNER JOIN PH_EncuestaEspecialDet ON PH_EncuestaEspecialResultados.Id_Pregunta_Encuesta = PH_EncuestaEspecialDet.Id_EncuestaEspecialDet"
	sql = sql & " WHERE"
	sql = sql & " PH_EncuestaEspecialResultados.Id_Hogar =" & idhogar
	sql = sql & " AND PH_EncuestaEspecialResultados.Id_EncuestaEspecial = " & idencuesta 
	sql = sql & " ORDER BY"
	sql = sql & " PH_EncuestaEspecialDet.Orden ASC"
	'
    rsx1.Open sql ,conexion
	'
	if not rsx1.eof then
		arrRespuestas = rsx1.GetRows()  ' Convert recordset to 2D Array
	end if
		'
	rsx1.Close
	Set rsx1 = Nothing
	'
	if IsArray(arrRespuestas) then
%>
	
	<div class="form-horizontal">
		<legend><i class="fas fa-check-double"></i>&nbsp;Resultados de la Encuesta:</legend>
		<table class="table table-hover table-fixed table-striped table-condensed table-sm" cellspacing="0">
			<thead>
				<tr>
					<th class="text-center" title="Pregunta Nro">DESDE</th>
					<th class="text-center" title="Pregunta">HASTA</th>
					<th class="text-center" title="Respuesta">STATUS</th>
					<th class="text-center" title="Respuesta">REALIZADA</th>
				</tr>
			</thead>
			<tbody>
			<tr>
				<td class="text-center"><%=fDesde%></td>
				<td class="text-center"><%=fHasta%></td>
				<td class="text-center"><%=sStatus%></td>
				<td class="text-center"><%=fRealizada%></td>
			</tr>
			</tbody>
		</table>

		
	<BR>
	<table class="table table-hover table-fixed table-striped table-bordered table-condensed" cellspacing="0">
		 <thead class="thead-light">
			<tr>
				<th class="text-center" title="Pregunta Nro">Pregunta Nro</th>
				<th class="text-center" title="Pregunta">Pregunta</th>
				<th class="text-center" title="Respuesta">Respuesta</th>
				<th class="text-center" title="Respuesta">idRespuesta</th>
			</tr>
		</thead>
		<tbody>
<%
					For i = 0 to ubound(arrRespuestas, 2)
							nroPregunta= arrRespuestas(0,i)
							pregunta   = arrRespuestas(1,i)
							respuesta  = REPLACE(arrRespuestas(2,i),"_"," ")
							respuestaid   = arrRespuestas(3,i)
%>
							<tr>
								<td class="text-center"><%=nroPregunta%></td>
								<td class="text-left"><%=pregunta%></td>
								<td class="text-left"><%=respuesta%></td>
								<td class="text-left"><%=respuestaid%></td>
							</tr>
<%
					next
%>
		</tbody>
	</table>
<%
	else
%>
	<div class="form-horizontal">
		<legend><i class="fas fa-check-double"></i>&nbsp;Resultados de la Encuesta:</legend>
		<table class="table table-hover table-fixed table-striped table-condensed table-sm" cellspacing="0">
			<thead>
				<tr>
					<th class="text-center" title="Pregunta Nro">DESDE</th>
					<th class="text-center" title="Pregunta">HASTA</th>
					<th class="text-center" title="Respuesta">STATUS</th>
					<th class="text-center" title="Respuesta">REALIZADA</th>
				</tr>
			</thead>
			<tbody>
			<tr>
				<td class="text-center"><%=fDesde%></td>
				<td class="text-center"><%=fHasta%></td>
				<td class="text-center"><%=sStatus%></td>
				<td class="text-center"><%=fRealizada%></td>
			</tr>
			</tbody>
		</table>

	</div>
		
		
	<table class="table table-hover table-fixed table-striped table-condensed table-sm" cellspacing="0">
		
		<tbody>
		<tr>
			<td class="text-center" colspan="3"><h3>ENCUESTA ASIGNADA PERO NO RESPONDIDA..!</h3></td>			
		</tr>
		</tbody>
	</table>
<%
	end if
	'
%>
	
	
	
