<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	' g_ConvCalcularTotalHogaresxConsumoPaso3.asp - 09abr21 - 29abr21
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	Response.Buffer = True
	'
	Dim rsRecordCount, iTotalHogaresCompradores, iTotalHogaresUnicaCompra, QrySql, iPorcentajeHogares, rsArray
    '
    ' Capturar las variables
    '
	idMeses = Request.QueryString("id_Mes")
	'idMeses="16,17,18,19,20,21,22,23,24,25,26,27,28" ' Trimestral
	'	
	' Calcular Total hogares del Mes 
	'	
	QrySql = vbnullstring
	QrySql = " SELECT" & _
	" Id_Hogar" & _
	" FROM" & _
	" PH_DataCrudaMensual" & _
	" WHERE" & _
	" PH_DataCrudaMensual.Id_Fabricante<>0" & _
	" AND" & _
	" Id_Semana in (" & idMeses & ")" & _
	" GROUP BY PH_DataCrudaMensual.Id_Hogar;"
	'
	Set rsRecordCount = Server.CreateObject("ADODB.recordset")
	Set rsRecordCount = conexion.Execute(QrySql)
	if not rsRecordCount.Eof then
		rsArray = rsRecordCount.GetRows() 
	  	iTotalHogares = UBound(rsArray, 2) + 1 	  
	else
	  iTotalHogares = 0
	end if
	'
	rsRecordCount.Close
	set rsRecordCount = nothing
	set rsArray = nothing
	'	
	' Ajuste del 11ABR
	'
	QrySql = vbnullstring
	QrySql = " SELECT" & _
	" PH_DataCrudaMensual.Id_Hogar" & _
	" FROM" & _
	" PH_DataCrudaMensual" & _
	" WHERE" & _
	" PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")" & _
	" AND" & _
	" PH_DataCrudaMensual.Id_Fabricante <> 0" & _
	" AND" & _
	" PH_DataCrudaMensual.Id_Categoria IN (1 , 3 , 12 , 22)" & _
	" GROUP BY" & _
	" PH_DataCrudaMensual.Id_Hogar" & _
	" ORDER BY" & _
	" PH_DataCrudaMensual.Id_Hogar;"
	'	
	'Response.Write QrySql '& "<BR><BR>"
	'Response.end
	'
	Set rsRecordCount = Server.CreateObject("ADODB.recordset")
	Set rsRecordCount = conexion.Execute(QrySql)
	if not rsRecordCount.Eof then
		rsArray = rsRecordCount.GetRows() 
	  	iTotalHogaresUnicaCompra = UBound(rsArray, 2) + 1 	  
	else
	  iTotalHogaresUnicaCompra = 0
	end if
	'
	rsRecordCount.Close
	set rsRecordCount = nothing
	set rsArray = nothing
	'
	if iTotalHogares > 0 then
    	iPorcentajeHogares = (iTotalHogaresUnicaCompra * 100) / iTotalHogares
		Response.Write FormatNumber(iPorcentajeHogares,2)
	else
		Response.Write 0
	end if
	'
	Response.flush
	Response.Clear
	'
%>