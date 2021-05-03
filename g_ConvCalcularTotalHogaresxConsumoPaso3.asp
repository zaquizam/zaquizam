<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	' g_ConvCalcularTotalHogaresxConsumoPaso3.asp - 09abr21 - 11abr21
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'
	Dim rsRecordCount, iTotalHogaresCompradores, iTotalHogaresUnicaCompra, QrySql, iPorcentajeHogares, rsArray
    '
    ' Capturar las variables
    '
	idMeses = Request.QueryString("id_Mes")	
	'	
	' Calcular Total hogares del Mes 
	'	
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " Id_Hogar"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_DataCrudaMensual"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante<>0"
	QrySql = QrySql & " AND"
	QrySql = QrySql & " Id_Semana in (" & idMeses & ")"
	'	
	' QrySql = QrySql & " (((PH_DataCrudaMensual.Id_Semana)=16"
	' QrySql = QrySql & " Or"
	' QrySql = QrySql & " (PH_DataCrudaMensual.Id_Semana)=17"
	' QrySql = QrySql & " Or"
	' QrySql = QrySql & " (PH_DataCrudaMensual.Id_Semana)=18"
	' QrySql = QrySql & " Or (PH_DataCrudaMensual.Id_Semana)=19"
	' QrySql = QrySql & " )"
	' QrySql = QrySql & " AND"
	
	QrySql = QrySql & " GROUP BY PH_DataCrudaMensual.Id_Hogar;"
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
	' Calcular Total hogares que compraron al menos un (refresco, te, agua o jugo)
	'	
	' QrySql = vbnullstring
	' QrySql = QrySql & " SELECT"
	' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar"
	' QrySql = QrySql & " FROM"
	' QrySql = QrySql & " PH_DataCrudaMensual"
	' QrySql = QrySql & " WHERE"
	' QrySql = QrySql & " (((PH_DataCrudaMensual.Id_Semana)=16"
	' QrySql = QrySql & " Or (PH_DataCrudaMensual.Id_Semana)=17"
	' QrySql = QrySql & " Or (PH_DataCrudaMensual.Id_Semana)=18"
	' QrySql = QrySql & " Or (PH_DataCrudaMensual.Id_Semana)=19"
	' QrySql = QrySql & " ) AND"
	' QrySql = QrySql & " ((PH_DataCrudaMensual.Id_Fabricante)<>0)"
	' QrySql = QrySql & " AND"
	' QrySql = QrySql & " ((PH_DataCrudaMensual.Id_Categoria)=1"
	' QrySql = QrySql & " Or"
	' QrySql = QrySql & " (PH_DataCrudaMensual.Id_Categoria)=3"
	' QrySql = QrySql & " Or"
	' QrySql = QrySql & " (PH_DataCrudaMensual.Id_Categoria)=12"
	' QrySql = QrySql & " Or"
	' QrySql = QrySql & " (PH_DataCrudaMensual.Id_Categoria)=22))"
	' QrySql = QrySql & " GROUP BY"
	' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar"
	' QrySql = QrySql & " ORDER BY PH_DataCrudaMensual.Id_Hogar;"
	' '
	'Ajuste del 11ABR
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_DataCrudaMensual"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria IN (1 , 3 , 12 , 22)"
	QrySql = QrySql & " GROUP BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar"
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
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
%>