<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_ValBuscarLlenarCmbOtrosProductos.asp - 22mar21
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'
	Dim QrySql, rsCategoria
	'
	' Buscar Datos de todas las Categorias Registrados
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " id_categoria,"
	QrySql = QrySql & " categoria"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_categoria"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " Ind_Activo = 1"
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " categoria ASC"
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsCategoria = Server.CreateObject("ADODB.recordset")
	rsCategoria.Open QrySql, conexion
	'
	if not rsCategoria.EOF then
		arrCategoria = rsCategoria.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	Response.ContentType = "application/json"
	'
	' Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrCategoria) then

		For i = 0 to ubound(arrCategoria, 2)
			'
			sTabla     =  chr(123) &  chr(34) & "id" 	 & chr(34) & ":" & chr(34) & arrCategoria(0,i) & chr(34) & chr(44)
			sTabla     =  sTabla   &  chr(34) & "nombre" & chr(34) & ":" & chr(34) & arrCategoria(1,i) & chr(34) & chr(125) & chr(44)
			sTablaJson =  sTablaJson & sTabla
			sTabla = vbnullstring
			'
		next

	else
		'Eof()
		sTabla    =   chr(123) &  chr(34) & "id" 		& chr(34) & ":"  & chr(34)  & "0" 			& chr(34) & chr(44)
		sTabla    =   sTabla   &  chr(34) & "nombre"    & chr(34) & ":"  & chr(34)  & "No Aplica" 	& chr(34) & chr(125) & chr(44)
		'
		sTablaJson = sTablaJson & sTabla
		sTabla = vbnullstring
		'
	end if
	''
	sTabla   = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Categoria"
	'
	JsonData = chr(123) & chr(34) & "data" & chr(34) & ":" & chr(91) & sTabla & chr(93) & chr(125)
	'
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'
	rsCategoria.Close
	Set rsCategoria = Nothing
	'
	conexion.close
	set conexion = nothing
	'	
%>