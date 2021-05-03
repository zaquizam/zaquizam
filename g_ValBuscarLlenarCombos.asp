<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
' g_ValBuscarLlenarCombos.asp
' 05ene21
'
Session.lcid=1034
Response.CodePage = 65001
Response.CharSet = "utf-8"
'
Dim idQuery, opcion,  QrySql
'
idQuery = Request.Form("id")
opcion  = Request.Form("opcion")
Buscar  = Request.Form("find")
'
IF (opcion=1) THEN
	'
	Dim rsCategoria
	'
	' Buscar Datos de todas las Categorias de Productos Registradas
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " PH_CB_Categoria.Id_Categoria AS id,"
	QrySql = QrySql & " PH_CB_Categoria.Categoria AS nombre"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_CB_Categoria"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_CB_Categoria.Ind_Activo = 1"	
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_CB_Categoria.Categoria ASC"
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsCategoria = Server.CreateObject("ADODB.recordset")
	rsCategoria.Open QrySql,conexion
	'
	if not rsCategoria.EOF then
		arrCategoria = rsCategoria.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	'Response.ContentType = "application/json"
	'
	' Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrCategoria) then

		For i = 0 to ubound(arrCategoria, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrCategoria(0,i)  & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrCategoria(1,i)  & chr(34) & chr(125)&chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla=""
			'
		next

	else
		'Eof()
		sTabla    =   chr(123) &  chr(34) & "id" 		& chr(34) & ":"  & chr(34)  & "0" 			& chr(34) & chr(44)
		sTabla    =   sTabla   &  chr(34) & "nombre"    & chr(34) & ":"  & chr(34)  & "No Aplica" 	& chr(34) & chr(125)&chr(44)
		''
		sTablaJson = sTablaJson & sTabla
		sTabla=""

	end if
	''
	sTabla = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Categoria"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
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
elseIF (opcion=2) THEN
	'
	Dim rsProductos
	'
	' Buscar Datos de todas las Productoss Registrados
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " PH_CB_Producto.CodigoBarra id,"
	QrySql = QrySql & " PH_CB_Producto.Producto AS nombre"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_CB_Producto"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_CB_Producto.Producto LIKE '%" & Buscar & '%'
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_CB_Producto.Id_Categoria = " & idQuery
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_CB_Producto.Ind_Activo = 1"
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_CB_Producto.Producto ASC"
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsProductos = Server.CreateObject("ADODB.recordset")
	rsProductos.Open QrySql,conexion
	'
	if not rsProductos.EOF then
		arrProductos = rsProductos.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	'Response.ContentType = "application/json"
	'
	' Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrProductos) then

		For i = 0 to ubound(arrCategoria, 2)
			'
			sTabla     =  chr(123) &  chr(34) & "id" 	 & chr(34) & ":" & chr(34) & arrProductos(0,i)  & chr(34) & chr(44)
			sTabla     =  sTabla   &  chr(34) & "nombre" & chr(34) & ":" & chr(34) & arrProductos(1,i)  & chr(34) & chr(125)&chr(44)
			sTablaJson =  sTablaJson & sTabla
			sTabla=""
			'
		next

	else
		'Eof()
		sTabla    =   chr(123) &  chr(34) & "id" 		& chr(34) & ":"  & chr(34)  & "0" 			& chr(34) & chr(44)
		sTabla    =   sTabla   &  chr(34) & "nombre"    & chr(34) & ":"  & chr(34)  & "No Aplica" 	& chr(34) & chr(125)&chr(44)
		''
		sTablaJson = sTablaJson & sTabla
		sTabla=""
		'
	end if
	''
	sTabla = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Categoria"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'
	rsProductos.Close
	Set rsProductos = Nothing
	'
	conexion.close
	set conexion = nothing
	'	
end if

'
%>