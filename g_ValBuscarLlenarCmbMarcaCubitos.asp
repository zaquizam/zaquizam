<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_ValBuscarLlenarCmbMarcaCubitos.asp - 22mar21
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'
	Dim QrySql, rsCubitos
	'
	' Buscar Datos de todas las Marcas de Cubitos
	'
	QrySql = vbnullstring
    QrySql = QrySql & " SELECT"
    QrySql = QrySql & " PH_CB_Marca.Id_Marca,"
    QrySql = QrySql & " PH_CB_Marca.Marca"
    QrySql = QrySql & " FROM"
    QrySql = QrySql & " PH_CB_Marca"
    QrySql = QrySql & " WHERE"
    QrySql = QrySql & " PH_CB_Marca.Id_Categoria = 7"
    QrySql = QrySql & " AND"
    QrySql = QrySql & " PH_CB_Marca.ind_Registrar_consumo = 1"
    QrySql = QrySql & " ORDER BY"
    QrySql = QrySql & " PH_CB_Marca.Marca ASC"
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsCubitos = Server.CreateObject("ADODB.recordset")
	rsCubitos.Open QrySql, conexion
	'
	if not rsCubitos.EOF then
		arrCubitos = rsCubitos.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	Response.ContentType = "application/json"
	'
	' Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrCubitos) then

		For i = 0 to ubound(arrCubitos, 2)
			'
			sTabla     =  chr(123) &  chr(34) & "id" 	 & chr(34) & ":" & chr(34) & arrCubitos(0,i) & chr(34) & chr(44)
			sTabla     =  sTabla   &  chr(34) & "nombre" & chr(34) & ":" & chr(34) & arrCubitos(1,i)  & chr(34) & chr(125) & chr(44)
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
	sTabla   = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cubitos"
	'
	JsonData = chr(123) & chr(34) & "data" & chr(34) & ":" & chr(91) & sTabla & chr(93) & chr(125)
	'
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'
	rsCubitos.Close
	Set rsCubitos = Nothing
	'
	conexion.close
	set conexion = nothing
	'	
%>