<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
' g_ValBuscarCadenaxConsumo.asp
' 03ene21
'
Session.lcid=1034
Response.CodePage = 65001
Response.CharSet = "utf-8"
'
Dim idQuery, opcion,  QrySql
'
idQuery = Request.Form("id")
opcion  = Request.Form("opcion")
'
IF (opcion=1) THEN

	Dim rsCadena
	'
	' Buscar Datos de todas las Cadenas Registrados
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " PH_Cadena.Id_Cadena AS id,"
	QrySql = QrySql & " PH_Cadena.Cadena AS nombre"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_Cadena"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_Cadena.Ind_Activo = 1"
	if( CInt(id) > 0 )then
		QrySql = QrySql & " AND"
		QrySql = QrySql & " PH_Cadena.Id_Canal =" & idQuery
	end if
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_Cadena.Cadena ASC"
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsCadena = Server.CreateObject("ADODB.recordset")
	rsCadena.Open QrySql,conexion
	'
	if not rsCadena.EOF then
		arrCadena = rsCadena.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	'Response.ContentType = "application/json"
	'
	'Crear Archivo Array Json
	'
	sTabla=""

	if IsArray(arrCadena) then

		For i = 0 to ubound(arrCadena, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrCadena(0,i)  & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrCadena(1,i)  & chr(34) & chr(125)&chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla=""
			'
		next

	else
		'Eof()
		sTabla    =   chr(123)&  chr(34) & "id" 			& chr(34)& ":" & chr(34) & "0" 			& chr(34) & chr(44)
		sTabla    =   sTabla &  chr(34) & "nombre"         & chr(34)& ":" & chr(34)  & "No Aplica" 	& chr(34) & chr(125)&chr(44)
		''
		sTablaJson = sTablaJson & sTabla
		sTabla=""

	end if
	''
	sTabla = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'
	rsCadena.Close
	Set rsCadena = Nothing
	'
	conexion.close
	set conexion = nothing
	'
elseIF (opcion=2) THEN

	Dim rsCanal, arrCanal
	'
	' Buscar Datos de todas las Cadenas Registrados
	'
	QrySql = vbnullstring
    QrySql = QrySql & " SELECT"
	QrySql = QrySql & " PH_Canal.Id_Canal AS id,"
	QrySql = QrySql & " PH_Canal.Canal AS nombre"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_Canal"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_Canal.Ind_Activo = 1"	
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_Canal.Canal ASC"
    '	
	Set rsCanal = Server.CreateObject("ADODB.recordset")
	rsCanal.Open QrySql, conexion
	'
	if not rsCanal.EOF then
    	arrCanal = rsCanal.GetRows()  ' Convert recordset to 2D Array
	end if
	'	
	rsCanal.Close
	Set rsCanal = Nothing
	'	
	' Response.ContentType = "application/json"
	'
	' Crear Archivo Array Json
	'
	sTabla=""

	if IsArray(arrCanal) then

		For i = 0 to ubound(arrCanal, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrCanal(0,i)  & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrCanal(1,i)  & chr(34) & chr(125)&chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla=""
			'
		next

	else
		'Eof()
		sTabla    =   chr(123)&  chr(34) & "id" 			& chr(34)& ":" & chr(34) & "0" 			& chr(34) & chr(44)
		sTabla    =   sTabla &  chr(34) & "nombre"         & chr(34)& ":" & chr(34)  & "No Aplica" 	& chr(34) & chr(125)&chr(44)
		''
		sTablaJson = sTablaJson & sTabla
		sTabla=""

	end if
	''
	sTabla = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'		
	conexion.close
	set conexion = nothing
	'
end if

'
%>