<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_MC_ValBuscarLlenarCmbMonedaPagoFactura.asp - 26feb21
	'
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'
	Dim QrySql, rsMoneda
	'
	' Buscar Datos de todas las Monedas Registrados
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " PH_Moneda.Id_moneda as id,"
	QrySql = QrySql & " PH_Moneda.Moneda AS nombre"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_Moneda"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_Moneda.Ind_Activo = 1"
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_Moneda.Moneda ASC"
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsMoneda = Server.CreateObject("ADODB.recordset")
	rsMoneda.Open QrySql, conexion
	'
	if not rsMoneda.EOF then
		arrMoneda = rsMoneda.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	Response.ContentType = "application/json"
	'
	' Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrMoneda) then

		For i = 0 to ubound(arrMoneda, 2)
			'
			sTabla     =  chr(123) &  chr(34) & "id" 	 & chr(34) & ":" & chr(34) & arrMoneda(0,i)  & chr(34) & chr(44)
			sTabla     =  sTabla   &  chr(34) & "nombre" & chr(34) & ":" & chr(34) & arrMoneda(1,i)  & chr(34) & chr(125) & chr(44)
			sTablaJson =  sTablaJson & sTabla
			sTabla=""
			'
		next

	else
		'Eof()
		sTabla    =   chr(123) &  chr(34) & "id" 		& chr(34) & ":"  & chr(34)  & "0" 			& chr(34) & chr(44)
		sTabla    =   sTabla   &  chr(34) & "nombre"    & chr(34) & ":"  & chr(34)  & "No Aplica" 	& chr(34) & chr(125) & chr(44)
		'
		sTablaJson = sTablaJson & sTabla
		sTabla=""
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
	rsMoneda.Close
	Set rsMoneda = Nothing
	'
	conexion.close
	set conexion = nothing
	'	
%>