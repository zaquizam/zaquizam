<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
' g_ValBuscarTipoComida.asp // 31ene21
'
Session.lcid = 1034
Response.CodePage = 65001
Response.CharSet = "utf-8"
'
Dim idQuery, QrySql, rsComida
'
' Buscar Datos de todas las Comidas Registrados
'
QrySql = vbnullstring
QrySql = QrySql & " SELECT"
QrySql = QrySql & " PH_TipoComida.Id_tipoComida AS id,"
QrySql = QrySql & " PH_TipoComida.Comida AS nombre"
QrySql = QrySql & " FROM"
QrySql = QrySql & " PH_TipoComida"
QrySql = QrySql & " WHERE"
QrySql = QrySql & " PH_TipoComida.Ind_Activo = 1"
QrySql = QrySql & " ORDER BY"
QrySql = QrySql & " PH_TipoComida.Comida ASC"
'
'Response.Write QrySql & "<BR><BR>"
'Response.end
'
Set rsComida = Server.CreateObject("ADODB.recordset")
rsComida.Open QrySql,conexion
'
if not rsComida.EOF then
	arrComida = rsComida.GetRows()  ' Convert recordset to 2D Array
end if
'
Response.ContentType = "application/json"
'
'Crear Archivo Array Json
'
sTabla = vbnullstring
if IsArray(arrComida) then

	For i = 0 to ubound(arrComida, 2)
		'		
		sTabla     =  chr(123) &  chr(34) & "id" 	  & chr(34) & ":" & arrComida(0,i) & chr(44)
        sTabla     =  sTabla   &  chr(34) & "nombre"  & chr(34) & ":" & chr(34) & arrComida(1,i)  & chr(34) & chr(125) & chr(44)
        sTablaJson =  sTablaJson & sTabla
		sTabla = vbnullstring		
		'
	Next

else
	'Eof()
	sTabla  =   chr(123) &  chr(34) & "id" 		& chr(34) & ":" & "0" & chr(44)
	sTabla  =   sTabla   &  chr(34) & "nombre"  & chr(34) & ":" & chr(34) & "No Aplica" & chr(34) & chr(125) & chr(44)	
	sTablaJson = sTablaJson & sTabla
	sTabla = vbnullstring		
	'
end if
''
sTabla = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Comida"
'JsonData = chr(123) & chr(34) & "data" & chr(34) & ":" & chr(91) & sTabla & chr(93) & chr(125)
JsonData	= 	chr(91) & sTabla & chr(93) '& chr(125)
Response.Write(JsonData)
'
' Cerrar conexiones
'
rsComida.Close
Set rsComida = Nothing
'
conexion.close
set conexion = nothing
'
%>