<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
Session.LCID = 8202 
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	ynum=Request.Form("id")

	dim gDatosSol
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_OcupacionVivienda, "	'0
	sql = sql & " OtroOcupacionVivienda, "	'1
	sql = sql & " Id_MontoVivienda "		'2
	sql = sql & " FROM "
	sql = sql & " PH_PanelHogar "
	sql = sql & " WHERE "
	sql = sql & " Id_PanelHogar = " & ynum
	'response.write "<br>220 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	'response.write "<br> Linea 223 " &
	'response.end
	iExiste = 0
	'
	if not rsx1.EOF then
		arrBloque4 = rsx1.GetRows()  ' Convert recordset to 2D Array
	end if
		'
	rsx1.Close
	Set rsPanelConsumo = Nothing
	sTabla=vbnullstring
	if IsArray(arrBloque4) then
			For i = 0 to ubound(arrBloque4, 2)
				sTabla = chr(123)&  chr(34) & "Id_OcupacionVivienda"	& chr(34)& ":" & chr(34) & arrBloque4(0,i)  & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "OtroOcupacionVivienda" 	& chr(34)& ":" & chr(34) & arrBloque4(1,i)	& chr(34) & chr(44)
				'
				sTabla = sTabla  &  chr(34) & "Id_MontoVivienda"        & chr(34)& ":" & chr(34) & arrBloque4(2,i) & chr(34) & chr(125)&chr(44)
				sTablaJson = sTablaJson & sTabla
				sTabla=vbnullstring
			next
			sTabla = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
			JsonData= chr(91) & sTabla & chr(93) '& chr(125)
		else
			'Eof()
			sTablaJson = sTablaJson & sTabla
			sTabla=vbnullstring
			JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
		end if
		Response.Write(JsonData)
		conexion.close
		set conexion = nothing
%>