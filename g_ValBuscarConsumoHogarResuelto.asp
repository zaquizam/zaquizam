<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	' g_ValBuscarConsumoHogarResuelto.asp // 20ene21 - 06abr21
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'
	Dim idConsumo, strSQL, rsRespuestaInvestigacion, arrRespInv
	'
	idConsumo = Request.QueryString("idConsumo")
	'
	Set rsRespuestaInvestigacion = CreateObject("ADODB.Recordset")
	rsRespuestaInvestigacion.CursorType = adOpenKeyset 
	rsRespuestaInvestigacion.LockType = 2 'adLockOptimistic 	
	'
	strSQL = vbnullString
	strSQL = strSQL & " SELECT"
	strSQL = strSQL & " PH_InvestigacionItems.InvestigacionItems,"
	strSQL = strSQL & " PH_Consumo_Investigar_Detalle.Observaciones_enviadas,"
	strSQL = strSQL & " PH_Consumo_Investigar_Detalle.Observaciones_recibidas"
	strSQL = strSQL & " FROM"
	strSQL = strSQL & " PH_Consumo_Investigar_Detalle"
	strSQL = strSQL & " INNER JOIN PH_InvestigacionItems ON"
	strSQL = strSQL & " PH_Consumo_Investigar_Detalle.Id_items_investigacion = PH_InvestigacionItems.Id_InvestigacionItems"
	strSQL = strSQL & " WHERE"
	strSQL = strSQL & " PH_Consumo_Investigar_Detalle.Validado = 0"
	strSQL = strSQL & " AND"
	strSQL = strSQL & " PH_Consumo_Investigar_Detalle.Resuelto = 1"
	strSQL = strSQL & " AND"
	strSQL = strSQL & " PH_Consumo_Investigar_Detalle.Caso_Cerrado = 1"
	strSQL = strSQL & " AND"
	strSQL = strSQL & " PH_Consumo_Investigar_Detalle.Id_Consumo = " & idConsumo
	''
	rsRespuestaInvestigacion.open strSQL, conexion	
	'
	'Response.write strSQL 
	'Response.End
	'
	Response.ContentType = "application/json"	
	'
	If not rsRespuestaInvestigacion.EOF  Then
    	arrRespInv = rsRespuestaInvestigacion.GetRows()  ' Convert recordset to 2D Array
	end if
	'	
	sTabla=vbnullstring
	
	if IsArray(arrRespInv) then
	
		For i = 0 to ubound(arrRespInv, 2)
		
							sTabla  =   chr(123) &  chr(34) & "motivo"     & chr(34) & ":" & chr(34) & RemoverSaltodeLinea(arrRespInv(0,i))  & chr(34) & chr(44)
			'
			sTabla    =   	sTabla  &   chr(34)  & "comentario" & chr(34)  & ":" & chr(34) & RemoverSaltodeLinea(arrRespInv(1,i)) & chr(34) & chr(44)
			
			sTabla    =   	sTabla  &   chr(34)  & "respuesta"  & chr(34)  & ":" & chr(34) & RemoverSaltodeLinea(arrRespInv(2,i)) & chr(34) & chr(125) & chr(44)
			
			sTablaJson = sTablaJson & sTabla
			sTabla=vbnullstring
			
		next				
		'
	else
		'Eof()
		'sTablaJson = sTablaJson & sTabla
		sTabla=vbnullstring
		'
						sTabla  =   chr(123) &  chr(34) & "motivo"     & chr(34) & ":" & chr(34) & "No Aplica" & chr(34) & chr(44)
		'
		sTabla    =   	sTabla  &   chr(34)  & "comentario" & chr(34)  & ":" & chr(34) & "No Aplica" & chr(34) & chr(44)
		
		sTabla    =   	sTabla  &   chr(34)  & "respuesta"  & chr(34)  & ":" & chr(34) & "No Aplica" & chr(34) & chr(125) & chr(44)
		'				
		sTablaJson = sTablaJson & sTabla
		sTabla=vbnullstring
		
		'JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	end if
	'	
	sTabla   = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData = chr(91) & sTabla & chr(93) 
	'JsonData = sTabla 
	Response.Write(JsonData)
	'
	rsRespuestaInvestigacion.close : set rsRespuestaInvestigacion = Nothing 
	conexion.Close : Set conexion = Nothing
	'
	
FUNCTION RemoverSaltodeLinea(byval str)
	'
	IF isNull(str) THEN
		str = ""
	END IF
	str = REPLACE(str,vbCr,"")			'Chr(13)
	str = REPLACE(str,vbLf,"")			'Chr(10)
	str = REPLACE(str,VbCrlf,"")		'Chr(13)+Chr(10)
	str = REPLACE(str,vbNewLine,"")		'vbNewLine
	str = REPLACE(str,vbFormFeed,"")	'Chr(12)
	str = REPLACE(str,vbTab,"")			'Chr(9)
	''
	RemoverSaltodeLinea = Trim(str)
	'
END FUNCTION	
	
	
	
%>