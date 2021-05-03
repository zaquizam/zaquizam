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
	sql = sql & " CodigoHogar, "			' 0
	sql = sql & " Id_Pais, "				' 1
	sql = sql & " Id_Estado, "				' 2
	sql = sql & " Id_Ciudad, "				' 3
	sql = sql & " Id_Municipio, "			' 4
	sql = sql & " Id_Parroquia, "			' 5
	sql = sql & " Calle, "					' 6
	sql = sql & " Edificio, "				' 7
	sql = sql & " Casa, "					' 8
	sql = sql & " Escalera, "				' 9
	sql = sql & " Piso, "					'10
	sql = sql & " Apto, "					'11
	sql = sql & " Barrio, "					'12
	sql = sql & " Referencia, "				'13
	sql = sql & " TelefonoLocal "			'14	
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
		arrBloque0 = rsx1.GetRows()  ' Convert recordset to 2D Array
	end if
		'
	rsx1.Close
	Set rsPanelConsumo = Nothing
	sTabla=vbnullstring
	if IsArray(arrBloque0) then
			For i = 0 to ubound(arrBloque0, 2)
				sTabla    =    chr(123)&  chr(34) & "CodigoHogar"	& chr(34)& ":" & chr(34) & arrBloque0(0,i)   	& chr(34) & chr(44)
				sTabla    =    sTabla &  chr(34) & "Id_Pais" 		& chr(34)& ":" & chr(34) & arrBloque0(1,i)		& chr(34) & chr(44)
				sTabla    =    sTabla &  chr(34) & "Id_Ciudad" 		& chr(34)& ":" & chr(34) & arrBloque0(3,i)		& chr(34) & chr(44)
				sTabla    =    sTabla &  chr(34) & "Id_Municipio" 	& chr(34)& ":" & chr(34) & arrBloque0(4,i)		& chr(34) & chr(44)
				sTabla    =    sTabla &  chr(34) & "Id_Parroquia" 	& chr(34)& ":" & chr(34) & arrBloque0(5,i)		& chr(34) & chr(44)
				sTabla    =    sTabla &  chr(34) & "Calle" 			& chr(34)& ":" & chr(34) & arrBloque0(6,i)		& chr(34) & chr(44)
				sTabla    =    sTabla &  chr(34) & "Edificio"	 	& chr(34)& ":" & chr(34) & arrBloque0(7,i)		& chr(34) & chr(44)
				sTabla    =    sTabla &  chr(34) & "Casa" 			& chr(34)& ":" & chr(34) & arrBloque0(8,i)		& chr(34) & chr(44)
				sTabla    =    sTabla &  chr(34) & "Escalera"	 	& chr(34)& ":" & chr(34) & arrBloque0(9,i)		& chr(34) & chr(44)
				sTabla    =    sTabla &  chr(34) & "Piso" 			& chr(34)& ":" & chr(34) & arrBloque0(10,i)		& chr(34) & chr(44)
				sTabla    =    sTabla &  chr(34) & "Apto"		 	& chr(34)& ":" & chr(34) & arrBloque0(11,i)		& chr(34) & chr(44)
				sTabla    =    sTabla &  chr(34) & "Barrio" 		& chr(34)& ":" & chr(34) & arrBloque0(12,i)		& chr(34) & chr(44)
				sTabla    =    sTabla &  chr(34) & "Referencia" 	& chr(34)& ":" & chr(34) & arrBloque0(13,i)		& chr(34) & chr(44)
				sTabla    =    sTabla &  chr(34) & "TelefonoLocal" 	& chr(34)& ":" & chr(34) & arrBloque0(14,i)		& chr(34) & chr(44)
				'
				sTabla    =    sTabla &  chr(34) & "Id_Estado"        & chr(34)& ":" & chr(34) & arrBloque0(2,i)    & chr(34) & chr(125)&chr(44)
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