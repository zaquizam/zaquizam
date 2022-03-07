<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	'
	' g_ConvCalcularCombinacionesxHogaresPasdo4.asp - 10abr21 - 29abr21
	' (Refresco = "1" - Agua = 2 - Te = 3 - Jugo = 4)
	'
	Session.lcid = 1034
	Response.CodePage = 65001	
	Server.ScriptTimeout = 10000
	Response.CharSet = "utf-8"
	Response.Buffer = True
	'	
	Dim QrySql, arrResultados(10000,3), idMeses
	Dim totalHogares, totalRefresco 
	Dim dataArray, rsx1
    '
	idMeses = Request.QueryString("id_Mes")	
	'idMeses="16,17,18,19,20,21,22,23,24,25,26,27,28" ' Trimestral
	'	
	' Calcular Total hogares del Mes 
	'
	set rsx1 = CreateObject("ADODB.Recordset")
	'rsx1.CursorType = adOpenKeyset 
	'rsx1.LockType   = 2 'adLockOptimistic 
	'
	QrySql = vbnullstring
	QrySql = " SELECT" & _
	" PH_DataCrudaMensual.Id_Hogar," & _
	" PH_DataCrudaMensual.Id_Categoria" & _
	" FROM" & _
	" PH_DataCrudaMensual" & _
	" WHERE" & _
	" PH_DataCrudaMensual.Id_Fabricante <> 0" & _
	" AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")" & _
	" GROUP BY" & _
	" PH_DataCrudaMensual.Id_Hogar," & _
	" PH_DataCrudaMensual.Id_Categoria" & _
	" HAVING" & _
	" PH_DataCrudaMensual.Id_Categoria IN (1, 3 ,12 , 22)" & _
	" ORDER BY" & _
	" PH_DataCrudaMensual.Id_Hogar;"
	'
	' Response.write QrySql & "<br>"
	'
    rsx1.Open QrySql, conexion, 0, 1		
	'
	if rsx1.eof then
		rsx1.close
	else		
		dataArray = rsx1.GetRows
		rsx1.close
		'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
	end if
	'	
	Set rsx1 = nothing
	conexion.close : Set conexion = Nothing
	'
	'Response.write "Recordset Ok<br>"
	'Response.end
	'	
	if IsArray(dataArray) then
		'
		'Response.write "Total: " & ubound(dataArray,2) + 1 
		totalHogares = ubound(dataArray,2) + 1 
		'Response.write "<br>"
		'
		' for iReg = 0 to ubound(dataArray,2)
			' Response.write "<br>" &  dataArray(0,iReg)  & "=>" & dataArray(1,iReg)
			' Response.Write "Hogar : " & dataArray(0,i) & " Categoria : " & dataArray(1, i) & "<BR>"
		' next
		'
		' Llenar la matriz Resultante con Ceros
		'	
		For i = 1 to 10000								
			arrResultados(i,0) = 0
			arrResultados(i,1) = 0
			arrResultados(i,2) = 0
			arrResultados(i,3) = 0
		Next
		'
		totalRefresco=0
		FOR  i = 0 to ubound(dataArray,2) 
			'
			 hogar  = dataArray(0,i)
			 Categ  = dataArray(1,i)
			' 
			' Response.Write "Hogar : " & hogar & " Categoria : " & Categ & "<BR>"
			' Response.Write "Tipo  : " & TypeName(hogar) & " Categoria : " & TypeName(Categ) & "<BR>"
			'
			if Categ = 1 then
				'Refresco
				arrResultados(hogar,0) = 1
				' Contador de Hogares que compraron refrescos 
				totalRefresco = totalRefresco + 1
			end if
			if Categ = 3 then
				'agua			
				arrResultados(hogar,1) = 1			
			end if
			if Categ = 12 then
				'te			
				arrResultados(hogar,2) = 1			
			end if
			if Categ = 22 then
				'jugo			
				arrResultados(hogar,3) = 1
			end if
			'			
		NEXT	
		'
		'response.write "<br> Hogares con Refresco " & totalRefresco				
		'		
		' Refresco = "1" - Agua = 2 - Te = 3 - Jugo = 4
		'
		totalHogares=totalRefresco
		'
		Contador=0
		For i = 1 to 10000										
			if arrResultados(i,0) = 1 and arrResultados(i,1) = 0 and arrResultados(i,2) = 0 and arrResultados(i,3) = 0 then
				Contador = Contador + 1
			end if			
		Next
		SoloRefresco = Contador * 100 / totalHogares
		'Response.Write " Solo Refresco " & contador
		'
		Contador=0
		For i = 1 to 10000											
			if arrResultados(i,0) = 1 and arrResultados(i,1)=1 and arrResultados(i,2)=0 and arrResultados(i,3)= 0 then
				Contador = Contador + 1
			end if
		Next
		RefrescoAgua = Contador * 100 / totalHogares
		'
		Contador=0
		For i = 1 to 10000											
			if arrResultados(i,0) = 1 and arrResultados(i,1)=0 and arrResultados(i,2)=0 and arrResultados(i,3)= 1 then
				Contador = Contador + 1
			end if
		Next
		RefrescoJugo = Contador * 100 / totalHogares
		'
		Contador=0
		For i = 1 to 10000								
			if arrResultados(i,0) = 1 and arrResultados(i,1)=0 and arrResultados(i,2)=1 and arrResultados(i,3)=0 then
				Contador = Contador + 1
			end if
		Next
		RefrescoTe = Contador * 100 / totalHogares
		'
		Contador=0
		For i = 1 to 10000								
			if arrResultados(i,0) = 1 and arrResultados(i,1)= 1 and arrResultados(i,2)=0 and arrResultados(i,3)= 1 then
				Contador = Contador + 1
			end if
		Next
		RefrescoJugoAgua = Contador * 100 / totalHogares
		'
		Contador=0
		 For i = 1 to 10000								
			if arrResultados(i,0) = 1 and arrResultados(i,1)=0 and arrResultados(i,2)= 1 and arrResultados(i,3)= 1 then
				Contador = Contador + 1
			end if
		Next
		RefrescoJugoTe = Contador * 100 / totalHogares
		'
		Contador=0
		For i = 1 to 10000								
			if arrResultados(i,0) = 1 and arrResultados(i,1)= 1 and arrResultados(i,2)= 1 and arrResultados(i,3)= "" then
				Contador = Contador + 1
			end if
		Next
		RefrescoAguaTe = Contador * 100 / totalHogares
		'
		Contador=0
		For i = 1 to 10000								
			if arrResultados(i,0) = 1 and arrResultados(i,1) = 1 and arrResultados(i,2) = 1 and arrResultados(i,3) = 1 then
				Contador = Contador + 1
			end if
		Next
		RefrescoAguaTeJugo = Contador * 100 / totalHogares
		'
		total = 0
		total = Round(SoloRefresco,2) + Round(RefrescoAgua,2) + Round(RefrescoJugo,2) + Round(RefrescoTe,2) + Round(RefrescoJugoAgua,2) + Round(RefrescoJugoTe,2) + Round(RefrescoAguaTe,2) + Round(RefrescoAguaTeJugo,2)		
		'
		' Graficar los resultados en tablas
		'
		Response.Write "<strong><table class='table table-borderless table-condensed table-hover' style=' margin: auto; width: 65% !important;'>"
		'Response.Write "<strong><table class='table table-borderless table-condensed table-hover' style=' margin: auto;'>"
			Response.Write "<thead>"
				Response.Write "<tr>"
					Response.Write "<th colspan='2' class='text-center text-danger'><i class='fas fa-check-double'></i>&nbsp;QUE PORCENTAJES DE HOGARES COMPRARON LAS SIGUIENTES COMBINACIONES&nbsp;</th>"	  
				Response.Write "</tr>"
		   Response.Write "</thead>"
		   Response.Write "<tbody>"
			  Response.Write "<tr>"
				 Response.Write "<td>Solo refresco</td>"
				 Response.Write "<td class='text-right text-primary'>" & FormatNumber(SoloRefresco,2) & " %</td>"
			  Response.Write "</tr>"
			  Response.Write "<tr>"
				Response.Write "<td>Refresco y Agua</td>"
				Response.Write "<td class='text-right text-primary'>" & FormatNumber(RefrescoAgua,2) & " %</td>"
			  Response.Write "</tr>"
			  Response.Write "<tr>"
				Response.Write "<td>Refresco y Jugo</td>"
				Response.Write "<td class='text-right text-primary'>" & FormatNumber(RefrescoJugo,2) & " %</td>"
			  Response.Write "</tr>"
			  Response.Write "<tr>"
				 Response.Write "<td>Refresco y T&eacute;</td>"
				 Response.Write "<td class='text-right text-primary'>" & FormatNumber(RefrescoTe,2) & " %</td>"
			  Response.Write "</tr>"
			  Response.Write "</tr>"
			  Response.Write "<tr>"
				 Response.Write "<td>Refresco, Jugos y Agua</td>"
				 Response.Write "<td class='text-right text-primary'>" & FormatNumber(RefrescoJugoAgua,2) & " %</td>"
			  Response.Write "</tr>"
			  Response.Write "<tr>"
				 Response.Write "<td>Refresco, Jugos y T&eacute;</td>"
				 Response.Write "<td class='text-right text-primary'>" & FormatNumber(RefrescoJugoTe,2) & " %</td>"
			  Response.Write "</tr>"
			  Response.Write "<tr>"
				 Response.Write "<td>Refresco, Agua y T&eacute;</td>"
				 Response.Write "<td class='text-right text-primary'>" & FormatNumber(RefrescoAguaTe,2) & " %</td>"
			  Response.Write "</tr>"
			  Response.Write "<tr>"
				 Response.Write "<td>Refresco, Agua, Jugos y T&eacute;</td>"
				 Response.Write "<td class='text-right text-primary'>" & FormatNumber(RefrescoAguaTeJugo,2) & " %</td>"
			  Response.Write "</tr>"
			  '
			  Response.Write "<tr style=' background: #F9F9F9;' >"
				 Response.Write "<td class='text-center text-primary'>TOTAL</td>"
				 Response.Write "<td class='text-right text-primary'>" & FormatNumber(total,2) & " %</td>"
			  Response.Write "</tr>"
			  '
		   Response.Write "</tbody>"
		Response.Write "</table></strong>"
		'	
	else		
		'
		' Graficar los resultados en tablas NO HAY DATOS
		'
		Response.Write "<table class='table table-borderless table-hover' style=' margin: auto; width: 50% !important;'>"
			Response.Write "<thead>"
				Response.Write "<tr>"
					Response.Write "<th colspan='3' class='text-center text-danger'><i class='fas fa-check-double'></i>&nbsp;QUE PORCENTAJES DE HOGARES COMPRARON LAS SIGUIENTES COMBINACIONES&nbsp;</th>"	  
				Response.Write "</tr>"
		   Response.Write "</thead>"
		   Response.Write "<tbody>"
			  Response.Write "<tr>"
				Response.Write "<th colspan='3' class='text-center text-primary'><strong>....NO HAY DATOS PARA EL MES SELECCIONADO....</strong></th>"
			  Response.Write "</tr>"
		   Response.Write "</tbody>"
		Response.Write "</table>"
		'		
	end if	
	'
	Response.Flush
	'
%>
