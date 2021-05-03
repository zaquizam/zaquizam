
<!--img  src="http://chart.apis.google.com/chart?cht=p3&chd=s:hW&chs=500x200&chl=Penetracion|World"-->


<%


Sub d_GraLinDat (g_idata, g_itit , g_ireg, g_iir)

  rr_sColFon2=Mid(rr_sColfon,2)


' Pagina de Google Chart
	sParGra ="http://chart.apis.google.com/chart"

' Tipo de Gráfico (cht)
	'sx	=  "?cht=p" ' Pie
	sx	=  "?cht=lc" ' Linea
	'sx	=  "?cht=bvg" ' barras
	sParGra = sParGra & sx

' Tamaño del gráfico (chs)
	sx = "&chs=600x300"	
	sx = "&chs=900x300"	
	sx = "&chs=800x375"	
	sx = "&chs=750x400"	
	sParGra = sParGra & sx
	
' Especificar el Background (chf)
	sx = "&chf=c,lg,0,76A4FB,1,ffffff,0|bg,s,"& rr_sColFon2
	sx = "&chf=c,lg,0,ffffff,1,ffffff,0|bg,s,"& rr_sColFon2
	'sx = "&chf=bg,s,"& rr_sColFon2
	if rr_ipExe=1 then sx = "&chf=bg,s,ffffff"
	sParGra = sParGra & sx
	
' Especificar las lineas del grid (chg)
	sx = "&chg=9.09,25,1,5"
	sParGra = sParGra & sx


' Titulo del Gráfico (chtt)(Espacio con + y salto de linea con | )
    sx=""
   
    for ia=1 to 1 'rr_nFil
         if g_ireg=0 and ia=1 then
            sy= " Total " & rr_sFil(ia,1) & ""
         else
            sy=rr_gTit(g_iReg,ia )
         end if   
        ' sy=mid(sy,3)
         sx=sx & sy 
         if ia<> rr_nFil then sx= sx & "|" 
    next    
	ilen=len(sx)
	'response.Write sx & "  len:=" & ilen & "<br>"

    sx=replace(sx,"+","&#43;")
	sx=replace(sx," ","+")
	sx=replace(sx,"á","&aacute;")
	sx=replace(sx,"é","&eacute;")
	sx=replace(sx,"í","&iacute;")
	sx=replace(sx,"ó","&oacute;")
	sx=replace(sx,"ú","&uacute;")
	sx=replace(sx,"´","")
	sx=replace(sx,"#","")
'	sx=replace(sx,"/","")

	
'	response.Write "<br>69" &  sx 
	sx = "&chtt=" &  sx
	sx = sx & "&chts=000000,18"
	sParGra = sParGra & sx



' Especificar la leyenda del Gráfico (chdl)
	sx="&chdl="
	sTit=rr_gTit(g_ireg,1)
	for iiFil=g_ireg to rr_iMaxReg
        if rr_gTit(iifil,1)<> sTit then exit for
        ix=ix+1
     '   if ix=5 then exit for
		sY=   rr_gTit(iiFil,2 )
		
		ilen=len(sy)
		for j=1 to ilen
			sz = mid(sy,j,1)
			if sz=" " then 
				sy = mid(sy,1,j-1) & "+" & mid (sy,j+1)
			end if
			if sz="%" then 
				sy = mid(sy,1,j-1) & "S." & mid (sy,j+1)
			end if		
		next
		sx = sx & sy & "|" 
	next
	ix=len(sx)
	sx=Mid(sx,1,ix-1)
	sx=replace(sx,"+"," ")
	sx=replace(sx," ","+")
	sx=replace(sx,"á","a")
	sx=replace(sx,"é","e")
	sx=replace(sx,"í","i")
	sx=replace(sx,"ó","o;")
	sx=replace(sx,"ú","u")
	sx=replace(sx,"´","")
	sx=replace(sx,"ñ","n")
	sx=replace(sx,"#","")
	'sx=replace(sx,"/","")
	
'response.write "<br>100 " & sx
	sParGra = sParGra & sx	
	

' Especificar ubicación de la leyenda r=derecha l=izquierda t= arriba b=abajo
	sx="&chdlp=r"
	sx="&chdlp=t"
	sParGra = sParGra & sx
' Especificar color de las lineas
	sx="&chco=0000ff,ff0000,FF9900,008000,ff00ff,33ffff,d2691e,00ff00,990066,c0c0c0,1e90ff,000000,ffff00"
	sParGra = sParGra & sx

' Calcular valor máximo del gráfico
	g_iMaxVal = 0
	g_iMinVal=999999999
	for iiFil=g_ireg to rr_iMaxReg
	for iiCol=1 to rr_ncol
	    for iiVar=1 to 1
'response.write "<br>125 " & g_iData(iifil,iicol,iiVar)	    
			if g_iData(iifil,iicol,iiVar)> g_iMaxVal then g_iMaxVal=g_iData(iifil,iicol,iiVar)
			if g_iData(iifil,iicol,iiVar)<>0 then 
			    if g_iData(iifil,iicol,iiVar)< g_iMinVal then g_iMinVal=g_iData(iifil,iicol,iiVar)
			end if    
		next
	next
	next	
	
'response.Write "<br>133 Valor Máximo:=" & g_iMaxVal
	if g_iMaxVal>1200 then
		ix=g_iMaxVal/500
		ix=int(ix)
		ix=ix+1
		g_iMaxVal= ix *500
	end if
	
	if g_iMaxVal>1000 and g_iMaxVal<1200 then g_iMaxVal=1200
	if g_iMaxVal>800 and g_iMaxVal<1000 then g_iMaxVal=1000
    if g_iMaxVal>799 and g_iMaxVal<900 then g_iMaxVal=900				
    if g_iMaxVal>699 and g_iMaxVal<800 then g_iMaxVal=800				
    if g_iMaxVal>599 and g_iMaxVal<700 then g_iMaxVal=700				
    if g_iMaxVal>499 and g_iMaxVal<600 then g_iMaxVal=600			
	if g_iMaxVal>399 and g_iMaxVal<500 then g_iMaxVal=500		
    if g_iMaxVal>299 and g_iMaxVal<400 then g_iMaxVal=400		
	if g_iMaxVal>199 and g_iMaxVal<300 then g_iMaxVal=300	
	if g_iMaxVal>102 and g_iMaxVal<200 then g_iMaxVal=200
	if g_iMaxVal>75 and g_iMaxVal<100 then g_iMaxVal=100
	if g_iMaxVal>50 and g_iMaxVal<75 then g_iMaxVal=75
	if g_iMaxVal<1 then g_iMaxVal=1
'	g_ImaxVal=100
'	response.Write "<br>138 Valor Máximo:=" & g_iMaxVal
'	response.Write "<br>138 Valor Míximo:=" & g_iMinVal
	g_iMAxVal=g_iMAxVal+g_iMinVal/2
	g_iMAxVal=int(g_iMAxVal)
    sx=cstr(g_iMAxVal)
    iLen=len(sx)-1
    sy=mid(sx,1,1)
    for i=1 to Ilen
        sy=sy&"0"
    next
    g_iMaxVal= cLNG(sy)
	'response.Write "<br>138 Valor Máximo:=" & g_iMaxVal  & " Ilen:=" & Ilen   & " sx:=" & sx 

	
'response.Write "<br>126" & sx &  " Maxreg:=" & rr_iMaxReg & "iireg:=" &  g_iReg &  " g_iir:=" & g_iir



' Especificar la data del Grafico (chd)
	sx = "&chd=t:"
	
	sTit=rr_gTit(g_ireg,1)
	for iiFil=g_ireg to rr_iMaxReg
      if rr_gTit(iifil,1)<> sTit then exit for
	for iiCol=1 to rr_nCol
	   ' if rr_sVar(iiVar,5)=1 then
	        if rr_gData(0,iiCol, 1)<>0 then 
	        
'response.write "<br> Data:=" & g_iData(iifil,iiCol, 1)       
	            iV=g_iData(iifil,iiCol, 1)
                iV= iV *  rr_sVar(1,1 )
            else
                iv=0
            end if        
	    ' else
	   '     iV= rr_gData(iifil,iiCol, iiVar) *  rr_sVar(iiVar,1 )
    	'end if   	
	    
		'	ix = g_iData(g_ireg,iicol,iiVar)
			ix=iv
          ' ix=99
			ix = ix/g_iMaxVal*100
'response.write "<br> " & Ix					
			'ix = ix
			ix = int(ix)
			if ix then 
			else
				ix =0
			end if
		'	response.Write "<br>ix:=" & ix
			if iiCol<> rr_ncol then sx = sx & ix & "," else sx = sx & ix & "|"
		next
	next
	ilen = len(sx)
	sx = mid(sx,1,ilen-1)
'response.Write "<br>181 Data:=" & sx 
	sParGra = sParGra & sx
'exit sub	

' Eje de X y Y (chxl)	
	sx = "&chxt=x,y,r,x"
	sParGra = sParGra & sx

' Nombre del Eje de las Y (chxl)
	sx="&chxl=0:|"
	sx3="3:|"
	For cc=1 to rr_nCol 'rr_iMaxPer
		'response.write "<br>196 " & rr_gData(0,cc,0)
		'sx1=Replace(rr_gData(0,cc,0)," 2015","-2015") 
		'sx1=Replace(sx1," <br>"," ")  
		sx1=rr_gData(0,cc,0)
        sj=sx1

		ilen = instr(sj,"<br>")
		if ilen<>0 then
			sy=mid(sj,1,ilen-1)
			sz=mid(sj,ilen+4)
		else
		    sy=sj
		    sz= " "
		end if	

		sx = sx & sy  & "|"
		sx3= sx3 + sz & "|"
	next

'response.Write "<br>210" & sx
'response.Write "<br>210" & sx3


' Nombre del Eje de las X	
	'sx = sx & "1:|0|500|1000|1500"
	sx = sx & "1:|0"
	g_iMaxVal =int(g_iMaxVal)
	ix= int(g_iMaxVal/4)
	sx = sx & "|" & ix
	iy =ix + ix
	sx = sx & "|" & iy
	iy = iy +ix
	sx = sx & "|" & iy
	sx = sx & "|" & g_iMaxVal
	'sx = sx & "|" & ix
	
	sx2="|2:||"
	
	ilen = len(sx3)
	sx3=mid(sx3,1,ilen-1)
	sx = sx & sx2 &  "|" & sx3
'response.Write "<br>235" & sx
	
	sParGra = sParGra & sx 

'
'	sx="&chf=c,ls,0,ffff00,0.2,FFFFF0,0.1"
'	sParGra = sParGra & sx
	
     sx=rr_sColFon
     if rr_ipExe=1 then sx="#ffffff"
%>	

	
	<table border="0" width="950px" id="Table4" bgcolor="<%=sx %>" align="center">
		<tr>
			
			<td align="center" ><img  src="<%=sParGra%>" /></td>
		</tr>
		<tr><td colspan="2" height="30"></td></tr>
	</table>
	    
<%
end Sub

Sub GraMue 

' Pagina de Google Chart
	sApi ="http://chart.apis.google.com/chart"

' Tipo de Gráfico (cht)
	sx	=  "?cht=lc"
	sx	=  "?cht=bvg"
	sApi = sApi & sx

' Tamaño del gráfico (chs)
	sx = "&chs=700x300"	
	'sx = "&chs=800x375"	
	sApi = sApi & sx
	
' Ancho de la barra
	sx= "&chbh=36,7"
	sApi = sApi & sx

' Especificar el Background (chf)
	sx = "&chf=c,lg,0,76A4FB,1,ffffff,0"
	sx = "&chf=c,lg,0,76A4FB,1,ffffff,0|bg,s,fffff0"
	'sx = "&chf=c,lg,0,4265A5,1,ffffff,0|bg,s,fffff0"
	'sx = "&chf=bg,s,fffff0"
	sApi = sApi & sx
	
' Especificar las lineas del grid (chg)
	sx = "&chg=9.09,25,1,5"
	sApi = sApi & sx

' Eje de X y Y (chxl)	
	sApi = sApi & "&chxt=x,y"
	
' Nombre del Eje de las Y (chxl)
	sx="&chxl=0:|"
	for ib= iMinCol to iMaxCol  
		sy=sNomPer(ib, 1)
		sx = sx & sNomPer(ib, 1) & "|"
	next
	sx = sx & "1:|0|500|1000|1500|2000"
	sApi = sApi & sx '"&chxl=0:|Mar|Apr|May|June|July|1:||50+Kb"
	
' Especificar color de las lineas
'	sx="&chco=0000ff,ff0000,FF9900,008000,ff00ff,33ffff,d2691e,00ff00,990066,c0c0c0,1e90ff,000000,ffff00"
	'sApi = sApi & sx
	
' Gradar Data (chd)
	chd = "&chd=t:"
	for ib= iMinCol to iMaxCol 
		ix = int(mDatMue(0,ib+2) /2000*100)
		'response.Write ix
		chd = chd & ix & ","
	next
	ilen = len(chd)
	chd = mid(chd,1,ilen-1)
	sApi = sApi & chd

' Titulo del Gráfico (chtt)(Espacio con + y salto de linea con | )	
	sx = "&chtt=Total+Venezuela|Muestra"
	sApi = sApi & sx
	sx= "&chts=000000,18"
	sApi = sApi & sx
		
'	sGra = sApi & cht & chd & chs & chxl & chxt & chtt & chf &chg
'	sGra = "http://chart.apis.google.com/chart?chs=200x125&chd=s:helloWorld&cht=lc&chxt=x,y&chxl=0:|Mar|Apr|May|June|July|1:||50+Kb"
'	response.Write "<br>" & sGra & "<br>" %>


	<table border=0 width=99% ID="Table3">
		<tr>
			<td width=20%></td>
			<td width=80%><img  src=<%=sApi%> > </td>
		</tr>
	</table>
	<br>

<%End Sub 



Sub MosGraInd ( iCua)

' Pagina de Google Chart
	sParGra ="http://chart.apis.google.com/chart"

' Tipo de Gráfico (cht)
	sx	=  "?cht=lc"
	'sx	=  "?cht=bvg"
	sParGra = sParGra & sx

' Tamaño del gráfico (chs)
	sx = "&chs=600x300"	
	sParGra = sParGra & sx
	
' Especificar el Background (chf)
	sx = "&chf=c,lg,0,76A4FB,1,ffffff,0|bg,s,fffff0"
	sx = "&chf=bg,s,fffff0"
	sx = "&chf=bg,s,fffff0"
	sParGra = sParGra & sx
	
' Especificar las lineas del grid (chg)
	sx = "&chg=9.09,25,1,5"
	sParGra = sParGra & sx

' Titulo del Gráfico (chtt)(Espacio con + y salto de linea con | )
	sx =sInd(iCua) 
	ilen=len(sx)
	'response.Write sx & "  len:=" & ilen & "<br>"
	for i=1 to ilen
		sy = mid(sx,i,1)
		if sy=" " then 
			sx = mid(sx,1,i-1) & "+" & mid (sx,i+1)
		end if	
		if sy="ó" then 
			sx = mid(sx,1,i-1) & "o" & mid (sx,i+1)
		end if	
		if sy="í" then 
			sx = mid(sx,1,i-1) & "i" & mid (sx,i+1)
		end if	
	next
	'response.Write sx & "<br>"
	sx = "&chtt=" &  sx
	sx = sx & "&chts=000000,18"
	sParGra = sParGra & sx

' Especificar la leyenda del Gráfico (chdl)
	sx="&chdl="
	for i=1 to iLin
	'    if iVisMar = 1 then
			if mDat(i,1,0) = 0 then  sy=sMar(mDat(i,0,0)) else sy=sNomPro(mDat(i,1,0))
	'	else
	'		sy =sInd(i)	
	'	end if	
		'response.Write sy & "<br>"
		ilen=len(sy)
		for j=1 to ilen
			sz = mid(sy,j,1)
			if sz=" " then 
				sy = mid(sy,1,j-1) & "+" & mid (sy,j+1)
			end if	
		next
		if i<> iLin then sx = sx & sy & "|" else sx = sx & sy 
	next
'	response.Write sx & "<br>"
	sParGra = sParGra & sx
	
' Especificar color de las lineas
	sx="&chco=0000ff,ff0000,FF9900,008000,ff00ff,33ffff,d2691e,00ff00,990066,c0c0c0,1e90ff,000000,ffff00"
	sParGra = sParGra & sx
	
' Calcular valor máximo del gráfico
	g_iMaxVal = 0
	For ia=1 to iLin
		for ib= iMinCol to iMaxCol
			if mDat(ia,ib+2,iCua)> g_iMaxVal then g_iMaxVal = mDat(ia,ib+2,iCua)
		next
	next	
	
	'response.Write "<br>Valor Máximo:=" & g_iMaxVal
	if g_iMaxVal>1200 then
		ix=g_iMaxVal/500
		ix=int(ix)
		ix=ix+1
		g_iMaxVal= ix *500
	end if
	if g_iMaxVal>1000 and g_iMaxVal<1200 then g_iMaxVal=1200
	if g_iMaxVal>800 and g_iMaxVal<1000 then g_iMaxVal=1000
	if g_iMaxVal>75 and g_iMaxVal<100 then g_iMaxVal=100
	if g_iMaxVal>50 and g_iMaxVal<75 then g_iMaxVal=75
	if g_iMaxVal>25 and g_iMaxVal<50 then g_iMaxVal=50
	if g_iMaxVal>20 and g_iMaxVal<25 then g_iMaxVal=25
	if g_iMaxVal>5 and g_iMaxVal<10 then g_iMaxVal=10
	if g_iMaxVal>3 and g_iMaxVal<5 then g_iMaxVal=5
	if g_iMaxVal>1 and g_iMaxVal<3 then g_iMaxVal=3
	if g_iMaxVal<1 then g_iMaxVal=1
response.Write "<br>Valor Máximo:=" & g_iMaxVal
	

' Especificar la data del Grafico (chd)
	sx = "&chd=t:"
	For ia=1 to iLin
		for ib= iMinCol to iMaxCol 
			ix = mDat(ia,ib+2,iCua)
			ix = ix/g_iMaxVal*100
			'ix = ix
			ix = int(ix)
			if ib<> ImaxCol then sx = sx & ix & "," else sx = sx & ix & "|"
		next
	next
	ilen = len(sx)
	sx = mid(sx,1,ilen-1)
	'response.Write sx
	sParGra = sParGra & sx
	
' Eje de X y Y (chxl)	
	sx = "&chxt=x,y"
	sParGra = sParGra & sx
	
' Nombre del Eje de las Y (chxl)
	sx="&chxl=0:|"
	for ib= iMinCol to iMaxCol  
		sy=sNomPer(ib, 1)
		sx = sx & sNomPer(ib, 1) & "|"
	next
' Nombre del Eje de las X	
	'sx = sx & "1:|0|500|1000|1500"
	sx = sx & "1:|0"
	g_iMaxVal =int(g_iMaxVal)
	ix= int(g_iMaxVal)/4
	sx = sx & "|" & ix
	iy =ix + ix
	sx = sx & "|" & iy
	iy = iy +ix
	sx = sx & "|" & iy
	sx = sx & "|" & g_iMaxVal
	'sx = sx & "|" & ix
	'response.Write sx	
	sParGra = sParGra & sx
'
'	sx="&chf=c,ls,0,ffff00,0.2,FFFFF0,0.1"
	sParGra = sParGra & sx
	
	

%>
	<br>
	<table border=0 width=99%>
		<tr>
			<td width=20%></td>
			<td width=80% ><img  src=<%=sParGra%> ></td>
		</tr>
	</table>
	<br>


<%
end Sub

Sub d_GraLin (g_idata, g_itit , g_ireg)

'g_iL1, g_iL2,g_iL3, g_iCol)

' Pagina de Google Chart
	sParGra ="http://chart.apis.google.com/chart"

' Tipo de Gráfico (cht)
	'sx	=  "?cht=p" ' Pie
	sx	=  "?cht=lc" ' Linea
	'sx	=  "?cht=bvg" ' barras
	sParGra = sParGra & sx

' Tamaño del gráfico (chs)
	sx = "&chs=600x300"	
	sParGra = sParGra & sx
	
' Especificar el Background (chf)
	sx = "&chf=c,lg,0,76A4FB,1,ffffff,0|bg,s,fffff0"
	sx = "&chf=bg,s,fffff0"
	if rr_ipExe=1 then sx = "&chf=bg,s,ffffff"
	sParGra = sParGra & sx
	
' Especificar las lineas del grid (chg)
	sx = "&chg=9.09,25,1,5"
	sParGra = sParGra & sx


' Titulo del Gráfico (chtt)(Espacio con + y salto de linea con | )
    sx=""
   
    for ia=1 to rr_nFil
         if g_ireg=0 and ia=1 then
            sy= " Total " & rr_sFil(ia,1) & ""
         else
            sy=rr_gTit(g_iReg,ia )
         end if   
        ' sy=mid(sy,3)
         sx=sx & sy 
         if ia<> rr_nFil then sx= sx & "|" 
    next    
	ilen=len(sx)
	'response.Write sx & "  len:=" & ilen & "<br>"

    sx=replace(sx,"+","&#43;")
	sx=replace(sx," ","+")
	sx=replace(sx,"á","&aacute;")
	sx=replace(sx,"é","&eacute;")
	sx=replace(sx,"í","&iacute;")
	sx=replace(sx,"ó","&oacute;")
	sx=replace(sx,"ú","&uacute;")
	sx=replace(sx,"´","")

	
'	response.Write sx & "<br>"
	sx = "&chtt=" &  sx
	sx = sx & "&chts=000000,18"
	sParGra = sParGra & sx



' Especificar la leyenda del Gráfico (chdl)
	sx="&chdl="
	for i=1 to rr_nVar
		sY=rr_sVar(i,0) 
 
		ilen=len(sy)
		if ilen>10 then 
		    sy=mid(sy,1,10)
		    iLen=10
		end if    
		for j=1 to ilen
			sz = mid(sy,j,1)
			if sz=" " then 
				sy = mid(sy,1,j-1) & "+" & mid (sy,j+1)
			end if
			if sz="%" then 
				sy = mid(sy,1,j-1) & "S." & mid (sy,j+1)
			end if		
		next
		if i<> rr_nVar then sx = sx & sy & "|" else sx = sx & sy 
	next


    response.Write "<br>573 rr_nVar:=" &  rr_nVar
	response.Write "<br>573" & sx
	sParGra = sParGra & sx	
	

' Especificar ubicación de la leyenda r=derecha l=izquierda t= arriba b=abajo
	sx="&chdlp=r"
	sParGra = sParGra & sx
' Especificar color de las lineas
	sx="&chco=0000ff,ff0000,FF9900,008000,ff00ff,33ffff,d2691e,00ff00,990066,c0c0c0,1e90ff,000000,ffff00"
	sParGra = sParGra & sx

' Calcular valor máximo del gráfico
	g_iMaxVal = 0
	for iiCol=1 to rr_ncol
	    for iiVar=1 to rr_nVar
			if g_iData(g_ireg,iicol,iiVar)> g_iMaxVal then g_iMaxVal=g_iData(g_ireg,iicol,iiVar)
		next
	next	
	
'	response.Write "<br>Valor Máximo:=" & g_iMaxVal
	if g_iMaxVal>1200 then
		ix=g_iMaxVal/500
		ix=int(ix)
		ix=ix+1
		g_iMaxVal= ix *500
	end if
	if g_iMaxVal>1000 and g_iMaxVal<1200 then g_iMaxVal=1200
	if g_iMaxVal>800 and g_iMaxVal<1000 then g_iMaxVal=1000
	if g_iMaxVal>102 and g_iMaxVal<200 then g_iMaxVal=200
	if g_iMaxVal>75 and g_iMaxVal<100 then g_iMaxVal=100
	if g_iMaxVal>50 and g_iMaxVal<75 then g_iMaxVal=75
	if g_iMaxVal>25 and g_iMaxVal<50 then g_iMaxVal=50
	if g_iMaxVal>20 and g_iMaxVal<25 then g_iMaxVal=25
	if g_iMaxVal>5 and g_iMaxVal<10 then g_iMaxVal=10
	if g_iMaxVal>3 and g_iMaxVal<5 then g_iMaxVal=5
	if g_iMaxVal>1 and g_iMaxVal<3 then g_iMaxVal=3
	if g_iMaxVal<1 then g_iMaxVal=1
response.Write "<br>613 Valor Máximo:=" & g_iMaxVal
	

' Especificar la data del Grafico (chd)
	sx = "&chd=t:"
	
	for iiVar=1 to rr_nVar
	

	for iiCol=1 to rr_nCol
	    if rr_sVar(iiVar,5)=1 then
	        if rr_gData(0,iiCol, iiVar)<>0 then 
	            iV=rr_gData(g_ireg,iiCol, iiVar)/rr_gData(0,iiCol, iiVar)*100
                iV= iV *  rr_sVar(iiVar,1 )
            else
                iv=0
            end if        
	     else
	        iV= rr_gData(g_ireg,iiCol, iiVar) *  rr_sVar(iiVar,1 )
    	end if   	
	    
		'	ix = g_iData(g_ireg,iicol,iiVar)
			ix=iv

			ix = ix/g_iMaxVal*100
'response.write "<br> " & Ix					
			'ix = ix
			ix = int(ix)
			if ix then 
			else
				ix =0
			end if
		'	response.Write "<br>ix:=" & ix
			if iiCol<> rr_ncol then sx = sx & ix & "," else sx = sx & ix & "|"
		next
	next
	ilen = len(sx)
	sx = mid(sx,1,ilen-1)
response.Write "<br>160 Data:=" & sx
	sParGra = sParGra & sx
	

' Eje de X y Y (chxl)	
	sx = "&chxt=x,y,r,x"
	sParGra = sParGra & sx

' Nombre del Eje de las Y (chxl)
	sx="&chxl=0:|"
	sx3="3:|"
	For cc=1 to rr_nCol 'rr_iMaxPer
		
		sx1=rr_gData(0,cc,0)  
        sj=sx1

		ilen = instr(sj," ")
		if ilen<>0 then
			sy=mid(sj,1,ilen-1)
			sz=mid(sj,ilen+1)
		else
		    sy=sj
		    sz= " "
		end if	

		sx = sx & sy  & "|"
		sx3= sx3 + sz & "|"
	next

'response.Write "<br>" & sx
' Nombre del Eje de las X	
	'sx = sx & "1:|0|500|1000|1500"
	sx = sx & "1:|0"
	g_iMaxVal =int(g_iMaxVal)
	ix= int(g_iMaxVal/4)
	sx = sx & "|" & ix
	iy =ix + ix
	sx = sx & "|" & iy
	iy = iy +ix
	sx = sx & "|" & iy
	sx = sx & "|" & g_iMaxVal
	'sx = sx & "|" & ix
	
	sx2="|2:||"
	
	ilen = len(sx3)
	sx3=mid(sx3,1,ilen-1)
	sx = sx & sx2 &  "|" & sx3
'response.Write "<br>" & sx
	
	sParGra = sParGra & sx 

'
'	sx="&chf=c,ls,0,ffff00,0.2,FFFFF0,0.1"
'	sParGra = sParGra & sx
	
     sx=rr_sColFon
     if rr_ipExe=1 then sx="#ffffff"
     
%>	

	
	<table border="0" width="950" id="Table5" bgcolor="<%=sx %>" align="center">
		<tr>
			<td width="20%"></td>
			<td width="80%" ><img  src="<%=sParGra%>" /></td>
		</tr>
		<tr><td colspan="2" height="30"></td></tr>
	</table>
	    
<%
end Sub

Sub GraMue 

' Pagina de Google Chart
	sApi ="http://chart.apis.google.com/chart"

' Tipo de Gráfico (cht)
	sx	=  "?cht=lc"
	sx	=  "?cht=bvg"
	sApi = sApi & sx

' Tamaño del gráfico (chs)
	sx = "&chs=700x300"	
	'sx = "&chs=800x375"	
	sApi = sApi & sx
	
' Ancho de la barra
	sx= "&chbh=36,7"
	sApi = sApi & sx

' Especificar el Background (chf)
	sx = "&chf=c,lg,0,76A4FB,1,ffffff,0"
	sx = "&chf=c,lg,0,76A4FB,1,ffffff,0|bg,s,fffff0"
	'sx = "&chf=c,lg,0,4265A5,1,ffffff,0|bg,s,fffff0"
	'sx = "&chf=bg,s,fffff0"
	sApi = sApi & sx
	
' Especificar las lineas del grid (chg)
	sx = "&chg=9.09,25,1,5"
	sApi = sApi & sx

' Eje de X y Y (chxl)	
	sApi = sApi & "&chxt=x,y"
	
' Nombre del Eje de las Y (chxl)
	sx="&chxl=0:|"
	for ib= iMinCol to iMaxCol  
		sy=sNomPer(ib, 1)
		sx = sx & sNomPer(ib, 1) & "|"
	next
	sx = sx & "1:|0|500|1000|1500|2000"
	sApi = sApi & sx '"&chxl=0:|Mar|Apr|May|June|July|1:||50+Kb"
	
' Especificar color de las lineas
'	sx="&chco=0000ff,ff0000,FF9900,008000,ff00ff,33ffff,d2691e,00ff00,990066,c0c0c0,1e90ff,000000,ffff00"
	'sApi = sApi & sx
	
' Gradar Data (chd)
	chd = "&chd=t:"
	for ib= iMinCol to iMaxCol 
		ix = int(mDatMue(0,ib+2) /2000*100)
		'response.Write ix
		chd = chd & ix & ","
	next
	ilen = len(chd)
	chd = mid(chd,1,ilen-1)
	sApi = sApi & chd

' Titulo del Gráfico (chtt)(Espacio con + y salto de linea con | )	
	sx = "&chtt=Total+Venezuela|Muestra"
	sApi = sApi & sx
	sx= "&chts=000000,18"
	sApi = sApi & sx
		
'	sGra = sApi & cht & chd & chs & chxl & chxt & chtt & chf &chg
'	sGra = "http://chart.apis.google.com/chart?chs=200x125&chd=s:helloWorld&cht=lc&chxt=x,y&chxl=0:|Mar|Apr|May|June|July|1:||50+Kb"
'	response.Write "<br>" & sGra & "<br>" %>


	<table border=0 width=99% ID="Table6">
		<tr>
			<td width=20%></td>
			<td width=80%><img  src=<%=sApi%> > </td>
		</tr>
	</table>
	<br>

<%End Sub 



Sub MosGraInd ( iCua)

' Pagina de Google Chart
	sParGra ="http://chart.apis.google.com/chart"

' Tipo de Gráfico (cht)
	sx	=  "?cht=lc"
	'sx	=  "?cht=bvg"
	sParGra = sParGra & sx

' Tamaño del gráfico (chs)
	sx = "&chs=600x300"	
	sParGra = sParGra & sx
	
' Especificar el Background (chf)
	sx = "&chf=c,lg,0,76A4FB,1,ffffff,0|bg,s,fffff0"
	sx = "&chf=bg,s,fffff0"
	sx = "&chf=bg,s,fffff0"
	sParGra = sParGra & sx
	
' Especificar las lineas del grid (chg)
	sx = "&chg=9.09,25,1,5"
	sParGra = sParGra & sx

' Titulo del Gráfico (chtt)(Espacio con + y salto de linea con | )
	sx =sInd(iCua) 
	ilen=len(sx)
	'response.Write sx & "  len:=" & ilen & "<br>"
	for i=1 to ilen
		sy = mid(sx,i,1)
		if sy=" " then 
			sx = mid(sx,1,i-1) & "+" & mid (sx,i+1)
		end if	
		if sy="ó" then 
			sx = mid(sx,1,i-1) & "o" & mid (sx,i+1)
		end if	
		if sy="í" then 
			sx = mid(sx,1,i-1) & "i" & mid (sx,i+1)
		end if	
	next
	'response.Write sx & "<br>"
	sx = "&chtt=" &  sx
	sx = sx & "&chts=000000,18"
	sParGra = sParGra & sx

' Especificar la leyenda del Gráfico (chdl)
	sx="&chdl="
	for i=1 to iLin
	'    if iVisMar = 1 then
			if mDat(i,1,0) = 0 then  sy=sMar(mDat(i,0,0)) else sy=sNomPro(mDat(i,1,0))
	'	else
	'		sy =sInd(i)	
	'	end if	
		'response.Write sy & "<br>"
		ilen=len(sy)
		for j=1 to ilen
			sz = mid(sy,j,1)
			if sz=" " then 
				sy = mid(sy,1,j-1) & "+" & mid (sy,j+1)
			end if	
		next
		if i<> iLin then sx = sx & sy & "|" else sx = sx & sy 
	next
'	response.Write sx & "<br>"
	sParGra = sParGra & sx
	
' Especificar color de las lineas
	sx="&chco=0000ff,ff0000,FF9900,008000,ff00ff,33ffff,d2691e,00ff00,990066,c0c0c0,1e90ff,000000,ffff00"
	sParGra = sParGra & sx
	
' Calcular valor máximo del gráfico
	g_iMaxVal = 0
	For ia=1 to iLin
		for ib= iMinCol to iMaxCol
			if mDat(ia,ib+2,iCua)> g_iMaxVal then g_iMaxVal = mDat(ia,ib+2,iCua)
		next
	next	
	
	'response.Write "<br>Valor Máximo:=" & g_iMaxVal
	if g_iMaxVal>1200 then
		ix=g_iMaxVal/500
		ix=int(ix)
		ix=ix+1
		g_iMaxVal= ix *500
	end if
	if g_iMaxVal>1000 and g_iMaxVal<1200 then g_iMaxVal=1200
	if g_iMaxVal>800 and g_iMaxVal<1000 then g_iMaxVal=1000
	if g_iMaxVal>75 and g_iMaxVal<100 then g_iMaxVal=100
	if g_iMaxVal>50 and g_iMaxVal<75 then g_iMaxVal=75
	if g_iMaxVal>25 and g_iMaxVal<50 then g_iMaxVal=50
	if g_iMaxVal>20 and g_iMaxVal<25 then g_iMaxVal=25
	if g_iMaxVal>5 and g_iMaxVal<10 then g_iMaxVal=10
	if g_iMaxVal>3 and g_iMaxVal<5 then g_iMaxVal=5
	if g_iMaxVal>1 and g_iMaxVal<3 then g_iMaxVal=3
	if g_iMaxVal<1 then g_iMaxVal=1
	'response.Write "<br>Valor Máximo:=" & g_iMaxVal
	

' Especificar la data del Grafico (chd)
	sx = "&chd=t:"
	For ia=1 to iLin
		for ib= iMinCol to iMaxCol 
			ix = mDat(ia,ib+2,iCua)
			ix = ix/g_iMaxVal*100
			'ix = ix
			ix = int(ix)
			if ib<> ImaxCol then sx = sx & ix & "," else sx = sx & ix & "|"
		next
	next
	ilen = len(sx)
	sx = mid(sx,1,ilen-1)
	'response.Write sx
	sParGra = sParGra & sx
	
' Eje de X y Y (chxl)	
	sx = "&chxt=x,y"
	sParGra = sParGra & sx
	
' Nombre del Eje de las Y (chxl)
	sx="&chxl=0:|"
	for ib= iMinCol to iMaxCol  
		sy=sNomPer(ib, 1)
		sx = sx & sNomPer(ib, 1) & "|"
	next
' Nombre del Eje de las X	
	'sx = sx & "1:|0|500|1000|1500"
	sx = sx & "1:|0"
	g_iMaxVal =int(g_iMaxVal)
	ix= int(g_iMaxVal)/4
	sx = sx & "|" & ix
	iy =ix + ix
	sx = sx & "|" & iy
	iy = iy +ix
	sx = sx & "|" & iy
	sx = sx & "|" & g_iMaxVal
	'sx = sx & "|" & ix
	'response.Write sx	
	sParGra = sParGra & sx
'
'	sx="&chf=c,ls,0,ffff00,0.2,FFFFF0,0.1"
	sParGra = sParGra & sx
	
	

%>
	<br>
	<table border=0 width=99%>
		<tr>
			<td width=20%></td>
			<td width=80% ><img  src=<%=sParGra%> ></td>
		</tr>
	</table>
	<br>


<%
end Sub

Sub MosGraMar (iTipGra,  iCua, idInd, ihInd )

' Pagina de Google Chart
	sParGra ="http://chart.apis.google.com/chart"

' Tipo de Gráfico (cht)
	if iTipGra=1 then sx	=  "?cht=lc" else sx	=  "?cht=bvs"
	sParGra = sParGra & sx

	
' Tamaño del gráfico (chs)
	sx = "&chs=600x300"	
	sx = "&chs=700x300"	
	sParGra = sParGra & sx

' Ancho de la barra
	if iTipGra=3 then
		sx= "&chbh=36,7"
		sParGra = sParGra & sx
	end if	
	
' Especificar el Background (chf)
	sx = "&chf=c,lg,0,76A4FB,1,ffffff,0|bg,s,fffff0"
	sx = "&chf=bg,s,fffff0"
	sx = "&chf=bg,s,fffff0"
	sParGra = sParGra & sx
	
' Especificar las lineas del grid (chg)
	sx = "&chg=9.09,25,1,5"
	sParGra = sParGra & sx

' Titulo del Gráfico (chtt)(Espacio con + y salto de linea con | )
	if ivPro = 0 then
		sx =sMar(mDat(iCua,0,0))
	else
		if mDat(iCua,1,0) <> 0 then
			sx = sNomPro(mDat(iCua,1,0))
		else
			sx =sMar(mDat(iCua,0,0))
		end if	
	end if	
	'response.Write sx
	ilen=len(sx)
	'response.Write sx & "  len:=" & ilen & "<br>"
	for i=1 to ilen
		sy = mid(sx,i,1)
		if sy=" " then 
			sx = mid(sx,1,i-1) & "+" & mid (sx,i+1)
		end if	
		if sy="ó" then 
			sx = mid(sx,1,i-1) & "o" & mid (sx,i+1)
		end if	
		if sy="í" then 
			sx = mid(sx,1,i-1) & "i" & mid (sx,i+1)
		end if	
	next
	'response.Write sx & "<br>"
	sx = "&chtt=" &  sx
	sx = sx & "&chts=000000,18"
	sParGra = sParGra & sx

' Especificar la leyenda del Gráfico (chdl)
	sx="&chdl="
	for i=idInd to ihInd
	   ' if iVisMar = 1 then
		'	if mDat(i,1,0) = 0 then  sy=sMar(mDat(i,0,0)) else sy=sNomPro(mDat(i,1,0))
	'	else
			sy =sInd(i)	
		'end if	
		'response.Write sy & "<br>"
		ilen=len(sy)
		for j=1 to ilen
			sz = mid(sy,j,1)
			if sz=" " then 
				sy = mid(sy,1,j-1) & "+" & mid (sy,j+1)
			end if	
			if sz="ó" then 
				sy = mid(sy,1,j-1) & "o" & mid (sy,j+1)
			end if	
			if sz="í" then 
				sy = mid(sy,1,j-1) & "i" & mid (sy,j+1)
			end if	
		next
		if i<> ihInd then sx = sx & sy & "|" else sx = sx & sy 
	next
'	response.Write sx & "<br>"
	sParGra = sParGra & sx
	
' Especificar color de las lineas
	sx="&chco=0000ff,ff0000,FF9900,008000,ff00ff,33ffff,d2691e,00ff00,990066,c0c0c0,1e90ff,000000,ffff00"
	sParGra = sParGra & sx
	
' Calcular valor máximo del gráfico
	g_iMaxVal = 0
	For ia=idInd to ihInd
		for ib= iMinCol to iMaxCol
			if mDat(iCua,ib+2,ia)> g_iMaxVal then g_iMaxVal = mDat(iCua, ib+2,ia)
		next
	next	
	if g_iMaxVal>1200 then
		ix=g_iMaxVal/500
		ix=int(ix)
		ix=ix+1
		g_iMaxVal= ix *500
	end if
	if g_iMaxVal>1000 and g_iMaxVal<1200 then g_iMaxVal=1200
	if g_iMaxVal>800 and g_iMaxVal<1000 then g_iMaxVal=1000
	if g_iMaxVal>75 and g_iMaxVal<100 then g_iMaxVal=100
	if g_iMaxVal>50 and g_iMaxVal<75 then g_iMaxVal=75
	if g_iMaxVal>25 and g_iMaxVal<50 then g_iMaxVal=50
	if g_iMaxVal>20 and g_iMaxVal<25 then g_iMaxVal=25
	if g_iMaxVal>5 and g_iMaxVal<10 then g_iMaxVal=10
	if g_iMaxVal>3 and g_iMaxVal<5 then g_iMaxVal=5
	if g_iMaxVal>1 and g_iMaxVal<3 then g_iMaxVal=3
	if g_iMaxVal<1 then g_iMaxVal=1
	if iTipGra=3 then g_iMaxVal = 100
'response.Write "<br>Valor Máximo:=" & g_iMaxVal
	

' Especificar la data del Grafico (chd)
	sx = "&chd=t:"
	For ia=idInd to ihInd
		for ib= iMinCol to iMaxCol 
			ix = mDat(iCua,ib+2, ia)
			ix = ix/g_iMaxVal*100
			'ix = ix
			ix = int(ix)
			if ib<> ImaxCol then sx = sx & ix & "," else sx = sx & ix & "|"
		next
	next
	ilen = len(sx)
	sx = mid(sx,1,ilen-1)
	'response.Write sx
	sParGra = sParGra & sx
	
' Eje de X y Y (chxl)	
	sx = "&chxt=x,y"
	sParGra = sParGra & sx
	
' Nombre del Eje de las Y (chxl)
	sx="&chxl=0:|"
	for ib= iMinCol to iMaxCol  
		sy=sNomPer(ib, 1)
		sx = sx & sNomPer(ib, 1) & "|"
	next
' Nombre del Eje de las X	
	'sx = sx & "1:|0|500|1000|1500"
	sx = sx & "1:|0"
	g_iMaxVal =int(g_iMaxVal)
	ix= int(g_iMaxVal)/4
	sx = sx & "|" & ix
	iy =ix + ix
	sx = sx & "|" & iy
	iy = iy +ix
	sx = sx & "|" & iy
	sx = sx & "|" & g_iMaxVal
	'sx = sx & "|" & ix
	'response.Write sx	
	sParGra = sParGra & sx
'
'	sx="&chf=c,ls,0,ffff00,0.2,FFFFF0,0.1"
	sParGra = sParGra & sx
	
	

%>
	<br>
	<table border=0 width=99% ID="Table2">
		<tr>
			<td width=20%></td>
			<td width=80% ><img  src=<%=sParGra%> ></td>
		</tr>
	</table>
	<br>


<%
end Sub

Sub MosGraMue (iTip, iDatGra(), iIniDat, iNumDat)
	Dim iFinDat
	iFindat = iIniDat +iNumDat -1


	

' Pagina de Google Chart
	sParGra ="http://chart.apis.google.com/chart"

' Tipo de Gráfico (cht)
	if iTip = 3 then sx	=  "?cht=lc" else sx	=  "?cht=bvs"
	sParGra = sParGra & sx
	
' Ancho de la barra
	sx= "&chbh=36,7"
	sParGra = sParGra & sx

' Tamaño del gráfico (chs)
	sx = "&chs=500x500"	
	sx = "&chs=700x300"	
	sParGra = sParGra & sx
	
' Especificar el Background (chf)
	sx = "&chf=c,lg,0,76A4FB,1,ffffff,0|bg,s,fffff0"
	sx = "&chf=bg,s,fffff0"
	sx = "&chf=bg,s,fffff0"
	sParGra = sParGra & sx
	
' Especificar las lineas del grid (chg)
	sx = "&chg=9.09,25,1,5"
	sParGra = sParGra & sx

' Titulo del Gráfico (chtt)(Espacio con + y salto de linea con | )
	sx =iDatGra(0,0)
	ilen=len(sx)
	'response.Write sx & "  len:=" & ilen & "<br>"
	for i=1 to ilen
		sy = mid(sx,i,1)
		if sy=" " then 
			sx = mid(sx,1,i-1) & "+" & mid (sx,i+1)
		end if	
		if sy="ó" then 
			sx = mid(sx,1,i-1) & "o" & mid (sx,i+1)
		end if	
		if sy="í" then 
			sx = mid(sx,1,i-1) & "i" & mid (sx,i+1)
		end if	
	next
	'response.Write sx & "<br>"

	sx = "&chtt=" &  sx
' Especificar Color y Font del Titulo		
	sx = sx & "&chts=000000,18"
	sParGra = sParGra & sx

' Especificar la leyenda del Gráfico (chdl)
	sx="&chdl="
	for i=iIniDat to iFinDat
		sy=iDatGra(i,0)
		'response.Write sy & "<br>"
		ilen=len(sy)
		for j=1 to ilen
			sz = mid(sy,j,1)
			if sz=" " then 
				sy = mid(sy,1,j-1) & "+" & mid (sy,j+1)
			end if
			if sz="ó" then 
				sy = mid(sy,1,j-1) & "o" & mid (sy,j+1)
			end if		
			if sz="í" then 
				sy = mid(sy,1,j-1) & "i" & mid (sy,j+1)
			end if		
		next
		if i<> iFinDat  then sx = sx & sy & "|" else sx = sx & sy 
	next
'	response.Write sx & "<br>"
	sParGra = sParGra & sx
	
' Especificar color de las lineas
	sx="&chco=0000ff,ff0000,FF9900,008000,ff00ff,33ffff,d2691e,00ff00,990066,c0c0c0,1e90ff,000000,ffff00"
	sParGra = sParGra & sx
	
' Calcular valor máximo del gráfico
	g_iMaxVal = 0
	for ia=iIniDat to iFinDat
		for ib= 1 to 12
			if iDatGra(ia,ib)> g_iMaxVal then g_iMaxVal = iDatGra(ia,ib)
		next
	next
	'response.Write "<br>395 g_iMaxVal " & g_iMaxVal	
	if g_iMaxVal>1200 then
		ix=g_iMaxVal/500
		ix=int(ix)
		ix=ix+1
		g_iMaxVal= ix *500
	end if
	if iTip = 3 then
		if g_iMaxVal>1000 and g_iMaxVal<1200 then g_iMaxVal=1200
		if g_iMaxVal>800 and g_iMaxVal<1000 then g_iMaxVal=1000
		if g_iMaxVal>75 and g_iMaxVal<100 then g_iMaxVal=100
		if g_iMaxVal>50 and g_iMaxVal<75 then g_iMaxVal=75
		if g_iMaxVal>25 and g_iMaxVal<50 then g_iMaxVal=50
		if g_iMaxVal>20 and g_iMaxVal<25 then g_iMaxVal=25
		if g_iMaxVal>5 and g_iMaxVal<10 then g_iMaxVal=10
		if g_iMaxVal<5 then g_iMaxVal=5
	else	
		g_iMaxVal = 100
	end if	
	

' Especificar la data del Grafico (chd)
	sx = "&chd=t:"
	for ia=iIniDat to iFinDat
		for ib= 1 to 12 
			ix = iDatGra(ia,ib)
			ix = ix/g_iMaxVal*100
			'ix = ix
			ix = int(ix)
			if ib<> iMaxCol then sx = sx & ix & "," else sx = sx & ix & "|"
		next
	next
	ilen = len(sx)
	sx = mid(sx,1,ilen-1)
	sParGra = sParGra & sx
	
' Eje de X y Y (chxl)	
	sx = "&chxt=x,y"
	sParGra = sParGra & sx
	
' Nombre del Eje de las X (chxl)
	sx="&chxl=0:|"
	for ib= 1 to 12  
		sy=iDatGra(0, ib)
		'response.Write ib & "  - " & sy &"<br>"
		sx = sx & sy & "|"
	next
	
' Nombre del Eje de las Y	
	sx = sx & "1:|0"
	g_iMaxVal =int(g_iMaxVal)
	ix= int(g_iMaxVal)/4
	sx = sx & "|" & ix
	iy =ix + ix
	sx = sx & "|" & iy
	iy = iy +ix
	sx = sx & "|" & iy
	sx = sx & "|" & g_iMaxVal
	'sx = sx & "|" & ix
	'response.Write sx	
	sParGra = sParGra & sx

' Sombra en el Gráfico
'	sx="&chf=c,ls,0,ffff00,0.2,FFFFF0,0.1"
'	sParGra = sParGra & sx

		
	
	

%>
	<br/>
	<table border="0" width="99%" id="Table1">
		<tr >
			<td width="15%"></td>
			<td align="center"><img  src="<%=sParGra%>" /></td>
		</tr>
	</table>
	


<%
Exit Sub
' Data del g´rafico	
	response.Write "<table>"	
	for i=0 to 6
		response.Write "<tr>"	
		response.Write "<td>" & i & "</td>"
		for j=0 to 12
			if i<>0 and j<> 0 then
				response.Write "<td>" & FormatNumber(iDatGra(i,j), 0,-1, 0, -1) & "</td>"
			else	
				response.Write "<td width =8% >" & iDatGra(i,j) & "</td>"
			end if	
		next
		response.Write "</tr>"	
	next	
	response.Write "</table>"
end Sub%>
