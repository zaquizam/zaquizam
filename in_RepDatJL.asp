<%
Dim rr_sColFon

' Actualizado 9/8/2016
' Se Modificó la variable a aplicar división , se colocó el el divisor en 7 y el numerador en el 6
' Agregó rr_tVar para número de variables del titulo
 %> 

 <STYLE type="text/css">


.rr_but			{font-family:arial,helvetica;font-weight:bold;width:100;height:100%;font-size: 7pt;text-align: center;background-color:#eeeeee;color: #000000;text-decoration: none;}
.rr_but a			{font-weight:bold;width:100%;height:100%;color: #000000;text-decoration: none;}
.rr_but a:hover	{font-size: 9pt;color: #800000;text-decoration: none;}

.rr_dim1			{font-family:Verdana;font-weight:normal;height:20;font-size:8pt;text-align: center;  vertical-align: middle;  color: #000000; text-decoration: none;}
.rr_dim1 a		    {font-weight:bold;width:100%;height:60;color: #000000;text-decoration: none;}
.rr_dim1 a:hover    {font-size: 9pt;color: #800000;text-decoration: none;}

.rr_dim4			{font-family:arial,helvetica;font-weight:normal;	font-size: 8pt;text-align: center;	 height:25;background-color:#ffffff;	color: #800000;text-decoration: none;}
.rr_dim4 a		    {																									color: #800000;text-decoration: none;}
.rr_dim4 a:hover	{								font-weight:bold;				font-size: 9pt;												background-color:#800000;	color: #ffffff;text-decoration: none;}

.rr_per			{font-family:verdana;font-weight:normal;	font-size: 10px;text-align: center;	 height:25;background-color:#ffffff;	color: #800000;text-decoration: none;}
.rr_tit			{font-family:Tahoma;font-weight:bold;	font-size: 10;text-align: center;background-color:#ffffff;	color: #000000;text-decoration: none;}

.rr_Var			    {font-family:Verdana;font-weight:normal;height:20;font-size:8pt;text-align: left;color: #000000;text-decoration: none;}

.rr_dat 			{font-family:Verdana;font-weight:normal;height:20;font-size: 8pt;text-align: center;color: #000000;text-decoration: none;text-indent:8}

.rr_cbo			{font-family:arial,helvetica;font-weight:normal;width:100%;font-size: 10pt;text-align: left;background-color:#ffffff;color: #000000;text-decoration: none; }
.rr_cbo a			{font-weight:bold;width:100%;height:100%;color: #000000;text-decoration: none;}
.rr_cbo a:hover	{font-size: 9pt;color: #800000;text-decoration: none;}
.rr_cbo2			{font-family:arial,helvetica;font-weight:normal;width:100%;font-size: 10pt;text-align: left;background-color:#ffffff;color: #000000;text-decoration: none; }
.rr_cbo2 a			{font-weight:bold;width:100%;height:100%;color: #000000;text-decoration: none;}
.rr_cbo2 a:hover	{font-size: 9pt;color: #800000;text-decoration: none;}

.rr_but4			{font-family:Verdana;font-weight:normal;width:100%;height:20;font-size: 8pt;text-align: center;background-color:#eeeccc;color: #000000;text-decoration: none;}
.rr_div			{text-align: center;background-color:<%=rr_sColFon%>; width:950;color: #000000;text-decoration: none;}
.rr_men1			{font-family:Tahoma;font-weight:normal;width:950;font-size:18pt;text-align: center;  vertical-align: middle;  color: #666666; text-decoration: none;}
.rr_hr			{background-color:<%=rr_sColFon%>; width:950;color: #800000;text-decoration: none;  height:1; }
.rr_but2    {display:block;font-weight:bold;color:#FFFFFF;background-color:#98bf21;width:120px;text-align:center;padding:4px;text-decoration:none;}
.rr_but2 a {display:block;font-weight:bold;color:#FFFFFF;background-color:#98bf21;width:120px;text-align:center;padding:4px;text-decoration:none;}
.rr_but2 a:hover{background-color:#7A991A;}



 	
 
    
</style>   
	

<!--#include File="in_RepGra.asp"-->

<%

'==========================================================================================
' Variables y Constantes
'==========================================================================================
	Dim rr_SqlNew
	Dim rr_SqlGru
	Dim rr_SqlFot ' Sql de las fotos
	
	Dim rr_SqlWhe
	Dim rr_rs1
	Dim rr_iLof
	Dim rr_nCom
	Dim rr_nFil
	Dim rr_nCol
	dim rr_sFil(5,10)
	Dim rr_sCom(20,10)	
	dim rr_sCol(10)
	Dim rr_sVar(99,10) ' Datos de las Variables	
	dim rr_sDim(30,12) 
	dim rr_iPad(20,5) ' codigo del Combo Padre	
	Dim rr_sP1, rr_sP2, rr_sP3, rr_sP4, rr_sP5, rr_sP6, rr_sP7, rr_sP8, sP9
	
	Dim rr_gData(5000,33,11)
	'Dim rr_gData(5000,13,11)
	Dim rr_gTit(5000,13)
	'Dim rr_gTit(5000,13)
	Dim rr_gPer
	Dim rr_nVar
	Dim rr_mVar ' Numero de VAriables a Mostrar
	Dim rr_tVar ' Numero de VAriables a Mostrar en el titulo

	Dim rr_iMaxreg ' Numero maximo de registros
   dim ipGra

	Dim rr_sPar(20,2)
	

	Dim rr_iUltPer
	Dim rr_iPriPer
	Dim rr_dPer
	dim rr_ipExe
	dim rr_ipGra
	dim rr_ipFot
	dim rr_ipDis
    dim rr_ipPag	
    dim rr_sDdd
    Dim rr_iGru
    Dim rr_iCombo ' 1=solo combo
    Dim rr_iDis ' 1= no hay cambio de diseño
    Dim rr_iwFoto ' 1=si 
    Dim rr_iwGra ' 1 = si
    dim rr_ipPer ' Codigo de Ultimo Periodo
    dim rr_iAnt
    dim rr_iPro
    dim rr_MesDes
    dim rr_MesHas
    dim rr_Prom
    Dim iColMas
    dim ipDes ' Para Desarrollo
    dim rr_sLink ' Link de la fila 


	Const rr_Espacio="&nbsp;"
	
    Set rr_dPer=Server.CreateObject("Scripting.Dictionary")
    
	
	

%>



<%	
Sub rr_LeePar
    
    rr_sPar(01,0)= Request.QueryString("rr_p1")
    rr_sPar(02,0)= Request.QueryString("rr_p2")
    rr_sPar(03,0)= Request.QueryString("rr_p3")
    rr_sPar(04,0)= Request.QueryString("rr_p4")
    rr_sPar(05,0)= Request.QueryString("rr_p5")
    rr_sPar(06,0)= Request.QueryString("rr_p6")
    rr_sPar(07,0)= Request.QueryString("rr_p7")
    rr_sPar(08,0)= Request.QueryString("rr_p8")
    rr_sPar(09,0)= Request.QueryString("rr_p9")
    rr_sPar(10,0)= Request.QueryString("rr_p10")
    rr_sPar(11,0)= Request.QueryString("rr_p11")
    rr_sPar(12,0)= Request.QueryString("rr_p12")
    rr_sPar(13,0)= Request.QueryString("rr_p13")
    rr_sPar(14,0)= Request.QueryString("rr_p14")
    rr_sPar(15,0)= Request.QueryString("rr_p15")
    rr_sPar(16,0)= Request.QueryString("rr_p16")
    rr_sPar(17,0)= Request.QueryString("rr_p17")
    rr_sPar(18,0)= Request.QueryString("rr_p18")
    rr_sPar(19,0)= Request.QueryString("rr_p19")
    rr_sPar(20,0)= Request.QueryString("rr_p20")
 
    
    rr_ipExe=Request.QueryString("rr_exe")
    if rr_ipExe="" then rr_ipExe=0
    rr_ipGra=Request.QueryString("rr_gra")
    if rr_ipGra="" then rr_ipGra=0
    rr_ipFot=Request.QueryString("rr_fot")
    if rr_ipFot="" then rr_ipFot=0
    rr_ipDis=Request.QueryString("rr_dis")
    if rr_ipDis="" then rr_ipDis=0
    rr_sDdd=Request.QueryString("rr_ddd")
    rr_ipPag=Request.QueryString("rr_pag")
    if rr_ipPag="" then rr_ipPag=1
    rr_ipPer=Request.QueryString("rr_ultper")

    ed_iMp=""
	ed_iMp=Request.QueryString("ed_mp")
	ed_iMs=""
	ed_iMs=Request.QueryString("ed_ms")



	for ias=1 to 20
	    do 
	        ixa =instr(rr_sPar(ias,0),"@@@")
	        if ixa<>0 then
	            rr_sPar(ias,0)= Mid(rr_sPar(ias,0),1,ixa-1) & " " & mid(rr_sPar(ias,0),ixa+3)
	        end if 
	   loop until ixa=0       
	   'response.write "<br> i:=" & i & " Par:=" & rr_sPar(ias,0) 
	next 
    rr_MesDes=Request.QueryString("desdemes")
    if rr_MesDes="" then rr_MesDes=rr_iPriPer
    
    rr_MesHas=Request.QueryString("hastames")
    if rr_MesHas="" then rr_MesHas=rr_iUltPer
    if rr_MesHas<rr_Mesdes then rr_MesHas=rr_iUltPer
    
    rr_Prom=Request.QueryString("prom")
    if rr_Prom = "" then rr_Prom = 0

    
end Sub	

Sub rr_CalPar
 
    dim tPar(20)
	sPar = ""
	CalPar
	
	for ias=1 to 20: tPar(ias)=rr_sPar(ias,0):next
	
	
	for ias=1 to 20
	    do 
	       ixa =instr(tPar(ias)," ")
	       if ixa<>0 then
	             tPar(ias)= Mid(tPar(ias),1,ixa-1) & "@@@" & mid(tPar(ias),ixa+1)
	           
	        end if 
	    loop until ixa=0        
	next    

	sPar =  sPar &  "&rr_p1=" & tPar(1)
	sPar =  sPar &  "&rr_p2=" & tPar(2)
	sPar =  sPar &  "&rr_p3=" & tPar(3)
    sPar =  sPar &  "&rr_p4=" & tPar(4)
    sPar =  sPar &  "&rr_p5=" & tPar(5)
    sPar =  sPar &  "&rr_p6=" & tPar(6)
    sPar =  sPar &  "&rr_p7=" & tPar(7)
    sPar =  sPar &  "&rr_p8=" & tPar(8)
    sPar =  sPar &  "&rr_p9=" & tPar(9)
    sPar =  sPar &  "&rr_p10=" & tPar(10)
    sPar =  sPar &  "&rr_p11=" & tPar(11)
    sPar =  sPar &  "&rr_p12=" & tPar(12)
    sPar =  sPar &  "&rr_p13=" & tPar(13)        
    sPar =  sPar &  "&rr_p14=" & tPar(14)
    sPar =  sPar &  "&rr_p15=" & tPar(15)
    sPar =  sPar &  "&rr_p16=" & tPar(16)
    sPar =  sPar &  "&rr_p17=" & tPar(17)        
    sPar =  sPar &  "&rr_p18=" & tPar(18)
    sPar =  sPar &  "&rr_p19=" & tPar(19)
    sPar =  sPar &  "&rr_p20=" & tPar(20)

   
    sPar =  sPar &  "&rr_exe=" & rr_ipExe
    sPar =  sPar &  "&rr_gra=" & rr_ipGra
    sPar =  sPar &  "&rr_fot=" & rr_ipFot
    sPar =  sPar &  "&rr_dis=" & rr_ipDis
    sPar =  sPar &  "&rr_ddd=" & rr_sDdd
    sPar =  sPar &  "&rr_pag=" & rr_ipPag
    sPar =  sPar &  "&rr_ultper=" & rr_ipPer
	sPar = sPar & "&ed_mp=" & ed_iMp
	sPar = sPar & "&ed_ms=" & ed_iMs	
	sPar = sPar & "&desdemes=" & rr_MesDes
	sPar = sPar & "&Hastames=" & rr_MesHas
	sPar = sPar & "&prom=" & rr_Prom
	
	'response.Write "<br>sPro:=" & sPar

end Sub


Sub rr_LeeCombo
    dim gTem
 
' Grabar en sPar1 la posición de la dimension    
    if rr_sDDD<>"" then
        for i=1 to 20
            ix = asc(mid(rr_sDdd,i,1))
            ix=ix-64
            if ix<>0 then
                rr_sPar(ix,1)=i
            else
                rr_sPar(ix,0)=""
            end if    
        next    
    end if    
    
   'response.write "<br> 220 " & rr_sDdd
    for i=1 to rr_nCom
        set rsx = CreateObject("ADODB.Recordset")
        rsx.CursorType = 1
        rsx.LockType = 1
        rsx.CacheSize=5000
              
           
            ip=asc(mid(rr_sDdd,i,1))-64
            if rr_iPad(ip,0)<>0 then
                 sFil=""  
                 for j=1 to rr_iPad(ip,0)  
                    sx=  chr(rr_iPad(ip,j)+64)
                    ix=rr_iPad(ip,j)
'response.write "<br>253 ip:=" &ip & " sx:=" & sx & rr_sPar(ix,1)
                    ipar=rr_sPar(ix,1)
                    if ipar<>0 then
                        if rr_sPar(ipar,0) ="" then rr_sPar(ipar,0)= "Todas"
                        if rr_sPar(ipar,0) <> "Todas" then
                        
                            sx=rr_sCom(ipar,10)
                            iy=instr(sx,".")
                            if iy<>0 then sx=mid(sx,iy+1)
                            sFil = sFil & " ((" & sx & ")='" & rr_sPar(ipar,0) & "') AND"
                        end if    
                    end if 
                next
               
                if sFil<>"" then 
                    ilen=len(sfil)
                    sFil=Mid(sFil,1,ilen-4)           

'response.write "<br><br>245 i:=" & i & " sFil=" & sFil                
                    rsx.filter=sFil    

                end if    
            end if
'response.write "<br><br> 244 " & " i:=" & i & " -" & rr_sCom(i,2)  
            sql = rr_sCom(i,2)

'response.write "<bR><br> 261"  & " i:=" & i & Sql
    	   rsx.Open sql ,conexion
		   
		    
		    if Not(rsx.EOF) then
			    gTem =rsx.GetRows
'response.Write "<br>=====203 i:=" & i &" Ubound:=" & ubound(gTem,2) & " TimeCombo:=" & Time:response.flush			    

			    rr_sCom(i,0)=""
			    s1=""
			    isw=0 
			    if rr_sCom(i,7)=1 then 
			        rr_sCom(i,0) = rr_sCom(i,0)  & "Todas" & chr(9) & "[Todas]" & chr(9)
			        s1="Todas"
                    if rr_sPar(i,0)="Todas" then isw=1			        
			    end if
			       
			    sx="****"   
			    for h=0 to ubound(gTem,2)
			        if gTem(0,h)<> sx then
           	            rr_sCom(i,0) = rr_sCom(i,0)  & gTem(0,h)  & chr(9) & gTem(1,h) & chr(9)
           	            if rr_sPar(i,0)=gTem(0,h) then isw=1
           	            ix=instr(rr_sPar(i,0),gTem(0,h))
           	            if ix<>0 then isw=1
		                if s1="" then s1=gTem(0,h)
		                sx=gtem(0,h)    
		            end if    
			    next
		     end if
		    rsx.close
		    Set rsx = Nothing
'response.Write "<br>=====228 i:=" & i &" s1" & s1 & "  sPar:=" & rr_sPar(i,0) &  " isw:=" & isw
		    'if rr_sPar(i,0)= "" then rr_sPar(i,0)=s1
		    if isw=0 then rr_sPar(i,0)=s1
        'response.Write "<br>=====236 i:=" & i &"  sPar:=" & rr_sPar(i,0) &  " isw:=" & isw
	
	next
End sub
Sub rr_LeeCol
   ' if rr_iCombo=1 then exit sub
    Dim rsx	
			    
' Leer descripcion del periodo
    set rsx = CreateObject("ADODB.Recordset")
    rsx.CursorType = 1
    rsx.LockType = 1
    rsx.CacheSize=5000
    sx = rr_sCol(2)
'response.write "<br> 346 " & sx
	rsx.Open sx ,conexion	
	rr_gPer =rsx.GetRows
   
	 for i=0 to ubound(rr_gPer,2)
	       iPos=iPos+1
	       'if iPos>12 then 
	           %>
	           <!--
    		        <br />
    		        <center>
			        <hr class="rr_hr" />
			        <font face= "Tahoma"  size="2" color="#000000" >
			        Advertencia: Se excedió el numero de columnas permitidas <br />
			        Se mostraran las primeras 12
    			
			        <hr class="rr_hr" />
    		        </center>
    		   -->
	           <% 
	           'rr_nCol=Ipos
	           'exit sub
	       'end if 
	       if isnumeric(rr_gPer(0,i)) then
'response.write "<br>347 Periodo:=" & rr_gPer(0,i)	& "POsi:=" & iPos	   
	            rr_dPer.Add rr_gPer(0,i),iPos
	            rr_gData(0,iPos,0) = rtrim(rr_gPer(1,i))
	       else
	            sx=rtrim(rr_gPer(0,i))
	            if rr_dPer.Exists(sx)=true then 
	                ipos=ipos-1
                else
                    rr_dPer.Add sx,iPos
                    rr_gData(0,iPos,0) = rtrim(rr_gPer(1,i))
                end if
	       end if   
	         
           
	  next
      rr_nCol=Ipos
        rsx.close
        set rsx=nothing
End Sub



 Sub VerSelector (iSel,iUbi )
        %>
        &nbsp;
        <select size="1" name="cboArea" id="Select1" onchange="location.href=this.options[this.selectedIndex].value" class="rr_cbo2">
            <% 
                for i=1 to 30
					if rr_sDim(i,0)  <> 0 then
					
					
						sX=rr_sDdd
						sb = chr(iSel+64)
						ix= instr(1,rr_sDdd,sb)
						
						sb2 = chr(i+64)
	    				ix2= instr(1,rr_sDdd,sb2)
	    				


						if ix2<>0 then
						    if isel<>0  then 
		    				    sX= Mid(sX,1,ix-1) & sb2  & Mid(sX,ix+1) 
			    			    sX= Mid(sX,1,ix2-1) & sb  & Mid(sX,ix2+1)
			    			    
			    			else
			    			    if iUbi=1 then
			    			        sX= Mid(sX,1,ix-1) & sb2  & Mid(sX,ix+1) 
			    			        sX= Mid(sX,1,ix2-1) & sb  & Mid(sX,ix2+1)
			    			        
			    			    else    
			    			        sx=replace (sx,sb2,"@")
			    			        
			    			        sC=mid(sx,1,20)
			    			        sC=replace(sC,"@","")
			    			        ilen=len(sC)
			    			        if ilen<20 then sc=sc&string(20-iLen,"@")
			    			      '  s6=sc
			    			        
			    		            sz=mid(sx, 21,5) 
			    		            sz=replace(sz,"@","")
	   	                            sz=sz&chr(i+64)
	   	                            ix=len(sz)
	   	                            sz=sz&string(5-ix,"@")
	   	                            
	   	                            sx=sC & sz &  mid(sx,26)			    			        
	   	                            
			    			    end if    
			    			end if        
			    		else

			    		end if	
						



						sF= rr_sDdd
						rr_sDdd = sX
					    rr_CalPar	
						rr_sDdd = sF
 
                        isw=0
					    if isel=0 and i=rr_sCol(6) then isw=1
					    if isel=0 and i=rr_sFil(1,6) and rr_nFil=1 then isw=1
					   
					    if isw=0 Then
					        if  rr_sDim(i,3)<>0 then%>
                                <option value="<%=sPar %>" <% if Isel = i then response.Write"selected"%>>
                                    <%=rr_sDim(i,1) %>
                                    
                                </option>
                           <%   end if
                        end if
                    End if
				Next 
' Opción [ninguno]
        if iUbi<>3 and isel=0 then 
        
            sb = chr(iSel+64)
	   	    sx=replace(rr_sDdd,sb,"@")
	   	    
	   	    if iUbi=1  then
	   	        sz=mid(sx, 1,20) 
	   	        sz=replace(sz,"@","")
	   	        iLen=len(sz)
	   	        if ilen<20 then
	   	            sz=sz&string(20-ilen,"@")
	   	            sx=sz & mid(sx,11)
	   	        end if    
	   	    else
	   	        sz=mid(sx, 21,5) 
	   	        sz=replace(sz,"@","")
	   	        iLen=len(sz)
	   	        if ilen<5 then
	   	            sz=sz&string(5-ilen,"@")
	   	            sx=mid(sx,1,5) & sz &  mid(sx,21)
	   	        end if    
	   	        
	   	    end if    
	   	    
	        sf = rr_sDdd 
	        rr_sDdd=sx
	        rr_CalPar
	        rr_sDdd = sF
	        
%>
           <option value="<%=sPar%>" <% if Isel = 0 then response.Write"selected"%>>
                <%="[Ninguno]" %>
            </option>
        
       <%end if%>
       </select>
<%	

End Sub

 
'==========================================================================================
' Mostrar los Combos
'==========================================================================================
Sub rr_VerCombo

   
%>

<table border="0" cellspacing="1" cellpadding="0" width="80%"  align="right"  id="Table3">
    <%
	for iii=1 to rr_nCom 
	
	%>
    <tr>
        <td width="50%" valign="middle" bgcolor="<%=rr_sColFon %>" align="right" height="20">
            
                <%if rr_ipDis= 1 and rr_ipExe=0 then
					
					if rr_sCom(iii,3)<>0 then
					    isel=rr_sCom(iii,6)
					    VerSelector iSel, 1
					else%>
                   <font face="Tahoma" size="2" color="#000000">
                    <b>
                    <%response.Write(rr_sCom(iii,1) & ":"  & rr_Espacio & rr_Espacio & rr_Espacio & rr_Espacio )%>
                    </b>
                     </font><%
					end if    
			    else%>
                
                     <font face="Tahoma" size="2" color="#000000">
                     <b>
                    <%response.Write(rr_sCom(iii,1) & ":"  & rr_Espacio & rr_Espacio & rr_Espacio & rr_Espacio )%>
                    </b>
                     </font>
				<%end if%>
        </td>
        <td width="50%" valign= "middle" bgcolor="#ffffff" align="left" height="20">
        
            
                <%
               
    'if rr_ipExe = 0 and rr_ipdis=0 then
    if rr_ipExe = 0 then
        if rr_ipDis=0 then
        

        %><select size="1" name="cboArea" id="Select4" onchange="location.href=this.options[this.selectedIndex].value" class="rr_cbo">
            
            <%



				sCam=rr_sCom(iii,0)
				isw=0
				sc1=""
				do
					ix= instr(1,sCam,chr(9))
					if ix<>0 then
						sC=mid(sCam,1,ix-1)
						sCam=mid(sCam,ix+1)
						ix = instr(1,sCam,chr(9))
						sT=mid(sCam,1,ix-1)
						sT=replace (sT,"<br>","")
	                    if sc1="" then
			                sc1=sc
			            end if

			          
			            sxp = rr_sPar(iii,0)
			            rr_sPar(iii,0) = Sc
			            ipP=rr_ipPag
			            rr_ipPag=1
		                rr_CalPar 
		                rr_ipPag=ipP
		                rr_sPar(iii,0) = sxP
			            
					    response.write "<option value='" & sPar & "'"
					   ' response.write "<br><br>" &  sPar & "<br><br>"
					    if sc =rr_sPar(iii,0) then 
					        response.write " selected "
					        rr_sPar(iii,2)=st
					        isw=1
					    end if
					    
    		           	  if isnumeric(sc) then
    		           	    if isnumeric(rr_sPar(iii,0)) then 
    		           	        ih=sc-rr_sPar(iii,0)
    		           	        if ih=0 then 
    		           	            response.write " selected "
    		           	            rr_sPar(iii,2)=st
    		           	            isw=1
    		           	        end if    
    		           	    end if
    		           	  end if      		           	      
                        response.Write(">" & st  &  "</option>") 
                      '  response.Write("<br>" &  st ) 
 						sCam=mid(sCam,ix+1)
					end if
				loop until ix =0
				
       			if  isw=0 then
				    
				  ' response.write "<br> De :=" & rr_sPar(iii,0) & " a " & sc1
				    rr_sPar(iii,0)=sc1
				end if 

%>
        </select>
        </td>
    </tr>
   
    <%  end if
        else%>
     
     
	<%		    sCam=rr_sCom(iii,0)
				isw=0
				sc1=""
				do
					ix= instr(1,sCam,chr(9))
					if ix<>0 then
						sC=mid(sCam,1,ix-1)
						sCam=mid(sCam,ix+1)
						ix = instr(1,sCam,chr(9))
						sT=mid(sCam,1,ix-1)
	                    if sc1="" then
			                sc1=sc
			            end if

	
					    if sc =rr_sPar(iii,0) then %>
					    
					    <b><p class="rr_tit"> 
					    <%=st %>
					        
					     <%   isw=1
					        exit do
					    end if
					    
    		           	  if isnumeric(sc) then
    		           	    if isnumeric(rr_sPar(iii,0)) then 
    		           	        ih=sc-rr_sPar(iii,0)
    		           	        if ih=0 then %>
					            <b><p class="rr_tit"> 
					                <%=st %>
					        
					            <%isw=1
					                exit do
    		           	        end if    
    		           	    end if
    		           	  end if      		           	      
                        
 						sCam=mid(sCam,ix+1)
					end if
				loop until ix =0     
     
      %>
     
     <% 'response.Write rr_sPar(iii,0)%></p></b>
     </td>
    </tr>
	<%end if%>
	
	<% 'response.write "<br>636 iii:= " & iii %>
    <% next %>

<%
' Nueva Dimension
    if rr_ipDis= 1 and rr_ipExe=0 then%>
    
        <tr><td>
       <%isel=rr_sCom(iii,6)
					VerSelector 0, 1 %>
        </td>    
        <td bgcolor="#eeeccc"></td>
        </tr>
<%
    end if
 %>    
    <% if rr_ipexe=0 then
'response.write "<br>749 rr_iDis:=" & rr_iDis & " combo:=" & rr_iCombo		   	
            if rr_iCombo<>1 then
'response.write "<br>749 rr_ipDis:=" & rr_ipDis			   
                if rr_iDis<>1 then %>
                <tr ><td colspan="2" width="428px" >
     		    <%if rr_ipDis=0 then 
     		        sx= "Cambiar Diseño" 
     		        ip=rr_ipdis
     		        rr_ipDis=1
     		        rr_CalPar
     		        rr_ipDis=ip
     		    else
         		    sx = "OK"
         		    ip=rr_ipdis
     	    	    rr_ipDis=0
     		        rr_CalPar
     		        rr_ipDis=ip
     		    end if
  
     		%>
                <form action="<%=sPar %>" method='post' id="Form5" style="margin-top:5px">
                    <input type="submit" name="but1" value="<%=sx%>" id="Submit2" class="rr_but4" />
                </td></tr>
                </form>
                
     <% End if    
        End if
     end if%>  
           
</table>

<!-- Listo -->
<%
End Sub









Sub rr_VerTit

      	iBor=0
	    if rr_ipExe=1 then iBor=1
  
          
        %>
        
        <table border="<%=iBor%>" width="98%" cellpadding="0" cellspacing="1" bgcolor="#c0c0c0" align="center" id="Table4">
          
    

            <!--- Primera Fila -->
            <tr  bgcolor="#ffffff">
                <% irp=rr_nFil+1
                iColMas = 0
                if rr_Prom = 1 then iColMas = 1
                if rr_ipDis=1 And rr_ipExe=0 then  irp=irp%>
                <td  height="30"   bgcolor="#ffffff" colspan="<%=irp%>">
                
                </td>
                    <td   colspan="<%=rr_nCol+iColmas%>" class="rr_dim4">
                    <% if rr_ipDis=1 And rr_ipExe=0 then %>
                        <%isel=asc(mid(rr_sDdd,26,1))-64%>
                        <%'VerSelector iSel, 3%>
                         <%=rr_sCol(1) %>
                    <% else %>
                   
                         <%=rr_sCol(1) %>
                    <% end If%> 
                 </td>
            </tr>
            <!--- Segunda Fila -->
            <tr  bgcolor="#ffffff">
                <% For cc=1 to rr_nFil %>
                        <!--- Titulo de  Columnas-->
                        <td width="<%=rr_sFil(cc,4)%>%" height="30"  class="rr_dim4">
                            <% if rr_ipDis=1 and rr_ipExe=0 then 
                                    isel=asc(mid(rr_sDdd,cc+20,1))-64
                                    isel=rr_sFil(cc,6)
                                    VerSelector iSel, 2
                                else %>
                                    <%=rr_sFil(cc,1)%>
                                <% end If%>
                        </td>
                <%next %>
<%
' Nueva Dimension
    if rr_ipDis= 1 and rr_ipExe=0 and rr_nfil<3 then%>
    
        <td width="15%">
       <%isel=rr_sCom(iii,6)
		VerSelector 0, 2 %>
        </td>    
      
        
<%
    end if
 %>                 
<% ' Titulo de las Variables %>                
                <td width="<%=rr_sVar(0,1)%>%" class="rr_dim4" >
                      <%=rr_sVar(0,0)%>
                </td>
                
                <%
                
' Titulo del periodo
                 for ib= 1 to rr_nCol%>
                 <td width="5%" align="center" class="rr_per">
                 <% sx= rr_gData(0,ib,0) 
                    sx= replace(sx," ","<br>")%>
                 <%= sx%>
                </td>
                <%
                next
                if rr_Prom = 1 then
                    %>
                    <td width="5%" align="center" class="rr_per">Promedio</td>
                    <%
                end if
                %>
                
            </tr>
            
            <% 
End Sub

Sub rr_SqlSub

' Construir SELECT y GROUP BY
    sx=""
    sx= rr_sCol(10) & ","



    iCam=0
	for iCC=1 to rr_nFil
        sx= sx & " " & rr_sFil(icc,10)
        icam=icam+1
        if icc<> rr_nFil then sx = sx & ","
	next
	'sx=rr_sqlGru  & sx 
	sqlS = "SELECT " & sx 
    if rr_ipdis=1 Then sqlS = "SELECT TOP 1 " & sx 	
	sqlG = "GROUP BY " & rr_sqlGru  &  sx 
'response.write "<br>857 " & sqls

' Agregar Variables    
    sx=""
    for iCC=1 to rr_nVar
	     sx= sx & rr_sVar(icc,10)
         icam=icam+1
         rr_sVar(icc,8)=icam  
        if icc<> rr_nVar then sx = sx  & ","
    next
'response.write "<br>857 " & sx    
    sqlS = sqls & "," & sx
'response.write "<br>857 " & sqls
' Buscar from    
     sqls = sqls &  " FROM " & rr_sDim(0,10)  
     rr_sqlFot = " SELECT FotoAsignada, T_MesAno.Periodo, Tienda FROM " & rr_sDim(0,10) 
     
 
' Construir el WHERE
	sql = rr_sqlWhe
	sql=" WHERE (" & sql
	for iCC=1 to rr_nCom
	    if rr_sPar(iCC,0)<>"Todas" then 
	        if rr_sPar(iCC,0)<>"" then 
	            select case rr_sCom(icc,4)
	                case 0 
                    sql= sql & " AND ((" & rr_sCom(icc,10) & ")='" & rr_sPar(iCC,0) & "')"
                 case 1
                    sql= sql & " AND ((" & rr_sCom(icc,10) & ") like'%" & rr_sPar(iCC,0) & "%')"
               
                End Select   
            end if    
        end if    
	next
'	ix=len(sql)
	'sql =mid(sql,1, ix-4)
	sqlW = sql & ") "
'response.write "<br><br>867 " & sqls
'response.write "<br><br>868 " & sqlw
'response.write "<br><br>" & sqlg

    rr_sqlNew =SqlS & sqlW &  SQLG
    rr_sqlfot= rr_sqlfot & sqlW & " GROUP By Fotoasignada, T_MesAno.Periodo, tienda, meses ORDER By meses DESC,tienda"


' Order By
    sx=" ORDER BY "	
    for i=1 to rr_nFil
        sx = sx  & rr_sFil(i,10)
        if i<> rr_nFil then sx = sx &  ","
     next
'response.write "<br><br>638" & rr_nFil
'response.write "<br><br>639" & sx
' Construir el Istrucción SQl
    rr_SqlNew = rr_SqlNew & sx    



end sub

'--------------------------------------------------------------------
' Leer registros
'--------------------------------------------------------------------
Sub rr_LeeDat

	Dim gTem
    Dim sSw
    
    rr_SqlSub


	set rr_rs1 = CreateObject("ADODB.Recordset")
	rr_rs1.CursorType = 1
	rr_rs1.LockType = 1
	rr_rs1.CacheSize=1000
    if rr_ipdis=1 then  
       rr_rs1.MaxRecords=1%>
    
     	
     	<table align="center" border="0" bgcolor="<%=rr_sColFon%>">
     	    <tr><td><hr class="rr_hr" /></td></tr>
     	
     	    <tr><td class="rr_men1"> Vista Previa</td></tr>
			<tr><td><hr class="rr_hr" /></td></tr>
         </table>
   <% end if    


if ipDes=1 then
    'response.Write "<br><br>888 " & rr_sqlNew		
    %>
    <br /><br /><br />
    <%
end if
	'response.Write "<br><br>888 " & rr_sqlNew	
	rr_rs1.Open rr_sqlNew,conexion

	if Not(rr_rs1.EOF) then 
		gTem=rr_rs1.GetRows
'response.Write "<br><br>675 registros" & rr_sqlNew		
		rr_Ilof=1

	else 
'response.Write "<br><br>465 " & rr_sqlNew	
		rr_ilof=0
		exit sub
	end if	

'response.write "<br> 520 time:=" & Time:response.flush			
' Guardar en rr_gData
    rr_iMaxReg=0
    sn1=""
    sn2=""
    
 
    ilof1=ubound(gtem,2)

    for ii=0 to iLof1
        ipp=rr_dPer.Item(gTem(0,ii))
       

if ipDes=1 then 
    response.write "<br>54 " 
    for i=0 to rr_rs1.fields.Count-1
            response.write  rr_rs1.Fields(i).name & "("& i&"):=" & gTem(i, ii) & "    "
    next
    response.write " ipoi:=" & Ipp
  
end if  
  'response.write "<br>902 0:=" & gTem(0, ii) & " 1:="  & gTem(1, ii) & " 2:="  & gTem(2, ii)& " ipp:=" & Ipp

'Calculo de Total
            

        if rr_sFil(1,5)=1 and rr_nFil>1 then
            ix=1 
            if sn1<>gTem(ix,ii) then
                rr_iMaxReg=rr_iMaxReg+1
                in1=rr_iMaxReg

                sx = gTem(ix,ii)
                'Quite dos asteriscos 05-06-2015 Luis Reyes
                if isnull(sx) then sx = ""
                rr_gTit(rr_iMaxReg,1)= sx
                rr_gTit(rr_iMaxReg,0)=1
                sn1=gTem(ix,ii)
            end if
        end if
        
        if rr_sFil(2,5)=1  and rr_nFil>2 then
            ix=1 
            sx2=gTem(ix,ii)
            ix2= 2 
            sx2=sx2 & gTem(ix2,ii)
            if sn2<>sx2 then
                rr_iMaxReg=rr_iMaxReg+1
                in2=rr_iMaxReg
                sx = gTem(ix,ii)
                'Quite dos asteriscos 05-06-2015 Luis Reyes
                if isnull(sx) then sx = ""
                rr_gTit(rr_iMaxReg,1)=sx
                rr_gTit(rr_iMaxReg,2)=gTem(ix2,ii)
                rr_gTit(rr_iMaxReg,0)=2
                sn2=sx2
            end if
        end if
        
                
        sx=""    
        for jj=1 to rr_nFil
            ix=jj  
           ' response.write "<br> 502" & rs1("Fabricante")
            sx= sx & gTem(ix,ii)
        next

        if sx<>sSw then    
            rr_iMaxReg=rr_iMaxReg+1
            if rr_iMaxReg>4999 then %>
    		    <br />
    		    <center>
			    <hr class="rr_hr" />
			    <font face= "tahoma"  size="2" color="#000000" >
			    Advertencia: Se excedió el numero de filas permitido <br />
			    Se mostraran las primeras 5000
			    </font>
				<hr class="rr_hr" />
			    </center>
            <%
                exit sub
            end if
'response.write "<br><br><br>1020 sx :=" &sx & " ssw:="& ssw & " rr_iMaxreg:=" & rr_iMaxreg            
       ' Ojo eliminado por JL 2/05/2016     rr_gTit(rr_iMaxReg,0)=rr_nFil
            ssW=sx
            for jj=1 to rr_nFil
                ix=jj 
                sx = gTem(ix,ii)
                'Quite dos asteriscos 05-06-2015 Luis Reyes
                if isnull(sx) then sx = ""
                rr_gTit(rr_iMaxReg,jj)= sx
            next
        end if    

        for vv=1 to  rr_nVar
            ipv= rr_sVar(vv,8) 

            iNum= gTem(ipv,ii)
'response.write "<br> Ipv:=" & ipv 
            if rr_sVar(vv,4)<>6 then
                rr_gData(rr_iMaxReg,iPp,vv) = iNum       
            else
                rr_gData(rr_iMaxReg,iPp,vv) = rr_gData(rr_iMaxReg,iPp,vv) +1
            end if    
            
            
            Select Case rr_sVar(vv,4) 
                case 4 ' Minimo
                 
                    if iNum<>0 then
                        if rr_gData(rr_iMaxReg,iPp,vv)<>0 then
                           if iNum< rr_gData(0,iPp,vv) then rr_gData(0,iPp,vv)= iNum
                        else
                            rr_gData(rr_iMaxReg,iPp,vv)= iNum
                            rr_gData(0,iPp,vv)= iNum
                        end if                
                        if rr_gData(in1,iPp,vv)<>0 then 
                           if iNum< rr_gData(in1,iPp,vv) then rr_gData(in1,iPp,vv)= iNum
                        else
                           rr_gData(in1,iPp,vv)= iNum
                        end if    
                        if rr_gData(in2,iPp,vv)<>0 then 
                            if iNum< rr_gData(in2,iPp,vv) then rr_gData(in2,iPp,vv)= iNum
                        else
                            rr_gData(in2,iPp,vv)= iNum
                        end if
                   
                      '  if rr_imaxreg= 3 then response.write "<br>" & rr_Imaxreg & " " & Ipp & " " &  iNum
                    end if    
                    
                case  5  ' Maximo
                    if iNum> rr_gData(in1,iPp,vv) then rr_gData(in1,iPp,vv)= iNum
                    if iNum> rr_gData(in2,iPp,vv) then rr_gData(in2,iPp,vv)= iNum
                    if iNum> rr_gData(0,iPp,vv) then rr_gData(0,iPp,vv)= iNum
                case 1, 2 ' Suma
                    rr_gData(0,iPp,vv)=  rr_gData(0,iPp,vv)+iNum
                    if in1<>0 then rr_gData(in1,iPp,vv)= rr_gData(in1,iPp,vv) + iNum
                    if in2<>0 then rr_gData(in2,iPp,vv)= rr_gData(in2,iPp,vv) + iNum
                    
               case 3  ' Promedio
             
                  '  if iNum> rr_gData(in1,iPp,vv) then rr_gData(in1,iPp,vv)= iNum
                  '  if iNum> rr_gData(in2,iPp,vv) then rr_gData(in2,iPp,vv)= iNum
                  
                        rr_gData(0,iPp,vv) = rr_gData(0,iPp,vv)+ iNum
                        rr_gData(0,iPp,10) = rr_gData(0,iPp,10) +1
                        rr_gTit(0,0)=1
'response.write "<br>1100 in1:=" & in1                       
                        if in1<>0 then
                            rr_gData(in1,iPp,vv)= rr_gData(in1,iPp,vv)+iNum
                            rr_gData(in1,iPp,10) = rr_gData(in1,iPp,10) +1
                         end if   
                        if in2<>0 then
                            rr_gData(in2,iPp,vv)= rr_gData(in2,iPp,vv)+iNum
                            rr_gData(in2,iPp,10) = rr_gData(in2,iPp,10) +1
                         end if   
              ' response.write "<br>1085 iNum:=" & iNum & " Data:=" & rr_gData(0,iPp,vv) &  " cuenta:=" & rr_gData(0,iPp,10)
                    
                  '  rr_gData(in2,0,vv)=-1
               case 6  ' Cuenta 1
                    rr_gData(0,iPp,vv)=  rr_gData(0,iPp,vv)+1
                    if in1<>0 then rr_gData(in1,iPp,vv)= rr_gData(in1,iPp,vv) + 1
                    if in2<>0 then rr_gData(in2,iPp,vv)= rr_gData(in2,iPp,vv) + 1   
               case else
            end select    
        next
      

  next
	rr_rs1.close
	set rr_rs1=nothing
  
   'response.write "<br> 824 Registros:=" & ubound(gtem,2) & " Lineas:=" & rr_iMAxreg & " time:" & time
End  sub
          



Sub rr_VerDat
   
	ix=0
	isw=0
	sn1="-1"
	ifl=0
    iiR=1
    if rr_sfil(1,5)=1 then iir=0
    if rr_ipDis=1 then iir=0 
	
    For iiReg=iiR to rr_iMaxReg 
        ifl=ifl+1
        if ifl=100 then
            response.flush
            ifl=0
        end if
        iswt=0
        if sn1<>rr_gTit(iiReg,1) then
            response.flush
            if sn1<>"-1" then
            'response.write "<br> 1034 time:=" & Time:response.flush			    
                 %>
                 
                </table><br/><br />
                
                <%
                 if rr_ipDis=1 then exit sub
             end if
			 'response.write "<br>1142 rr_ipGra:="  & rr_ipGra
			 'response.write "<br>1142 rr_ipdis:="  & rr_ipdis
            if rr_ipGra=1 and rr_ipdis=0 then 
                'response.write "<br>1162 Paso"
				sx=rr_sColFon
                if rr_ipExe=1 then sx="#ffffff"
                vGrafico iireg
                %>
                
            <%end if 
            rr_VerTit
            iswt=1
            sn1=rr_gTit(iiReg,1)
        else
            if rr_ipExe=0 then
             sx=rr_sColFon
            
%>

        
            <tr   bgcolor="#ffffff" >        
            <td height="5" colspan="<%=rr_Nfil+rr_ncol+iColMas+1 %>"></td>
            </tr>      

<%          end if
        end if    
        if iswt=1 then
            rr_xVar=rr_tVar
            iswt=0
        else    
            rr_xVar=rr_mVar
        end if    
        isw=0

        For iiVar=1 to rr_xVar   
            isTot=0
            if rr_gTit(iiReg,0)<>0 then isTot=1
		   ' if rr_gTit(iiReg,0)<>0 then
		        ix=ix+1%>
		        
                <tr height="20" bgcolor="#ffffff" onmouseover="this.style.backgroundColor='#eeeeee' "onmouseout="this.style.backgroundColor='#ffffff'">
                <!--tr height="20" bgcolor="#ffffff" -->
                    <%
' Escribir Titulo del registro
                    if isw=0 then
                    if rr_gTit(iiReg,0)=12 then%>
                       <td  class="rr_dim1" rowspan="<%=rr_xVar    %>"  colspan="1" valign="middle"  align="center" >
                            <%=rr_gTit(iiReg,1 )  %>
                       </td>
                       <td  class="rr_dim1" rowspan="<%=rr_xVar    %>"  colspan="2" valign="middle"  align="center" >
                            <%=rr_gTit(iiReg,2 )  %>
                       </td>
                     
                        <%isw=1
                     else
                          for ixz= 1 to rr_nFil
                                if iiReg=0 and ixz=1 then
                                    if  rr_sDim(0,1)<>"" then
                                        rr_gTit(iiReg,ixz )= rr_sDim(0,1)
                                    else    
                                        rr_gTit(iiReg,ixz )= " Total " & rr_sFil(ixz,1) & ""
                                    end if    
                                  ' response.write "<br>1192 " & rr_gTit(iiReg,ixz ) & " " & rr_nFil
                                end if    
                                 %>
                                <td  onmouseover="this.style.backgroundColor='#eeeeee' "onmouseout="this.style.backgroundColor='#ffffff'" class="rr_dim1" rowspan="<%=rr_xVar    %>" valign="middle"  align="center" >
                                    <%  'sp=Spro
                                       ' sPro="ms_rReporte5p.asp"
                                       ' rr_calpar
                                        'spro=sp
                                     %>
                                  <p style="text-align: inherit" > <%=  rr_gTit(iiReg,ixz )%> </p>
                                </td>
                         <%next
                        
                        isw=1
                    end if
                     end if
' Escribir Texto de Variable                                        
                     %>  
                     <% if rr_ipDis=1 And rr_ipExe=0 and rr_nfil<3 then %>
                     <td bgcolor="#eeeccc"></td>
                     <% end if %>
                     
                     
                     <td class="rr_Var"   valign="middle"  style="padding:0 0 0 5px"   > 
                           <%=rr_sVar(iiVar,0 )%> 
                     </td>     
               	<%
 ' Escribir la data 
 		   
 		            iPro=0
 		            iFre=0
        			for iiCol=1 to rr_nCol
        			
				%>
                <td class="rr_dat"  >
                    <%
                         
							if rr_gData(iiReg,iiCol, iiVar) then
								if rr_gData(iiReg,iiCol, iiVar)=0  then 
									response.Write "-"
									'response.Write "No Disponible"
								else
								    ' Share
								    if rr_sVar(iiVar,5)=1 then%><b><%
								        iV=rr_gData(iiReg,iiCol, iiVar)/rr_gData(0,iiCol, iiVar)*100
								        iV= iV *  rr_sVar(iiVar,1 )
								    elseif rr_sVar(iiVar,5)=2 then
								        ix=rr_gData(iiReg,iiCol, rr_sVar(iiVar,7))
								        'response.write "<br>1218 " & ix & " Var7:=" & rr_sVar(iiVar,7)& " Var6:=" & rr_sVar(iiVar,6)& "<br>"
								        if ix<>0 then
									        iV=rr_gData(iiReg,iiCol, rr_sVar(iiVar,6))/ix
    'response.write "<br>1221 " & " Var6:=" & rr_gData(iiReg,iiCol, rr_sVar(iiVar,6))& " Var7:=" & ix & "<br>"									        
									    else    
									        iv=0
									    end if        
								        iV= iV *  rr_sVar(iiVar,1 )
								    else
								        ' Promedio
    								    if rr_sVar(iiVar,4)= 3 then

    								        if istot<>0 then 
    								     'response.write "<br>1257 =" & rr_gData(iireg,iiCol, 10) & "="     								        
    								            rr_gData(iiReg,iiCol, iiVar)=rr_gData(iiReg,iiCol, iiVar)/rr_gData(iireg,iiCol, 10)
    								        
    								        end if
									    end if
									    iV= rr_gData(iiReg,iiCol, iiVar) *  rr_sVar(iiVar,1 ) 
									   ' response.write "<br>1175 iv:=" & iv
									end if 
									
									ixDec=rr_sVar(iiVar,2)
									if iv>99.98 and ixdec>1 then ixdec=1
								    if iv then
								        iPro=iPro +Iv
								        IFre=iFre+1
								        response.Write FormatNumber(iV,ixDec, -1, 0, -1)  
									  
								    else
								        response.Write "-"
                                    end if									    
								end if	
							else
							    response.Write " -" 
							end if
							if rr_ipExe = 0 then response.Write "&nbsp;"	
                    %>
                </td>
                
                <%
                
                next
                if iColMas = 1 then
                    %>
                    <td class="rr_dat" align="right" >
                    <%
                    if iPro<>0 then 
                        iv=iPro/iFre
                    else
                        iv=0
                    end if
                    if iv then
                        response.Write FormatNumber(iV,ixDec, -1, 0, -1)
                    else
				        response.Write "-"
                    end if                        
                    if rr_ipExe = 0 then response.Write "&nbsp;"
                    %>
                    </td>
                    <%
                 end if
                 %>
                 
         </tr>
            <%	
			Next

			%>

                   <%
            
            response.flush
                   
    next%>

</table>            

            <%
                        
end Sub

Sub rr_VerVis
    if rr_ipExe=1 then exit sub

    sx=rr_sColFon
            if rr_ipExe=1 then sx="#ffffff"
            iw=200
            if rr_iwFoto=1 then iw=300      
        %>
        <table border="0" width="<%=Iw%>" bgcolor="<%=sx %>" cellspacing="0" align="right" style="margin-top:5px">
            <% 
                'if rr_iCombo<>1 then
                if rr_iwGra=1 then
                ig = rr_ipGra
				if rr_ipGra=0 then
					sx ="Ver Gráficos" 
					rr_ipGra= 1 
				else 
					sx ="Solo Data" 
					rr_ipGra=0
				end if	
				rr_CalPar
				rr_ipGra=ig
            %>
            
            <tr>
            
                <td width="33%">
                <form action="<%=sPar %>" method="post"  id="Form4">
                    <input type="submit" name="but1" value="<%=sx%>" id="Submit4" class="rr_but"/>
                      
                </td>
            </form>  
            <%else %>
                <td width="33%"></td>
            <%end if	

' Ver Fotos
            if rr_iwFoto=1 then
                ig = rr_ipFot
				if rr_ipFot=0 then
					sx ="Ver Imagen" 
					rr_ipFot= 1 
				else 
					sx ="Ocultar Imagen" 
					rr_ipFot=0
				end if	
				rr_CalPar
				rr_ipFot=ig
            %>
            
            
                <td width="33%">
                    
                        <form action="<%=sPar %>" method="post"  id="Form6">
                        <input type="submit" name="but1" value="<%=sx%>" id="Submit6" class="rr_but"/>
                    
                </td>
                </form>
 
            <%end if	
            
            
					iE = rr_ipExe
					iD=ipDis
					iSW=iswVar
					ipDis=0
					rr_ipExe=1
					rr_CalPar
					ipDis=iD
					iswVar=isw
					rr_ipExe=IE %>
            
            <td width="33%">
            <form action="<%=sPar %>" method='POST' id="Form7">
                <input type="submit" name="but1" value="Exportar a Excel" id="Submit7" class="rr_but"/>
            </td>
            </form>
            </tr>
            
        </table>
        <%
end Sub



 Sub rr_VerPie

        
        '<!--#include Virtual="/rtmaster/lib/inPie.asp"-->%>
        <br />
        <font  face="Verdana" color="#dddddd" size="1">
			<% response.write INIvUsu 
			    response.write "<br>" & iper
			    response.write "<br>" & idusu
			 Response.Write "<br>" & Session.SessionID 
			 Response.Write("<br>Ac" & Application("vactivos") )%>
			 </font>
<%
End Sub 

Sub rr_LeeDdd
' Dimensiones por defecto
'response.write "<br> sdd:=" & rr_sDdd
    if rr_sDdd="" then 
        for i=1 to 30
            if rr_sDim(i,0)=1 then
               s1= s1 & chr(i+64)
             '  response.write "<br>i:=" & i & " d=" & rr_sDim(i,0)  & S1
            end if 
            if rr_sDim(i,0)=2 then
               s2= s2 & chr(i+64)
            end if   
            if rr_sDim(i,0)=3 then
               s3= s3 & chr(i+64)
            end if               
        next
         
        ix=len(s1) 
        ix=20-ix
        if ix<>0 then s1 = s1 & string(ix,"@")
        ix=len(s2) 
        ix=5-ix
        if ix<>0 then s2 = s2 & string(ix,"@")
        ix=len(s3) 
        ix=5-ix
        if ix<>0 then s3 = s3 & string(ix,"@")
         rr_sDdd=s1&s2&s3

    end if 
'response.write "<br> sdd:=" & rr_sDdd
' Chequear si existen todas las dimensiones    
    for i=1 to 30
        if rr_sDim(i,0)<>0 then
            sx=chr(i+64)
            ix=instr(rr_sDdd,sx)
            if ix=0 then
                sd=rr_sDdd
                iU=instr(rr_sDdd,"@")
                sd=mid(sd,1,iu-1) & sx & mid(sd,iu+1)
            'response.write "<br> i:=" & i & " - " & rr_sDim(i,0) & " " & sx & "  " & sd
                rr_sDdd=sd
            end if
        end if
    next


end Sub
Sub rr_Dimensiones

    rr_LeeDdd
'response.write "<br>1492 " & rr_sDdd   
' Llenar Combo	
	ix=0
	for i=1 to 20
	    sx=Mid(rr_sDdd,i,1)
        iL=Asc(sx)-64
        if rr_sDim(il,0)=0 then
             il=0
             rr_sDdd=replace (rr_sDdd,sx,"@")
        end if     
        if il>0 then
            ix=ix+1    

            rr_sCom(ix,01)=rr_sDim(il,01)  ' Nombre
    '        response.write "<br> 1157 i:=" & i & " il=" & il & " " &  rr_sCom(ix,01)
            rr_sCom(ix,02)=rr_sDim(il,02)  ' Sql
            rr_sCom(ix,03)=rr_sDim(il,03)  ' 1=No se puede cambiar la dimension
            rr_sCom(ix,07)=rr_sDim(il,07)  ' Totales en Combo
            rr_sCom(ix,06)=il              ' Letra
            rr_sCom(ix,10)=rr_sDim(il,10)  ' NombreCampo
            rr_sCom(ix, 9) = rr_sPar(ix,0) ' Parametros
            rr_sCom(ix,04)=rr_sDim(il,06)  ' COmparar Igual o like
        end if    
	next    
	
    rr_nCom=ix
 
' Llenar Fila
	ix=0
	for i=21 to 25
	    iL=Asc(Mid(rr_sDdd,i,1))-64
	    if il>0 then
	        ix=ix+1
	        rr_sFil(ix,01)=rr_sDim(il,01)  ' Nombre
	        rr_sFil(ix,04)=rr_sDim(il,04)  ' Ancho
	        rr_sFil(ix,05)=rr_sDim(il,05)  ' Total
	        rr_sFil(ix,06)=il              ' Letra
	        rr_sFil(ix,10)=rr_sDim(il,10)  ' NombreCampo
	    end if    
	next    
	rr_nFil=ix

' Llenar Columna
	ix=0
    iL=Asc(Mid(rr_sDdd,26,1))-64
    if il<>0 then
        ix=ix+1
        rr_sCol(1)=rr_sDim(iL,1)     ' Nombre
        rr_sCol(2)=rr_sDim(iL,2)     ' Sql
        rr_sCol(6)=il                ' Letra
        rr_sCol(10)=rr_sDim(iL,10)   ' NombreCampo
        rr_ilof=1

' Link de Periodo        
        if rr_sCol(1)="Periodo" then 
           if rr_ipExe<>1 then
           
                if rr_iAnt<>0 then 
                    ixS=rr_ipPer
                    rr_ipPer=rr_iAnt
                    rr_calpar
                    'sx="<a href='" & sPar & "' title='Periodo anterior " & rr_iAnt & "' ><< </a>"
                    sx="<a href='" & sPar & "' title='Periodo anterior " & rr_iAnt & "' ></a>"
                    rr_ipPer=ixS
                    rr_sCol(1) = sx & rr_sCol(1)
                end if
        
                
                if rr_iPro<>0 then
                    ixS=rr_ipPer
                    rr_ipPer=rr_iPro
                    rr_calpar
                    rr_ipPer=ixS
                    'sx="<a href='" & sPar & "' title='Próximo Periodo" & rr_iPro &   "' > >></a>"
                    sx="<a href='" & sPar & "' title='Próximo Periodo" & rr_iPro &   "' ></a>"
                    rr_sCol(1) =  rr_sCol(1) & sx
                end if   
                
            end if        
        end if       
    else%>
    		<br />
    		<center>
			<hr class="rr_hr" />
			<font face= "Tahoma"  size="5" color="<%=cGris2%> " >
			Error: Debe haber al menos una columna para mostrar
			<hr class="rr_hr" />
			</center>
    <%exit sub
    end if    
	 
end sub


Sub rr_Main 
    rr_ParDat  
    if rr_mVar="" then rr_mVar=rr_nVar  
    if rr_tVar="" then rr_tVar=rr_mVar 

  
    rr_Dimensiones
'response.write "<br> Time11:=" & Time:response.flush		
    if rr_ilof = 0 then exit sub
    rr_ilof = 0
    sx="#ffffff"
    
    'response.write "<br>181 rr_Mesdes:=" & rr_Mesdes
    'response.write "<br>182 rr_Mesdes:=" & rr_MesHas
    'response.write "<br>183 ipPer:=" & iPer
   
    if rr_ipExe=1 then sx="#ffffff"
    %>
    
    
    <table width="100%"  border="0" align="center" cellspacing="0" bgcolor="<%=rr_sColFon%>"  ID="Table2">
    <tr>
        <td>
        
	      <% rr_Vervis%>
	    
    </td></tr>
    <tr>
    <td width="70%" align="center"><%=rr_sTit %>
        </td>
        </tr>
    </table>
	   
    <%
    'response.write "<br> leecombo Time:=" & Time:response.flush		 
    if rr_ipdis=0 then rr_LeeCombo 
    rr_LeeCol
    sx="#ffffff"
    if rr_ipExe=1 then sx="#ffffff"
    'Luis Reyes 12-08-2014
    %>
  
 
	
				
	
	<% 
       if rr_ipDis=0 then 
            rr_LeeDat
       else
            rr_ilof = 1     
       end if   
         
       %>
		    <center>
            <%sx=rr_sColFon
            if rr_ipExe=1 then sx="#ffffff"%>		    
                <div style="background-color:<%=sx%>; width:100%">
		    
                    <div   style="background-color:<%=rr_sColFon%>; width: 98%;padding: 0px 0px 0px 0px; margin-left:auto; margin-right:auto;   ">
    
                        <%if rr_ipExe=0 then %>
                        <div style="width:50%; float:left;vertical-align:middle;padding:3px 0px 3px 0px;"    >
                            <% 'rr_VerMeses %>
		                </div>
		                <%end if %>
                        <div style="width:50%;float:left;vertical-align:middle;  margin-right:auto"   >
                            <div><% rr_VerCombo%></div>
                        </div>   
        
	                </div>
	            <% if ipGra=1 then %>
	                <%vGraBar %>
	            <%end if %>
		        
		        <%if rr_ilof = 1	 then  %>
  	                <div style="font-family:Tahoma; float:none; font-size:10px;width:98%; text-align:right; background-color:<%=rr_sColFon%>; margin-top:15px; color:#000000">
     	                <%' Nota: (*) Valor entre 0 y 0.4 %>
     	            </div>       
      		        <br />  
		        
		            <% sBg="#cccccc"
		            if rr_ipFot=1 and rr_ipdis=0 then rr_VerFotos rr_sqlFot,3,5,120,90, sbg%>
		            <%
					rr_VerDat
					%>
		            <br />
		        
		        
	            <%else%>
    	            <table align="center" border="0" bgcolor="<%=rr_sColFon%>">
     	                <tr><td><hr class="rr_hr" /></td></tr>
     	                <tr><td class="rr_men1">No hay data para Mostrar</td></tr>
			            <tr><td><hr class="rr_hr" /></td></tr>
                    </table>			
	            <%end if%>
	           </div>
		        <br/>
		        </center> 
        <br/>

<%
End Sub







Sub rr_VerMeses
exit sub
    ipExe=0
    CalPar
    Dim RsBus

    set RsBus = CreateObject("ADODB.Recordset")
    RsBus.CursorType = 0
    RsBus.LockType = 1
    sql = ""
    sql = sql  & " SELECT "
    sql = sql  & " Id_Perfil_Cliente, "
    sql = sql  & " Periodo_Pub_Mens_Desde, "
    sql = sql  & " Periodo_Pub_Mens_Hasta "
    sql = sql  & " FROM "
    sql = sql  & " O_Perfil_Cliente "
    sql = sql  & " WHERE "
    sql = sql  & " Id_Perfil_Cliente = " & iPer
    RsBus.Open sql ,conexion
    rr_MesDesTot = RsBus.fields("Periodo_Pub_Mens_Desde")
    rr_MesHasTot = RsBus.fields("Periodo_Pub_Mens_Hasta")
    if INIvUsu = 3 then rr_MesHasTot = rr_MesHasTot + 1
    RsBus.close
 
%>

<table style="width:460px; height:50px; background:#ffffff; border: 1px solid #ccdbe4; " align="center"  border="0">
    <tr><td align="center" class="rr_dim2" style=" font-weight:bold" >Período</td>
    <td align="center" class="rr_dim1">Desde:<br /> 
    
    
    <select size="1" name="DesdeMes" id="Select2"  onchange="location.href=this.options[this.selectedIndex].value"  style="font-family:Verdana;font-weight:normal;height:20;font-size:8pt; vertical-align: middle;">
    
    <%
        'for i=rr_iPriper to rr_iUltPer
        for i=rr_MesDesTot to rr_MesHasTot
                z=int(i/12)
                M=i-z*12
                if m=0 then 
                    m=12
                    z = z - 1
                end if
                
                ipExe=0
                iM=rr_MesDes
                rr_MesDes= i
                rr_CalPar
                rr_mesDes=im
                iSi=rr_MesDes-i
                %>
                <option value="<%=sPar%>"
                <%if iSi=0 then response.write " selected " %>
                >
                <%
                sL=Ucase(mid(MonthName(m),1,1)) & mid(MonthName(m),2,2)
                response.write   sL & " - " & z  '& "--" & rr_iMes(ix) & " -- " & iz & smes
                
                %>
                
                </option>
                
                
        <%
            response.write  "<br>"
        next
    %>
    </select>
    </td>
    
  <td  align="center" class="rr_dim1">Hasta:<br />
    
    
    <select size="1" name="HastaMes" id="Select5"  onchange="location.href=this.options[this.selectedIndex].value" style="font-family:Verdana;font-weight:normal;height:20;font-size:8pt; vertical-align: middle;">
    
    <%
        'for i=rr_iPriper to rr_iUltPer
        for i=rr_MesDesTot to rr_MesHasTot
                z=int(i/12)
                M=i-z*12
                if m=0 then 
                    m=12
                    z = z -1 
                end if
                iM=rr_MesHas
                rr_MesHas= i
                rr_CalPar
                rr_MesHas=im
                iSi=rr_MesHas-i
                 %>
                <option value="<%=sPar %>"
                <%if iSi=0 then response.write " selected " %>
                >
                <%
                sL=Ucase(mid(MonthName(m),1,1)) & mid(MonthName(m),2,2)
                
                response.write  sL  & " - " & z  '& "--" & rr_iMes(ix) & " -- " & iz & smes
                
                %>
                
                </option>
                
                
        <%
            response.write  "<br>"
        next
    %>
    </select>
    
    </td>
  <td  align="center" class="rr_dim1"><br />Promedio:
    
    <select size="1" name="Promedio" id="Select3"  onchange="location.href=this.options[this.selectedIndex].value" style="font-family:Verdana;font-weight:normal;height:20;font-size:8pt; vertical-align: middle;">
                <%
                iM = rr_Prom
                rr_Prom = 1
                rr_CalPar
                rr_Prom = iM
                %>
                <option value="<%=sPar %>"
                <%if rr_Prom=1 then response.write " selected " %>
                >Si
                </option>
                <%
                iM = rr_Prom
                rr_Prom = 0
                rr_CalPar
                rr_Prom = iM
                %>
                <option value="<%=sPar %>"
                <%if rr_Prom=0 then response.write " selected " %>
                >No
                </option>
                <%
            response.write  "<br>"
    %>
    </select>
    </td>    
    
    
    
    
    </tr>    
</table>

<%
    'response.write "<br>2044 Promedio:= " & rr_Prom
    exit sub

	rr_ipExe=0
    rr_CalPar
'response.write sPar
' Leer CheckBox
	sx="option2"
	ipCom=0
	ix=0
	for each valor in request.Form(sx)
		ix=ix+1
		response.Write  ix & "-" & (valor)  & "<br>"
		rr_iMes(ix) = valor
		sMes=smes &"@" & Valor
	next
	
    if ix=0 then 
        for i=rr_iPriper to rr_iUltPer
            ix=ix+1
            rr_Imes(ix)=i
            sMes=smes &"@" & i
        next
     end if   

    %>
<div style=" background:#ffffff; text-align:right; width:300px">
<form action="<%=sPar %>" method='post' id="Form9">    
<font face= "verdana" size="2" color="#000066" >


<%
        ix=0
        for i=rr_iPriper to rr_iUltPer
            
                z=int(i/12)
                M=i-z*12
                if m=0 then m=12
                ix=ix+1
                'rr_Imes(ix)=i
                sx="@" & i
                iz=instr(sMes,sx)
                'iz=rr_imes(ix)-i
                response.write   mid(MonthName(m),1,3) & "-" & z  '& "--" & rr_iMes(ix) & " -- " & iz & smes
                
                %>
                <input type="checkbox" name="option2" value="<%=i%>" 
                 <% if iz<>0 then response.write "checked"%>
                 > 
                
        <%
            response.write  "<br>"
        next
     
    
%>
       <input type="submit" />
</font>
    
</form>
</div>

<%
end sub

%>
