<!--encabezado--> 
<% 
'dim CodigoPerfil
'dim CodigoProfit
CodigoPerfil = Session("idPerfil")
CodigoProfit = Session("filtro1")
'response.write "<br>8 link syntonex:= " & Session("linksyn") & " = " & Session("menusyn")
'response.write "<br>9 link jvg:= " & Session("linkjvg") & " = " & Session("menujvg")
'response.write "<br>9 CodigoPerfil:= " & CodigoPerfil
sLinkEmp = "default.asp"
if not isnull(Session("linksyn") ) then sLinkEmp = Session("linksyn") & "?edpas=1" & "&smenu=" & Session("menusyn")

sLinkCli = "default.asp"
if not isnull(Session("linkjvg") ) then sLinkCli = Session("linkjvg") & "?edpas=1" & "&smenu=" & Session("menujvg")
	sLinkPrecios = "Sys_rPrecios.asp?x=1&smenu=Reporte%20de%20Precios&edpas=1&edcla=999999&edpag=1&edbus=&edcol=&edord=&ed_fil=&ed_mp=1&ed_ms=4&cc_p1=&cc_p2=&cc_p3=&cc_p4=&cc_p5=&cc_p6=&cc_p7=&cc_p8=&cc_p9=&ed_des=0 "	
	'sLinkEmp = "http://atenas.conexionremota.com.ve/ph_mPanelHogares.asp?x=1&smenu=Hogares%20/%20Socio%20Demografico&edpas=1&edcla=999999&edpag=1&edbus=&edcol=&edord=&ed_fil=&ed_mp=41&ed_ms=46&cc_p1=&cc_p2=&cc_p3=&cc_p4=&cc_p5=&cc_p6=&cc_p7=&cc_p8=&cc_p9=&ed_des=0"
	if Session("perusu") <> 5 and Session("perusu") <> 7 and Session("perusu") <> 8 then 
		sLinkEmp = "http://atenas.pricetrack.com.ve.192-185-6-37.hgws18.hgwin.temp.domains/ph_mPanelHogares.asp?x=1&smenu=Hogares%20/%20Socio%20Demografico&edpas=1&edcla=999999&edpag=1&edbus=&edcol=&edord=&ed_fil=&ed_mp=41&ed_ms=46&cc_p1=&cc_p2=&cc_p3=&cc_p4=&cc_p5=&cc_p6=&cc_p7=&cc_p8=&cc_p9=&ed_des=0"
		sLinkCli = ""
	else
		sLinkEmp = "https://atenasconsultores.com/"
		sLinkCli = ""
	end if
	iCli = cint(Session("idCliente"))
	sLogo = ""
	'response.write "<br> iCli:= " & iCli
	select case iCli
		case 3
			sLogo = "Pepsico.png"
		case 6
			sLogo = "Sindoni.png"
		case 7
			sLogo = "Genica.png"
		case 8
			sLogo = "dimassilogo.png"
		case 9
			sLogo = "Cargill.jpg"
		case 10
			sLogo = "ScJohnson.png"
		case 11
			sLogo = "pepsico-logo.png"
		case 12
			sLogo = "cocacola.png"
		case 14
			sLogo = "Iancarina.jpg"
		case 15
			sLogo = "Fisa.png"
		case 16
			sLogo = "Colgate-Palmolive.png"
		case 17
			sLogo = "Nestle.jpg"
		case 19
			sLogo = "logoBaron.jpg"
		case 20	
			sLogo = "centralelpalmar.png"
		case 21
			sLogo = "eltunal.jpg"
		case 24
			sLogo = "Heinz-logo.png"
		case 27
			sLogo = "Pharsana.png"
		case 23
			sLogo = "LogoBotalon.png"
		case 28
			sLogo = "Alimex.jpg"
			
			
			
	end select 
	
	
	
	'sLinkPrecios = ""
	'response.write  "Link:=" & sLinkEmp
%>
<br>
<div class="row">
    <div class="col-md-12">
        <div class="row">
            <header>
                <div class="col-md-4">
					<div class="pull-center">
						<a href="<%=sLinkEmp%>" title="<%=Session("menusyn")%>"><img alt="Logo de la Empresa" src="images/logo/LogoAtenasNew02.jpeg" class="img-responsive left-block"></a>
					</div>
				</div>
                <div class="col-md-4">
					<h5 class="text-center"><strong>Usuario: <%=Session("NomApe")%>	</strong></h5>
					<!--<a href="default.asp"><img alt="Logo de la Empresa" src="images/logoEdgewell_opt.png"></a>-->
                </div>
				<%
				if Session("perusu") = 5 or Session("perusu") = 7  or Session("perusu") = 8 then 
				%>
				<div class="col-md-4">
					<div class="pull-right">
							<a href=""><img alt="Logo de la Empresa" src="images/logo/<%=sLogo%>" class="img-responsive center-block" > </a>
					</div>  		
				</div>  																	
				<%
				end if
				%>
            </header>
        </div>       
		<center>
					
		</center>		
    </div>
</div>

<br>