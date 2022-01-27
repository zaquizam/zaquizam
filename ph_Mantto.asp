<%@Language="VBScript"%>
<!DOCTYPE HTML>
<html >
<head>
	<title>Atenas | Mantenimiento</title>
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
	<link rel="icon" href="favicon.ico" type="image/x-icon"> 
	<link href="css/bootstrap.min.css" rel="stylesheet" type="text/css" />	
</head>
<body topmargin="0">
	<!--#include file="estiloscss.asp"-->
	<!--#include file="meta.asp"-->
	<!--#include file="encabezado.asp"-->
	<!--#include file="nn_subN.asp"-->
	<!--#include file="in_DataEN.asp"-->
	<%		
		Apertura		
		' Parámetros del Manteniemiento		
		LeePar		
		if ed_iPas<>4 then 
			Encabezado
		end if    	
		'		
		Session.lcid		= 1034
		Response.CodePage 	= 65001
		Response.CharSet 	= "utf-8"
		'		
	%>
	
	<div class="container-fluid">
       
	   <div class="row">
	   
          <div class="col-md-12">
		  
            <div class="text-xs-center text-lg-center">           
		   		<img src="images/mantto.png" class="img-responsive" style="margin:15px auto;" onclick="location.href = 'pr_mInicio.asp';" />
				<button type="button" class="btn btn-danger center-block" onclick="location.href = 'pr_mInicio.asp';">Inicio</button>
           </div>
		   
          </div>
		  
       </div>
	   
  	</div>	
							
	<%conexion.close%>

</body>
</html>

<script src="https://kit.fontawesome.com/9d7cfbccc5.js" crossorigin="anonymous"></script>
<script src="js/jquery-3.1.1.min.js"></script>
<script src="js/bootstrap.min.js"></script>

<script>
	
	$(function() {		
		redirectTime = "15000";
		redirectURL = "pr_mInicio.asp";
		function timedRedirect() {
			setTimeout("location.href = redirectURL;",redirectTime);
		}
		timedRedirect();			
	});
	
</script>
