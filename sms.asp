
<!DOCTYPE HTML>
<html>
<head>

	<title>Envio SMS Gratis</title>
	<meta http-equiv="Content-Type" content="text/html; ISO-8859-1">  
    <link href="np.css" rel="stylesheet" type="text/css" media="screen" />
     
</head>

<script type="text/javascript">
var myWindow;

function EnviarMensajeTexto(slink) 
{
	var httpRequest;
	var sx="https://site.albertext.com/api/messages/save-message";
	sx= sx + slink; 
//	alert("Entro:="+sx);
	
	if (window.XMLHttpRequest)
	{
		//El explorador implementa la interfaz de forma nativa
		httpRequest = new XMLHttpRequest();
	} 
	else if (window.ActiveXObject)
	{
		//El explorador permite crear objetos ActiveX
		try {
			httpRequest = new ActiveXObject("MSXML2.XMLHTTP");
		} catch (e) {
			try {
				httpRequest = new ActiveXObject("Microsoft.XMLHTTP");
			} catch (e) {}
		}
	}
	if (!httpRequest)
	{
		alert("No ha sido posible enviar los mensaje de texto a celular");
	}
	else
	{
		httpRequest.open("POST",sx,false);
		httpRequest.send();
		//Set httpRequest = nothing
	}
	//alert("Enviar:="+sx);
	//Set httpRequest = nothing
	//alert("Envio los Mensajes de Texto"+sx);
}
</script>
<body>
<%

        sx=""
 
        
        sx="?user=Atenas&token=XZM2tfgOW0tscbRETqh91H7TKcT19NES&phone=584242110998&text=Prueba"
      '  response.write "<br> sx=" &  sx
        %>
        
        <script>
            EnviarMensajeTexto('<%=sx%>');
        </script>
        
       <% sSms="https://site.albertext.com/api/messages/save-message/"&sx %>
        <a href=<%=sSms %>  > Envio</a>           
      


</body>
</html>
