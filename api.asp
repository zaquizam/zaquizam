<% 
	dim urlApi
	if(request.servervariables("remote_addr")=request.servervariables("local_addr")) Then
		urlApi = "http://localhost:3000/api"
	else
		urlApi = "http://216.198.73.34:3000/api"
	end if
	Session("urlApi")  = urlApi

%>