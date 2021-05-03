$('document').ready(function() { 
	/* actualizado 25jun17*/
    /* validacion */
	
    $("#login-form").validate({
				
		rules: {
			password: { required: true, },
			email: { required: true, email: true },
		},
		messages:{
            password:{ required: "FAVOR INDIQUE SU PASSWORD" },
            email: "FAVOR INDIQUE SU EMAIL",
        },    
		submitHandler: submitForm	
		
    }); /* Fin validacion Datos*/
    
	function submitForm(){		
		
		/* Stop form from submitting normally */
		event.preventDefault();
		
		 /* Clear result div*/
		$("#error").html('');
		
		/* login enviado DATOS ENVIADOS*/		
        var data = $("#login-form").serialize();
				
        $.ajax({
            type : 'POST',
            url  : 'pr_login_proceso.asp',
            data : data,
            beforeSend: function(){	
                $("#error").fadeOut();
                $("#btn-login").html('<span class="glyphicon glyphicon-transfer"></span> &nbsp; enviando ...');
            },
            success : function(response){						
                if (response=="usuario") {					
					$("#btn-login").html('<img src="images/ajax-loader3.gif" /> &nbsp; Espere ...!');
                    setTimeout(' window.location.href = "pr_mInicio.asp"; ',1500);					
				}
				else if (response=="auditor") {
					$("#btn-login").html('<img src="images/ajax-loader3.gif" /> &nbsp; Ingresando ..!');					
                    setTimeout(' window.location.href = "pr_mInicioAdt.asp"; ',1500);					
				}
				else {
					$("#error").fadeIn(2000, function(){						
						$("#error").html('<div class="alert alert-danger"> <span class="glyphicon glyphicon-info-sign"></span> &nbsp; '+response+' !</div>');
						$("#btn-login").html('<span class="glyphicon glyphicon-log-in"></span> &nbsp; Ingresar');                                                       
					});
				}
            }
        });
        return false;
    }			
});