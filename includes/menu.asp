<nav class="navbar navbar-default ">
	<div class="container-fluid">
		<!-- Brand and toggle get grouped for better mobile display -->
		<!-- Creado: 105abr7 -- Actualizado: 24oct17-->
		<div class="navbar-header">
			<button type="button" class="navbar-toggle collapsed" data-toggle="collapse" data-target="#bs-example-navbar-collapse-1" aria-expanded="false">
				<span class="sr-only">Toggle navigation</span>
				<span class="icon-bar"></span>
				<span class="icon-bar"></span>
				<span class="icon-bar"></span>
			</button>	
			<a class="navbar-brand" href="pr_mInicioAdt.asp"><i class='glyphicon glyphicon-home'></i></a>
		</div>	
		<!-- Collect the nav links, forms, and other content for toggling -->
		<div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
			<ul class="nav navbar-nav">
				<li class="<%= sActivar_Solicitar%>"><a href="pr_mSolicitar.asp"><i class="glyphicon glyphicon-calendar"></i> Solicitar <span class="sr-only">(current)</span></a></li>				
				<li class="<%= sActivar_Pendientes%>"><a href="pr_mPendientes.asp"><i class="glyphicon glyphicon-list"></i> Aprobar</a></li>
				<li class="<%= sActivar_Soporte%>"><a href="pr_mSoporte.asp"><i class="glyphicon glyphicon-check"></i> Registrar Soportes</a></li>						
			</ul>
			<ul class="nav navbar-nav navbar-right">				
				<li>
					<a href="pr_mInicioCta.asp"><i class="glyphicon glyphicon-lock"></i> Password</a>
				</li>
				<li>
					<a href="pr_mMaestros.asp"><i class="glyphicon glyphicon-calendar"></i> Maestros</a>											
				</li>
				<li>
					<a href="pr_mInicioAdtSalir.asp"><i class="glyphicon glyphicon-off"></i> Salir</a>
				</li>	
			</ul>		
		</div><!-- /.navbar-collapse -->
	</div><!-- /.container-fluid -->
</nav>