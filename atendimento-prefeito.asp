<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cpf.asp" -->
<!--#include file="logout.asp" -->
<!--#include file="daee_restrict_access.asp" -->
<%
	Response.CharSet = "UTF-8"

	Dim objCon
	Set objCon = Server.CreateObject("ADODB.Connection")
		objCon.Open MM_cpf_STRING
%>
<!DOCTYPE html>
<html>
<head>
	<title>:: DAEE ::</title>
	<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
	<link rel="stylesheet" type="text/css" href="css/bootstrap-flaty.min.css">
	<link rel="stylesheet" type="text/css" href="css/daee.css">
	<link rel="stylesheet" href="//maxcdn.bootstrapcdn.com/font-awesome/4.3.0/css/font-awesome.min.css">
	<script type="text/javascript" src="//code.jquery.com/jquery-1.11.2.min.js"></script>
	<script type="text/javascript" src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js"></script>
	<script type="text/javascript" src="http://maps.googleapis.com/maps/api/js?libraries=places&sensor=false"></script>
	<script type="text/javascript" src="js/jquery.number.min.js"></script>
	<script type="text/javascript" src="js/fullscreen.js"></script>
	<style type="text/css">
		.container {
			width: 100%;
		}

		.container .row {
			overflow-y: hidden;
		}

		.sidebar {
			overflow-y: scroll;
			padding-top: 10px;
		}

		.sidebar .find, .sidebar .result {
			padding-top: 10px; 
		}

		.mapa {
			padding-left: 0px;
			padding-right: 0px;
		}

		.mapa #map-canvas {
			width: 100%;
			height: 600px;
			bottom:0;
			overflow:hidden;
		}

		#modalInfoMunicipio {
			overflow-y: scroll !important;
		}

		.btn-text-left {
			text-align: left !important;
		}

		.table-responsive {
			margin-top: 10px;
		}
	</style>
	<script type="text/javascript">
		var map;
		var markersInMap = [];

		function initMap(){
			// objeto de preferências e configurações do mapa do google maps
			var mapOptions = {
				streetViewControl: false,
				scrollWheel: false,
				zoom: 7,
				mapTypeId: google.maps.MapTypeId.TERRAIN
			};

			// inicializa o google maps and adiciona-o ao html
			map = new google.maps.Map(document.getElementById("map-canvas"), mapOptions);
			var initialLocation = new google.maps.LatLng(-22.5455869,-48.6355226);
			map.setCenter(initialLocation);
		}

		function setAllMap(map){
			for(var i=0; i< markersInMap.length; i++) {
				markersInMap[i].setMap(map);
			}
		}

		function clearMarkers() {
			setAllMap(null);
		}

		function deleteMarkers() {
			clearMarkers();
			markersInMap = [];
		}

		function addMunicipioMarkersToMap(markers){
			$.each(markers, function(i, item){
				if(item.latlog != null){
					var latlog 		= item.latlog.split(",");
					var latitude 	= ((latlog[0]).indexOf("-") >= 0) ? -Math.abs(latlog[0]) : Math.abs(latlog[0]);
					var longitude 	= ((latlog[1]).indexOf("-") >= 0) ? -Math.abs(latlog[1]) : Math.abs(latlog[1]);

					var location = new google.maps.LatLng(latitude, longitude);

					var marker = new google.maps.Marker({
						icon: 'img/location-primary.png',
						map: map,
						position: location,
						title: item.municipio
					});

					markersInMap.push(marker);

					var content = ''+
						'<strong>'+ item.municipio +'</strong>'+
						'<br/>'+
						'<strong>Prefeito: </strong>'+ item.prefeito +
						'<br/>'+
						'<strong>População em 2010 (IBGE): </strong> '+ item.pop2010 +
						'<br/>'+
						'<strong>Projeção p/ 2030 (hab):</strong> '+ item.pop2030 +
						'<br/>'+
						'<strong class="text-'+ item.classProgram +'">'+ item.textProgram +'</strong>'+
						'<br/><br/>'+
						'<a class="btn btn-sm btn-block btn-info" href="informacao-municipio-resumida.asp?cod_municipio='+ item.cod_municipio +'">'+
							'+ Informações'+
						'</a>';
					var infowindow = new google.maps.InfoWindow();

					google.maps.event.addListener(marker,'click', (function(marker,content,infowindow){
						return function() {
							infowindow.setContent(content);
							infowindow.open(map,marker);
						};
					})(marker,content,infowindow));
				}
			});
		}

		function addEmpreendimentoMarkersToMap(markers){
			$.each(markers, function(i, item){
				if(item.latlog != null){
					item.latlog 	= replaceAll("°", "", item.latlog);
					var latlog 		= item.latlog.split(",");
					var latitude 	= ((latlog[0]).indexOf("-") >= 0) ? -Math.abs(latlog[0]) : Math.abs(latlog[0]);
					var longitude 	= ((latlog[1]).indexOf("-") >= 0) ? -Math.abs(latlog[1]) : Math.abs(latlog[1]);

					var location = new google.maps.LatLng(latitude, longitude);

					var marker = new google.maps.Marker({
						icon: 'img/location-'+ item.classStatus +'.png',
						map: map,
						position: location,
						title: item.nme_municipio +" - "+ item.nme_empreendimento
					});

					markersInMap.push(marker);

					link = "";

					switch(item.cod_situacao_interna){
						case 39: // Concluída
							link = "ficha-tecnica-obra-concluida.asp?cod_empreendimento=" + item.cod_empreendimento;
							break;
						case 40: 
						case 42: // Em atendimento
						case 43:
							link = "ficha-tecnica-obra.asp?cod_empreendimento=" + item.cod_empreendimento;
							break;
						case 44: // Programado para atendimento
						case 41: // Não atendido
							link = "ficha-tecnica-obra-programada.asp?cod_empreendimento=" + item.cod_empreendimento;
							break;
						case 45: // Atendimento potencial
							link = "";
							break;
					}

					var btnInformacoes = "";
					var btnZoom = '<a href="#" class="btn btn-sm btn-block btn-primary btn-zoom-in" data-latitude="'+ latitude +'" data-longitude="'+ longitude +'">Zoom na Obra</a>';
					var btnVoltarZoom = '<a href="#" class="btn btn-sm btn-block btn-primary btn-zoom-out hide">Voltar</a>';

					if(link != "")
						btnInformacoes = '<br/><br/><a href="'+ link +'" class="btn btn-sm btn-block btn-info btn-dados-municipio">+ Informações</a>'

					var content = ''+
					'<strong>'+ item.nme_municipio +" - "+ item.nme_empreendimento +'</strong>'+
					'<br/>'+
					'<strong>População em 2010 (IBGE): </strong> '+ $.number(item.pop2010, 0, ",", ".") +
					'<br/>'+
					'<strong>Projeção p/ 2030 (hab):</strong> '+ $.number(item.pop2030, 0, ",", ".") +
					btnInformacoes +
					btnZoom + 
					btnVoltarZoom;
					
					var infowindow = new google.maps.InfoWindow();

					google.maps.event.addListener(marker,'click', (function(marker,content,infowindow){
						return function() {
							infowindow.setContent(content);
							infowindow.open(map,marker);
							addButtonZoomClickListener();
						};
					})(marker,content,infowindow));
				}
			});
		}

		function addButtonZoomClickListener() {
			$("a.btn-zoom-in").on("click", function(){
				var itemData = $(this).data();
				var location = new google.maps.LatLng(itemData.latitude, itemData.longitude);
				
				map.setMapTypeId(google.maps.MapTypeId.SATELLITE);
				map.setCenter(location);
				map.setZoom(18);
				
				$("a.btn-zoom-out").removeClass("hide");
				$(this).addClass("hide");
			});

			$("a.btn-zoom-out").on("click", function(){
				var initialLocation = new google.maps.LatLng(-22.5455869,-48.6355226);
			
				map.setMapTypeId(google.maps.MapTypeId.TERRAIN);
				map.setCenter(initialLocation);
				map.setZoom(7);

				$("a.btn-zoom-in").removeClass("hide");
				$(this).addClass("hide");
			});
		}

		function zoomTo(level) {
			google.maps.event.addListener(map, 'zoom_changed', function () {
				zoomChangeBoundsListener = google.maps.event.addListener(map, 'bounds_changed', function (event) {
					if (this.getZoom() > level && this.initialZoom == true) {
						this.setZoom(level);
						this.initialZoom = false;
					}
					google.maps.event.removeListener(zoomChangeBoundsListener);
				});
			});
		}

		function addButtonInfoClickListener() {
			$(".btn-dados-municipio").on("click", function(){
				var codMunicipio = $(this).closest("tr").data().codMunicipio;
				loadDataMunicipio(codMunicipio);
				loadEmpreendimentosByMunicipio(codMunicipio);
			});
		}

		function listMunicipiosTable(municipios) {
			$.each(municipios, function(i, item) {
				var itemLayout = '<tr data-cod-municipio="'+ item.cod_municipio +'">'+
									'<td class="text-middle">'+
										'<i class="fa fa-'+ item.iconProgram +' text-'+ item.classProgram +'" data-toggle="tooltip" data-placement="right" title="'+ item.textProgram +'"></i> '+ item.municipio +
									'</td>'+
									'<td width="32">'+
										'<a class="btn btn-sm btn-info" href="informacao-municipio-resumida.asp?cod_municipio='+ item.cod_municipio +'"><i class="fa fa-info-circle"></i></a>'+
									'</td>'+
								'</tr>';
				$("table.table-municipios tbody").append(itemLayout);
			});

			$('[data-toggle="tooltip"]').tooltip();

			addButtonInfoClickListener();
		}

		function listEmpreendimentosTable(localidades) {
			$.each(localidades, function(i, item) {
				var link = "";

				switch(item.cod_situacao_interna){
					case 39: // Concluída
						link = "ficha-tecnica-obra-concluida.asp?cod_empreendimento=" + item.cod_empreendimento;
						break;
					case 40: 
					case 42: // Em atendimento
					case 43:
						link = "ficha-tecnica-obra.asp?cod_empreendimento=" + item.cod_empreendimento;
						break;
					case 44: // Programado para atendimento
					case 41: // Não atendido
						link = "ficha-tecnica-obra-programada.asp?cod_empreendimento=" + item.cod_empreendimento;
						break;
					case 45: // Atendimento potencial
						link = "";
						break;
				}

				var itemLayout = '<tr>'+
									'<td class="text-middle">'+ item.nme_empreendimento + '</td>'+
									'<td class="text-middle">'+ item.dsc_situacao_interna + '</td>'+
									'<td class="text-middle"><a href="'+ link +'" class="btn btn-sm btn-primary">Ficha Técnica</a></td>'+
								'</tr>';
				$("table.table-localidades tbody").append(itemLayout);
			});
		}

		function setLayersHeight() {
			$(".sidebar").css("height", $(window).height()-60)
			$("#map-canvas").css("height", $(window).height()-60)
		}

		function clearTableMunicipiosData(){
			$("table.table-municipios tbody tr").remove();
		}

		function clearTableEmpreendimentosData(){
			$("table.table-localidades tbody tr").remove();
		}

		function findMunicipios(nmeMunicipio) {
			$.ajax({
				url: "query-to-json-util.asp",
				method: "POST",
				data: {
					sql: 'SELECT * FROM c_status_municipios_programa WHERE municipio LIKE "'+ nmeMunicipio +'%"'
				},
				beforeSend: function() {
					if(!$(".alert-danger").hasClass("hide"))
						$(".alert-danger").addClass("hide");
					
					if(!$("#btn-clear").hasClass("hide"))
						$("#btn-clear").addClass("hide");

					clearTableMunicipiosData();
					deleteMarkers();
					$("#modalLoading").modal("show");
				},
				success: function(data, textStatus, jqXHR){
					data = JSON.parse(data);

					if(data.length > 0) {
						var items = [];

						$.each(data, function(i, item) {
							pop2030 = item.qtd_populacao_urbana_2010 * 1.25;
							pop2030 = parseFloat((pop2030/100).toFixed(0));
							pop2030 = parseFloat((pop2030 * 100).toFixed(0));

							var markerData = {
								cod_municipio: item.id_predio,
								municipio: item.municipio,
								prefeito: "",
								pop2010: item.qtd_populacao_urbana_2010,
								pop2030: pop2030,
								latlog: item.latitude_longitude,
								classProgram: (item.cod_status == 1) ? "success" : "danger",
								iconProgram: (item.cod_status == 1) ? "check" : "times",
								textProgram: item.dsc_status
							};

							items.push(markerData);
						});

						addMunicipioMarkersToMap(items);
						listMunicipiosTable(items);
						$("#btn-clear").removeClass("hide");
					}
					else
						$(".alert-danger").removeClass("hide");

					$("#modalLoading").modal("hide");
				},
				error: function(jqXHR, textStatus, errorThrown){
					console.log(jqXHR, textStatus, errorThrown);
				}
			});
		}

		function adjustNumLayout() {
			$.each($(".num"), function(i, item){
				$(item).val($.number($(item).val(), 0, ",", "."));
				$(item).text($.number($(item).text(), 0, ",", "."));
			});
		}

		function adjustVlrLayout() {
			$.each($(".vlr"), function(i, item){
				// $(item).val($.number($(item).val(), 0, ",", "."));
				if($(item).text() != "")
					$(item).text("R$ " + $.number($(item).text(), 2, ",", "."));
			});
		}

		function adjustPrcLayout() {
			$.each($(".prc"), function(i,item){
				$(item).text($.number(( parseFloat($(item).text()) * 100 ), 2, ",", ".") + "%");
			});
		}

		function loadDataMunicipio(codMunicipio){
			$.ajax({
				url: "query-to-json-util.asp",
				method: "POST",
				data: {
					sql: "SELECT * FROM c_lista_predios WHERE id_predio = " + codMunicipio
				},
				beforeSend: function() {
					$("#modalLoading").modal("show");
				},
				success: function(data, textStatus, jqXHR){
					data = JSON.parse(data);

					if(data.length > 0) {
						var dadosMunicipio = data[0];

						var pop2030;
							pop2030 = dadosMunicipio.qtd_populacao_urbana_2010 * 1.25;
							pop2030 = parseFloat((pop2030/100).toFixed(0));
							pop2030 = parseFloat((pop2030 * 100).toFixed(0));

						$("#txt-dsc-municipio").val(dadosMunicipio['nme_prefeitura']);
						$("#txt-nme-bacia-daee").val(dadosMunicipio['bacia_daee']);
						$("#txt-nme-bacia-secretaria").val(dadosMunicipio['bacia_secretaria']);
						$("#txt-pop-2010").val(dadosMunicipio['qtd_populacao_urbana_2010']);
						$("#txt-pop-2030").val(pop2030);
						$("#txt-nme-prefeito").val(dadosMunicipio['nme_prefeito']);
						$("#txt-nme-partido").val(dadosMunicipio['nme_partido']);
						$("#txt-atendido-sabesp").val(dadosMunicipio['dsc_concessao']);

						adjustNumLayout();

						$("#btn-ficha-completa").attr("href", "informacao-municipio.asp?cod_municipio="+codMunicipio);

						$("#modalLoading").modal("hide");
						$("#modalInfoMunicipio").modal("show");
					}
					else
						alert("Nenhuma informação encontrada!");
				},
				error: function(jqXHR, textStatus, errorThrown){
					console.log(jqXHR, textStatus, errorThrown);
				}
			});
		}

		function replaceAll(find, replace, str) {
			return str.replace(new RegExp(find, 'g'), replace);
		}

		function loadEmpreendimentosByMunicipio(codMunicipio) {
			$.ajax({
				url: "query-to-json-util.asp",
				method: "POST",
				data: {
					sql: "SELECT * FROM c_lista_pi WHERE id_predio = " + codMunicipio
				},
				beforeSend: function() {
					clearTableEmpreendimentosData();
				},
				success: function(data, textStatus, jqXHR){
					data = JSON.parse(data);

					if(data.length > 0) {
						var items = [];

						$.each(data, function(i, item) {
							var itemData = {
								cod_empreendimento: item.PI,
								nme_empreendimento: item.nome_empreendimento,
								cod_situacao_interna: item.cod_situacao,
								dsc_situacao_interna: item.dsc_situacao_interna
							};

							items.push(itemData);
						});

						listEmpreendimentosTable(items);
					}
				},
				error: function(jqXHR, textStatus, errorThrown){
					console.log(jqXHR, textStatus, errorThrown);
				}
			});
		}

		function loadEmpreendimentos(sql) {
			if(sql.indexOf("AND") != -1) {
				$.ajax({
					url: "query-to-json-util.asp",
					method: "POST",
					data: {
						sql: sql
					},
					beforeSend: function() {
						clearTableMunicipiosData();
						deleteMarkers();
						$("#modalLoading").modal("show");
					},
					success: function(data, textStatus, jqXHR){
						data = JSON.parse(data);

						if(data.length > 0) {
							var items = [];

							$.each(data, function(i, item) {
								var pop2030;
								pop2030 = item.qtd_populacao_urbana_2010 * 1.25;
								pop2030 = parseFloat((pop2030/100).toFixed(0));
								pop2030 = parseFloat((pop2030 * 100).toFixed(0));

								var markerData = {
									cod_empreendimento: item.PI,
									nme_empreendimento: item.nome_empreendimento,
									nme_municipio: item.municipio,
									pop2010: item.qtd_populacao_urbana_2010,
									pop2030: pop2030,
									latlog: item.latitude_longitude,
									cod_situacao_interna: item.cod_situacao
								};

								switch(item.cod_situacao){
									case 39:
										markerData.classStatus = "success";
										break;
									case 40:
									case 42:
									case 43:
										markerData.classStatus = "warning";
										break;
									case 44:
										markerData.classStatus = "info";
										break;
									case 41:
										markerData.classStatus = "danger";
										break;
									case 45:
										markerData.classStatus = "primary";
										break;
								}

								items.push(markerData);
							});

							addEmpreendimentoMarkersToMap(items);
						}

						$("#modalLoading").modal("hide");
					},
					error: function(jqXHR, textStatus, errorThrown){
						console.log(jqXHR, textStatus, errorThrown);
					}
				});
			}
			else {
				clearTableMunicipiosData();
				deleteMarkers();
			}
		}

		$(function() {
			setLayersHeight();
			initMap();

			$(window).on("resize", function(){
				setLayersHeight();
			});

			$("#btn-clear").on("click", function(){
				$("#txt-municipio").val("");
				clearTableMunicipiosData();
				deleteMarkers();
				$("#btn-clear").addClass("hide");
			});

			$("#form-search").on("submit", function(e) {
				e.preventDefault();
				$("#txt-municipio").closest(".form-group").removeClass("has-error");
				var nmeMunicipio = $("#txt-municipio").val();
				if(nmeMunicipio != "")
					findMunicipios(nmeMunicipio);
				else
					$("#txt-municipio").closest(".form-group").addClass("has-error");
			});

			var sqlFiltroSituacao = "SELECT * FROM c_lista_pi WHERE 1 = 1 ";
			var hasAndSqlQuery = false;

			$("#chk-concluidas").on("click", function(){
				var context = "";

				if(!hasAndSqlQuery){
					hasAndSqlQuery = true;
					context = " AND cod_situacao = 39";
				}
				else
					context = " OR cod_situacao = 39";

				if(this.checked)
					sqlFiltroSituacao += context;
				else{
					if(sqlFiltroSituacao.indexOf(" AND cod_situacao = 39") != -1)
						context = " AND cod_situacao = 39";
					else if(sqlFiltroSituacao.indexOf(" OR cod_situacao = 39") != -1)
						context = " OR cod_situacao = 39";

					sqlFiltroSituacao = sqlFiltroSituacao.replace(context, "");

					if(sqlFiltroSituacao.indexOf("AND") == -1){
						sqlFiltroSituacao = sqlFiltroSituacao.replace("OR", "AND");
					}

					if(sqlFiltroSituacao.indexOf("AND") == -1){
						hasAndSqlQuery = false;
					}
				}

				loadEmpreendimentos(sqlFiltroSituacao);
			});

			$("#chk-em-andamento").on("click", function(){
				var context = "";

				if(!hasAndSqlQuery){
					hasAndSqlQuery = true;
					context = " AND cod_situacao IN (40,42,43)";
				}
				else
					context = " OR cod_situacao IN (40,42,43)";

				if(this.checked)
					sqlFiltroSituacao += context;
				else{
					if(sqlFiltroSituacao.indexOf(" AND cod_situacao IN (40,42,43)") != -1)
						context = " AND cod_situacao IN (40,42,43)";
					else if(sqlFiltroSituacao.indexOf(" OR cod_situacao IN (40,42,43)") != -1)
						context = " OR cod_situacao IN (40,42,43)";

					sqlFiltroSituacao = sqlFiltroSituacao.replace(context, "");

					if(sqlFiltroSituacao.indexOf("AND") == -1){
						sqlFiltroSituacao = sqlFiltroSituacao.replace("OR", "AND");
					}

					if(sqlFiltroSituacao.indexOf("AND") == -1){
						hasAndSqlQuery = false;
					}
				}
				
				loadEmpreendimentos(sqlFiltroSituacao);
			});

			$("#chk-atendimento-programado").on("click", function(){
				var context = "";

				if(!hasAndSqlQuery){
					hasAndSqlQuery = true;
					context = " AND cod_situacao = 44";
				}
				else
					context = " OR cod_situacao = 44";

				if(this.checked)
					sqlFiltroSituacao += context;
				else{
					if(sqlFiltroSituacao.indexOf(" AND cod_situacao = 44") != -1)
						context = " AND cod_situacao = 44";
					else if(sqlFiltroSituacao.indexOf(" OR cod_situacao = 44") != -1)
						context = " OR cod_situacao = 44";

					sqlFiltroSituacao = sqlFiltroSituacao.replace(context, "");

					if(sqlFiltroSituacao.indexOf("AND") == -1){
						sqlFiltroSituacao = sqlFiltroSituacao.replace("OR", "AND");
					}

					if(sqlFiltroSituacao.indexOf("AND") == -1){
						hasAndSqlQuery = false;
					}
				}
				
				loadEmpreendimentos(sqlFiltroSituacao);
			});

			$("#chk-nao-atendida").on("click", function(){
				var context = "";

				if(!hasAndSqlQuery){
					hasAndSqlQuery = true;
					context = " AND cod_situacao = 41";
				}
				else
					context = " OR cod_situacao = 41";

				if(this.checked)
					sqlFiltroSituacao += context;
				else{
					if(sqlFiltroSituacao.indexOf(" AND cod_situacao = 41") != -1)
						context = " AND cod_situacao = 41";
					else if(sqlFiltroSituacao.indexOf(" OR cod_situacao = 41") != -1)
						context = " OR cod_situacao = 41";

					sqlFiltroSituacao = sqlFiltroSituacao.replace(context, "");

					if(sqlFiltroSituacao.indexOf("AND") == -1){
						sqlFiltroSituacao = sqlFiltroSituacao.replace("OR", "AND");
					}

					if(sqlFiltroSituacao.indexOf("AND") == -1){
						hasAndSqlQuery = false;
					}
				}
				
				loadEmpreendimentos(sqlFiltroSituacao);
			});

			$("#chk-atendimento-potencial").on("click", function(){
				var context = "";

				if(!hasAndSqlQuery){
					hasAndSqlQuery = true;
					context = " AND cod_situacao = 45";
				}
				else
					context = " OR cod_situacao = 45";

				if(this.checked)
					sqlFiltroSituacao += context;
				else{
					if(sqlFiltroSituacao.indexOf(" AND cod_situacao = 45") != -1)
						context = " AND cod_situacao = 45";
					else if(sqlFiltroSituacao.indexOf(" OR cod_situacao = 45") != -1)
						context = " OR cod_situacao = 45";

					sqlFiltroSituacao = sqlFiltroSituacao.replace(context, "");

					if(sqlFiltroSituacao.indexOf("AND") == -1){
						sqlFiltroSituacao = sqlFiltroSituacao.replace("OR", "AND");
					}

					if(sqlFiltroSituacao.indexOf("AND") == -1){
						hasAndSqlQuery = false;
					}
				}
				
				loadEmpreendimentos(sqlFiltroSituacao);
			});
		});
	</script>
</head>
<body id="body">
	<nav class="navbar navbar-default navbar-fixed-top">
		<div class="container-fluid">
			<div class="navbar-header">
				<button type="button" class="navbar-toggle collapsed" data-toggle="collapse" data-target="#bs-example-navbar-collapse-1">
					<span class="sr-only">Toggle navigation</span>
					<span class="icon-bar"></span>
					<span class="icon-bar"></span>
					<span class="icon-bar"></span>
				</button>
				<a class="navbar-brand" href="#">Programa Água Limpa | Atendimento a Prefeitos</a>
			</div>

			<div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
				<ul class="nav navbar-nav navbar-right">
					<li><a href="javascript:window.history.back();"><i class="fa fa-chevron-left"></i> Voltar</a></li>
					<li><a href="#" class="expand"><i class="fa fa-expand"></i>&nbsp;&nbsp;Tela Cheia</a></li>
					<li><a href="<%= MM_Logout %>" class="sign-out"><i class="fa fa-sign-out"></i> Sair do Sistema</a></li>
				</ul>
			</div>
		</div>
	</nav>

	<div class="container">
		<div class="row">
			<div class="col-xs-3 sidebar">
				<div class="form-group">
					<label class="control-label">Relatórios Consolidados:</label>
				</div>

				<div class="form-group">
					<label class="control-label sr-only"></label>
					<a href="resumo-situacao.asp?rep_universo_programa=sim" class="btn btn-info btn-block btn-sm btn-text-left"><i class="fa fa-file-text-o"></i> Universo do Programa</a>
				</div>

				<div class="form-group">
					<label class="control-label sr-only"></label>
					<a href="resumo-situacao.asp?rep_universo_atendimento_programa=sim" class="btn btn-info btn-block btn-sm btn-text-left"><i class="fa fa-file-text-o"></i> Universo de Atendimento</a>
				</div>

				<div class="form-group">
					<label class="control-label">Filtro por Situação:</label>
				</div>

				<div class="checkbox">
					<label>
						<input id="chk-concluidas" type="checkbox"><i class="fa fa-map-marker text-success"></i> Localidades c/ obras concluídas
					</label>
				</div>

				<div class="checkbox">
					<label>
						<input id="chk-em-andamento" type="checkbox"><i class="fa fa-map-marker text-warning"></i> Localidades c/ obras em andamento
					</label>
				</div>

				<div class="checkbox">
					<label>
						<input id="chk-atendimento-programado" type="checkbox"><i class="fa fa-map-marker text-info"></i> Localidades c/ atendimento programado
					</label>
				</div>

				<div class="checkbox">
					<label>
						<input id="chk-nao-atendida" type="checkbox"><i class="fa fa-map-marker text-danger"></i> Localidades não atendidas
					</label>
				</div>

				<div class="checkbox">
					<label>
						<input id="chk-atendimento-potencial" type="checkbox"><i class="fa fa-map-marker text-primary"></i> Localidades c/ atendimento potencial
					</label>
				</div>

				<form id="form-search" class="form" role="form">
					<div class="form-group find">
						<label class="control-label">Pesquisa de Muncípios:</label>
						<input class="form-control" id="txt-municipio" placeholder="Digite o nome do muncípio"></input>
					</div>

					<div class="form-group">
						<label class="control-label sr-only">buscar</label>
						<button type="submit" class="btn btn-block btn-primary"><i class="fa fa-search"></i> Buscar</button>
					</div>
				</form>

				<div class="form-group result">
					<label class="control-label">Resultado da Pesquisa:</label>
					<div class="alert alert-danger hide"><i class="fa fa-warning"></i> Nenhum município encontrado</div>
					<table class="table table-municipios">
						<tbody></tbody>
					</table>
					<button type="button" id="btn-clear" class="btn btn-danger btn-sm btn-block hide"><i class="fa fa-trash-o"></i> Limpar pesquisa</button>
				</div>
			</div>

			<div class="col-xs-9 mapa">
				<div id="map-canvas"></div>
			</div>
		</div>
	</div>

	<div class="modal fade" id="modalLoading" tabindex="-1" role="dialog" aria-labelledby="modalLoadingLabel" aria-hidden="true">
		<div class="modal-dialog modal-sm">
			<div class="modal-content">
				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
					<h4 class="modal-title" id="modalLoadingLabel">Aguarde!</h4>
				</div>
				<div class="modal-body">
					<i class="fa fa-spinner fa-spin"></i> Buscando informações...
				</div>
			</div>
		</div>
	</div>

	<div class="modal fade" id="modalInfoMunicipio" tabindex="-1" role="dialog" aria-labelledby="modalInfoMunicipioLabel" aria-hidden="true">
		<div class="modal-dialog">
			<div class="modal-content">
				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
					<h4 class="modal-title" id="modalInfoMunicipioLabel">Informações do Município</h4>
				</div>
				<div class="modal-body">
					<div class="row">
						<div class="col-xs-12">
							<div class="form-group">
								<label class="control-label">Município</label>
								<input class="form-control input-sm" readonly="readonly" id="txt-dsc-municipio">
							</div>
						</div>
					</div>

					<div class="row">
						<div class="col-xs-6">
							<label class="control-label">Dir. de Bacia - DAEE</label>
							<input class="form-control input-sm" readonly="readonly" id="txt-nme-bacia-daee">
						</div>

						<div class="col-xs-6">
							<label class="control-label">UGRHI</label>
							<input class="form-control input-sm" readonly="readonly" id="txt-nme-bacia-secretaria">
						</div>
					</div>

					<div class="row">
						<div class="col-xs-6">
							<label class="control-label">População Urbana 2010 (IBGE)</label>
							<input class="form-control input-sm num" readonly="readonly" id="txt-pop-2010">
						</div>

						<div class="col-xs-6">
							<label class="control-label">Projeção de População (2030)</label>
							<input class="form-control input-sm num" readonly="readonly" id="txt-pop-2030">
						</div>
					</div>

					<div class="row">
						<div class="col-xs-9">
							<label class="control-label">Nome do Prefeito</label>
							<input class="form-control input-sm" readonly="readonly" id="txt-nme-prefeito">
						</div>

						<div class="col-xs-3">
							<label class="control-label">Partido</label>
							<input class="form-control input-sm" readonly="readonly" id="txt-nme-partido">
						</div>
					</div>

					<div class="row">
						<div class="col-xs-5">
							<label class="control-label">Atendido por</label>
							<input class="form-control input-sm" readonly="readonly" id="txt-atendido-sabesp">
						</div>
					</div>

					<div class="row">
						<div class="col-xs-12">
							<div class="table-responsive">
								<table class="table table-localidades table-bordered table-condensed table-striped table-hover">
									<thead>
										<tr class="info">
											<th>Localidade</th>
											<th>Situação</th>
											<th width="50"></th>
										</tr>
									</thead>
									<tbody></tbody>
								</table>
							</div>
						</div>
					</div>
				</div>
				<div class="modal-footer">
					<a id="btn-ficha-completa" class="btn btn-info">Detalhes do Município</a>
					<button type="button" class="btn btn-default" data-dismiss="modal">Fechar</button>
				</div>
			</div>
		</div>
	</div>
</body>
</html>